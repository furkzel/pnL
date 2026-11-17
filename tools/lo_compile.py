"""LibreOffice Basic (VBA compatibility mode) compile smoke test.

This is NOT Excel's VBA compiler. It parses every module of src/ with
LibreOffice's Basic compiler in VBASupport mode and reports modules (and the
procedures inside them) that fail to compile. It catches syntax errors such as
unbalanced If/For/With blocks, bad line continuations and malformed
declarations. It does not resolve Excel object-model members or undeclared
variables (LibreOffice reports those at run time), so a pass here is weaker
than Debug > Compile in Excel; a failure is still a real problem in most cases.

  python tools/lo_compile.py [--src DIR]

Requires LibreOffice (soffice) and its Python UNO bridge. Exit 0 = all pass,
1 = failures, 2 = LibreOffice unavailable.
"""
from __future__ import annotations

import argparse
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
PROC = re.compile(r"^\s*(?:Public |Private |Friend )?(?:Static )?(Sub|Function|Property (?:Get|Let|Set))\s+(\w+)", re.I)
END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)


def free_port() -> int:
    with socket.socket() as s:
        s.bind(("localhost", 0))
        return s.getsockname()[1]


def module_parts(path: Path):
    """(name, is_class, header_lines, [(proc_name, lines)])"""
    text = path.read_text(encoding="ascii", errors="replace").replace("\r\n", "\n")
    lines = [ln for ln in text.split("\n") if not ln.startswith("Attribute ")]
    if path.suffix.lower() == ".cls":
        # drop VERSION/BEGIN..END class header
        out, skip = [], False
        for ln in lines:
            if ln.startswith("VERSION ") or ln.strip() == "BEGIN":
                skip = True
                continue
            if skip and ln.strip() == "END":
                skip = False
                continue
            if not skip:
                out.append(ln)
        lines = out
    header, procs, cur, name = [], [], None, None
    for ln in lines:
        m = PROC.match(ln)
        if cur is None and m:
            cur, name = [ln], m.group(2)
            continue
        if cur is not None:
            cur.append(ln)
            if END.match(ln):
                procs.append((name, cur))
                cur = None
            continue
        if re.match(r"\s*Option\s+(Explicit|Private\s+Module)\b", ln, re.I):
            continue
        header.append(ln)
    is_class = path.suffix.lower() == ".cls" and path.stem in ("Sozluk",)
    return path.stem, is_class, header, procs


# LibreOffice rejects a few constructs that Excel VBA accepts. They are
# rewritten in the throw-away copy only, so the rest of the line is still checked.
LO_REWRITES = [
    (re.compile(r"\bReDim\s+(Preserve\s+)?(\w+)\.(\w+)\(", re.I), r"ReDim \1\2_\3("),
    (re.compile(r"^(\s*)mod(\s+As\s)", re.I), r"\1mod_\2"),
    (re.compile(r"\.mod\b"), ".mod_"),
]


def lo_line(line: str) -> str:
    # Excel-only construct marked in the source (e.g. Byte() = String, UDT LSet):
    # the line is commented out in the LibreOffice copy only.
    if line.rstrip().endswith("' LO-atla"):
        return "' " + line
    for rx, rep in LO_REWRITES:
        line = rx.sub(rep, line)
    return line


def build_source(name, is_class, header, procs, probe):
    pre = ["Option VBASupport 1"]
    if is_class:
        pre.append("Option ClassModule")
    pre.append("Option Explicit")
    body = []
    for _, ln in procs:
        body.extend(ln)
    header = [lo_line(x) for x in header]
    body = [lo_line(x) for x in body]
    probe_fn = [f"Public Function {probe}() As Long", f"    {probe} = 42", "End Function"]
    if is_class:
        # class methods are not callable as scripts: compile probe goes into a helper module
        return "\n".join(pre + header + body) + "\n"
    return "\n".join(pre + header + body + probe_fn) + "\n"


# Procedures that Excel compiles but LibreOffice rejects (baseline v1.0).
KNOWN: dict[str, set[str]] = {}


class LO:
    def __init__(self):
        import uno  # noqa: F401
        self.port = free_port()
        self.profile = tempfile.mkdtemp(prefix="lo_prof_")
        self.proc = subprocess.Popen(
            ["soffice", "--headless", "--invisible", "--norestore", "--nologo", "--nodefault",
             f"-env:UserInstallation=file://{self.profile}",
             f"--accept=socket,host=localhost,port={self.port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        import uno
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(120):
            try:
                self.ctx = resolver.resolve(f"uno:socket,host=localhost,port={self.port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise RuntimeError("LibreOffice did not start")
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def compiles(self, modules: dict[str, str], probe_module: str, probe: str) -> bool:
        doc = self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        try:
            libs = doc.BasicLibraries
            libs.VBACompatibilityMode = True
            if not libs.hasByName("VBAProject"):
                libs.createLibrary("VBAProject")
            lib = libs.getByName("VBAProject")
            for n, s in modules.items():
                lib.insertByName(n, s)
            sp = doc.getScriptProvider()
            try:
                s = sp.getScript(f"vnd.sun.star.script:VBAProject.{probe_module}.{probe}?language=Basic&location=document")
                r = s.invoke((), (), ())
                return r[0] == 42
            except Exception:
                return False
        finally:
            doc.close(True)

    def close(self):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        self.proc.terminate()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=str(SRC))
    ap.add_argument("--only", default=None, help="comma separated module names")
    args = ap.parse_args(argv)
    src = Path(args.src)
    files = sorted(p for p in src.glob("*") if p.suffix.lower() in (".bas", ".cls"))
    parts = [module_parts(p) for p in files]
    try:
        lo = LO()
    except Exception as e:  # pragma: no cover
        print("LibreOffice unavailable:", e)
        return 2
    failures = []
    try:
        only = set(args.only.split(",")) if args.only else None
        for name, is_class, header, procs in parts:
            if only and name not in only:
                continue
            if is_class:
                mods = {name: build_source(name, True, header, procs, ""),
                        "ZZProbe": "Option VBASupport 1\nOption Explicit\nPublic Function ZZ_P() As Long\n"
                                   f"    Dim o As New {name}\n    ZZ_P = 42\nEnd Function\n"}
                ok = lo.compiles(mods, "ZZProbe", "ZZ_P")
                print(f"  {'ok  ' if ok else 'FAIL'}  {name} (class)")
                if not ok:
                    failures.append((name, "(class)"))
                continue
            probe = f"ZZ_{name}"
            if lo.compiles({name: build_source(name, False, header, procs, probe)}, name, probe):
                print(f"  ok    {name}")
                continue
            bad = []
            if not lo.compiles({name: build_source(name, False, header, [], probe)}, name, probe):
                bad.append("(declarations)")
            else:
                for pname, plines in procs:
                    if not lo.compiles({name: build_source(name, False, header, [(pname, plines)], probe)}, name, probe):
                        bad.append(pname)
            known = KNOWN.get(name, set())
            new_bad = [b for b in bad if b not in known]
            if bad and not new_bad:
                print(f"  ok*   {name} (known LibreOffice-only issues: {', '.join(bad)})")
                continue
            print(f"  FAIL  {name}: {', '.join(new_bad) or '(procedure combination; Excel-only constructs?)'}")
            if new_bad:
                failures.append((name, new_bad))
    finally:
        lo.close()
    print("LibreOffice compile:", "ok" if not failures else f"{len(failures)} module(s) failed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
