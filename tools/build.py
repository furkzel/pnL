"""
build.py - inject src/ into an Excel workbook, compile it, optionally run the
bootstrapper, and save an .xlsm.

Requirements: Windows, Excel, pywin32, and in Excel
  File > Options > Trust Center > Macro Settings >
  "Trust access to the VBA project object model" (AccessVBOM = 1).

Examples
--------
  # Fresh workbook, built from scratch and set up (Panel, Rapor, Kilavuz ...)
  python tools/build.py --new dist/AktarimPaneli_v34.xlsm

  # Upgrade an existing v33 workbook. The input is NEVER modified; a copy is
  # written to --out and SifirdanKur migrates the panel inside the copy.
  python tools/build.py --upgrade "_/Kitap1.xlsm" --out dist/Kitap1_v34.xlsm

  # Only check that every source file is pure ASCII
  python tools/build.py --check

Exit code 0 = success, 1 = failure (compile error, missing setup, ...).
"""
from __future__ import annotations

import argparse
import shutil
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlauto  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--new", metavar="OUT.xlsm", help="create a new workbook")
    g.add_argument("--upgrade", metavar="IN.xlsm", help="upgrade a copy of an existing workbook")
    g.add_argument("--check", action="store_true", help="only run the ASCII check")
    ap.add_argument("--out", metavar="OUT.xlsm", help="output path for --upgrade")
    ap.add_argument("--no-setup", action="store_true", help="do not run SifirdanKur after injecting")
    ap.add_argument("--visible", action="store_true", help="show the Excel window while building")
    args = ap.parse_args(argv)

    problems = xlauto.check_ascii()
    if problems:
        print("ASCII CHECK FAILED:")
        print("\n".join(problems))
        return 1
    print("ascii check: ok")
    if args.check:
        return 0

    if args.upgrade:
        src_wb = Path(args.upgrade).resolve()
        if not args.out:
            ap.error("--upgrade needs --out")
        out = Path(args.out).resolve()
        if out == src_wb:
            ap.error("--out must differ from the input (the input is never modified)")
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_wb, out)
    else:
        out = Path(args.new).resolve()
        out.parent.mkdir(parents=True, exist_ok=True)
        if out.exists():
            out.unlink()

    with xlauto.ExcelSession(visible=args.visible) as s:
        if args.upgrade:
            wb = s.xl.Workbooks.Open(str(out), UpdateLinks=0)
        else:
            wb = s.xl.Workbooks.Add()
            # keep a single sheet; SifirdanKur names/creates the rest
            while wb.Worksheets.Count > 1:
                wb.Worksheets(wb.Worksheets.Count).Delete()

        for line in xlauto.inject_sources(wb):
            print("  " + line)

        ok, msgs = xlauto.compile_project(s, wb)
        print("compile:", "ok" if ok else "FAILED")
        if not ok:
            print(msgs)
            others = [c.Name for c in wb.VBProject.VBComponents
                      if c.Type in (xlauto.VBEXT_STD, xlauto.VBEXT_CLASS)
                      and not xlauto._is_project_component(c.Name)]
            if others:
                print("Other (non-project) modules in this workbook that may reference removed v33 code:",
                      ", ".join(others))
            return 1

        if not args.no_setup:
            t0 = time.time()
            s.run(wb, "SifirdanKur", timeout=900)
            print(f"SifirdanKur: done in {time.time() - t0:.1f}s")
            print(s.watcher.dump())
            names = [ws.Name for ws in wb.Worksheets]
            print("sheets:", names)
            if "Panel" not in names:
                print("SETUP FAILED: no Panel sheet")
                return 1

        if args.upgrade:
            wb.Save()
        else:
            wb.SaveAs(str(out), FileFormat=xlauto.XL_OPENXML_MACRO)
        print("saved:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
