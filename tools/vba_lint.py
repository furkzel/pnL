"""Static VBA checks that need no Excel (portable, stdlib only).

Excel's Debug > Compile rejects, among others:
  - an identifier that is not declared anywhere (Option Explicit),
  - a project procedure called with too many / too few arguments,
  - a Sub called with parentheses around several arguments ("Foo(a, b)" as a statement),
  - a Private procedure of another module,
  - a name declared Public in two standard modules (ambiguous name),
  - a duplicate declaration inside one procedure.

This module approximates those rules on the source text. Identifiers that are
not declared in the project (VBA runtime, Excel/Office object model and
constants) must appear in tools/vba_builtins.txt. That list was derived from
the v1.0 source, which compiled in Excel; adding a name to it is a deliberate,
reviewed act.

  python tools/vba_lint.py [--src DIR] [--derive-builtins]
"""
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
BUILTINS_FILE = Path(__file__).resolve().parent / "vba_builtins.txt"

KEYWORDS = {w.lower() for w in """
If Then Else ElseIf End Sub Function Dim As ByVal ByRef Optional ParamArray For To Step Next Each In
Do Loop While Until Wend Select Case Is With Set Let Get Property Exit Call Private Public Friend Static
Const New Nothing Not And Or Xor Mod Like True False Empty Null On Error Resume GoTo GoSub Return ReDim
Preserve Option Explicit Compare Base Module Type Enum Me Long LongLong LongPtr Integer String Double Boolean
Variant Object Date Currency Byte Single Any LSet RSet Erase Stop TypeOf Global Declare Lib Alias
Imp Eqv AddressOf WithEvents Implements Event RaiseEvent Open Close Output Input Append Binary Random Access
Read Write Lock Shared Line Print Width Seek Name Kill MkDir RmDir ChDir ChDrive Decimal Text
""".split()}

# Words Excel's compiler rejects as variable / parameter names (case-insensitive).
RESERVED = {w.lower() for w in """
And As Boolean ByRef Byte ByVal Call Case Const Currency Date Declare Dim Do Double Each Else ElseIf End Enum
Erase Error Event Exit False For Friend Function Get Global GoSub GoTo If Implements In Integer Is Let Like Long
LongLong Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Public
RaiseEvent ReDim Resume Return RSet Select Set Single Static Stop String Sub Then To True Type TypeOf Until
Variant Wend While With WithEvents Xor
""".split()}

PROC_RE = re.compile(r"^\s*(?:(Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*(\((.*)\))?", re.I)
END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*[$%&#!@]?")


def strip_code(line: str) -> str:
    """Blank string literals, drop comments ('... and Rem)."""
    out, i, n = [], 0, len(line)
    while i < n:
        ch = line[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if line[j] == '"':
                    if j + 1 < n and line[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append('""')
            i = j + 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    s = "".join(out)
    if re.match(r"\s*Rem\b", s, re.I):
        return ""
    return s


def logical_lines(text: str):
    """Yield (first_line_no, code) with continuations joined."""
    buf, start = [], None
    for n, raw in enumerate(text.replace("\r\n", "\n").split("\n"), 1):
        if raw.startswith("Attribute "):
            continue
        code = strip_code(raw)
        if start is None:
            start = n
        if code.rstrip().endswith(" _") or code.rstrip() == "_":
            buf.append(code.rstrip()[:-1])
            continue
        buf.append(code)
        yield start, " ".join(buf)
        buf, start = [], None
    if buf:
        yield start, " ".join(buf)


def split_top(s: str, sep: str = ","):
    parts, depth, cur = [], 0, []
    for ch in s:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == sep and depth == 0:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return parts


@dataclass
class Param:
    name: str
    optional: bool
    paramarray: bool


@dataclass
class Proc:
    module: str
    name: str
    kind: str
    private: bool
    params: list
    line: int
    body: list = field(default_factory=list)   # (line_no, code)


@dataclass
class Module:
    name: str
    kind: str            # std | class | doc
    procs: list = field(default_factory=list)
    decls: dict = field(default_factory=dict)  # lower name -> (public?, line)
    decl_lines: list = field(default_factory=list)


def parse_params(s: str):
    params = []
    if not s or not s.strip():
        return params
    for p in split_top(s):
        p = p.strip()
        if not p:
            continue
        opt = bool(re.match(r"Optional\b", p, re.I))
        pa = bool(re.search(r"\bParamArray\b", p, re.I))
        p2 = re.sub(r"^(Optional|ByVal|ByRef|ParamArray)\s+", "", p, flags=re.I)
        p2 = re.sub(r"^(Optional|ByVal|ByRef|ParamArray)\s+", "", p2, flags=re.I)
        m = IDENT_RE.match(p2)
        if m:
            params.append(Param(m.group(0).rstrip("$%&#!@"), opt, pa))
    return params


def declared_names(code: str):
    """Names declared by a Dim/Const/Public/Private/Static/ReDim statement (after the keyword)."""
    m = re.match(r"\s*(?:(?:Public|Private|Global|Dim|Static)\s+(?:Const\s+|WithEvents\s+)?|Const\s+)(.*)$", code, re.I)
    if not m:
        return []
    rest = m.group(1)
    if re.match(r"(Sub|Function|Property|Enum|Type|Declare)\b", rest, re.I):
        return []
    names = []
    for part in split_top(rest):
        part = part.strip()
        mm = IDENT_RE.match(part)
        if mm:
            names.append(mm.group(0).rstrip("$%&#!@"))
    return names


def parse_module(path: Path) -> Module:
    text = path.read_text(encoding="ascii", errors="replace")
    stem = path.stem
    kind = "std"
    if path.suffix.lower() == ".cls":
        kind = "class" if "VERSION 1.0 CLASS" in text else "doc"
    if kind == "class":
        text = re.sub(r"(?s)^VERSION.*?\nEND\n", "", text)
    mod = Module(stem, kind)
    cur = None
    in_enum = in_type = False
    for n, code in logical_lines(text):
        if not code.strip():
            continue
        if cur is None:
            if in_enum:
                if re.match(r"\s*End\s+Enum\b", code, re.I):
                    in_enum = False
                else:
                    m = IDENT_RE.match(code.strip())
                    if m:
                        mod.decls[m.group(0).lower()] = (enum_public, n)
                continue
            if in_type:
                if re.match(r"\s*End\s+Type\b", code, re.I):
                    in_type = False
                continue
            m = re.match(r"\s*(?:(Public|Private)\s+)?Enum\s+(\w+)", code, re.I)
            if m:
                enum_public = (m.group(1) or "Public").lower() != "private"
                mod.decls[m.group(2).lower()] = (enum_public, n)
                in_enum = True
                continue
            m = re.match(r"\s*(?:(Public|Private)\s+)?Type\s+(\w+)", code, re.I)
            if m:
                mod.decls[m.group(2).lower()] = ((m.group(1) or "Public").lower() != "private", n)
                in_type = True
                continue
            m = re.match(r"\s*(?:(Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)", code, re.I)
            if m:
                mod.decls[m.group(2).lower()] = ((m.group(1) or "Public").lower() != "private", n)
                continue
            m = PROC_RE.match(code)
            if m:
                vis = (m.group(1) or "Public").lower()
                kindp = re.sub(r"\s+", " ", m.group(2)).title()
                cur = Proc(stem, m.group(3), kindp, vis == "private", parse_params(m.group(5) or ""), n)
                continue
            if re.match(r"\s*Option\b", code, re.I):
                continue
            names = declared_names(code)
            pub = bool(re.match(r"\s*(Public|Global)\b", code, re.I))
            for nm in names:
                if nm.lower() in mod.decls:
                    mod.decl_lines.append((n, f"duplicate module-level declaration '{nm}'"))
                mod.decls[nm.lower()] = (pub, n)
            continue
        if END_PROC_RE.match(code):
            mod.procs.append(cur)
            cur = None
            continue
        cur.body.append((n, code))
    return mod


class Project:
    def __init__(self, src: Path):
        files = sorted(p for p in src.glob("*") if p.suffix.lower() in (".bas", ".cls"))
        self.modules = [parse_module(p) for p in files]
        self.procs: dict[str, list[Proc]] = {}
        for m in self.modules:
            for p in m.procs:
                self.procs.setdefault(p.name.lower(), []).append(p)
        self.classes = {m.name.lower() for m in self.modules if m.kind == "class"}
        self.module_names = {m.name.lower() for m in self.modules}

    def public_names(self):
        """lower name -> list of (module, kind) visible project-wide."""
        out: dict[str, list] = {}
        for m in self.modules:
            if m.kind != "std":
                continue
            for nm, (pub, _) in m.decls.items():
                if pub:
                    out.setdefault(nm, []).append((m.name, "decl"))
            for p in m.procs:
                if not p.private:
                    out.setdefault(p.name.lower(), []).append((m.name, "proc"))
        return out


def identifiers(code: str):
    """(name, start, end) of identifiers that are not member accesses, labels or named arguments."""
    code = re.sub(r"&[Hh][0-9A-Fa-f]+&?|&[Oo][0-7]+&?", " 0 ", code)
    code = re.sub(r"#[^#\n]{1,30}#", " 0 ", code) if re.search(r"#\d", code) else code
    res = []
    for m in IDENT_RE.finditer(code):
        s, e = m.start(), m.end()
        before = code[:s].rstrip()
        if before.endswith(".") or before.endswith("!"):
            continue
        if re.match(r"\s*:=", code[e:]):
            continue
        if s > 0 and (code[s - 1].isdigit() or code[s - 1] in "_"):
            continue
        res.append((m.group(0), s, e))
    return res


def lint(src: Path = SRC, builtins: set[str] | None = None, collect_unknown: bool = False):
    proj = Project(src)
    if builtins is None:
        builtins = load_builtins()
    problems: list[str] = []
    unknown: dict[str, str] = {}
    pubs = proj.public_names()

    for nm, owners in pubs.items():
        mods = {o[0] for o in owners}
        if len(mods) > 1:
            problems.append(f"ambiguous public name '{nm}' in {', '.join(sorted(mods))}")

    for m in proj.modules:
        for n, msg in m.decl_lines:
            problems.append(f"{m.name}.bas:{n}: {msg}")
        seen = {}
        for p in m.procs:
            key = (p.name.lower(), p.kind if p.kind.startswith("Property") else "")
            if key in seen:
                problems.append(f"{m.name}:{p.line}: duplicate procedure '{p.name}'")
            seen[key] = True

        for p in m.procs:
            local: dict[str, int] = {}
            for prm in p.params:
                if prm.name.lower() in RESERVED:
                    problems.append(f"{m.name}:{p.line}: {p.name}: reserved word '{prm.name}' used as a parameter name")
                if prm.name.lower() in local:
                    problems.append(f"{m.name}:{p.line}: {p.name}: duplicate parameter '{prm.name}'")
                local[prm.name.lower()] = p.line
            labels = set()
            for n, code in p.body:
                lm = re.match(r"\s*([A-Za-z_]\w*):(?!=)(\s|$)", code)
                if lm and lm.group(1).lower() not in KEYWORDS:
                    labels.add(lm.group(1).lower())
                for stmt in split_statements(code):
                    names = declared_names(stmt)
                    if re.match(r"\s*ReDim\b", stmt, re.I):
                        names = []
                    for nm in names:
                        if nm.lower() in RESERVED:
                            problems.append(f"{m.name}:{n}: {p.name}: reserved word '{nm}' used as a variable name")
                        if nm.lower() in local:
                            problems.append(f"{m.name}:{n}: {p.name}: duplicate declaration '{nm}'")
                        local[nm.lower()] = n
            if p.kind == "Function" or p.kind == "Property Get":
                local.setdefault(p.name.lower(), p.line)

            for n, code in p.body:
                for stmt in split_statements(code):
                    check_statement(proj, m, p, n, stmt, local, labels, builtins, problems, unknown)
    if collect_unknown:
        return problems, unknown
    return problems


def split_statements(code: str):
    """Split a logical line at ':' statement separators (not ':=')."""
    parts, depth, cur, i = [], 0, [], 0
    while i < len(code):
        ch = code[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == ":" and depth == 0 and not code[i + 1:i + 2] == "=":
            parts.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    parts.append("".join(cur))
    # single-line If: also check the statements after Then / Else
    out = []
    for s in parts:
        m = re.match(r"(\s*(?:ElseIf|If)\b.*?\bThen\b)(.*)$", s, re.I)
        if m and m.group(2).strip():
            out.append(m.group(1))
            rest = m.group(2)
            em = re.split(r"\bElse\b", rest, maxsplit=1, flags=re.I)
            out.extend(em)
        else:
            out.append(s)
    return [x for x in out if x.strip()]


def resolve(proj: Project, m: Module, name: str, local: dict):
    nl = name.rstrip("$%&#!@").lower()
    if nl in local:
        return "local"
    if nl in m.decls:
        return "module"
    if any(p.name.lower() == nl for p in m.procs):
        return "proc"
    if nl in proj.public_names():
        return "public"
    if nl in proj.classes or nl in proj.module_names:
        return "module"
    return None


def check_statement(proj, m, p, n, stmt, local, labels, builtins, problems, unknown):
    s = stmt.strip()
    if not s:
        return
    if re.match(r"(On\s+Error\s+GoTo|GoTo|GoSub|Resume)\s+\w+", s, re.I):
        return
    lm = re.match(r"([A-Za-z_]\w*):$", s)
    if lm:
        return
    toks = identifiers(stmt)
    prev_word = ""
    for name, st, en in toks:
        nl = name.lower()
        after_as = re.search(r"\b(As|New|TypeOf\s+\w+\s+Is)\s*$", stmt[:st], re.I)
        if after_as:
            continue
        if nl.rstrip("$%&#!@") in KEYWORDS or nl in KEYWORDS:
            continue
        if nl in labels:
            continue
        r = resolve(proj, m, name, local)
        if r is None:
            if nl.rstrip("$%&#!@") in builtins or nl in builtins:
                continue
            unknown.setdefault(nl, f"{m.name}:{n}")
            problems.append(f"{m.name}:{n}: {p.name}: undeclared identifier '{name}'")
            continue
    check_calls(proj, m, p, n, stmt, local, problems)


def target_proc(proj, m, name, local):
    nl = name.lower()
    if nl in local or nl in m.decls:
        return None
    cands = proj.procs.get(nl, [])
    own = [c for c in cands if c.module == m.name]
    if own:
        return own[0]
    std = [c for c in cands if any(x.name == c.module and x.kind == "std" for x in proj.modules)]
    if not std:
        return None
    return std[0]


def check_arity(m, p, n, tp: Proc, args: list[str], problems, how: str):
    if any(prm.paramarray for prm in tp.params):
        return
    named = {}
    positional = []
    for a in args:
        mm = re.match(r"\s*([A-Za-z_]\w*)\s*:=", a)
        if mm:
            named[mm.group(1).lower()] = True
        else:
            positional.append(a)
    if len(args) == 1 and not args[0].strip():
        positional = []
    total = len(tp.params)
    required = sum(1 for prm in tp.params if not prm.optional)
    if tp.kind == "Property Let" or tp.kind == "Property Set":
        return
    if len(positional) > total:
        problems.append(f"{m.name}:{n}: {p.name}: {tp.name} called with {len(positional)} args, takes {total} ({how})")
        return
    for i, prm in enumerate(tp.params):
        if not prm.optional and i >= len(positional) and prm.name.lower() not in named:
            problems.append(f"{m.name}:{n}: {p.name}: {tp.name} missing argument '{prm.name}' ({how})")
            return
    for nm in named:
        if nm not in {x.name.lower() for x in tp.params}:
            problems.append(f"{m.name}:{n}: {p.name}: {tp.name} has no parameter '{nm}'")


def paren_args(s: str, start: int):
    """s[start] == '(' -> (args list, end index) for the matching ')'."""
    depth, i = 0, start
    while i < len(s):
        if s[i] == "(":
            depth += 1
        elif s[i] == ")":
            depth -= 1
            if depth == 0:
                return split_top(s[start + 1:i]), i
        i += 1
    return split_top(s[start + 1:]), len(s)


def check_calls(proj, m, p, n, stmt, local, problems):
    s = stmt.strip()
    # statement call
    sm = re.match(r"(Call\s+)?([A-Za-z_]\w*)(.*)$", s, re.I)
    stmt_target = None
    if sm and sm.group(2).lower() not in KEYWORDS:
        rest = sm.group(3)
        is_assign = re.match(r"\s*(\(.*?\))?\s*=", rest) is not None and not sm.group(1)
        if not is_assign and not rest.startswith(".") and not rest.startswith("!"):
            tp = target_proc(proj, m, sm.group(2), local)
            if tp is not None:
                stmt_target = sm.start(2)
                if tp.private and tp.module != m.name:
                    problems.append(f"{m.name}:{n}: {p.name}: calls Private {tp.module}.{tp.name}")
                r = rest.strip()
                if sm.group(1):
                    args = paren_args(r, 0)[0] if r.startswith("(") else []
                    check_arity(m, p, n, tp, args, problems, "Call")
                elif r.startswith("("):
                    args, end = paren_args(r, 0)
                    tail = r[end + 1:].strip()
                    if tail.startswith(","):
                        check_arity(m, p, n, tp, [r[:end + 1]] + split_top(tail[1:]), problems, "statement")
                    elif len(args) > 1 and not tail:
                        problems.append(f"{m.name}:{n}: {p.name}: '{tp.name}(a, b)' used as a statement (needs Call or no parentheses)")
                    elif not tail:
                        check_arity(m, p, n, tp, [r], problems, "statement")
                else:
                    check_arity(m, p, n, tp, split_top(r) if r else [""], problems, "statement")
    # function calls inside expressions
    for fm in re.finditer(r"(?<![.\w!])([A-Za-z_]\w*)\s*\(", stmt):
        if stmt_target is not None and fm.start(1) == stmt_target:
            continue
        name = fm.group(1)
        if name.lower() in KEYWORDS:
            continue
        tp = target_proc(proj, m, name, local)
        if tp is None or tp.kind == "Sub":
            continue
        if tp.private and tp.module != m.name:
            problems.append(f"{m.name}:{n}: {p.name}: calls Private {tp.module}.{tp.name}")
        args, _ = paren_args(stmt, fm.end() - 1)
        check_arity(m, p, n, tp, args, problems, "expression")
    # function used without parentheses as a value: skip (zero-arg functions)


def load_builtins() -> set[str]:
    if not BUILTINS_FILE.exists():
        return set()
    out = set()
    for ln in BUILTINS_FILE.read_text(encoding="ascii").splitlines():
        ln = ln.split("#", 1)[0].strip()
        if ln:
            out.add(ln.lower())
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=str(SRC))
    ap.add_argument("--derive-builtins", action="store_true",
                    help="print identifiers unresolved in --src (use on a source that compiled in Excel)")
    args = ap.parse_args(argv)
    if args.derive_builtins:
        _, unknown = lint(Path(args.src), builtins=set(), collect_unknown=True)
        for k in sorted(unknown):
            print(f"{k:32s} # {unknown[k]}")
        return 0
    problems = lint(Path(args.src))
    for x in problems:
        print(x)
    print("vba lint:", "ok" if not problems else f"{len(problems)} problem(s)")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
