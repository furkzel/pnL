"""Portable VBA source gates; no Excel or pywin32 imports."""
from pathlib import Path
import re

SRC = Path(__file__).resolve().parents[1] / 'src'


def code_only(line):
    """Remove VBA strings and comments, respecting doubled quotation marks."""
    return re.sub(r'"(?:[^"]|"")*"', '""', line).split("'", 1)[0]


def check_ascii(src=SRC):
    problems = []
    files = sorted(p for p in src.glob('*') if p.suffix.lower() in ('.bas', '.cls'))
    if not files:
        return [f'{src}: no VBA sources found']
    for path in files:
        for n, line in enumerate(path.read_bytes().split(b'\n'), 1):
            if any(b not in (9, 13) and not 32 <= b <= 126 for b in line):
                problems.append(f'{path.name}:{n}: non-printable ASCII source')
    return problems


def check_sources(src=SRC, *, palette=True):
    problems = check_ascii(src)
    for path in sorted(src.glob('*')):
        if path.suffix.lower() not in ('.bas', '.cls'):
            continue
        lines = path.read_text(encoding='ascii', errors='replace').splitlines()
        code = [code_only(line) for line in lines]
        if not any(re.fullmatch(r'\s*Option Explicit\s*', line, re.I) for line in code):
            problems.append(f'{path.name}: missing Option Explicit')
        continuation = 0
        procedure = None
        locked = unlocked = False
        for n, line in enumerate(code, 1):
            if palette and path.name != 'M01_Ayarlar.bas' and re.search(r'\bRGB\s*\(', line, re.I):
                problems.append(f'{path.name}:{n}: RGB literal outside palette')
            if re.search(r'\bApplication\.Transpose\b|\bCreateObject\s*\(', line, re.I):
                problems.append(f'{path.name}:{n}: forbidden runtime dependency/Transpose')
            if re.search(r'\.Cells\.Clear\s*$', line, re.I):
                problems.append(f'{path.name}:{n}: unbounded Cells.Clear')
            continuation = continuation + 1 if line.rstrip().endswith(' _') else 0
            if continuation > 24:
                problems.append(f'{path.name}:{n}: more than 24 continuations')
            start = re.match(r'\s*(?:Public |Private |Friend )?(?:Sub|Function|Property (?:Get|Let|Set)) (\w+)', line, re.I)
            if start:
                procedure = start[1]
                locked = unlocked = False
            if procedure:
                locked |= bool(re.search(r'\bEkranKilitle\b', line))
                unlocked |= bool(re.search(r'\bEkranAc\b', line))
                if re.match(r'\s*End (Sub|Function|Property)\s*$', line, re.I):
                    if locked and not unlocked and procedure != 'EkranKilitle':
                        problems.append(f'{path.name}:{n}: {procedure} locks without EkranAc')
                    procedure = None
    return problems


def check_contrast(src=SRC):
    """WCAG normal-text contrast for the semantic pairs used by status UI.
    This verifies declared pairs, not rendered Excel/OS theme overrides.
    """
    palette = (src / 'M01_Ayarlar.bas').read_text(encoding='ascii')
    colors = {name: tuple(map(int, (r, g, b))) for name, r, g, b in re.findall(
        r'Case (\w+):\s*Palet = RGB\((\d+),\s*(\d+),\s*(\d+)\)', palette)}
    pairs = [('prMetin', 'prZemin'), ('prIkincilMetin', 'prYuzey'),
             ('prOk', 'prOkZemin'), ('prDikkat', 'prDikkatZemin'),
             ('prEngel', 'prEngelZemin'), ('prOneriMetin', 'prOneriZemin'),
             ('prZemin', 'prButonAna'), ('prZemin', 'prButonYesil'),
             ('prZemin', 'prButonKirmizi'), ('prSeritYazi', 'prSeritZemin')]
    def luminance(rgb):
        linear = [v / 255 / 12.92 if v / 255 <= .04045 else ((v / 255 + .055) / 1.055) ** 2.4 for v in rgb]
        return sum(v * weight for v, weight in zip(linear, (.2126, .7152, .0722)))
    problems = []
    for fg, bg in pairs:
        if fg not in colors or bg not in colors:
            problems.append(f'missing contrast token: {fg}/{bg}')
            continue
        lo, hi = sorted((luminance(colors[fg]), luminance(colors[bg])))
        ratio = (hi + .05) / (lo + .05)
        if ratio < 4.5:
            problems.append(f'contrast {fg}/{bg}: {ratio:.2f}:1 < 4.5:1')
    return problems
