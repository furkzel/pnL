"""
export.py - copy the VBA code of a workbook back into src/ (round trip).

Use this after a quick fix was made directly in the Excel VBA editor, then
review the change with `git diff src/` before committing.

  python tools/export.py dist/AktarimPaneli_v34.xlsm

Only project components are exported (M0*_*.bas, Sozluk.cls, the Panel sheet
module -> src/Panel.cls, ThisWorkbook -> src/ThisWorkbook.cls). The ASCII check
runs afterwards; a non-ASCII character typed in the VBA editor is reported.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlauto  # noqa: E402


def main(argv=None) -> int:
    argv = argv if argv is not None else sys.argv[1:]
    if len(argv) != 1:
        print(__doc__)
        return 1
    path = Path(argv[0]).resolve()
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(path), ReadOnly=True)
        for line in xlauto.export_sources(wb):
            print("  " + line)
        wb.Close(False)
    problems = xlauto.check_ascii()
    if problems:
        print("WARNING - non-ASCII characters (VBA editor locale!):")
        print("\n".join(problems))
        return 1
    print("ascii check: ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
