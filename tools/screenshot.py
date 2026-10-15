"""
screenshot.py - render the Panel (populated with demo data) and the Kilavuz to
PNG files in docs/ for UI review and the README.

Rendering happens INSIDE a private, invisible Excel instance: the ranges are
exported to PDF (Range.ExportAsFixedFormat) and rasterised with pdftoppm
(Poppler / TeX Live / MiKTeX). Nothing is captured from the screen and the
clipboard is not touched.

  python tools/screenshot.py [--panel dist/AktarimPaneli_v34.xlsm]
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlauto  # noqa: E402
import test_e2e as t  # noqa: E402

XL_LANDSCAPE = 2


def render(ws, address: str, out_png: Path, work: Path, dpi: int = 110):
    ps = ws.PageSetup
    ps.Orientation = XL_LANDSCAPE
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 10
    ps.PrintArea = address
    pdf = work / (out_png.stem + ".pdf")
    ws.ExportAsFixedFormat(Type=0, Filename=str(pdf), IgnorePrintAreas=False, OpenAfterPublish=False)
    ps.PrintArea = ""
    exe = shutil.which("pdftoppm")
    if not exe:
        print("pdftoppm not found; PDF left at", pdf)
        return
    subprocess.run([exe, "-png", "-r", str(dpi), "-singlefile", str(pdf), str(out_png.with_suffix(""))], check=True)
    print("saved", out_png)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--panel", default=str(xlauto.ROOT / "dist" / "AktarimPaneli_v34.xlsm"))
    args = ap.parse_args()
    docs = xlauto.ROOT / "docs"
    docs.mkdir(exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="apb_shot_"))
    panel = work / "AktarimPaneli_demo.xlsm"
    shutil.copy2(args.panel, panel)

    with xlauto.ExcelSession(visible=False) as s:
        xl = s.xl
        wbP = xl.Workbooks.Open(str(panel))
        src = t.make_source(xl, work)
        mst = t.make_master(xl, work)
        pnl = wbP.Worksheets("Panel")
        wbP.Activate()
        pnl.Activate()
        pnl.Cells(10, 3).Value = src.Name
        pnl.Cells(11, 3).Value = "Sayfa1"
        pnl.Cells(14, 3).Value = mst.Name
        pnl.Cells(15, 3).Value = "Ana"
        s.run(wbP, "Yenile")
        pnl.Cells(19, 3).Value = "D - TUTAR"
        pnl.Cells(20, 3).Value = "B - HASAR TARIHI"
        pnl.Cells(21, 3).Value = "C - IHBAR TARIHI"
        pnl.Cells(8, 3).Value = "Evet"
        pnl.Cells(9, 4).Value = "I - RAPOR_DONEMI"
        pnl.Cells(9, 5).Value = "EYLUL 2026"
        s.run(wbP, "OtomatikEslestir")
        s.run(wbP, "YeniPanelEkle")
        wbP.Activate()
        pnl.Activate()
        s.run(wbP, "PaneleGit", 1)
        render(pnl, "$A$1:$R$50", docs / "panel_v34.png", work)

        klv = wbP.Worksheets(1)
        render(klv, "$A$1:$D$48", docs / "kilavuz_v34.png", work)
        wbP.Close(False)


if __name__ == "__main__":
    main()
