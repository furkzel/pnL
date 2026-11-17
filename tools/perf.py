"""
perf.py - timing of Yenile / On Kontrol / Aktar on large files.

  python tools/perf.py [--rows 200000] [--master 300000] [--cols 30] [--panels 3]
  python tools/perf.py --master 870000 --analiz       # Analiz: 870k rows, <= 1000 groups

v1.1 prints, per step, the wall time and the SURE (ms) the macro wrote to
Rapor, so K9-K14 / reconciliation cost can be compared with v1.0 runs of the
same command (run both builds back to back on the same machine).

Builds the data in memory-friendly chunks inside a private Excel instance.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlauto  # noqa: E402


def fill(ws, rows: int, cols: int, start_row: int, seed: int):
    chunk = 50000
    r = 0
    while r < rows:
        n = min(chunk, rows - r)
        data = []
        for i in range(n):
            k = r + i + seed
            row = [f"{k:08d}", 1000.5 + (k % 997), f"SIRKET{k % 7}", 45000 + (k % 400)]
            row += [f"v{j}_{k % 50}" for j in range(cols - 4)]
            data.append(row)
        ws.Range(ws.Cells(start_row + r, 1), ws.Cells(start_row + r + n - 1, cols)).Value = data
        r += n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=200000)
    ap.add_argument("--master", type=int, default=300000)
    ap.add_argument("--cols", type=int, default=30)
    ap.add_argument("--panels", type=int, default=3)
    ap.add_argument("--panel", default=str(xlauto.ROOT / "dist" / "AktarimPaneli_v1.1.xlsm"))
    ap.add_argument("--analiz", action="store_true", help="also time Analiz (group by SIRKET, sum TUTAR) on the master")
    a = ap.parse_args()

    work = Path(tempfile.mkdtemp(prefix="apb_perf_"))
    panel = work / "AktarimPaneli_perf.xlsm"
    shutil.copy2(a.panel, panel)
    hdr = ["POL NO", "TUTAR", "SIRKET", "TARIH"] + [f"ALAN {j}" for j in range(a.cols - 4)]

    with xlauto.ExcelSession() as s:
        xl = s.xl
        t0 = time.time()
        src = xl.Workbooks.Add()
        ws = src.Worksheets(1)
        ws.Name = "Veri"
        ws.Range(ws.Cells(1, 1), ws.Cells(1, a.cols)).Value = [hdr]
        ws.Range(ws.Cells(2, 1), ws.Cells(a.rows + 1, 1)).NumberFormat = "@"
        fill(ws, a.rows, a.cols, 2, 0)
        ws.Range(ws.Cells(2, 4), ws.Cells(a.rows + 1, 4)).NumberFormat = "dd.mm.yyyy"
        src.SaveAs(str(work / "Kaynak.xlsx"), FileFormat=51)
        mst = xl.Workbooks.Add()
        wm = mst.Worksheets(1)
        wm.Name = "Ana"
        wm.Range(wm.Cells(1, 1), wm.Cells(1, a.cols)).Value = [hdr]
        wm.Range(wm.Cells(1, 1), wm.Cells(1, 3)).Interior.Color = 0x50D092
        fill(wm, a.master, a.cols, 2, 10_000_000)
        mst.SaveAs(str(work / "Ana.xlsx"), FileFormat=51)
        print(f"data built in {time.time() - t0:.1f}s  (source {a.rows} x {a.cols}, master {a.master} x {a.cols})")

        wbP = xl.Workbooks.Open(str(panel))
        pnl = wbP.Worksheets("Panel")
        for k in range(2, a.panels + 1):
            s.run(wbP, "YeniPanelEkle")
        for k in range(1, a.panels + 1):
            c = 1 + (k - 1) * 6 + 2
            pnl.Cells(10, c).Value = src.Name
            pnl.Cells(11, c).Value = "Veri"
            pnl.Cells(14, c).Value = mst.Name
            pnl.Cells(15, c).Value = "Ana"

        t0 = time.time()
        s.run(wbP, "Yenile")
        t_yenile = time.time() - t0
        print(f"Yenile ({a.panels} panels): {t_yenile:.2f}s")

        s.run(wbP, "PaneleGit", 1)
        t0 = time.time()
        s.run(wbP, "OtomatikEslestir")
        print(f"OtomatikEslestir: {time.time() - t0:.2f}s")
        pnl.Cells(19, 3).Value = "B - TUTAR"
        pnl.Cells(8, 3).Value = "Evet"

        t0 = time.time()
        s.run(wbP, "OnKontrol")
        print(f"OnKontrol: {time.time() - t0:.2f}s")

        t0 = time.time()
        s.run(wbP, "Aktar")
        t_aktar = time.time() - t0
        last = wm.Cells(wm.Rows.Count, 1).End(-4162).Row
        print(f"Aktar ({a.rows} rows x {a.cols} cols, incl. dialogs): {t_aktar:.2f}s   master last row {last}")
        print(s.watcher.dump(len(s.watcher.events) - 2)[-900:])

        t0 = time.time()
        s.run(wbP, "GeriAl")
        print(f"GeriAl: {time.time() - t0:.2f}s")

        if a.analiz:
            an = wbP.Worksheets("Analiz")
            s.run(wbP, "AnalizAc")
            an.Range("D5").Value = None
            an.Range("D6").Value = next(an.Cells(r, 31).Value for r in range(1, 40)
                                        if "Ana" in str(an.Cells(r, 31).Value or ""))
            an.Range("D7").Value = "C - SIRKET"
            an.Range("D9").Value = "Toplam"
            an.Range("D10").Value = "B - TUTAR"
            t0 = time.time()
            s.run(wbP, "AnalizCalistir")
            print(f"Analiz ({a.master} rows, group SIRKET, sum TUTAR): {time.time() - t0:.2f}s")
            an.Range("D7").Value = "A - POL NO"       # high cardinality -> sort-based grouping
            an.Range("D9").Value = "Adet"
            t0 = time.time()
            s.run(wbP, "AnalizCalistir")
            print(f"Analiz high cardinality ({a.master} distinct keys): {time.time() - t0:.2f}s")
        wbP.Close(False)
        src.Close(False)
        mst.Close(False)


if __name__ == "__main__":
    main()
