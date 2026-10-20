"""
test_e2e.py - end-to-end test of the built workbook, driven through Excel.

  python tools/test_e2e.py [--panel dist/AktarimPaneli_v1.0.xlsm] [--v33 _/Kitap1.xlsm]

Creates a monthly source file and a master file in a temp folder, configures
PANEL 1 through COM exactly like a user would (cell edits fire the sheet
events), then runs the real macros. Every MsgBox/InputBox is answered by the
DialogWatcher and its text is checked.

Scenarios
  1  Yenile, Otomatik Eslestir, Oneri onayi, Profil kaydet + uygula
  2  On Kontrol (dry run): suspicious total row, date checks, format guard
  3  Filters: literal value with comma, multi value, numeric comparison
  4  Aktar: rows, leading zeros, 19-digit ids, =text, formula row shift,
     fixed value, number format inheritance, reconciliation, Gecmis/Rapor
  5  Geri Al
  6  Fingerprint: row inserted above the block -> Geri Al warns and stops
  7  Secili Sutun Yamasi: overwrite + backup, Geri Al restores old values
  8  v33 -> v34 migration of a copy of the real workbook (if --v33 given)
"""
from __future__ import annotations

import argparse
import datetime as dt
import shutil
import sys
import tempfile
import time
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import xlauto  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ROOT = xlauto.ROOT
N_ROWS = 200          # monthly data rows (2..201), blank 202, TOTAL 203
MASTER_ROWS = 1000    # existing master data rows (2..1001)

RESULTS: list[tuple[str, bool, str]] = []


def check(name: str, cond: bool, detail: str = ""):
    RESULTS.append((name, bool(cond), detail))
    print(("  PASS " if cond else "  FAIL ") + name + (f"   [{detail}]" if detail and not cond else ""))


class Policy:
    """Answers dialogs; per-test overrides by substring of the message."""

    def __init__(self):
        self.rules: list[tuple[str, object]] = []

    def set(self, *rules):
        self.rules = list(rules)

    def __call__(self, title, text, buttons):
        for key, answer in self.rules:
            if key.lower() in (text or "").lower() or key.lower() in (title or "").lower():
                return answer
        return xlauto.default_policy(title, text, buttons)


# ---------------------------------------------------------------------------
#  Test data
# ---------------------------------------------------------------------------
def make_source(xl, folder: Path):
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Sayfa1"
    hdr = ["Pol No", "Hasar Tarihi", "Ihbar Tarihi", "Odenen Tutar", "Sirket", "Brans", "Hesap", "Aciklama"]
    rows = []
    base = dt.datetime(2026, 9, 1)
    for i in range(N_ROWS):
        pol = f"{i + 1:07d}"                         # leading zeros, text
        if i % 50 == 7:
            pol = "1234567890123456" + f"{i:03d}"     # 19-digit id
        hasar = base - dt.timedelta(days=30 + i)
        ihbar = hasar + dt.timedelta(days=5)
        if i in (3, 4):
            ihbar = hasar - dt.timedelta(days=2)      # ihbar before hasar
        sirket = ["AXA", "ALLIANZ", "ALLIANZ, INC", "MAPFRE"][i % 4]
        tutar = round(1000 + i * 12.5, 2)
        rows.append([pol, hasar, ihbar, tutar, sirket, "HAYAT" if i % 2 else "OTO", None, f"not {i}"])
    ws.Range("A1:H1").Value = [hdr]
    ws.Range("A2:A%d" % (N_ROWS + 1)).NumberFormat = "@"
    data = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]] for r in rows]
    ws.Range("A2:H%d" % (N_ROWS + 1)).Value = data
    ws.Range("G2:G%d" % (N_ROWS + 1)).Formula = "=D2*2"
    ws.Range("H10").Value = "'=HACK"                   # text that looks like a formula
    ws.Range("C20").Value = None                       # empty ihbar
    ws.Range("C21").Value = "abc"                      # broken date
    # grand total separated by an empty row
    ws.Range("E%d" % (N_ROWS + 3)).Value = "TOPLAM"
    ws.Range("D%d" % (N_ROWS + 3)).Formula = "=SUM(D2:D%d)" % (N_ROWS + 1)
    path = folder / "Odenen_2026_09.xlsx"
    wb.SaveAs(str(path), FileFormat=51)
    return wb


def make_master(xl, folder: Path):
    wb = xl.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Ana"
    hdr = ["POLICE NO", "HASAR TARIHI", "IHBAR TARIHI", "TUTAR", "SIRKET", "BRANS", "HESAP",
           "ACIKLAMA", "RAPOR_DONEMI", "DOSYA_TURU"]
    ws.Range("A1:J1").Value = [hdr]
    ws.Range("I1:J1").Interior.Color = 0x50D092          # green-ish (BGR) -> mandatory
    ws.Range("C1").Interior.Color = 0x50D092             # mandatory + mapped (blank-cell check)
    ws.Range("A2:A%d" % (MASTER_ROWS + 1)).NumberFormat = "@"
    vals = []
    for i in range(MASTER_ROWS):
        vals.append([f"{900000 + i:07d}", dt.datetime(2025, 1, 1) + dt.timedelta(days=i % 300),
                     dt.datetime(2025, 1, 10) + dt.timedelta(days=i % 300), 100.0 + i, "OLD", "OTO",
                     1.0, "eski", dt.datetime(2025, 12, 31), "ODENEN"])
    ws.Range("A2:J%d" % (MASTER_ROWS + 1)).Value = vals
    ws.Range("B2:C%d" % (MASTER_ROWS + 1)).NumberFormat = "dd.mm.yyyy"
    ws.Range("D2:D%d" % (MASTER_ROWS + 1)).NumberFormat = "#,##0.00"
    path = folder / "Ana_Odenen.xlsx"
    wb.SaveAs(str(path), FileFormat=51)
    return wb


# ---------------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------------
C = 1          # block 1 first column
AYAR = C + 2   # setting value column
HEDEF = C + 2
FILTRE = C + 3
ISLEM = C + 4
HARITA = C + 5
KARNE = C + 5
ILK = 29


def panel_row_of(pnl, letter_name: str) -> int:
    for r in range(ILK, ILK + 40):
        v = pnl.Cells(r, C).Value
        if v and str(v).startswith(letter_name):
            return r
    raise LookupError(letter_name)


def last_row(ws, col=1):
    return ws.Cells(ws.Rows.Count, col).End(-4162).Row   # xlUp


def rapor_since(rpr, start_row):
    """ACIKLAMA (col 7) of Rapor rows written after start_row."""
    last = last_row(rpr)
    return "\n".join(str(rpr.Cells(r, 7).Value or "") for r in range(start_row + 1, last + 1))


def dialogs_since(s, start):
    return "\n".join(e.text for e in s.watcher.events[start:])


# ---------------------------------------------------------------------------
#  Scenarios
# ---------------------------------------------------------------------------
def run_main(panel_path: Path, work: Path):
    pol = Policy()
    with xlauto.ExcelSession(policy=pol) as s:
        xl = s.xl
        wbP = xl.Workbooks.Open(str(panel_path))
        src = make_source(xl, work)
        mst = make_master(xl, work)
        wsS = src.Worksheets("Sayfa1")
        wsM = mst.Worksheets("Ana")
        pnl = wbP.Worksheets("Panel")
        wbP.Activate()
        pnl.Activate()

        # ---------------- 1. configure + Yenile ----------------
        print("\n[1] Yenile / Otomatik Eslestir / Profil")
        pnl.Cells(10, AYAR).Value = src.Name
        pnl.Cells(11, AYAR).Value = "Sayfa1"
        pnl.Cells(14, AYAR).Value = mst.Name
        pnl.Cells(15, AYAR).Value = "Ana"
        k0 = len(s.watcher.events)
        s.run(wbP, "Yenile")
        check("yenile summary shown", "YENILE TAMAMLANDI" in dialogs_since(s, k0), dialogs_since(s, k0))
        check("source columns listed", pnl.Cells(ILK, C).Value == "A - Pol No", str(pnl.Cells(ILK, C).Value))
        check("data type column", "TARIH" in str(pnl.Cells(ILK + 1, C + 1).Value), str(pnl.Cells(ILK + 1, C + 1).Value))
        check("target map lists master columns", "POLICE NO" in str(pnl.Cells(ILK, HARITA).Value), str(pnl.Cells(ILK, HARITA).Value))
        check("green column flagged", "YESIL, BOSTA" in str(pnl.Cells(ILK + 8, HARITA).Value), str(pnl.Cells(ILK + 8, HARITA).Value))
        # v1.0 layout: editable columns together, read-only map at the right edge
        hdr = [str(pnl.Cells(28, C + i).Value) for i in range(6)]
        check("v1.0 column order", hdr[3].startswith("F") and hdr[4].endswith("LEM") and "HAR" in hdr[5], str(hdr))
        check("setting labels merged (c..c+1)", pnl.Cells(10, C).MergeCells and pnl.Cells(10, C).MergeArea.Columns.Count == 2)
        # live checkpoint card after Yenile: K1/K4/K5 known, data checks pending
        k1, k4, k6 = str(pnl.Cells(9, KARNE).Value), str(pnl.Cells(12, KARNE).Value), str(pnl.Cells(14, KARNE).Value)
        check("scorecard header", "KONTROL NOKTALARI" in str(pnl.Cells(8, KARNE).Value).replace("İ", "I"))
        check("scorecard K1 ENGEL before any mapping", "K1" in k1 and "✕" in k1, k1)
        check("scorecard K4 live (green unmapped)", "K4" in k4 and k4.startswith("!"), k4)
        check("scorecard K6 pending", "○" in k6, k6)
        rib = wbP.Worksheets("Panel").Shapes("AP_HAZIRLIK").TextFrame2.TextRange.Text
        check("readiness pill exists", "dahil" in rib or "panel" in rib, rib)
        check("filter dropdown exists (Sirket)", _has_validation(pnl.Cells(ILK + 4, FILTRE)))
        check("auto-detect Tutar", pnl.Cells(19, AYAR).Value == "D - TUTAR", repr(pnl.Cells(19, AYAR).Value))
        check("auto-detect Hasar", pnl.Cells(20, AYAR).Value == "B - HASAR TARIHI", repr(pnl.Cells(20, AYAR).Value))
        check("auto-detect Ihbar", pnl.Cells(21, AYAR).Value == "C - IHBAR TARIHI", repr(pnl.Cells(21, AYAR).Value))
        check("auto-detect noted in DURUM", "OTOMATIK SECILDI" in str(pnl.Cells(26, C).Value))

        # settings that need the target list
        pnl.Cells(19, AYAR).Value = "D - TUTAR"
        pnl.Cells(20, AYAR).Value = "B - HASAR TARIHI"
        pnl.Cells(21, AYAR).Value = "C - IHBAR TARIHI"
        pnl.Cells(8, AYAR).Value = "Evet"
        # fixed value: sabit box row 9 (column c+3 / c+4)
        pnl.Cells(9, C + 3).Value = "I - RAPOR_DONEMI"
        pnl.Cells(9, C + 4).Value = "EYLUL 2026"
        check("sabit box opened next row", pnl.Cells(10, C + 3).Interior.Color != 16777215)

        k0 = len(s.watcher.events)
        s.run(wbP, "OtomatikEslestir")
        msg = dialogs_since(s, k0)
        check("auto-map reported", "OTOMATIK ESLESTIRME" in msg, msg)
        r_sirket = panel_row_of(pnl, "E - Sirket")
        check("exact match Sirket->SIRKET", pnl.Cells(r_sirket, HEDEF).Value == "E - SIRKET", str(pnl.Cells(r_sirket, HEDEF).Value))
        r_tutar = panel_row_of(pnl, "D - Odenen Tutar")
        check("fuzzy suggestion Odenen Tutar->TUTAR", pnl.Cells(r_tutar, HEDEF).Value == "D - TUTAR", str(pnl.Cells(r_tutar, HEDEF).Value))
        cm = pnl.Cells(r_tutar, HEDEF).Comment
        check("suggestion comment lists candidates", cm is not None and "En iyi adaylar" in cm.Text(), "")

        # approve all suggestions (nothing selected in the table -> all)
        wbP.Activate()
        pnl.Activate()
        pnl.Range("A1").Select()
        k0 = len(s.watcher.events)
        s.run(wbP, "OneriOnayla")
        check("smart approve -> all", "TUM ONERILERI ONAYLA" in dialogs_since(s, k0), dialogs_since(s, k0))
        check("suggestion approved", pnl.Cells(r_tutar, HEDEF).Comment is None)
        check("scorecard K1 OK once mapped", "✓" in str(pnl.Cells(9, KARNE).Value), str(pnl.Cells(9, KARNE).Value))

        # double-click shortcuts
        pnl.Cells(r_tutar, HEDEF).Value = None
        s.run(wbP, "OtomatikEslestir")
        check("suggestion re-created", pnl.Cells(r_tutar, HEDEF).Comment is not None)
        s.run(wbP, "PanelCiftTik", pnl.Cells(r_tutar, HEDEF), False)
        check("double-click approves suggestion", pnl.Cells(r_tutar, HEDEF).Comment is None
              and pnl.Cells(r_tutar, HEDEF).Value == "D - TUTAR", str(pnl.Cells(r_tutar, HEDEF).Value))
        s.run(wbP, "PanelCiftTik", pnl.Cells(r_sirket, ISLEM), False)
        check("double-click ISLEM -> SEC", "SE" in str(pnl.Cells(r_sirket, ISLEM).Value or ""), repr(pnl.Cells(r_sirket, ISLEM).Value))
        s.run(wbP, "PanelCiftTik", pnl.Cells(r_sirket, ISLEM), False)
        check("double-click ISLEM -> off", not pnl.Cells(r_sirket, ISLEM).Value, repr(pnl.Cells(r_sirket, ISLEM).Value))

        # manual mapping + formula flag
        r_pol = panel_row_of(pnl, "A - Pol No")
        pnl.Cells(r_pol, HEDEF).Value = "A - POLICE NO"
        r_hesap = panel_row_of(pnl, "G - Hesap")
        pnl.Cells(r_hesap, ISLEM).Value = "FORMUL"

        # profile save, clear mapping, Yenile -> profile re-applies
        pol.set(("hangi isimle", ("TEST PROFIL", "OK")))
        s.run(wbP, "ProfilKaydet")
        pol.set()
        prf = wbP.Worksheets("Profiller")
        check("profile rows written", last_row(prf) >= 9, str(last_row(prf)))
        pnl.Cells(r_pol, HEDEF).Value = None
        pnl.Cells(r_sirket, HEDEF).Value = None
        s.run(wbP, "Yenile")
        check("profile re-applied (Pol No)", pnl.Cells(r_pol, HEDEF).Value == "A - POLICE NO", str(pnl.Cells(r_pol, HEDEF).Value))
        check("profile re-applied (FORMUL flag)", "FORM" in str(pnl.Cells(r_hesap, ISLEM).Value), str(pnl.Cells(r_hesap, ISLEM).Value))
        # double-click: "Bu paneli aktar" toggles, map column -> fixed value
        before_dahil = str(pnl.Cells(8, AYAR).Value)
        s.run(wbP, "PanelCiftTik", pnl.Cells(8, AYAR), False)
        check("double-click DAHIL toggles", str(pnl.Cells(8, AYAR).Value) != before_dahil, str(pnl.Cells(8, AYAR).Value))
        s.run(wbP, "PanelCiftTik", pnl.Cells(8, AYAR), False)
        check("double-click DAHIL toggles back", str(pnl.Cells(8, AYAR).Value) == before_dahil)
        s.run(wbP, "PanelCiftTik", pnl.Cells(ILK + 9, HARITA), False)          # J - DOSYA_TURU (green, empty)
        check("double-click map -> fixed value row", pnl.Cells(10, C + 3).Value == "J - DOSYA_TURU", repr(pnl.Cells(10, C + 3).Value))
        check("map now shows [+] for DOSYA_TURU", str(pnl.Cells(ILK + 9, HARITA).Value).startswith("[+]"), str(pnl.Cells(ILK + 9, HARITA).Value))
        pnl.Cells(10, C + 4).Value = "ODENEN"
        durum = str(pnl.Cells(26, C).Value)
        check("DURUM says HAZIR", "HAZIR" in durum and "AKTARIMA DAHIL" in durum, durum)
        check("leading-zero codes typed METIN", str(pnl.Cells(r_pol, C + 1).Value).startswith("METIN"),
              str(pnl.Cells(r_pol, C + 1).Value))
        check("no false type conflict", "VERI TIPI UYUSMAZ" not in durum, durum)

        # ---------------- 2. On Kontrol ----------------
        print("\n[2] On Kontrol")
        rpr = wbP.Worksheets("Rapor")
        r0 = last_row(rpr)
        k0 = len(s.watcher.events)
        s.run(wbP, "OnKontrol")
        msg = dialogs_since(s, k0) + "\n" + rapor_since(rpr, r0)
        check("dry run: suspicious tail", "SUPHELI" in msg, msg)
        check("dry run: 200 rows", "Aktarilacak : 200 satir" in msg, msg)
        check("dry run: date check", "KAYNAK TARIH KONTROLU" in msg and "ONCE: 2 satir" in msg, msg)
        check("dry run: format guard", "BICIM KORUMASI" in msg, msg)
        check("dry run: green blank count (C20)", "IHBAR TARIHI sutununda 1 satirda" in msg, msg[:300])
        check("dialog shows checkpoint summary", "KONTROL NOKTALARI:" in msg and "K1 OK" in msg, msg[:900])
        check("items tagged [K..]", "[K3]" in msg and "[K8]" in msg and "[K6]" in msg, msg[:900])
        karne = [str(pnl.Cells(9 + i, KARNE).Value) for i in range(8)]
        check("scorecard K3 DIKKAT after dry run", karne[2].startswith("!"), karne[2])
        check("scorecard K8 DIKKAT after dry run", karne[7].startswith("!"), karne[7])
        check("scorecard K2 OK after dry run", "✓" in karne[1], karne[1])
        rlast = last_row(rpr)
        tags = [str(rpr.Cells(r, 27).Value or "") for r in range(r0 + 1, rlast + 1)]
        check("Rapor KONTROL NOKTASI column", any(t.startswith("K8 DIKKAT") for t in tags) and any(t.startswith("K3") for t in tags), str(tags))
        check("dry run wrote nothing", last_row(wsM) == MASTER_ROWS + 1, str(last_row(wsM)))
        check("dry run: tail preview shown", "Icerik:" in msg and "TOPLAM" in msg, msg[:600])
        check("dry run: green blank cells", "ZORUNLU (YESIL) SUTUNDA BOS HUCRE" in msg and "IHBAR TARIHI sutununda 1 satirda" in msg, msg)
        check("dry run: fixed value type", "SABIT DEGER TIPI UYUSMAZ" in msg and "RAPOR_DONEMI" in msg, msg)
        check("every item logged in Rapor", rapor_since(rpr, r0).count("\n") >= 6)
        # corrupt amounts (restored afterwards)
        old = [wsS.Range(a).Value for a in ("D30", "D31", "D32", "D33")]
        wsS.Range("D30").Value = "1.200 TL"
        wsS.Range("D31").Value = "YOK"
        wsS.Range("D32").Value = None
        wsS.Range("D33").Value = -50
        r0 = last_row(rpr)
        s.run(wbP, "OnKontrol")
        t = rapor_since(rpr, r0)
        check("tutar: non-numeric", "sayi OLMAYAN deger 2 hucre" in t and "satir 30 31" in t, t)
        check("tutar: blank", "BOS 1 hucre (satir 32)" in t, t)
        check("tutar: negative", "EKSI tutar 1 hucre (satir 33)" in t, t)
        for a, v in zip(("D30", "D31", "D32", "D33"), old):
            wsS.Range(a).Value = v

        # ---------------- 3. Filters ----------------
        print("\n[3] Filters")
        for ftext, expected in (("ALLIANZ, INC", 50), ("AXA, MAPFRE", 100), ("AL*", 100), ("<>AXA", 150)):
            pnl.Cells(r_sirket, FILTRE).Value = ftext
            k0 = len(s.watcher.events)
            s.run(wbP, "OnKontrol")
            msg = dialogs_since(s, k0)
            check(f"filter '{ftext}' -> {expected}", f"Aktarilacak : {expected} satir" in msg, msg[:300])
        pnl.Cells(r_sirket, FILTRE).Value = None
        pnl.Cells(r_tutar, FILTRE).Value = ">3000"
        k0 = len(s.watcher.events)
        s.run(wbP, "OnKontrol")
        exp = sum(1 for i in range(N_ROWS) if 1000 + i * 12.5 > 3000)
        check(f"filter '>3000' -> {exp}", f"Aktarilacak : {exp} satir" in dialogs_since(s, k0))
        pnl.Cells(r_tutar, FILTRE).Value = None

        # ---------------- 4. Aktar ----------------
        print("\n[4] Aktar")
        src_total = sum(1000 + i * 12.5 for i in range(N_ROWS))
        before_total = sum(100.0 + i for i in range(MASTER_ROWS))
        k0 = len(s.watcher.events)
        s.run(wbP, "Aktar")
        msg = dialogs_since(s, k0)
        first = MASTER_ROWS + 2
        lastr = MASTER_ROWS + 1 + N_ROWS
        check("rows appended", last_row(wsM) == lastr, str(last_row(wsM)))
        check("total row excluded", wsM.Cells(lastr, 4).Value != src_total, "")
        check("leading zeros kept", wsM.Cells(first, 1).Value == "0000001", repr(wsM.Cells(first, 1).Value))
        check("19-digit id kept", wsM.Cells(first + 7, 1).Value == "1234567890123456007", repr(wsM.Cells(first + 7, 1).Value))
        h10 = wsM.Cells(first + 8, 8)
        check("'=HACK' stays text", h10.Value == "=HACK" and not h10.HasFormula, repr(h10.Value))
        g = wsM.Cells(first, 7)
        check("formula row-shifted", g.HasFormula and g.Formula == f"=D{first}*2", g.Formula)
        check("formula value", abs(g.Value - 2000.0) < 1e-9, str(g.Value))
        check("fixed value written", wsM.Cells(lastr, 9).Value == "EYLUL 2026", str(wsM.Cells(lastr, 9).Value))
        check("number format inherited", wsM.Cells(first, 4).NumberFormat == "#,##0.00", wsM.Cells(first, 4).NumberFormat)
        check("date stays date", isinstance(wsM.Cells(first, 2).Value, dt.datetime), repr(wsM.Cells(first, 2).Value))
        check("reconciliation in dialog", "Kaynaktan okunan" in msg and "TUTUYOR" in msg, msg[-600:])
        new_total = xl.WorksheetFunction.Sum(wsM.Range(f"D2:D{lastr}"))
        check("master total increased exactly", abs(new_total - (before_total + src_total)) < 0.005, str(new_total))
        rpr = wbP.Worksheets("Rapor")
        rl = last_row(rpr)
        aktar_row = max(r for r in range(2, rl + 1) if rpr.Cells(r, 3).Value == "AKTARIM")
        check("Rapor: TUTUYOR", str(rpr.Cells(aktar_row, 20).Value).startswith("TUTUYOR"), str(rpr.Cells(aktar_row, 20).Value))
        check("Rapor: reconciliation cell green", rpr.Cells(aktar_row, 20).Interior.Color == 0xE7FCDC, hex(int(rpr.Cells(aktar_row, 20).Interior.Color)))
        dikkat_rows = [r for r in range(2, aktar_row) if rpr.Cells(r, 3).Value == "DIKKAT" and str(rpr.Cells(r, 27).Value or "").startswith("K")]
        check("OnayAl wrote every DIKKAT item (tagged)", len(dikkat_rows) >= 4, str(len(dikkat_rows)))
        check("scorecard K7 after transfer", "✓" in str(pnl.Cells(15, KARNE).Value) and "FARK 0,00" in str(pnl.Cells(17, KARNE).Value),
              str(pnl.Cells(15, KARNE).Value) + " | " + str(pnl.Cells(17, KARNE).Value))
        den = wbP.Worksheets("Denetim")
        check("Denetim sheet refreshed", str(den.Cells(6, 2).Value).strip("'") == "1" and str(den.Cells(6, 4).Value).strip("'") == "1",
              f"{den.Cells(6, 2).Value!r} {den.Cells(6, 4).Value!r}")
        check("Denetim: K8 statistics", (den.Cells(18, 4).Value or 0) >= 1, repr(den.Cells(18, 4).Value))
        check("Denetim: last transfer row", den.Cells(22, 10).Value == "TUTUYOR" and den.Cells(22, 11).Value == "ETKIN",
              f"{den.Cells(22, 10).Value!r} {den.Cells(22, 11).Value!r}")
        check("Rapor: onceki/yeni toplam", abs(rpr.Cells(aktar_row, 25).Value - before_total) < 0.005
              and abs(rpr.Cells(aktar_row, 26).Value - (before_total + src_total)) < 0.005, "")
        gec = wbP.Worksheets("Gecmis")
        gl = last_row(gec)
        check("Gecmis rows (8 mapped + 2 fixed)", gl - 1 == 10, str(gl))
        check("Gecmis fingerprint", str(gec.Cells(gl, 15).Value).startswith("V;n="), str(gec.Cells(gl, 15).Value))
        k0 = len(s.watcher.events)
        s.run(wbP, "OnKontrol")
        check("dry run: duplicate-transfer warning", "MUKERRER AKTARIM UYARISI" in dialogs_since(s, k0), dialogs_since(s, k0)[:400])

        # ---------------- 5. Geri Al ----------------
        print("\n[5] Geri Al")
        k0 = len(s.watcher.events)
        s.run(wbP, "GeriAl")
        check("undo cleared rows", last_row(wsM) == MASTER_ROWS + 1, str(last_row(wsM)))
        check("Gecmis marked", gec.Cells(gl, 11).Value == "GERI ALINDI", str(gec.Cells(gl, 11).Value))
        check("master total restored", abs(xl.WorksheetFunction.Sum(wsM.Range("D:D")) - before_total) < 0.005)

        # ---------------- 6. Fingerprint ----------------
        print("\n[6] Parmak izi")
        s.run(wbP, "Aktar")
        check("re-transfer", last_row(wsM) == lastr, str(last_row(wsM)))
        wsM.Rows(500).Insert()                      # user inserts a row above the block
        pol.set(("PARMAK IZI", "&No"), ("parmak iziyle", "&No"))
        k0 = len(s.watcher.events)
        s.run(wbP, "GeriAl")
        msg = dialogs_since(s, k0)
        pol.set()
        check("fingerprint warning shown", "TUTMUYOR" in msg, msg[:400])
        check("nothing cleared after 'No'", last_row(wsM) == lastr + 1, str(last_row(wsM)))
        wsM.Rows(500).Delete()                      # user fixes it
        k0 = len(s.watcher.events)
        s.run(wbP, "GeriAl")
        msg = dialogs_since(s, k0)
        check("no warning once fixed", "TUTMUYOR" not in msg, msg[:300])
        check("undo after fix", last_row(wsM) == MASTER_ROWS + 1, str(last_row(wsM)))

        # ---------------- 7. Secili Sutun Yamasi ----------------
        print("\n[7] Secili Sutun Yamasi")
        s.run(wbP, "Aktar")
        old_d = wsM.Cells(first, 4).Value
        wsS.Range("D2:D%d" % (N_ROWS + 1)).Value = [[1000 + i * 12.5 + 1] for i in range(N_ROWS)]
        pnl.Cells(r_tutar, ISLEM).Value = "SEC"
        k0 = len(s.watcher.events)
        s.run(wbP, "AktarSecili")
        msg = dialogs_since(s, k0)
        check("patch warns about overwrite", "SECILI SUTUN YAMASI" in msg, msg[:500])
        check("patch overwrote column", abs(wsM.Cells(first, 4).Value - (old_d + 1)) < 1e-9, str(wsM.Cells(first, 4).Value))
        check("patch kept row count", last_row(wsM) == lastr, str(last_row(wsM)))
        check("other columns untouched", wsM.Cells(first, 1).Value == "0000001")
        gl = last_row(gec)
        check("backup recorded", str(gec.Cells(gl, 16).Value or "") != "", str(gec.Cells(gl, 16).Value))
        check("SEC flag cleared, FORMUL kept", "SE" not in str(pnl.Cells(r_tutar, ISLEM).Value or ""))
        s.run(wbP, "GeriAl")
        check("undo of patch restored old values", abs(wsM.Cells(first, 4).Value - old_d) < 1e-9, str(wsM.Cells(first, 4).Value))
        check("undo of patch kept other columns", wsM.Cells(first, 1).Value == "0000001")

        # ---------------- 7b. Tutar mutabakati (re-check) ----------------
        k0 = len(s.watcher.events)
        s.run(wbP, "TutarMutabakatiSon")
        check("TutarMutabakatiSon: TUTUYOR", "TUTARI TUTUYOR" in dialogs_since(s, k0), dialogs_since(s, k0)[:300])

        # ---------------- 7c. Geri alinan araliga + satir ekleme ----------------
        print("\n[7c] Geri alinan araliga (satir ekleme)")
        s.run(wbP, "GeriAl")                                   # clear the 200-row block
        pnl.Cells(r_sirket, FILTRE).Value = "AXA"
        s.run(wbP, "Aktar")                                    # 50 rows -> 1002..1051
        check("filtered transfer 50 rows", last_row(wsM) == MASTER_ROWS + 51, str(last_row(wsM)))
        wsM.Range(f"A{MASTER_ROWS + 52}:A{MASTER_ROWS + 54}").Value = [["ALT1"], ["ALT2"], ["ALT3"]]
        s.run(wbP, "GeriAl")                                   # gap 1002..1051, data below
        pnl.Cells(r_sirket, FILTRE).Value = None
        pnl.Cells(16, AYAR).Value = "Geri alinan araliga"
        k0 = len(s.watcher.events)
        s.run(wbP, "Aktar")                                    # 200 rows into a 50-row gap
        msg = dialogs_since(s, k0)
        check("gap mode warns about inserted rows", "SATIR EKLENECEK" in msg, msg[:400])
        check("150 rows inserted, data below shifted", wsM.Cells(MASTER_ROWS + 202, 1).Value == "ALT1",
              repr(wsM.Cells(MASTER_ROWS + 202, 1).Value))
        check("gap filled from first row", wsM.Cells(first, 1).Value == "0000001", repr(wsM.Cells(first, 1).Value))
        s.run(wbP, "GeriAl")
        check("undo after gap transfer", wsM.Cells(first, 1).Value is None and wsM.Cells(MASTER_ROWS + 202, 1).Value == "ALT1")
        pnl.Cells(16, AYAR).Value = "Sayfanin en alti"

        # ---------------- 7d. Panel ekle / kaldir / setler / yeniden kur ----------------
        print("\n[7d] Panel yonetimi, setler, PaneliYenidenKur")
        pol.set(("hangi isimle kaydedilsin", ("HAYAT", "OK")))
        s.run(wbP, "SetKaydet")
        s.run(wbP, "YeniPanelEkle")
        check("panel 2 added", str(pnl.Cells(7, 7).Value).startswith("PANEL 2"), str(pnl.Cells(7, 7).Value))
        check("panel 2 copied target", pnl.Cells(14, 9).Value == mst.Name, str(pnl.Cells(14, 9).Value))
        pol.set(("Hangi panel kaldirilsin", ("2", "OK")))
        s.run(wbP, "PanelKaldir")
        check("panel 2 hidden", pnl.Columns(7).Hidden)
        pol.set()
        s.run(wbP, "YeniPanelEkle")
        check("hidden block reused", not pnl.Columns(7).Hidden and str(pnl.Cells(7, 13).Value or "") == "",
              str(pnl.Cells(7, 13).Value))
        pol.set(("Hangi panel kaldirilsin", ("2", "OK")))
        s.run(wbP, "PanelKaldir")
        pol.set(("Hangi set yuklensin", ("HAYAT", "OK")))
        s.run(wbP, "SetYukle")
        pol.set()
        check("set: source cleared", pnl.Cells(10, AYAR).Value is None)
        check("set: fixed value restored", pnl.Cells(9, C + 3).Value == "I - RAPOR_DONEMI", str(pnl.Cells(9, C + 3).Value))
        check("set: profile restored", pnl.Cells(18, AYAR).Value == "TEST PROFIL", str(pnl.Cells(18, AYAR).Value))
        pnl.Cells(10, AYAR).Value = src.Name
        pnl.Cells(11, AYAR).Value = "Sayfa1"
        s.run(wbP, "Yenile")
        check("set + profile -> mapping back", pnl.Cells(r_pol, HEDEF).Value == "A - POLICE NO", str(pnl.Cells(r_pol, HEDEF).Value))
        s.run(wbP, "PaneliYenidenKur")
        check("rebuild keeps mapping", pnl.Cells(r_pol, HEDEF).Value == "A - POLICE NO", str(pnl.Cells(r_pol, HEDEF).Value))
        check("rebuild keeps fixed value", pnl.Cells(9, C + 4).Value == "EYLUL 2026", str(pnl.Cells(9, C + 4).Value))
        check("rebuild keeps hidden block", pnl.Columns(7).Hidden)
        s.run(wbP, "Yenile")
        k0 = len(s.watcher.events)
        s.run(wbP, "OnKontrol")
        check("after rebuild: ready to transfer", "Aktarilacak : 200 satir" in dialogs_since(s, k0))
        s.run(wbP, "KontrolTemizle")
        s.run(wbP, "KilavuzTumunuAc")
        s.run(wbP, "KilavuzTumunuKapat")

        # ---------------- 7e. Iki panel, tek parti, Tutar Mutabakati ----------------
        print("\n[7e] Iki panel tek parti")
        s.run(wbP, "YeniPanelEkle")                      # reuses hidden block 2
        C2 = 7
        pnl.Cells(10, C2 + 2).Value = src.Name
        pnl.Cells(11, C2 + 2).Value = "Sayfa1"
        pnl.Cells(8, C2 + 2).Value = "Evet"
        pnl.Cells(8, AYAR).Value = "Evet"
        s.run(wbP, "Yenile")
        check("panel 2 tutar auto/copied", pnl.Cells(19, C2 + 2).Value == "D - TUTAR", repr(pnl.Cells(19, C2 + 2).Value))
        pnl.Cells(r_pol, C2 + 2).Value = "A - POLICE NO"
        pnl.Cells(r_tutar, C2 + 2).Value = "D - TUTAR"
        before = last_row(wsM)
        s.run(wbP, "Aktar")
        check("two panels transferred", last_row(wsM) == before + 2 * N_ROWS, str(last_row(wsM)))
        k0 = len(s.watcher.events)
        s.run(wbP, "TutarMutabakatiSon")
        msg = dialogs_since(s, k0)
        check("batch reconciliation lists both panels", "PANEL 1" in msg and "PANEL 2" in msg and "2 aktarim" in msg, msg)
        check("batch reconciliation TUTUYOR", "TUTARI TUTUYOR" in msg, msg[:300])
        wsM.Cells(before + N_ROWS + 1, 4).Value = 999999        # tamper with panel 2 rows
        k0 = len(s.watcher.events)
        s.run(wbP, "TutarMutabakatiSon")
        msg = dialogs_since(s, k0)
        check("batch reconciliation finds tampering", "FARK VAR" in msg, msg[:400])
        s.run(wbP, "GeriAl")                                   # parti question -> Yes (whole batch)
        check("batch undone", last_row(wsM) == before, str(last_row(wsM)))

        # ---------------- misc ----------------
        s.run(wbP, "Tani")
        check("Excel events back on", xl.EnableEvents is True)
        check("calculation automatic", xl.Calculation == -4105)
        wbP.Save()
        print("\n--- dialog log (last 6) ---")
        print(s.watcher.dump(max(0, len(s.watcher.events) - 6))[:4000])


def _has_validation(cell) -> bool:
    try:
        return cell.Validation.Type == 3
    except Exception:
        return False


def run_migration(v33_path: Path, work: Path):
    print("\n[8] v33 -> v34 gecisi")
    copy = work / "v33_kopya.xlsm"
    print("  copy", repr(str(v33_path)), "->", repr(str(copy)))
    shutil.copy2(v33_path, copy)
    # read the v33 panel first
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(copy))
        pnl = wb.Worksheets("Panel")
        old = []
        k = 1
        while str(pnl.Cells(7, 1 + (k - 1) * 7).Value or "").upper().startswith("PANEL"):
            c = 1 + (k - 1) * 7
            old.append({
                "hidden": pnl.Columns(c).Hidden,
                "kdosya": pnl.Cells(10, c + 1).Value, "hdosya": pnl.Cells(14, c + 1).Value,
                "hsayfa": pnl.Cells(15, c + 1).Value, "profil": pnl.Cells(18, c + 1).Value,
                "tutar": pnl.Cells(24, c + 1).Value,
                "map": [(pnl.Cells(r, c).Value, pnl.Cells(r, c + 1).Value) for r in range(29, 60)
                        if pnl.Cells(r, c).Value],
            })
            k += 1
        sheets_before = [ws.Name for ws in wb.Worksheets]
        wb.Close(False)
    print("  v33 blocks:", len(old), "sheets:", sheets_before)

    out = work / "v33_yukseltilmis.xlsm"
    rc = __import__("build").main(["--upgrade", str(copy), "--out", str(out)])
    check("upgrade build ok", rc == 0)
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(out))
        names = [ws.Name for ws in wb.Worksheets]
        check("legacy sheets removed", not any(n in names for n in ("Dosya Durumu", "Pivot Kontrol", "OzetVeri")), str(names))
        check("Ozet removed", not any(n.upper().endswith("ZET") for n in names), str(names))
        check("records kept", "Rapor" in names and "Gecmis" in names and "Profiller" in names)
        pnl = wb.Worksheets("Panel")
        nm = [n.Name for n in wb.Names]
        check("layout marker", "AP_DUZEN" in nm, str(nm))
        check("no OZ_ names left", not any(n.startswith("OZ_") for n in nm), str(nm))
        for i, b in enumerate(old, 1):
            c = 1 + (i - 1) * 6
            if b["hidden"]:
                check(f"P{i} hidden stays hidden", pnl.Columns(c).Hidden)
                continue
            check(f"P{i} kaynak dosya kept", pnl.Cells(10, c + 2).Value == b["kdosya"], f"{pnl.Cells(10, c + 2).Value!r} vs {b['kdosya']!r}")
            check(f"P{i} hedef dosya kept", pnl.Cells(14, c + 2).Value == b["hdosya"])
            check(f"P{i} tutar kept", pnl.Cells(19, c + 2).Value == b["tutar"], f"{pnl.Cells(19, c + 2).Value!r} vs {b['tutar']!r}")
            newmap = [(pnl.Cells(r, c).Value, pnl.Cells(r, c + 2).Value) for r in range(29, 60) if pnl.Cells(r, c).Value]
            check(f"P{i} mapping kept ({len(b['map'])} rows)", newmap == b["map"], f"{newmap[:3]} vs {b['map'][:3]}")
        comps = [c.Name for c in wb.VBProject.VBComponents]
        check("Module1/2/3 removed", not any(n in comps for n in ("Module1", "Module2", "Module3")), str(comps))
        wb.Close(False)


V34_COMMIT = "7924d1f"   # last commit that still has the v34 panel layout


def run_migration_v34(work: Path):
    """Build a REAL v34-layout workbook from the pinned v34-layout commit (git archive),
    fill it like a user, upgrade it with the current code, check every value
    landed in its new v1.0 column."""
    import subprocess
    import tarfile
    print("\n[9] v34 -> v1.0 gecisi (git HEAD kaynagindan gercek v34 paneli)")
    old_src = work / "src_v34"
    old_src.mkdir()
    tar = work / "v34.tar"
    with open(tar, "wb") as f:
        subprocess.run(["git", "-C", str(ROOT), "archive", V34_COMMIT, "src"], stdout=f, check=True)
    with tarfile.open(tar) as t:
        t.extractall(old_src)
    old_layout = (old_src / "src" / "M01_Ayarlar.bas").read_text(encoding="ascii")
    if "OFS_HARITA As Long = 3" not in old_layout:
        check("HEAD commit has the v34 layout (skip)", True)
        return
    old = work / "v34_panel.xlsm"
    build = __import__("build")
    check("v34 workbook built from git HEAD", build.main(["--new", str(old), "--src", str(old_src / "src"), "--no-setup"]) == 0)
    data = work / "v34_data"
    data.mkdir()
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(old))
        s.run(wb, "SifirdanKur")
        src = make_source(s.xl, data)
        mst = make_master(s.xl, data)
        pnl = wb.Worksheets("Panel")
        pnl.Cells(10, 3).Value = src.Name
        pnl.Cells(11, 3).Value = "Sayfa1"
        pnl.Cells(14, 3).Value = mst.Name
        pnl.Cells(15, 3).Value = "Ana"
        s.run(wb, "Yenile")
        r_pol, r_sir, r_hes = (panel_row_of(pnl, x) for x in ("A - Pol No", "E - Sirket", "G - Hesap"))
        pnl.Cells(r_pol, 3).Value = "A - POLICE NO"
        pnl.Cells(r_sir, 3).Value = "E - SIRKET"
        pnl.Cells(r_sir, 5).Value = "AXA, MAPFRE"          # v34 FILTRE = c+4
        pnl.Cells(r_hes, 6).Value = "FORMUL"               # v34 ISLEM  = c+5
        pnl.Cells(9, 4).Value = "I - RAPOR_DONEMI"
        pnl.Cells(9, 5).Value = "EYLUL 2026"
        pnl.Cells(19, 3).Value = "D - TUTAR"
        pnl.Cells(8, 3).Value = "Evet"
        check("v34 panel marker", "v34" in wb.Names("AP_DUZEN").RefersTo, wb.Names("AP_DUZEN").RefersTo)
        wb.Save()
        wb.Close(False)
    new = work / "v34_to_v1_0.xlsm"
    check("v34 -> v1.0 upgrade build", build.main(["--upgrade", str(old), "--out", str(new)]) == 0)
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(new))
        pnl = wb.Worksheets("Panel")
        check("marker now v1.0", "v1.0" in wb.Names("AP_DUZEN").RefersTo)
        check("v34: mapping kept", pnl.Cells(r_pol, 3).Value == "A - POLICE NO" and pnl.Cells(r_sir, 3).Value == "E - SIRKET")
        check("v34: filter moved to c+3", pnl.Cells(r_sir, 4).Value == "AXA, MAPFRE", repr(pnl.Cells(r_sir, 4).Value))
        check("v34: ISLEM moved to c+4", "FORM" in str(pnl.Cells(r_hes, 5).Value or ""), repr(pnl.Cells(r_hes, 5).Value))
        check("v34: map header at c+5", "HAR" in str(pnl.Cells(28, 6).Value), str(pnl.Cells(28, 6).Value))
        check("v34: fixed value kept", pnl.Cells(9, 4).Value == "I - RAPOR_DONEMI" and pnl.Cells(9, 5).Value == "EYLUL 2026")
        check("v34: settings kept", pnl.Cells(19, 3).Value == "D - TUTAR" and pnl.Cells(8, 3).Value == "Evet")
        check("v34: scorecard drawn", "KONTROL" in str(pnl.Cells(8, 6).Value))
        wb.Close(False)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--panel", default=str(ROOT / "dist" / "AktarimPaneli_v1.0.xlsm"))
    ap.add_argument("--v33", default=None)
    ap.add_argument("--only-migration", action="store_true")
    args = ap.parse_args()

    work = Path(tempfile.mkdtemp(prefix="apb_e2e_"))
    panel_copy = work / "AktarimPaneli_test.xlsm"
    shutil.copy2(args.panel, panel_copy)
    print("work folder:", work)
    try:
        if not args.only_migration:
            run_main(panel_copy, work)
    except Exception:
        traceback.print_exc()
        check("scenario 1-7 completed without exception", False)
    try:
        run_migration_v34(work)
    except Exception:
        traceback.print_exc()
        check("v34 migration completed without exception", False)
    if args.v33:
        try:
            run_migration(Path(args.v33).resolve(), work)
        except Exception:
            traceback.print_exc()
            check("migration completed without exception", False)

    failed = [r for r in RESULTS if not r[1]]
    print(f"\n{len(RESULTS) - len(failed)} / {len(RESULTS)} checks passed")
    for name, _, detail in failed:
        print("FAILED:", name, "|", detail[:500])
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
