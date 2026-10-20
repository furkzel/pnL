"""
release.py - build the three download packages for a release.

  python tools/release.py [--version v1.0]

Output in dist/release/:
  AktarimPaneli_<v>.xlsm          workbook + macros (ready to use)
  AktarimPaneli_<v>_makrosuz.xlsx  same workbook WITHOUT macros (for networks
                                   that block macro files on download/e-mail)
  AktarimPaneli_<v>_makrolar.zip   only the VBA modules (CRLF, ready for the
                                   VBA editor's File > Import) + KURULUM.txt
  SHA256SUMS.txt                   checksums to verify the downloads

Split install = xlsx + zip: open the xlsx, import the modules, save as .xlsm.
"""
from __future__ import annotations

import argparse
import hashlib
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build  # noqa: E402
import xlauto  # noqa: E402

KURULUM = """AKTARIM PANELI {v} - MAKROLARI ELLE KURMA
=========================================

Bu paket, makro iceren dosyalari indirmeyi engelleyen ag/e-posta
politikalari icindir. Excel dosyasi (makrosuz .xlsx) ve makrolar (bu .zip)
ayri indirilir, sizin bilgisayarinizda birlestirilir.

1) AktarimPaneli_{v}_makrosuz.xlsx dosyasini acin.
   Dosya > Farkli Kaydet > "Excel Makro Etkin Calisma Kitabi (*.xlsm)".

2) Alt+F11 ile VBA editorunu acin.
   File > Import File... ile bu klasordeki dosyalari TEK TEK ice aktarin:
       Sozluk.cls
       M01_Ayarlar.bas ... M10_Denetim.bas  (10 dosya)

3) Sol panelde "Panel" sayfasinin kod sayfasina (parantez icinde "Panel"
   yazan satir) cift tiklayin. Panel_sayfa_kodu.txt dosyasini Not
   Defteri ile acin, TAMAMINI kopyalayip yapistirin.

4) Ayni seyi "ThisWorkbook" (Turkce Excel'de "BuCalismaKitabi") icin
   ThisWorkbook_kodu.txt ile yapin.

5) Debug > Compile VBAProject  (hata cikmamali).

6) Dosyayi kaydedin, kapatip yeniden acin, "Icerigi Etkinlestir"e basin.
   Panel butonlari artik calisir. Ayrinti: Kilavuz sayfasi.

GUVENLIK
- Dosyalarin butunlugunu SHA256SUMS.txt ile dogrulayabilirsiniz
  (PowerShell: Get-FileHash dosya -Algorithm SHA256).
- Kodun tamami ACIK metindir; ice aktarmadan once okuyabilirsiniz.
- Makrolar disari baglanmaz, dosya silmez, yalnizca sizin actiginiz
  Excel dosyalarinda calisir.
"""


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


TRACES = (b"fozel", b"furkan", b"furkzel")


def audit(p: Path) -> list[str]:
    """Search every member of an OOXML package for the builder's identity."""
    bad = []
    with zipfile.ZipFile(p) as z:
        for n in z.namelist():
            low = z.read(n).lower()
            for t in TRACES:
                if t in low:
                    bad.append(f"{p.name}: {n} contains {t.decode()}")
    return bad


def crlf(text: str) -> bytes:
    return text.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="v1.0")
    a = ap.parse_args(argv)
    v = a.version
    out = xlauto.ROOT / "dist" / "release"
    out.mkdir(parents=True, exist_ok=True)
    xlsm = out / f"AktarimPaneli_{v}.xlsm"
    xlsx = out / f"AktarimPaneli_{v}_makrosuz.xlsx"
    zp = out / f"AktarimPaneli_{v}_makrolar.zip"

    # 1) merged workbook (built + set up from src/)
    if build.main(["--new", str(xlsm)]) != 0:
        return 1

    # 2) macro-free copy: same sheets, panel and guide; VBA project dropped
    if xlsx.exists():
        xlsx.unlink()
    with xlauto.ExcelSession() as s:
        wb = s.xl.Workbooks.Open(str(xlsm))
        xlauto.scrub_personal_info(wb)
        wb.SaveAs(str(xlsx), FileFormat=51)
        wb.Close(False)
    xlauto.strip_package_identity(xlsx)
    print("saved:", xlsx)

    # 3) macros only
    src = xlauto.SRC
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        for f in [src / "Sozluk.cls"] + sorted(src.glob("M*.bas")):
            z.writestr(f"makrolar/{f.name}", crlf(f.read_text(encoding="ascii")))
        z.writestr("makrolar/Panel_sayfa_kodu.txt", crlf((src / "Panel.cls").read_text(encoding="ascii")))
        z.writestr("makrolar/ThisWorkbook_kodu.txt", crlf((src / "ThisWorkbook.cls").read_text(encoding="ascii")))
        z.writestr("makrolar/KURULUM.txt", crlf(KURULUM.format(v=v)))
    print("saved:", zp)

    # Distribution files must be free of local identity traces
    bad = audit(xlsm) + audit(xlsx)
    for src_f in sorted(src.glob("*")):
        low = src_f.read_bytes().lower()
        bad += [f"{src_f.name} contains {t.decode()}" for t in TRACES if t in low]
    if bad:
        print("IDENTITY TRACES FOUND:")
        print(chr(10).join(bad))
        return 1
    print("identity audit: clean")

    sums = "".join(f"{sha256(p)}  {p.name}" + chr(10) for p in (xlsm, xlsx, zp))
    (out / "SHA256SUMS.txt").write_text(sums, encoding="ascii", newline=chr(10))
    print(sums)

    # Keep the download page's checksum table in sync with the packages
    page = xlauto.ROOT / "docs" / "index.html"
    if page.exists():
        html = page.read_text(encoding="utf-8")
        for f in (xlsm, xlsx, zp):
            html = re.sub(r"(<td>" + re.escape(f.name) + r"</td><td class=\"hash\">)[0-9a-f]{64}",
                          lambda m, h=sha256(f): m.group(1) + h, html)
        page.write_text(html, encoding="utf-8")
        print("updated checksums in", page)
    return 0


if __name__ == "__main__":
    sys.exit(main())
