"""
release.py - build the three download packages for a release.

  python tools/release.py [--version v1.1]
  python tools/release.py --version v1.1 --macros-only [--out docs/v1.1]
      Portable (no Excel): only the macro zip + SHA256SUMS.txt, with
      reproducible timestamps, for a pre-release that has no built workbook.

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
ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
ZIP_TIME = (2026, 11, 17, 0, 0, 0)   # fixed member timestamps: identical bytes on every build

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
       M01_Ayarlar.bas ... M16_Bellek.bas  (15 dosya)

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


KURULUM_ONSURUM = """AKTARIM PANELI {v} ON SURUM - MAKROLARI MEVCUT KITABA KURMA
===========================================================

DURUM: {v} kaynak kodu tamamdir ama HENUZ EXCEL'DE DERLENMEDI ve uctan
uca testleri calistirilmadi. Bu paket gozden gecirme ve deneme icindir;
aylik kapanista kararli surum v1.0'dir. Denemeyi bir KOPYA uzerinde yapin.

1) v1.0 calisma kitabinizin (AktarimPaneli_v1.0.xlsm) bir KOPYASINI alin
   ve kopyayi acin. (Yeni baslayanlar: sayfadaki v1.0 .xlsm'i indirin.)

2) Alt+F11 ile VBA editorunu acin. Soldaki agacta Sozluk ve M01_... ile
   baslayan TUM modulleri sag tik > Remove (Export: Hayir) ile kaldirin.

3) File > Import File... ile bu klasordeki dosyalari ice aktarin:
       Sozluk.cls
       M01_Ayarlar.bas ... M16_Bellek.bas  (15 dosya)

4) "Panel" sayfasinin kod sayfasini (parantez icinde "Panel" yazan satir)
   acin, icini silip Panel_sayfa_kodu.txt dosyasinin TAMAMINI yapistirin.

5) Ayni seyi "ThisWorkbook" ("BuCalismaKitabi") icin ThisWorkbook_kodu.txt
   ile yapin.

6) Debug > Compile VBAProject. HATA CIKARSA durun ve hatanin goruldugu
   satiri bildirin (bu surum Excel'de henuz derlenmedi).

7) Alt+F8 > SifirdanKur. Panel v1.1 duzenine tasinir (ayarlar, eslesmeler,
   sabit degerler, Rapor, Gecmis, Profiller, Setler korunur); Ayarlar ve
   Analiz sayfalari kurulur. Kaydedin.

GUVENLIK
- Dosyalarin butunlugunu SHA256SUMS.txt ile dogrulayin
  (PowerShell: Get-FileHash dosya -Algorithm SHA256).
- Kodun tamami ACIK metindir; ice aktarmadan once okuyabilirsiniz.
- Makrolar disari baglanmaz. Sildigi tek dosya, aktarim sirasinda ana
  dosyanin yanina kendisinin yazdigi .aplock kilit dosyasidir.
"""


def macros_zip(zp: Path, v: str, kurulum: str) -> None:
    def add(z, name, data):
        info = zipfile.ZipInfo(name, ZIP_TIME)
        info.compress_type = zipfile.ZIP_DEFLATED
        info.external_attr = 0o644 << 16
        z.writestr(info, data)
    with zipfile.ZipFile(zp, "w") as z:
        for f in [SRC / "Sozluk.cls"] + sorted(SRC.glob("M*.bas")):
            add(z, f"makrolar/{f.name}", crlf(f.read_text(encoding="ascii")))
        add(z, "makrolar/Panel_sayfa_kodu.txt", crlf((SRC / "Panel.cls").read_text(encoding="ascii")))
        add(z, "makrolar/ThisWorkbook_kodu.txt", crlf((SRC / "ThisWorkbook.cls").read_text(encoding="ascii")))
        add(z, "makrolar/KURULUM.txt", crlf(kurulum.format(v=v)))


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
    ap.add_argument("--version", default="v1.1")
    ap.add_argument("--macros-only", action="store_true",
                    help="portable: only the macro zip and its checksum (pre-release, no Excel)")
    ap.add_argument("--out", default=None, help="output folder (default dist/release)")
    a = ap.parse_args(argv)
    v = a.version
    if a.macros_only:
        from source_checks import check_sources, check_contrast
        from vba_lint import lint
        problems = check_sources() + check_contrast() + lint()
        if problems:
            print("SOURCE CHECK FAILED:\n" + "\n".join(problems))
            return 1
        out = Path(a.out).resolve() if a.out else ROOT / "dist" / "release"
        out.mkdir(parents=True, exist_ok=True)
        zp = out / f"AktarimPaneli_{v}_makrolar.zip"
        macros_zip(zp, v, KURULUM_ONSURUM)
        (out / "SHA256SUMS.txt").write_text(f"{sha256(zp)}  {zp.name}\n", encoding="ascii")
        print("saved:", zp, sha256(zp))
        return 0
    import build  # noqa: F401  Windows/COM only from here on
    import xlauto
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
