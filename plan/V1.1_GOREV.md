# Agent Görev Tanımı — Aktarım Paneli v1.1 "Kurumsal Sürüm"

## Rolün
Kurumsal bir sigorta şirketinde 30 yıllık tecrübeye sahip **aktüerya geliştiricisi** ve **bilgi tasarımcısı** olarak
çalışıyorsun. Elindeki ürün `furkzel/pnL` deposundaki **Aktarım Paneli v1.0** (saf Excel VBA): aylık hasar dosyalarını
(ödenen, muallak) 870 000+ satırlık birikimli ana dosyalara sütun eşlemesiyle aktarır, K1–K8 kontrol noktalarıyla denetler,
mutabakat yapar, geri alınabilir şekilde kaydeder. Hedefin bu ürünü **v1.1**'e taşımak: eksiksiz, kompakt, minimalist,
yüksek işlevli, kullanıcı dostu ve "iş bitirici" bir Excel mimarisi.

## Önce oku (zorunlu)
1. `plan/INCELEME_v1.0.md`: tüm bulgular, Yapılacaklar Listesi (TS, BT, PN, DN, SY, UX, AL), motor bulguları M1–M15,
   yeni kontrol noktaları K9–K14, metrik kataloğu ve Analiz sayfası fizibilitesi. **Bu belge senin iş emrindir.**
2. `README.md` → "Rules for maintainers" (9 kural). Bu kurallar **bağlayıcıdır**.
3. `src/M01_Ayarlar.bas` (yerleşim ve palet), `src/M06_Aktarim.bas` (plan, kontrol ve yazma), `src/M07_Kayit.bas`
   (geri alma ve log), `src/M10_Denetim.bas` (karne ve Denetim).

## Değiştirilemez kısıtlar
- VBA kaynağı **%100 yazdırılabilir ASCII**; hücre metinleri `TR("#s#i#g")` ile.
- `Scripting.Dictionary`, `RegExp`, ActiveX, eklenti **yok**. `Sozluk`, diziler ve şekiller kullan. Yerel `ChartObjects`
  serbesttir (ActiveX değildir, Mac'te çalışır).
- Büyük sayfalar `SutunOku` ile sütun sütun okunur; `Application.Transpose` yasak; menülü sayfada `Cells.Clear` yasak.
- Her giriş makrosu `EkranKilitle`/`EkranAc` çiftiyle çalışır.
- Yerleşim yalnızca `M01_Ayarlar`'da yaşar. Blok yerleşimi değişirse `DUZEN_DEGERI`'ni artır ve
  `M08_Kurulum.BlokOku`'ya eski ofsetleri öğret. **v33, v34 ve v1.0 çalışma kitapları kayıpsız v1.1'e geçmeli**
  (ayarlar, eşlemeler, sabit değerler, filtreler, öneriler, Rapor, Gecmis, Profiller, Setler).
- Mevcut 154 E2E kontrolünün hiçbiri silinmez ya da gevşetilmez; değişen davranış için kontrol güncellenir ve gerekçesi yazılır.

## Tek seferde teslim edilecekler (Definition of Done)

### A. Tasarım sistemi ve görsel dil
- [ ] `Palet()` tek renk kaynağı; `M01` dışında `RGB(` literali kalmaz. `build.py --check` bunu denetler.
- [ ] Semantik palet (INCELEME §2.1 tokenları veya gerekçeli daha iyisi): nötr skala + 1 marka + OK/DİKKAT/ENGEL + öneri.
      **Yeşil yalnızca "OK"** demektir. Zorunlu sütun glif ve kenar çizgisiyle, öneri morla gösterilir.
- [ ] Tek yazı ailesi ve ölçek (başlık 12 / bölüm 10 kalın / gövde 9.5–10 / dipnot 8.5), tüm sayfalarda.
- [ ] Tüm metin/zemin çiftleri ≥ 4.5:1 kontrast (build'de otomatik kontrol).
- [ ] Boş tablo satırlarında kenarlık yok; panel gökkuşağı zeminleri kaldırılmış.

### B. Şerit ve buton şeması
- [ ] Üç seviyeli hiyerarşi: birincil (yalnız sıradaki adım), ikincil (kenarlık), üçüncül ("Diğer ▾" menüsü).
- [ ] Adım durumu: tamamlanan `✓`, sıradaki dolgu, sonraki pasif. AKTAR, ENGEL varken pasif görünür.
- [ ] Geri Al ayrı grupta, kırmızı kenarlıklı hayalet stil. Görünür buton sayısı ≤ 8.
- [ ] HAZIRLIK pili okunur ("2/3 hazır · 1 !"), kontrastı uygun.

### C. Panel, Denetim, sayfalar
- [ ] DURUM rengi metinle aynı anlamda; "Hayır" paneller soluk ve alarm dışı; karne başlığında özet rozet.
- [ ] Denetim: KPI'lar **sayı** (yerel biçim `#.##0,00`), 12 aylık trend grafiği, K oranları, A4 yatay tek sayfa baskı,
      veri kalitesi skoru kartı.
- [ ] Geliştirici, Yedek ve Listeler `xlSheetVeryHidden`. Sayfa sırası ve sekme renkleri INCELEME §5'e uygun.
- [ ] `Protect UserInterfaceOnly:=True` korumaları; `Workbook_Open`'da yeniden uygulanır.
- [ ] Rapor/Gecmis: dondurulmuş başlık, otomatik filtre, Türkçe görünen başlıklar.
- [ ] Yayın paketi temiz: boş log, derleme kullanıcı/tarih izi yok, çalışma kitabı çıktısı < 15 sayfa.
- [ ] Bilgilendirme amaçlı MsgBox'lar panel içi bildirim şeridine taşınmış (hedef ≤ 30 MsgBox); karar diyalogları modal kalır.

### D. Motor ve kontrol noktaları
- [ ] M1–M15'in tamamı giderilmiş, en azından: çoklu tutar sütunu (≥ 4) ve **adet mutabakatı**; yerel ayardan bağımsız
      sayı ayrıştırıcı; `Currency` birikimi; hedefteki hata hücrelerinin sayılması; tam aralık karması (FNV-1a/CRC32)
      ile parmak izi; içerik tabanlı K2; makul tarih aralığı; tabakalı tip örneklemi; kaynak soy bilgisi ve TOCTOU kontrolü;
      ana dosya kilidi; Yedek saklama politikası.
- [ ] **K9–K14** (INCELEME §3.4) `KEkle` altyapısıyla eklenmiş: karneye, Rapor'a, Denetim'e ve HAZIRLIK'a akar.
      Eşikler (z-skor, tamlık oranı, büyük hasar, makul tarih aralığı) panelde veya bir **Ayarlar kartında** düzenlenebilir;
      sihirli sayı yok.
- [ ] Metrik kataloğundaki (INCELEME §3.5) metrikler hesaplanır ve Rapor/Denetim'e yazılır; adım süreleri ms olarak loglanır.
- [ ] Performans: `perf.py` (200k × 30, 300k ana dosya, 3 panel) v1.0'dan kötü değil; Yenile ≤ 6 s hedefi.
      K9–K14'ün ek maliyeti raporlanır.

### E. Analiz sayfası (`M11_Analiz.bas` + `Analiz` sayfası)
- [ ] İstek kartı: Kaynak (Aylık / Ana dosya / Son parti) · Satır boyutu · Sütun boyutu (opsiyonel) · Ölçü
      (Adet, Toplam, Ortalama, Min, Maks, Tekil, P50, P95) · Filtre (panel filtre sözdizimi, aynı motor) · Dönem.
- [ ] Çıktı: sonuç tablosu (≤ 50 satır + "Diğer"), bir yerel grafik, 3 maddelik otomatik yorum.
- [ ] Hazır şablonlar A1–A8; kullanıcı şablonu kaydet/yükle (Setler deseni).
- [ ] Yalnız istenen sütunlar okunur; 870k satır, ≤ 1 000 grup, tek boyut ≤ 6 s; yüksek kardinalitede sıralama tabanlı gruplama.

### F. Altyapı, test, belge
- [ ] `build.py --check` pywin32'siz (Linux'ta) çalışır; GitHub Actions ile ASCII, renk ve statik kontroller koşar.
- [ ] `test_e2e.py`'ye her yeni K noktası, metrik, Analiz şablonu, koruma ve geçiş (v1.0 → v1.1) için kontroller eklenir.
- [ ] README, Kılavuz ("1 dakikada başla" kartı dahil), `docs/index.html` ve ekran görüntüleri v1.1'e güncellenir;
      sürüm tablosuna v1.1 satırı eklenir.

## Çalışma ortamı notu
`build.py`, `test_e2e.py`, `perf.py` ve `screenshot.py` **Windows + Excel + pywin32** gerektirir. Ortamında Excel yoksa:
statik kapıları (ASCII, renk, Option Explicit, EkranKilitle/EkranAc eşleşmesi) çalıştır, E2E kontrollerini yaz ama
"çalıştırılamadı" diye açıkça raporla, VBA'yı derleme hatası çıkarmayacak titizlikle (tip bildirimleri, satır devamı ≤ 24,
`IIf` + `Nothing` tuzağı) yaz. Asla çalıştırmadığın bir testi "geçti" diye raporlama.

## Çalışma biçimi
1. **Plan:** Kod yazmadan önce `plan/V1.1_PLAN.md` dosyasına kısa bir uygulama planı yaz: modül bazında değişiklik listesi,
   yeni yerleşim (satır/sütun ofsetleri), geçiş stratejisi, risk listesi.
2. **Sıra:** D (motor ve kontroller) → A/B (tasarım sistemi ve şerit) → C (sayfalar) → E (Analiz) → F (test ve belge).
   Motoru önce bitir: tasarım, motorun ürettiği durumları gösterir.
3. **Küçük, mantıklı commit'ler**, her biri derlenebilir durumda. Commit mesajları İngilizce, kısa.
4. **Sonuç raporu:** `plan/V1.1_SONUC.md`: yapılanlar (ID bazında ✓/✗), ölçümler (önce/sonra), çalıştırılamayan testler,
   bilinçli olarak ertelenenler ve nedenleri.

## Yaratıcı özgürlük alanın
Bu belge bir **taban**, tavan değil. Aşağıdakilerde kendi yargını kullanmaya ve beni şaşırtmaya yetkin var:
- INCELEME'deki herhangi bir öneriden **daha iyi** bir çözüm görürsen uygula; neden daha iyi olduğunu sonuç raporunda
  bir paragrafla savun.
- Renk tokenlarını, buton geometrisini, kart düzenini, ikon/glif dilini, Denetim ve Analiz kompozisyonunu bir tasarımcı
  olarak yeniden yorumlayabilirsin. Tek ölçüt: bir aktüer ekranı 3 saniyede okuyup doğru kararı verebilmeli.
- Listede olmayan aktüeryal kontroller (ör. hasar bazında kümülatif ödenen monotonluğu, muallak/ödenen oranı sıçraması,
  Benford sapması, kapanmış hasarın yeniden açılması) veya Analiz şablonları (ör. chain-ladder gelişim faktörleri,
  frekans/şiddet ayrışımı) öner ve **değerli buluyorsan uygula**. Maliyetini ölç ve raporla.
- Mimariyi sadeleştirmek için modül birleştirme/bölme yapabilirsin; kurallar ve geçiş uyumluluğu korunduğu sürece.

**Sınır:** Kısıtlar bölümü pazarlık konusu değildir. Kullanıcı verisini sessizce değiştiren, kontrolü atlatan veya
denetim izini zayıflatan hiçbir "kolaylık" kabul edilmez. Emin olmadığın bir aktüeryal yorumda varsayımını açıkça yaz.

## Kalite çıtası
Ortaya çıkan dosya açıldığında şu izlenimi vermeli: *"Bunu, bu işi 30 yıldır yapan ve her ay kapanışta bu ekranın
başında oturan biri tasarlamış."* Her piksel ve her kontrol bir soruya cevap vermeli; cevap vermeyen her şey kaldırılmalı.
