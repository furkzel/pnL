# Aktarım Paneli v1.0 — Ürün, Tasarım ve Motor İncelemesi

Kapsam: `src/` (14 197 satır VBA + araçlar), yayındaki `AktarimPaneli_v1.0.xlsm` / `_makrosuz.xlsx`
(openpyxl ile yapı, LibreOffice ile baskı çıktısı), `docs/*.png` ekran görüntüleri.
Not: E2E testleri Windows + Excel gerektirdiği için bu ortamda çalıştırılamadı; motor bulguları kaynak okumasına dayanır.

---

## 0. Tek paragrafta teşhis

Motor, sektör standardının üstünde: sütun sütun okuma, biçim koruması, 4'lü mutabakat, parmak izli iki aşamalı geri alma
ve K1–K8 etiketli denetim izi gerçek bir aktüerya disiplini taşıyor. Zayıf halka **sunum katmanı**: renkler anlam taşımıyor
(yeşil 6 farklı şey demek), 18 eşit ağırlıklı buton "sıradaki adım" sorusunu cevaplamıyor, Denetim sayfası sayıları metin olarak
tutuyor, son kullanıcı dosyasında 12 580 satırlık kaynak kod görünür duruyor. Kontrol tarafında tek tutar sütunu, dosya
adına dayalı mükerrer tespiti ve zayıf parmak izi, "kusursuz" iddiasının önündeki üç somut engel.

---

## 1. Son kullanıcı + tasarımcı eleştirisi

### 1.1 Renk ve görsel dil
- **Palet Tailwind "slate" tabanlı ve tutarlı başlıyor**, ama `M01_Ayarlar.Palet` dışında **64 sabit `RGB()`** var
  (M09, M10, M08, M07, M04). Tek kaynak ilkesi kırılmış; Denetim ve Kılavuz kendi tonlarını kullanıyor.
- **Yeşil 6 anlama geliyor:** zorunlu sütun, "Evet", SABİT DEĞERLER başlığı, FİLTRE zemini, ✓ OK, AKTAR butonu.
  **Turuncu 3 anlama:** öneri, DİKKAT, HAZIRLIK pili. Kullanıcı rengi okuyarak karar veremiyor.
- **Panelde 8+ pastel zemin** (krem, nane, pembe, açık yeşil, lavanta, sarı, gri, beyaz) aynı ekranda → "gökkuşağı" etkisi;
  minimalizmin tam tersi. Zemin rengi yalnızca *durum* için kullanılmalı, *bölge* için değil.
- **Çelişki:** Panel 1 `DURUM: HAZIR ...` satırı pembe (hata tonu) zemin üzerinde. Metin "hazır", renk "hata" diyor.
- **Dahil olmayan panel (Hayır)** kırmızı ✕ K1 ENGEL gösteriyor; aktarıma girmeyen panel alarm vermemeli (soluk/gri olmalı).
- Tipografi 8–9 pt Segoe UI: 1080p / %100 yakınlaştırmada okunabilirlik sınırında; Mac'te Segoe UI yok (yedek fonta düşüyor).
  Kayıt sayfaları (Rapor, Gecmis, Profiller, Setler) ise tema fontu "Gövde Yazı Tipi" kullanıyor → iki ayrı yazı ailesi.

### 1.2 Buton şeması (şerit)
- **18 buton, 4 dolgu rengi, hepsi aynı görsel ağırlıkta.** 1-2-3 mavi, 4 yeşil, Geri Al kırmızı, gerisi koyu gri.
  Mavi aynı zamanda "AKTİF: PANEL" göstergesi ve gezinme okları → birincil renk aşırı yüklü.
- **Yıkıcı eylem (Geri Al) birincil eylemle (AKTAR) aynı boyut ve yoğunlukta, hemen altında.** Yanlış tık riski; sektör pratiği:
  tehlikeli eylem = kenarlıklı/hayalet stil, ayrı grupta.
- AKTAR, HAZIRLIK kırmızıyken bile yeşil ve tıklanabilir görünüyor; durumla bağlantısı yok.
- "Set Sil" kardeşlerinden küçük punto; boyut ızgarası tutarsız.
- 1→2→3→4 numaralandırma iyi fikir, ama **"şu an neredeyim / sıradaki ne"** görselleşmiyor (tamamlanan adımda ✓ yok).
- HAZIRLIK pili turuncu zemin üzerinde 8 pt beyaz yazı: kontrast düşük (≈3:1, WCAG AA altı).

### 1.3 Panel ekranı (bilgi mimarisi)
- 500 boş satırın tamamına kenarlık çizili → veri yokken bile ızgara gürültüsü; dolu satır/boş satır ayrımı kayboluyor.
- Kesilen metinler: "4 AKTAR: aktarım no 1 · FARK 0,00 (20.1…", "Canlı: K1 K4 K5 · tam kontrol: 3 Ön Kontrol (18:44)".
- Ayar kartı, SABİT DEĞERLER ve KONTROL NOKTALARI üç ayrı başlık stiliyle çiziliyor (beyaz/nane/koyu lacivert).
- Karne (K1–K8) panelin en güçlü öğesi ama metin satırı olarak duruyor; tek bakışta skor/özet (ör. "6/8 ✓ · 2 !") yok.
- Hedef haritası çok değerli, ancak zorunlu sütun **yeşil dolgu** ile işaretleniyor; ana dosya renklerinin kopyası olduğu için
  panelin kendi semantiğiyle çakışıyor.

### 1.4 Denetim sayfası
- **KPI değerleri metin** (`"'" & deger`, `M10_Denetim.bas:602`) → toplanamaz, filtrelenemez, biçimlenemez.
- **Sayı biçimi en-US** (`448,750.00`) iken arayüz Türkçe ve Rapor "FARK 0,00" yazıyor; ondalık ayıracı karışık.
- Trend yok, grafik yok, dönem kırılımı yok; K1–K8 tablosunda oran (DİKKAT/çalıştırma) yok, yalnızca adet.
- Yayın dosyasında **derleme izi kalmış**: "20.10.2026 18:41 · AKTUERYA".
- Baskı ayarı yok; "tek sayfa yönetici özeti" iddiası çıktıda karşılanmıyor.

### 1.5 Kılavuz, Geliştirici ve kayıt sayfaları
- **Geliştirici sayfası son kullanıcı dosyasında görünür ve 12 580 satır kaynak kod içeriyor.** Çalışma kitabının PDF çıktısı
  **271 sayfa**; kurumsal dağıtımda hem IP hem algı sorunu.
- Kılavuz 17 akordiyon başlığı, koyu lacivert bloklar: tam ekran ağır. "1 dakikada başla" özeti ve arama yok.
- Sayfa adları karışık: `Kılavuz` (Türkçe harf) / `Gecmis`, `Gelistirici` (ASCII). Rapor başlıkları `ACIKLAMA`, `TARIH`.
- Hiçbir sayfada koruma yok: "elle düzenlemeyin" denen Gecmis/Denetim düzenlenebilir.
- Rapor/Gecmis'te otomatik filtre, dondurulmuş başlık, tablo stili, sütun genişliği standardı yok.
- Setler'de `SABIT SUTUN (v31)` gibi eski sürüm sütunları kullanıcıya görünüyor.

### 1.6 Etkileşim
- **113 MsgBox + 5 InputBox.** Modal diyalog yağmuru; yanıtlar tekrar okunamaz, kopyalanamaz.
- MsgBox metinleri ASCII ("Hicbir sey yazilmadi"): Türk kurumsal kullanıcıya "çeviri yazılım" izlenimi.
- Çift tıklama kısayolları güçlü ama keşfedilemez (yalnızca not/Kılavuz).

---

## 2. Yapılacaklar Listesi

Öncelik: **P0** yayın kalitesini doğrudan bozan · **P1** ürünü "kurumsal" seviyeye taşıyan · **P2** fark yaratan.

### 2.1 Tasarım sistemi (TS)
| ID | Öncelik | İş |
|---|---|---|
| TS-01 | P0 | Tüm `RGB()` sabitlerini `Palet()`'e taşı; `M01` dışında renk literali kalmasın (build kapısına grep kuralı ekle). |
| TS-02 | P0 | Semantik palet: 1 marka + 3 durum + 1 öneri + nötr gri skalası (aşağıdaki tokenlar). Zemin rengi yalnızca durum için. |
| TS-03 | P0 | Yeşilin tekil anlamı = OK. Zorunlu sütun → `●` glif + sol kenar çizgisi; Evet/Hayır → anahtar glifi (`● Evet` / `○ Hayır`); filtre/işlem bölge zeminleri → beyaz. |
| TS-04 | P1 | Öneri rengi turuncu yerine **mor/indigo** (DİKKAT amber'ı ile karışmasın). |
| TS-05 | P1 | Tipografi ölçeği: başlık 12 / bölüm 10 kalın / gövde 9.5–10 / dipnot 8.5; tüm sayfalarda tek aile (Segoe UI, yedek Calibri/Aptos). |
| TS-06 | P1 | Boş tablo satırlarında kenarlık yok; yalnız dolu satırlara ince alt çizgi (zebra yok). |
| TS-07 | P1 | Kontrast: tüm metin/zemin çiftleri ≥ 4.5:1 (build'de otomatik kontrol). |

Önerilen tokenlar (açık tema): Zemin `#FFFFFF` · Yüzey `#F8FAFC` · Çizgi `#E2E8F0` · Metin `#0F172A` · İkincil metin `#475569` ·
Marka `#1D4ED8` · OK `#15803D` / zemin `#DCFCE7` · DİKKAT `#B45309` / zemin `#FEF3C7` · ENGEL `#B91C1C` / zemin `#FEE2E2` ·
Öneri `#6D28D9` / zemin `#EDE9FE` · Pasif `#94A3B8`.

### 2.2 Buton şeması ve şerit (BT)
| ID | Öncelik | İş |
|---|---|---|
| BT-01 | P0 | 3 seviye: **Birincil** (dolgu, yalnızca *sıradaki adım*), **İkincil** (kenarlık), **Üçüncül** (yalnız metin / "Diğer ▾" menüsü). |
| BT-02 | P0 | Dinamik adım göstergesi: tamamlanan adım `✓`, sıradaki adım dolgu marka rengi, sonraki adımlar pasif. |
| BT-03 | P0 | AKTAR yalnızca HAZIRLIK yeşil/amber iken aktif görünür; ENGEL'de pasif gri + ipucu. |
| BT-04 | P0 | Geri Al: kırmızı kenarlıklı hayalet buton, AKTAR'dan ayrı grupta ve sağda; iki adımlı onay korunur. |
| BT-05 | P1 | Görünür buton sayısı ≤ 8. Set Yükle/Kaydet/Sil, Profil Kaydet, Tanı, Seçili Sütun Yaması → "Diğer ▾" şekil menüsü. |
| BT-06 | P1 | Tek boy ızgarası: birincil 26 px, ikincil 22 px yükseklik; eşit iç boşluk; aynı punto. |
| BT-07 | P1 | HAZIRLIK pili: durum rengi zeminde koyu metin (ör. amber zemin + `#78350F`), 9 pt kalın; "2/3 hazır · 1 dikkat". |

### 2.3 Panel ekranı (PN)
| ID | Öncelik | İş |
|---|---|---|
| PN-01 | P0 | DURUM satırı rengi metinle aynı semantikte (HAZIR = OK, EKSİK = DİKKAT, ENGEL = kırmızı). |
| PN-02 | P0 | "Hayır" paneller soluk (gri başlık, karne `–`), HAZIRLIK hesabına ve alarma girmez. |
| PN-03 | P1 | Karne başlığında özet rozet: `6 ✓ · 2 ! · 0 ✕`. |
| PN-04 | P1 | Kesilen metinler: karne notu iki satıra veya hücre notuna; tarih `gg.aa ss:dd`. |
| PN-05 | P1 | Kart başlıklarını tek stilde birleştir (nötr yüzey + kalın koyu metin + ince alt çizgi). |
| PN-06 | P2 | Çift tıklanabilir hücrelerde imleç ipucu: başlıkta `⇆` / `✓` glifi. |

### 2.4 Denetim (DN)
| ID | Öncelik | İş |
|---|---|---|
| DN-01 | P0 | KPI'ları **sayı** olarak yaz, biçimi `NumberFormatLocal` ile ver (`#.##0,00`); metin önekini kaldır. |
| DN-02 | P0 | Yayın paketinde Denetim, Rapor, Gecmis boş ve "henüz aktarım yok" durumunda olsun; derleme kullanıcı/tarih izi temizlensin. |
| DN-03 | P1 | 12 aylık trend: aktarılan satır ve tutar (yerel ChartObject, çizgi + sütun), mutabakat oranı. |
| DN-04 | P1 | K tablosunda oran sütunu (DİKKAT / Ön Kontrol sayısı) ve mini çubuk (veri çubuğu koşullu biçimi). |
| DN-05 | P1 | Baskı: A4 yatay, 1 sayfaya sığdır, üst bilgi/alt bilgi (dosya, tarih, sayfa). |
| DN-06 | P2 | "Veri kalitesi skoru" (0–100) KPI kartı (bkz. 3.5). |

### 2.5 Sayfalar, koruma, dağıtım (SY)
| ID | Öncelik | İş |
|---|---|---|
| SY-01 | P0 | Geliştirici sayfası `xlSheetVeryHidden` veya yayın paketinden tamamen çıkar (kaynak zaten zip'te). |
| SY-02 | P0 | Koruma: `Protect UserInterfaceOnly:=True` (Panel etiketleri, Rapor, Gecmis, Denetim, Profiller, Setler); Workbook_Open'da yeniden uygula. |
| SY-03 | P1 | Sayfa sırası ve sekme renkleri: Panel · Analiz · Denetim │ Rapor · Gecmis │ Profiller · Setler │ Kılavuz. |
| SY-04 | P1 | Rapor/Gecmis: dondurulmuş başlık, otomatik filtre, sabit genişlik, Türkçe başlık etiketleri (iç sütun adları ASCII kalabilir). |
| SY-05 | P1 | Setler'den v31 sütunlarını gizle; geçiş kodu okumaya devam etsin. |
| SY-06 | P1 | Kılavuz: en üstte "1 dakikada başla" kartı (4 adım, 4 satır), akordiyon başlıkları açık zemin + koyu metin. |
| SY-07 | P2 | Tüm sayfalarda baskı alanı; çalışma kitabının toplam çıktısı < 15 sayfa. |

### 2.6 Etkileşim (UX)
| ID | Öncelik | İş |
|---|---|---|
| UX-01 | P1 | Bilgilendirme MsgBox'larını panel içi **bildirim şeridine** taşı (son 5 mesaj, zaman damgalı, Rapor'a bağlantılı). Yalnız karar gerektirenler modal kalsın. Hedef: 113 → ≤ 30. |
| UX-02 | P1 | Diyaloglarda Türkçe karakter: Türkçe kod sayfasında `TR()` ile, değilse ASCII'ye düş (`Application.International` kontrolü). |
| UX-03 | P2 | Klavye kısayolları (`Application.OnKey`): Ctrl+Shift+1..4 adımlar. |

### 2.7 Altyapı (AL)
| ID | Öncelik | İş |
|---|---|---|
| AL-01 | P0 | `build.py --check` pywin32'siz çalışsın (xlauto'yu tembel içe aktar); renk-literali ve ASCII kapısı Linux'ta da koşsun. |
| AL-02 | P1 | GitHub Actions: ASCII kapısı, renk kapısı, VBA statik kontroller (Option Explicit, EkranKilitle/EkranAc eşleşmesi). |
| AL-03 | P1 | Yenile performans gerilemesi (7.5 s → 8.2 s) kök nedeni; hedef ≤ 6 s (3 panel, 200k × 30). |
| AL-04 | P1 | Her makro süresini Rapor'a yaz (`SURE_MS`), perf.py ile karşılaştırmalı tablo. |

---

## 3. Motor değerlendirmesi (aktüerya veri bilimi gözüyle)

### 3.1 Doğru yapılanlar
- Sütun sütun okuma, `Value2` geri dönüşü, `Transpose` yasağı: 870k satırda bellek ve tip güvenliği doğru.
- Kuru çalışma (Ön Kontrol) ile gerçek aktarımın **aynı plan kodunu** kullanması: kontrol/uygulama ayrışması yok.
- Onaydan sonra satır ekleme; iptalde iz bırakmama.
- Biçim koruması (baştaki sıfır, 16+ hane, `=` metni) ve harfle başlayan metin için hızlı yol.
- 4'lü mutabakat (kaynak = yazılan, Δhedef = yazılan − ezilen) kavramsal olarak doğru kurgulanmış.
- Her bulgu K-etiketli, Rapor'a kısaltılmadan yazılıyor: denetim izi ASOP 23 ruhuna uygun.

### 3.2 Bulgular ve riskler
| # | Bulgu | Yer | Etki | Öneri |
|---|---|---|---|---|
| M1 | Tek tutar sütunu mutabakatı. Ödenen, muallak, rücu, masraf ayrı ayrı mutabık edilmiyor. | `Yaz` | Yüksek | Panel başına N tutar sütunu (en az 4), her biri 4'lü mutabakat. |
| M2 | **Satır sayısı mutabakatı yok** (kaynak filtreli = yazılan = Δana dosya satır). | `Yaz` | Yüksek | Adet mutabakatı K7'ye eklenmeli. |
| M3 | `SayiDegeri` metin sayıyı `IsNumeric`/`CDbl` ile çeviriyor: yerel ayara bağlı. TR'de `"1.5"` → 15, `"1,234.56"` belirsiz; `"(12)"`, `"1e5"` kabul. | `M02:413` | Yüksek | Açık ayrıştırıcı: ondalık/binlik ayıracı veriden tespit, belirsizse DİKKAT. |
| M4 | Hedef toplamı `AGGREGATE(9,6)` hataları yok sayıyor: ana dosyadaki `#YOK` tutarlar mutabakatı gizler. | `M06:2155` | Orta | Hata hücre sayısını ayrıca say, >0 ise DİKKAT. |
| M5 | Tolerans mutlak 0,005 ve `Double` birikim. | `M06:2050` | Düşük–Orta | `Currency` ile topla (4 ondalık, kesin), kuruşa yuvarla; tolerans = max(0,005; 1e-9 × |toplam|) raporlansın. |
| M6 | Parmak izi: adet + toplam + ilk/son değer. Ortadaki metin hücresi veya toplamı koruyan değişiklik yakalanmıyor. | `M07:323` | Orta | Tüm aralık üzerinde FNV-1a / CRC32 karması (VBA'da `Currency`/`Double` modüler aritmetik). |
| M7 | K2 mükerrer tespiti dosya/sayfa **adına** dayalı; dosya yeniden adlandırılınca atlanır. | `M07:743` | Yüksek | İçerik parmak izi (satır sayısı + tutar toplamı + anahtar sütun karması) + anahtar bazlı ana dosya taraması. |
| M8 | `TarihSayisi` 1..2 958 465 arası her sayıyı tarih sayıyor (tutar 1500 → 1904 tarihi geçer). | `M02:442` | Orta | Makul aralık: 01.01.1990 – rapor dönemi sonu + 1 yıl; dışı "aralık dışı tarih". |
| M9 | Tip örneklemi ilk 300 satır; filtre adayları ilk 10 000 satır. Dosya sonu anomalileri görünmez. | `M01:193-195` | Orta | Tabakalı örnek: baş 100 + orta 100 + son 100 + rastgele 100. |
| M10 | Zorunlu sütun tespiti renk eşiğine dayalı (G > R+12, G > B+12). | `M02:1311` | Orta | Açık meta veri: Profilde `ZORUNLU` bayrağı; renk yalnız varsayılan öneri. |
| M11 | Ön Kontrol ile Aktar arasında kaynak dosya değişirse fark edilmiyor (TOCTOU). | `AktarIc` | Orta | Ön Kontrol'de kaynak karması sakla, Aktar'da karşılaştır (K14). |
| M12 | Eşzamanlılık kilidi yok; iki kullanıcı aynı ana dosyaya yazabilir. | genel | Orta | Ana dosya yanında `.lock` (kullanıcı, zaman), bayatlık süresi. |
| M13 | `AktarCalistir` `On Error Resume Next` ile sarılı; `AktarIc` dışı hatalar sessiz. | `M06:467` | Düşük | Hata numarası + prosedür adını Rapor'a yaz. |
| M14 | Yedek sayfası sınırsız büyüyor (geri alınmış/eski partiler). | `M07` | Düşük | Saklama politikası (son N parti) + Denetim'de dosya boyutu metriği. |
| M15 | Kaynak soy bilgisi yok: dosya yolu, boyut, değiştirilme zamanı, karma Rapor'a yazılmıyor. | `RaporYaz` | Orta | Rapor'a `KAYNAK_KARMA`, `KAYNAK_BOYUT`, `KAYNAK_DEGISTIRILME` sütunları. |

### 3.3 Mevcut K1–K8: değerlendirme
| K | Güçlü yön | Eksik | Yükseltme |
|---|---|---|---|
| K1 | Başlık yeniden doğrulama, çift hedef ENGEL | — | Şema kayması K12'ye ayrılsın. |
| K2 | Kullanıcıya varsayılan "Hayır" ile soruyor | Ad tabanlı | İçerik + anahtar tabanlı (M7). |
| K3 | Çoğunluk son satırı + alt blok ayracı + önizleme | Genel toplam satırı boşluksuz gelirse | "Toplam satırı" sezgisi: tutarı üstündeki satırların toplamına eşit satır; `TOPLAM` anahtar kelimesi. |
| K4 | Eşleşmemiş + boş yeşil hücre | Renk bağımlı, eşik yok | Meta veri + sütun bazlı tamlık eşiği (ör. ≥ %99,5). |
| K5 | Tip çatışması, sabit değer tipi | İlk 300 satır | Tabakalı örnek (M9). |
| K6 | Kapsamlı | — | Ondalık hane > 2 olan tutarlar bilgi notu. |
| K7 | 4'lü mutabakat | Tek sütun, adet yok, yerel ayar riski, hata gizleme | M1–M5. Eksi tutar: sütun bazlı işaret kuralı (rücu eksi olabilir). |
| K8 | İhbar < hasar, boş, bozuk, metin tarih | Dönem ve makul aralık yok | K10'a genişlet. |

### 3.4 Yeni kontrol noktaları (K9–K14)
| K | Ad | Ne kontrol edilir | Seviye |
|---|---|---|---|
| **K9** | Anahtar tekilliği | Hasar/dosya no (+hareket tarihi +tutar) aylık dosyada tekil mi; aynı anahtar ana dosyanın son 12 ayında var mı | DİKKAT / ENGEL (birebir aynı satır) |
| **K10** | Dönem tutarlılığı | Hasar/ihbar/ödeme tarihi ≤ rapor dönemi sonu; gelecek tarih; makul aralık dışı; ödeme < ihbar | DİKKAT |
| **K11** | Makullük ve aykırı değer | Satır sayısı ve toplam tutar, son 12 ayın ortalamasına göre z-skor (|z| > 3 DİKKAT); büyük hasar eşiği; ortalama hasar Δ% | DİKKAT |
| **K12** | Şema ve kategori kayması | Profile göre yeni/kaybolan sütun; branş/şirket gibi kategorik alanda ana dosyada hiç görülmemiş yeni kod veya bu ay kaybolan kod | DİKKAT |
| **K13** | Katmanlı mutabakat | Kırılım bazında (branş × şirket) adet ve tutar: kaynak = yazılan; birden çok tutar sütunu | ENGEL (fark ≠ 0) |
| **K14** | Kaynak bütünlüğü ve soy | Kaynak karması Ön Kontrol ↔ Aktar arasında değişmedi; dosya boyutu/değiştirilme zamanı Rapor'a | ENGEL (değiştiyse) |

Uygulama notu: README kuralı #8 (`KEkle p, n, ...`) K9–K14'e doğrudan genişler; `K_SAYI` 8 → 14, karne satırları
`KARNE_ILK..` yeniden yerleşir → `DUZEN_DEGERI` artırılmalı ve `BlokOku` eski yerleşimi öğrenmeli.

### 3.5 Raporlanması gereken teknik metrikler
| Alan | Metrik | Tanım | Eşik / hedef | Nerede |
|---|---|---|---|---|
| Tamlık | Dolu oranı | sütun bazında dolu / toplam | zorunlu ≥ %99,5 | Analiz A1, K4 |
| Geçerlilik | Geçerli tip oranı | tipine uyan / dolu | ≥ %99,9 | A1, K5 |
| Tekillik | Anahtar tekillik oranı | tekil anahtar / satır | = %100 | K9 |
| Mutabakat | Tutar farkı | Σkaynak − Σyazılan (her tutar sütunu) | 0,00 | K7, K13, Denetim |
| Mutabakat | Adet farkı | n kaynak(filtreli) − n yazılan − Δsatır | 0 | K7 |
| Mutabakat | Başarı oranı | TUTUYOR / aktarım | %100 | Denetim |
| Makullük | Satır z-skoru | (n_ay − μ12) / σ12 | |z| ≤ 3 | K11, A2 |
| Makullük | Tutar z-skoru | (Σ_ay − μ12) / σ12 | |z| ≤ 3 | K11, A2 |
| Dağılım | Tutar çeyrekleri | min, P50, P95, P99, maks | — | A1, A6 |
| Gecikme | İhbar gecikmesi | ihbar − hasar (gün): ort., P50, P90 | P90 trendi | A5 |
| Kayma | Yeni / kaybolan kod | kategori kümesi farkı | 0 (bilgi) | K12 |
| Süreç | İlk geçiş verimi | bulgusuz Ön Kontrol / Ön Kontrol | ↑ | Denetim |
| Süreç | Geçersiz kılma oranı | "Yine de devam" / DİKKAT | ↓ (denetim için kritik) | Denetim |
| Süreç | Geri alma oranı | geri alınan / aktarım | ↓ | Denetim |
| Süreç | Eşleştirme kapsamı | eşleşen / kaynak sütun; oto / manuel; öneri kabul oranı | — | Panel, Denetim |
| Performans | Adım süresi | Yenile, Ön Kontrol, Aktar (ms), satır/sn | 200k×30 Aktar ≤ 15 s | Rapor, perf |
| Sağlık | Dosya boyutu, Yedek boyutu | MB | uyarı > 50 MB | Denetim |
| Bileşik | Veri kalitesi skoru | 0–100: tamlık %30, geçerlilik %25, tekillik %20, makullük %15, mutabakat %10 | ≥ 95 | Denetim, Panel |

---

## 4. "Analiz" sayfası: fizibilite ve durum

**Amaç:** Sisteme aktarılacak (veya aktarılmış) veriyi, kullanıcının seçtiği boyut, ölçü ve filtreyle okuyup tablo, grafik
ve kısa yorum olarak raporlamak.

### 4.1 Mevcut altyapının hazırlık ölçümü
| Yetenek | Mevcut parça | Hazırlık |
|---|---|---|
| Büyük sayfa okuma | `SutunOku`, `VeriOku` | %90 |
| Tip çıkarımı | `VeriTipiSay`, `TipKarari` | %75 |
| Filtre motoru (aynı sözdizimi) | `FiltreAyristir`, `SatirSec` (Private → Public yapılmalı) | %85 |
| Gruplama / toplama | `Sozluk.Artir` (yalnız sayaç) | %25 |
| Çıktı biçimleme | `KpiKart`, `BaslikSatiri` (M10) | %50 |
| İstek formu (menüler) | `MenuKur`, `ListeYaz` | %60 |
| Şablon saklama | Setler/Profiller deseni | %50 |
| Grafik | yok (kodda hiç `ChartObjects` yok) | %0 |
| Test | `xlauto`, `test_e2e` deseni | %70 |
| **Ağırlıklı toplam** | | **≈ %55** |

**Karar: Yapılabilir, orta risk.** Mimari hazır; üç eksik parça: group-by motoru, grafik katmanı, istek formu.
Tahmini kapsam: yeni `M11_Analiz.bas` ≈ 1 500–2 000 satır, 25–35 yeni E2E kontrolü.

### 4.2 Teknoloji seçimi
| Seçenek | Artı | Eksi | Karar |
|---|---|---|---|
| VBA yerel (Sozluk + diziler) + yerel ChartObject | Citrix/Mac uyumlu, kural setine uygun, tam kontrol | Kod hacmi | **Seçilen** |
| PivotTable (VBA ile oluşturma) | Hızlı, tanıdık | Pivot önbelleği dosyayı şişirir, dış bağlantı kırılganlığı | Opsiyonel "Pivot'a gönder" düğmesi |
| Power Query | Güçlü | Kilitli masaüstü / Mac kısıtı, M kodu bakımı | Hayır |
| `GROUPBY` / `PIVOTBY` | Formülle canlı | Yalnız yeni M365 | Hayır |

### 4.3 Performans ve bellek bütçesi
- Yalnız istenen sütunlar okunur (870k × 3 sütun ≈ 45 MB Variant; 30 sütunun tamamı ≈ 420 MB, 32-bit Excel'de riskli).
- `Sozluk` Collection tabanlı; `Exists` hata yakalamalı → çok yüksek kardinalitede yavaş. Hedef: 870k satır, ≤ 1 000 grup,
  tek boyut ≤ 6 s. Kardinalite > 5 000 ise sıralama tabanlı gruplamaya geç.

### 4.4 Sayfa tasarımı
- Üst: **istek kartı** (Kaynak: Aylık dosya / Ana dosya / Son parti · Satır boyutu · Sütun boyutu (opsiyonel) · Ölçü:
  Adet/Toplam/Ortalama/Min/Maks/Tekil/P50/P95 · Filtre (panel sözdizimi) · Dönem) + `Çalıştır` + `Şablon kaydet`.
- Orta: sonuç tablosu (≤ 50 satır + "Diğer"), yanında tek grafik (çubuk/çizgi, marka rengi + nötr).
- Alt: otomatik yorum, 3 madde (ör. "Bu ay toplam tutar 12 ay ortalamasının +2,4σ üstünde; artışın %71'i KASKO").
- Hazır şablonlar:
  - A1 Veri profili
  - A2 Ay-ay karşılaştırma (z-skor)
  - A3 Branş × şirket kırılımı
  - A4 Hasar gelişim üçgeni (kaza dönemi × gelişim, kümülatif ödenen, link oranları)
  - A5 İhbar gecikmesi dağılımı
  - A6 Büyük hasar Top-N
  - A7 Aktarım etkisi (önce/sonra)
  - A8 Kontrol noktası trendi

---

## 5. Hedef mimari (özet)

```
Şerit: [1 Yenile ✓] [2 Eşleştir ✓] [3 Ön Kontrol ●] [4 AKTAR (pasif)]  ·  Diğer ▾  ·  ◀ Panel 1/3 ▶  ·  HAZIRLIK 2/3 · 1 !  ·  ⟲ Geri Al
Sayfalar: Panel · Analiz · Denetim │ Rapor · Gecmis │ Profiller · Setler │ Kılavuz   (Yedek, Listeler, Geliştirici: very hidden)
Kontroller: K1–K14, her biri Rapor'a etiketli, karne + Denetim + Analiz A8'e akar
Renk: nötr + 1 marka + OK/DİKKAT/ENGEL + öneri (mor). Yeşil = yalnızca OK.
```
