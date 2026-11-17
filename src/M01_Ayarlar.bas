Attribute VB_Name = "M01_Ayarlar"
Option Explicit
Option Private Module

' ============================================================
' AKTARIM PANELI  v1.0  -  M01_Ayarlar
'
' Tum SABITLER, ENUM'lar ve GLOBAL durum burada durur. Baska hicbir
' modulde "sihirli sayi" (Cells(i, 8), Cells(2, 341) ...) yazilmaz;
' yerlesim degisecekse SADECE bu dosya degisir.
'
' MODUL HARITASI (ayrinti: README.md ve Gelistirici sayfasi)
'   Sozluk          (cls) Scripting.Dictionary yerine yerel sozluk
'   M01_Ayarlar     sabitler, enum'lar, global durum, renk paleti
'   M02_Yardimci    metin/Turkce, dosya/sayfa bulma, okuma, veri tipi,
'                   ekran kilidi (EkranKilitle / EkranAc)
'   M03_Panel       panel geometrisi, buton seridi, blok cizimi,
'                   panel ekle/kaldir, gezinme
'   M04_Yenile      Yenile, menuler, DURUM, SABIT DEGERLER, hedef
'                   haritasi, VERI TIPI, hucre degisikligi yonlendirici
'   M05_Eslestirme  otomatik eslestirme, oneriler, profiller, setler
'   M06_Aktarim     aktarim motoru: on kontrol (ENGEL/DIKKAT), filtre,
'                   bicim korumasi, yazma, tutar mutabakati
'   M07_Kayit       Geri Al, Gecmis, Rapor, yedek, parmak izi
'   M08_Kurulum     SifirdanKur, PaneliYenidenKur, v33/v34 gecisi, Tani
'   M10_Denetim     Denetim sayfasi (yonetici ozeti), kontrol noktasi adlari
'   M09_Kilavuz     Kilavuz ve Gelistirici sayfalari
'   Panel (sayfa)   Worksheet_Change / SelectionChange olaylari
'   ThisWorkbook    Workbook_Open: Excel durumunu onarir
'
' VBA KURALLARI (bu projeyi degistirecek kisi icin):
'   - Kaynak kod %100 ASCII. Hucreye Turkce harf gerekiyorsa TR("#s#i")
'     kullanilir (M02_Yardimci). MsgBox metinleri ASCII kalir.
'   - Karsilastirmada UCase/LCase DEGIL, BH() / Normalize() kullan.
'   - Scripting.Dictionary, RegExp, ActiveX YOK. Sozluk sinifini kullan.
'   - Buyuk sayfa tek seferde okunmaz: SutunOku ile SUTUN SUTUN okunur.
'   - Diziye yazarken Application.Transpose KULLANILMAZ.
'   - Menulu (Data Validation) sayfada ws.Cells.Clear KULLANILMAZ;
'     AlanTemizle ile belirli araliklar temizlenir.
'   - Her giris makrosu EkranKilitle / EkranAc ciftiyle calisir ve
'     hata yolunda da EkranAc'a ugrar.
'   - Modul duzeyi bildirimler ilk prosedurden ONCE yazilir.
'   - Bir ifadede en fazla 24 satir devami ( _ ).
'   - MsgBox ~1024 karakterden sonrasini keser: Sigdir kullan.
'   - IIf iki dali da hesaplar; Nothing olabilecek nesnede kullanma.
' ============================================================

Public Const SURUM As String = "v1.1"
Public Const PANEL_YAZI_TIPI As String = "Calibri"
Public Const PANEL_YAZI_BOYUT As Double = 9.5

' Yazi olcegi (tek aile, dort basamak + gosterge rakami). Font.Size
' yalnizca bu sabitlerle verilir (build kapisi denetler).
Public Const YZ_BASLIK As Double = 12        ' sayfa / blok basligi
Public Const YZ_BOLUM As Double = 10         ' bolum basligi (kalin), birincil buton
Public Const YZ_GOVDE As Double = 9.5        ' govde, tablo, ikincil buton
Public Const YZ_DIPNOT As Double = 8.5       ' not, aciklama, alt yazi
Public Const YZ_GOSTERGE As Double = 16      ' KPI rakami (yalniz Denetim kartlari)

' Panel duzeni isareti (calisma kitabi adi). v33 panelinde yoktur.
Public Const DUZEN_ADI As String = "AP_DUZEN"
Public Const DUZEN_DEGERI As String = "v1.1"
Public Const DUZEN_V10 As String = "v1.0"      ' v1.0 (ve v35): ayni tablo/ayar ofsetleri

' ------------------------------------------------------------
'  SAYFA ADLARI  (Kilavuz adi Turkce harf icerdigi icin fonksiyon)
' ------------------------------------------------------------
Public Const SAYFA_PANEL As String = "Panel"
Public Const SAYFA_RAPOR As String = "Rapor"
Public Const SAYFA_PROFIL As String = "Profiller"
Public Const SAYFA_SET As String = "Setler"
Public Const SAYFA_GECMIS As String = "Gecmis"
Public Const SAYFA_GELISTIRICI As String = "Gelistirici"
Public Const SAYFA_YEDEK As String = "Yedek"
Public Const SAYFA_DENETIM As String = "Denetim"
Public Const SAYFA_LISTE_ONEK As String = "Listeler"
Public Const SAYFA_ANALIZ As String = "Analiz"
Public Const SAYFA_AYARLAR As String = "Ayarlar"
Public Const SAYFA_BELLEK As String = "Bellek"          ' gizli: durum ve bildirim bellegi
Public Const SAYFA_SABLON As String = "AnalizSablon"    ' gizli: kullanici analiz sablonlari

' ------------------------------------------------------------
'  PANEL YERLESIMI  (v1.0: blok basina 6 sutun)
'
'  Satir  1-5 : buton seridi (dondurulmus)
'  Satir  7   : blok basligi "PANEL k - dosya / sayfa"
'  Satir  8-21: ayarlar (etiket c..c+1 birlesik, deger c+2)
'  Satir  8-24: SABIT DEGERLER kutusu (sutun c+3, deger c+4)
'  Satir  8-17: KONTROL NOKTALARI karnesi K1..K8 (c+5)
'  Satir 26   : DURUM
'  Satir 28   : tablo basligi, 29-530: eslestirme tablosu
'
'  Tablo sutunlari (c = blogun ilk sutunu):
'   c+0 KAYNAK SUTUN | c+1 VERI TIPI | c+2 HEDEF SUTUN |
'   c+3 FILTRE | c+4 ISLEM (SEC / FORMUL) | c+5 HEDEF HARITASI (salt okunur)
'   v1.0: kullanicinin yazdigi sutunlar yan yana, salt okunur harita en sagda.
' ------------------------------------------------------------
Public Const MAX_BLOK As Long = 10
Public Const BLOK_GENISLIK As Long = 6

Public Const OFS_KAYNAK As Long = 0
Public Const OFS_TIP As Long = 1
Public Const OFS_HEDEF As Long = 2
Public Const OFS_FILTRE As Long = 3
Public Const OFS_ISLEM As Long = 4
Public Const OFS_HARITA As Long = 5

Public Const OFS_ETIKET As Long = 0         ' ayar etiketi
Public Const OFS_AYAR As Long = 2           ' ayar degeri (menulu hucre)
Public Const OFS_SABIT_SUTUN As Long = 3    ' SABIT DEGERLER: ana dosya sutunu
Public Const OFS_SABIT_DEGER As Long = 4    ' SABIT DEGERLER: deger
Public Const OFS_KARNE As Long = 5          ' KONTROL NOKTALARI karnesi

Public Const SERIT_SON_SATIR As Long = 5     ' 1-5 serit; 6'dan itibaren kayar
Public Const SATIR_BLOK_BASLIK As Long = 7
Public Const SATIR_DAHIL As Long = 8
Public Const SATIR_KAYNAK_BOLUM As Long = 9
Public Const SATIR_KAYNAK_DOSYA As Long = 10
Public Const SATIR_KAYNAK_SAYFA As Long = 11
Public Const SATIR_SON_SATIR As Long = 12
Public Const SATIR_HEDEF_BOLUM As Long = 13
Public Const SATIR_HEDEF_DOSYA As Long = 14
Public Const SATIR_HEDEF_SAYFA As Long = 15
Public Const SATIR_YAZIM As Long = 16
Public Const SATIR_KONTROL_BOLUM As Long = 17
Public Const SATIR_PROFIL As Long = 18
Public Const SATIR_TUTAR As Long = 19
Public Const SATIR_HASAR As Long = 20
Public Const SATIR_IHBAR As Long = 21
Public Const SATIR_ANAHTAR As Long = 22      ' v1.1: K9 anahtar (hasar/dosya no)
Public Const SATIR_KIRILIM As Long = 23      ' v1.1: K12/K13 kirilim (brans/sirket)
Public Const SATIR_DONEM As Long = 24        ' v1.1: K10 rapor donemi sonu
Public Const SATIR_AYAR_SON As Long = 24
Public Const SATIR_V10_AYAR_SON As Long = 21 ' v1.0 ayar karti son satiri

Public Const SABIT_BASLIK As Long = 8
Public Const SABIT_ILK As Long = 9
Public Const SABIT_SON As Long = 23
Public Const SABIT_NOT As Long = 24
Public Const SABIT_EN_FAZLA As Long = 15

' KONTROL NOKTALARI karnesi (c+5): 8 baslik + rozet, 9..22 K1..K14,
' 23 veri kalitesi skoru, 24 son kontrol notu (v1.0: 9..16 K1..K8, 17 not)
Public Const KARNE_BASLIK As Long = 8
Public Const KARNE_ILK As Long = 9
Public Const KARNE_KALITE As Long = 23
Public Const KARNE_NOT As Long = 24
Public Const K_SAYI As Long = 14

' Kontrol noktasi durumlari (artan onem sirasiyla)
Public Const KD_BEKLIYOR As String = "-"       ' henuz calismadi
Public Const KD_YOK As String = "YOK"          ' bu aktarimda uygulanmaz
Public Const KD_OK As String = "OK"
Public Const KD_DIKKAT As String = "DIKKAT"
Public Const KD_ENGEL As String = "ENGEL"

Public Const SATIR_DURUM As Long = 26
Public Const BASLIK_SATIR As Long = 28
Public Const ILK_SATIR As Long = 29
Public Const SON_SATIR As Long = 530
Public Const EN_FAZLA_SUTUN As Long = 300    ' kaynak/hedef sayfada okunan sutun siniri

' Panel temizlenirken bakilan en sag sutun (v33: 20 blok x 7 sutun)
Public Const TEMIZLIK_SON_SUTUN As Long = 160

' ------------------------------------------------------------
'  v33 YERLESIMI (yalnizca gecis / PaneliYenidenKur icin)
' ------------------------------------------------------------
Public Const V33_BLOK_GENISLIK As Long = 7
Public Const V33_MAX_BLOK As Long = 20
Public Const V33_OFS_HEDEF As Long = 1
Public Const V33_OFS_FILTRE As Long = 3
Public Const V33_OFS_GERIAL As Long = 4
Public Const V33_OFS_FORMUL As Long = 5
Public Const V33_SABIT_SUTUN_OFS As Long = 2
Public Const V33_SABIT_DEGER_OFS As Long = 3
Public Const V33_SATIR_HASAR As Long = 22
Public Const V33_SATIR_IHBAR As Long = 23
Public Const V33_SATIR_TUTAR As Long = 24
' v34 (6 sutun, harita ortada)
Public Const V34_OFS_FILTRE As Long = 4
Public Const V34_OFS_ISLEM As Long = 5
Public Const V31_SABIT_SUTUN As Long = 19     ' v31 tek sabit deger satirlari
Public Const V31_SABIT_DEGER As Long = 20

' ------------------------------------------------------------
'  Listeler<k> (gizli) SAYFASI SUTUNLARI
' ------------------------------------------------------------
Public Const LST_DOSYALAR As Long = 1        ' acik dosyalar
Public Const LST_KAYNAK_SAYFALAR As Long = 2
Public Const LST_HEDEF_SAYFALAR As Long = 3
Public Const LST_HEDEF_SUTUNLAR As Long = 4  ' "F - TUTAR" bicimi
Public Const LST_KAYNAK_TIP As Long = 5      ' kaynak sutun tipi (satir = sutun no)
Public Const LST_DURUM As Long = 6           ' 1 notlar, 2 yesil toplam, 3 yesil bos, 4 tip uyusmaz
Public Const LST_PROFILLER As Long = 7
Public Const LST_KALAN_HEDEF As Long = 8     ' henuz kullanilmamis hedefler
Public Const LST_EVET_HAYIR As Long = 9
Public Const LST_YAZIM As Long = 10
Public Const LST_ISLEM As Long = 11
Public Const LST_FILTRE_ILK As Long = 12     ' 12 .. 12+299: sutun basina filtre adaylari
Public Const LST_SABIT_ILK As Long = 320     ' 320 .. 334: sabit satiri basina deger menusu
Public Const LST_KAYNAK_TIP_METIN As Long = 338
Public Const LST_HEDEF_TIP As Long = 340     ' ana dosya tip onbellegi (satir = sutun no)
Public Const LST_RENK As Long = 341          ' 341 renk, 342 renkli mi, 343/1 sutun sayisi

Public Const LST_DURUM_NOT As Long = 1
Public Const LST_DURUM_YESIL_TOPLAM As Long = 2
Public Const LST_DURUM_YESIL_BOS As Long = 3
Public Const LST_DURUM_TIP_UYUSMAZ As Long = 4

' ------------------------------------------------------------
'  ESIKLER
' ------------------------------------------------------------
Public Const FILTRE_TARAMA As Long = 10000   ' filtre adaylari icin okunan en fazla satir
Public Const FILTRE_SINIRI As Long = 300     ' bundan fazla farkli deger -> menu kurulmaz
Public Const TIP_ORNEK As Long = 300         ' veri tipi icin bakilan satir
Public Const SABIT_TARAMA As Long = 50000    ' sabit deger menusu: ana dosyanin son N satiri
Public Const SABIT_MENU_SINIRI As Long = 300
Public Const SATIR_UYARI_ESIGI As Long = 50000
Public Const MSG_SINIR As Long = 950
Public Const MSGBOX_SINIR As Long = 1010      ' MsgBox metninin gercek siniri ~1024
Public Const BENZERLIK_ESIGI As Double = 0.7
Public Const KUYRUK_PENCERE As Long = 30
Public Const KUYRUK_EN_FAZLA As Long = 15
Public Const KONTROL_ORNEK As Long = 5       ' on kontrolde gosterilen ornek satir sayisi
Public Const TABAKA_PARCA As Long = 100      ' M9: bas / orta / son / rastgele parca boyu
Public Const TABAKA_TOHUM As Long = 20260917 ' M9: rastgele parca icin sabit tohum (tekrarlanabilir)
Public Const BILDIRIM_TUTULAN As Long = 200  ' Bellek'te tutulan son bildirim
Public Const BILDIRIM_GOSTERILEN As Long = 5 ' serit listesinde gosterilen
Public Const KILIT_UZANTI As String = ".aplock"
Public Const PARA_SINIR As Double = 900000000000000#   ' Currency guvenli sinir (~9,2E14)

' Veri kalitesi skoru agirliklari (toplam 100; DN-06). Uygulanamayan
' bilesenin agirligi digerlerine oransal dagitilir.
Public Const KS_TAMLIK As Double = 30
Public Const KS_GECERLILIK As Double = 25
Public Const KS_TEKILLIK As Double = 20
Public Const KS_MAKULLUK As Double = 15
Public Const KS_MUTABAKAT As Double = 10

' ------------------------------------------------------------
'  DUZENLENEBILIR ESIKLER (Ayarlar sayfasi; M15_Esik)
'  Burada yalnizca VARSAYILANLAR durur. Etkin deger Esik(ek...) ile
'  Ayarlar sayfasindan okunur; bos/gecersizse varsayilan kullanilir.
' ------------------------------------------------------------
Public Enum EsikKod
    ekTarihAltYil = 1       ' K10/M8: makul tarih alt siniri (yil)
    ekTarihUstAy            ' K10/M8: ust sinir = donem sonu + N ay
    ekZSkor                 ' K11: |z| esigi
    ekZMinAy                ' K11: z icin gereken en az gecmis ay
    ekBuyukHasar            ' K11: buyuk hasar tutari (0 = kapali)
    ekBenfordMad            ' K11: Benford ilk hane MAD esigi (0 = kapali)
    ekBenfordMinN           ' K11: Benford icin en az satir
    ekK9Pencere             ' K9: ana dosyanin taranan son satir sayisi
    ekK9Ornek               ' K9: ana dosyada aranan ornek satir
    ekK9KopyaEngel          ' K9: birebir kopya satir ENGEL mi (1) DIKKAT mi (0)
    ekK12Pencere            ' K12: kategori kaymasi icin taranan son satir
    ekTamlikHedef           ' kalite: tamlik hedefi (%)
    ekKaliteHedef           ' kalite skoru hedefi (0-100)
    ekYedekParti            ' M14: yedegi saklanan son parti sayisi
    ekKilitDakika           ' M12: kilit bayatlik suresi (dakika)
    ekMutabakatSutun        ' M1: mutabakat yapilan en fazla tutar sutunu
    ekMutabakatTolerans     ' K7: mutlak tolerans (para birimi)
    ekDosyaMB               ' Denetim: dosya boyutu uyari esigi (MB)
    ekAnalizGrup            ' Analiz: karma gruplama siniri (ustu siralama)
    ekAnalizSatir           ' Analiz: tabloda gosterilen en fazla satir
End Enum
Public Const ESIK_SAYI As Long = 20

Public Const ESV_TARIH_ALT_YIL As Double = 1990
Public Const ESV_TARIH_UST_AY As Double = 12
Public Const ESV_Z_SKOR As Double = 3
Public Const ESV_Z_MIN_AY As Double = 3
Public Const ESV_BUYUK_HASAR As Double = 10000000
Public Const ESV_BENFORD_MAD As Double = 0.015
Public Const ESV_BENFORD_MIN_N As Double = 1000
Public Const ESV_K9_PENCERE As Double = 250000
Public Const ESV_K9_ORNEK As Double = 200
Public Const ESV_K9_KOPYA_ENGEL As Double = 1
Public Const ESV_K12_PENCERE As Double = 250000
Public Const ESV_TAMLIK_HEDEF As Double = 99.5
Public Const ESV_KALITE_HEDEF As Double = 95
Public Const ESV_YEDEK_PARTI As Double = 20
Public Const ESV_KILIT_DAKIKA As Double = 120
Public Const ESV_MUTABAKAT_SUTUN As Double = 8
Public Const ESV_MUTABAKAT_TOLERANS As Double = 0.005
Public Const ESV_DOSYA_MB As Double = 50
Public Const ESV_ANALIZ_GRUP As Double = 5000
Public Const ESV_ANALIZ_SATIR As Double = 50

' Ayarlar sayfasi yerlesimi
Public Const AYR_BASLIK_SATIR As Long = 5
Public Const AYR_ILK As Long = 6                ' ESIK_SAYI satir: 6..25
Public Const AYR_KOD As Long = 2
Public Const AYR_AD As Long = 3
Public Const AYR_DEGER As Long = 4
Public Const AYR_VARSAYILAN As Long = 5
Public Const AYR_ACIKLAMA As Long = 6
Public Const AYR_Z_BASLIK As Long = 28          ' ZORUNLU SUTUNLAR (M10 meta veri)
Public Const AYR_Z_ILK As Long = 30
Public Const AYR_Z_SON As Long = 129
Public Const AYR_Z_SAYFA As Long = 2            ' ana dosya sayfasi (bos = hepsi)
Public Const AYR_Z_SUTUN As Long = 3            ' ana dosya basligi

' Analiz sayfasi: istek karti (etiket B, deger D), yorum, tablo, grafik
Public Const AN_ETIKET_SUTUN As Long = 2
Public Const AN_DEGER_SUTUN As Long = 4
Public Const AN_SATIR_SABLON As Long = 5
Public Const AN_SATIR_KAYNAK As Long = 6
Public Const AN_SATIR_SATIR As Long = 7       ' satir boyutu
Public Const AN_SATIR_SUTUN As Long = 8       ' sutun boyutu (opsiyonel)
Public Const AN_SATIR_OLCU As Long = 9
Public Const AN_SATIR_OLCU_SUTUN As Long = 10
Public Const AN_SATIR_FILTRE_SUTUN As Long = 11
Public Const AN_SATIR_FILTRE As Long = 12
Public Const AN_SATIR_DONEM_SUTUN As Long = 13
Public Const AN_SATIR_DONEM_BAS As Long = 14
Public Const AN_SATIR_DONEM_SON As Long = 15
Public Const AN_ILK_ALAN As Long = 5
Public Const AN_SON_ALAN As Long = 15
Public Const AN_YORUM_ILK As Long = 18          ' 3 madde: 18..20
Public Const AN_TABLO_ILK As Long = 23          ' sonuc tablosu basligi
Public Const AN_TABLO_SUTUN As Long = 2
Public Const AN_EN_FAZLA_SUTUN As Long = 12     ' sutun boyutunda gosterilen en fazla deger
Public Const AN_GRAFIK As String = "AP_AnalizGrafik"
Public Const AN_LST_SUTUN As Long = 30          ' istek karti menuleri (gizli liste sutunlari 30..)

' Bellek sayfasi (gizli): anahtar/deger ve bildirim halkasi
Public Const BLK_ANAHTAR As Long = 1
Public Const BLK_DEGER As Long = 2
Public Const BLK_ZAMAN As Long = 3
Public Const BLD_ZAMAN As Long = 5
Public Const BLD_SEVIYE As Long = 6
Public Const BLD_BASLIK As Long = 7
Public Const BLD_METIN As Long = 8
Public Const BLD_RAPOR As Long = 9

' Eski surumlerin ana dosyaya koydugu gecici olcut isareti (harita/listeden dusulur)
Public Const KONTROL_BASLIK As String = "_KONTROL_FARK"

' ------------------------------------------------------------
'  RAPOR SAYFASI SUTUNLARI (olay gunlugu)
' ------------------------------------------------------------
Public Const R_TARIH As Long = 1
Public Const R_KULLANICI As Long = 2
Public Const R_OLAY As Long = 3
Public Const R_PANEL As Long = 4
Public Const R_PARTI As Long = 5
Public Const R_AKTARIM As Long = 6
Public Const R_ACIKLAMA As Long = 7
Public Const R_KARAR As Long = 8
Public Const R_KAYNAK_DOSYA As Long = 9
Public Const R_KAYNAK_SAYFA As Long = 10
Public Const R_HEDEF_DOSYA As Long = 11
Public Const R_HEDEF_SAYFA As Long = 12
Public Const R_KAYNAK_SON As Long = 13
Public Const R_SATIR As Long = 14
Public Const R_SUTUN As Long = 15
Public Const R_HEDEF_ILK As Long = 16
Public Const R_HEDEF_SON As Long = 17
Public Const R_FILTRE As Long = 18
Public Const R_SABIT As Long = 19
Public Const R_MUTABAKAT As Long = 20
Public Const R_AYLIK_TOPLAM As Long = 21
Public Const R_TUTAR_SUTUN As Long = 22
Public Const R_EXCEL_KULLANICI As Long = 23
Public Const R_YAZILAN_TOPLAM As Long = 24   ' v34
Public Const R_HEDEF_ONCEKI As Long = 25     ' v34
Public Const R_HEDEF_YENI As Long = 26       ' v34
Public Const R_KONTROL As Long = 27          ' v1.0: kontrol noktasi (K1..K14)
Public Const R_KAYNAK_YOL As Long = 28       ' v1.1: kaynak soyu (M15)
Public Const R_KAYNAK_BOYUT As Long = 29
Public Const R_KAYNAK_ZAMAN As Long = 30
Public Const R_KAYNAK_KARMA As Long = 31     ' K14 kaynak izi
Public Const R_ICERIK_IZI As Long = 32       ' K2 icerik izi (M7)
Public Const R_SURE_MS As Long = 33          ' adim suresi (AL-04)
Public Const R_ADET As Long = 34             ' adet mutabakati (M2)
Public Const R_KALITE As Long = 35           ' veri kalitesi skoru 0-100
Public Const R_METRIK As Long = 36           ' metrikler: ad=deger;...
Public Const R_SON As Long = 36
Public Const R_V10_SON As Long = 27

' ------------------------------------------------------------
'  GECMIS SAYFASI SUTUNLARI (Geri Al defteri)
' ------------------------------------------------------------
Public Const G_PARTI As Long = 1
Public Const G_AKTARIM As Long = 2
Public Const G_TARIH As Long = 3
Public Const G_HEDEF_DOSYA As Long = 4
Public Const G_HEDEF_SAYFA As Long = 5
Public Const G_HEDEF_SUTUN As Long = 6       ' HARF olarak
Public Const G_KAYNAK_SUTUN As Long = 7
Public Const G_ILK As Long = 8
Public Const G_SON As Long = 9
Public Const G_SATIR_SAYISI As Long = 10
Public Const G_DURUM As Long = 11            ' AKTIF / GERI ALINDI
Public Const G_GERI_ALMA_TARIHI As Long = 12
Public Const G_HEDEF_ADI As Long = 13
Public Const G_KULLANICI As Long = 14
Public Const G_IMZA As Long = 15             ' v34: parmak izi
Public Const G_YEDEK As Long = 16            ' v34: uzerine yazilan eski degerlerin yedegi
Public Const G_SON_SUTUN As Long = 16

Public Const DURUM_AKTIF As String = "AKTIF"
Public Const DURUM_GERI_ALINDI As String = "GERI ALINDI"

' ------------------------------------------------------------
'  PROFILLER SAYFASI SUTUNLARI
' ------------------------------------------------------------
Public Const P_AD As Long = 1
Public Const P_KAYNAK As Long = 2
Public Const P_HEDEF As Long = 3
Public Const P_FILTRE As Long = 4
Public Const P_FORMUL As Long = 5
Public Const P_DIGER As Long = 6
Public Const P_SIRA As Long = 7
Public Const P_YER As Long = 8
Public Const P_TIP As Long = 9
Public Const P_HEDEF_SIRA As Long = 10
Public Const P_KAYIT As Long = 11
Public Const P_SON As Long = 11

' ------------------------------------------------------------
'  SETLER SAYFASI SUTUNLARI
' ------------------------------------------------------------
Public Const ST_AD As Long = 1
Public Const ST_PANEL As Long = 2
Public Const ST_HEDEF_DOSYA As Long = 3
Public Const ST_HEDEF_SAYFA As Long = 4
Public Const ST_PROFIL As Long = 5
Public Const ST_ESKI_SABIT_SUTUN As Long = 6  ' v31
Public Const ST_ESKI_SABIT_DEGER As Long = 7  ' v31
Public Const ST_TUTAR As Long = 8
Public Const ST_HASAR As Long = 9
Public Const ST_IHBAR As Long = 10
Public Const ST_YAZIM As Long = 11
Public Const ST_SABITLER As Long = 12
Public Const ST_ANAHTAR As Long = 13          ' v1.1
Public Const ST_KIRILIM As Long = 14          ' v1.1
Public Const ST_DONEM As Long = 15            ' v1.1
Public Const ST_SON As Long = 15

' Serit geometrisi (point). Birincil = 4 adim (yalniz SIRADAKI dolgulu),
' ikincil = kenarlikli (Analiz, Denetim, Diger), ucuncul = "Diger" menusu,
' tehlikeli = Geri Al (kirmizi hayalet, ayri grupta, en sagda).
Public Const SERIT_SOL As Double = 10
Public Const SERIT_ADIM_GENISLIK As Double = 112
Public Const SERIT_BUTON_YUKSEKLIK As Double = 26
Public Const SERIT_IKINCIL_GENISLIK As Double = 84
Public Const SERIT_IKINCIL_YUKSEKLIK As Double = 22
Public Const SERIT_BUTON_ARALIK As Double = 6
Public Const SERIT_GRUP_ARALIK As Double = 18
Public Const SERIT_MENU_GENISLIK As Double = 190
Public Const SERIT_MENU_YUKSEKLIK As Double = 22
Public Const SERIT_MENU_ONEK As String = "AP_MENU_"
Public Const SERIT_ADIM_SAYI As Long = 4

Public Const DN_TREND_SATIR As Long = 10
Public Const DN_TREND_SUTUN As Long = 15
Public Const DN_TREND_GRAFIK_SUTUN As Long = 8
Public Const DN_TREND_AY As Long = 12
Public Const DN_TREND_GRAFIK As String = "AP_DenetimTrend"

' ------------------------------------------------------------
'  ENUM'LAR
' ------------------------------------------------------------
Public Enum YazimModu
    ymSayfaSonu = 1         ' "Sayfanin en alti"   (normal aylik aktarim)
    ymSutunSonu = 2         ' "Sutunun kendi sonu"
    ymGeriAlinan = 3        ' "Geri alinan araliga"
End Enum

' BaslikDogrula donusu
Public Enum BaslikDurumu
    bdYerinde = 0
    bdYerDegisti = 1
    bdYok = 2
    bdBelirsiz = 3
End Enum

' AktarCalistir donusu
Public Enum AktarimSonucu
    asHata = -1
    asIptal = 0
    asTamam = 1
End Enum

Public Enum PaletRenk
    prSeritZemin = 1
    prSeritYazi
    prSeritGrupYazi
    prButonAna
    prButonYesil
    prButonKirmizi
    prButonIkincil
    prButonKenar
    prBlokBaslik
    prBlokBaslikAktif
    prBolumYazi
    prEtiketYazi
    prAyarZemin
    prAyarKenar
    prDahilEvet
    prDahilHayir
    prSabitBaslik
    prSabitDolu
    prSabitBosSatir
    prSabitEksik
    prDurumNotr
    prDurumHazir
    prDurumUyari
    prDurumHata
    prDurumOneri
    prTabloBaslik
    prKaynakZemin
    prHedefZemin
    prFiltreZemin
    prIslemZemin
    prOneri
    prTipUyusmaz
    prTipUyusmazYazi
    prCerceve
    prZemin
    prYuzey
    prCizgi
    prMetin
    prIkincilMetin
    prMarka
    prMarkaZemin
    prOk
    prOkZemin
    prDikkat
    prDikkatZemin
    prEngel
    prEngelZemin
    prOneriMetin
    prOneriZemin
    prPasifMetin
End Enum

' ------------------------------------------------------------
'  GLOBAL DURUM
' ------------------------------------------------------------
Public g_SonAcmaHatasi As String           ' dosya acilirken olusan son hata
Public g_MutabakatOzet As String           ' sirali aktarimda panel basina sonuc
Public g_OkumaUyarisi As String            ' Value2'ye dusulen sutunlar (##### tarih)
Public g_SabitOnbellek As Sozluk           ' sabit deger menuleri (bir Yenile boyunca)
Public g_SonAktifBlok As Long              ' serit gostergesi icin

' EkranKilitle / EkranAc durumu (M02_Yardimci)
Public g_KilitDerinlik As Long
Public g_EskiEkran As Boolean
Public g_EskiOlay As Boolean
Public g_EskiUyari As Boolean
Public g_EskiHesap As Long
Public g_HesapDurdu As Boolean


' ============================================================
'  RENK PALETI  (TEK KAYNAK; baska modulde RGB yok - build kapisi denetler)
'
'  Anlam katmani: 6 notr + 1 marka + 3 durum + 1 oneri. Zemin rengi
'  yalnizca DURUM icin kullanilir, bolge icin kullanilmaz.
'    YESIL  = yalnizca OK       AMBER = DIKKAT     KIRMIZI = ENGEL
'    MOR    = oneri (onay bekliyor)                MAVI    = marka / siradaki adim
'  Her metin/zemin cifti >= 4.5:1 (WCAG AA; tools/source_checks.py).
'  Asagidaki rol adlari (prSerit..., prDurum...) tokenlara takma addir.
' ============================================================
Public Function Palet(ByVal renk As PaletRenk) As Long

    Select Case renk
        ' --- tokenlar ---
        Case prZemin: Palet = RGB(255, 255, 255)
        Case prYuzey: Palet = RGB(248, 250, 252)
        Case prTabloBaslik: Palet = RGB(241, 245, 249)
        Case prCizgi: Palet = RGB(226, 232, 240)
        Case prCerceve: Palet = RGB(203, 213, 225)
        Case prMetin: Palet = RGB(15, 23, 42)
        Case prIkincilMetin: Palet = RGB(71, 85, 105)
        Case prPasifMetin: Palet = RGB(100, 116, 139)
        Case prMarka: Palet = RGB(29, 78, 216)
        Case prMarkaZemin: Palet = RGB(219, 234, 254)
        Case prOk: Palet = RGB(21, 128, 61)
        Case prOkZemin: Palet = RGB(220, 252, 231)
        Case prDikkat: Palet = RGB(146, 64, 14)
        Case prDikkatZemin: Palet = RGB(254, 243, 199)
        Case prEngel: Palet = RGB(185, 28, 28)
        Case prEngelZemin: Palet = RGB(254, 226, 226)
        Case prOneriMetin: Palet = RGB(109, 40, 217)
        Case prOneriZemin: Palet = RGB(237, 233, 254)
        ' --- rol takma adlari ---
        Case prSeritZemin: Palet = Palet(prYuzey)
        Case prSeritYazi: Palet = Palet(prMetin)
        Case prSeritGrupYazi: Palet = Palet(prIkincilMetin)
        Case prButonAna: Palet = Palet(prMarka)
        Case prButonYesil: Palet = Palet(prOk)
        Case prButonKirmizi: Palet = Palet(prEngel)
        Case prButonIkincil: Palet = Palet(prIkincilMetin)
        Case prButonKenar: Palet = Palet(prCerceve)
        Case prBlokBaslik: Palet = Palet(prMetin)
        Case prBlokBaslikAktif: Palet = Palet(prMarka)
        Case prBolumYazi: Palet = Palet(prMarka)
        Case prEtiketYazi: Palet = Palet(prIkincilMetin)
        Case prAyarZemin: Palet = Palet(prZemin)
        Case prAyarKenar: Palet = Palet(prCizgi)
        Case prDahilEvet: Palet = Palet(prZemin)
        Case prDahilHayir: Palet = Palet(prZemin)
        Case prSabitBaslik: Palet = Palet(prTabloBaslik)
        Case prSabitDolu: Palet = Palet(prZemin)
        Case prSabitBosSatir: Palet = Palet(prYuzey)
        Case prSabitEksik: Palet = Palet(prEngelZemin)
        Case prDurumNotr: Palet = Palet(prYuzey)
        Case prDurumHazir: Palet = Palet(prOkZemin)
        Case prDurumUyari: Palet = Palet(prDikkatZemin)
        Case prDurumHata: Palet = Palet(prEngelZemin)
        Case prDurumOneri: Palet = Palet(prOneriZemin)
        Case prKaynakZemin: Palet = Palet(prYuzey)
        Case prHedefZemin: Palet = Palet(prZemin)
        Case prFiltreZemin: Palet = Palet(prZemin)
        Case prIslemZemin: Palet = Palet(prZemin)
        Case prOneri: Palet = Palet(prOneriZemin)
        Case prTipUyusmaz: Palet = Palet(prEngelZemin)
        Case prTipUyusmazYazi: Palet = Palet(prEngel)
        Case Else: Palet = RGB(255, 255, 255)
    End Select

End Function
