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

Public Const SURUM As String = "v1.0"

' Panel duzeni isareti (calisma kitabi adi). v33 panelinde yoktur.
Public Const DUZEN_ADI As String = "AP_DUZEN"
Public Const DUZEN_DEGERI As String = "v1.0"

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
Public Const SATIR_AYAR_SON As Long = 21

Public Const SABIT_BASLIK As Long = 8
Public Const SABIT_ILK As Long = 9
Public Const SABIT_SON As Long = 23
Public Const SABIT_NOT As Long = 24
Public Const SABIT_EN_FAZLA As Long = 15

' KONTROL NOKTALARI karnesi (c+5): 8 baslik, 9..16 K1..K8, 17 son kontrol notu
Public Const KARNE_BASLIK As Long = 8
Public Const KARNE_ILK As Long = 9
Public Const KARNE_NOT As Long = 17
Public Const K_SAYI As Long = 8

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
Public Const R_KONTROL As Long = 27          ' v1.0: kontrol noktasi (K1..K8)
Public Const R_SON As Long = 27

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
'  RENK PALETI  (tek yerden degistirilir)
' ============================================================
Public Function Palet(ByVal renk As PaletRenk) As Long

    Select Case renk
        Case prSeritZemin:      Palet = RGB(15, 23, 42)
        Case prSeritYazi:       Palet = RGB(241, 245, 249)
        Case prSeritGrupYazi:   Palet = RGB(148, 163, 184)
        Case prButonAna:        Palet = RGB(37, 99, 235)
        Case prButonYesil:      Palet = RGB(22, 163, 74)
        Case prButonKirmizi:    Palet = RGB(220, 38, 38)
        Case prButonIkincil:    Palet = RGB(51, 65, 85)
        Case prButonKenar:      Palet = RGB(100, 116, 139)
        Case prBlokBaslik:      Palet = RGB(71, 85, 105)
        Case prBlokBaslikAktif: Palet = RGB(29, 78, 216)
        Case prBolumYazi:       Palet = RGB(29, 78, 216)
        Case prEtiketYazi:      Palet = RGB(51, 65, 85)
        Case prAyarZemin:       Palet = RGB(255, 251, 235)
        Case prAyarKenar:       Palet = RGB(203, 213, 225)
        Case prDahilEvet:       Palet = RGB(187, 247, 208)
        Case prDahilHayir:      Palet = RGB(254, 226, 226)
        Case prSabitBaslik:     Palet = RGB(167, 243, 208)
        Case prSabitDolu:       Palet = RGB(220, 252, 231)
        Case prSabitBosSatir:   Palet = RGB(240, 253, 244)
        Case prSabitEksik:      Palet = RGB(254, 226, 226)
        Case prDurumNotr:       Palet = RGB(241, 245, 249)
        Case prDurumHazir:      Palet = RGB(220, 252, 231)
        Case prDurumUyari:      Palet = RGB(254, 243, 199)
        Case prDurumHata:       Palet = RGB(254, 226, 226)
        Case prDurumOneri:      Palet = RGB(255, 237, 213)
        Case prTabloBaslik:     Palet = RGB(226, 232, 240)
        Case prKaynakZemin:     Palet = RGB(248, 250, 252)
        Case prHedefZemin:      Palet = RGB(254, 252, 232)
        Case prFiltreZemin:     Palet = RGB(240, 253, 244)
        Case prIslemZemin:      Palet = RGB(245, 243, 255)
        Case prOneri:           Palet = RGB(255, 200, 120)
        Case prTipUyusmaz:      Palet = RGB(255, 199, 206)
        Case prTipUyusmazYazi:  Palet = RGB(156, 0, 6)
        Case prCerceve:         Palet = RGB(148, 163, 184)
        Case Else:              Palet = RGB(255, 255, 255)
    End Select

End Function
