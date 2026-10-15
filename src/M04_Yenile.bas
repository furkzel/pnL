Attribute VB_Name = "M04_Yenile"
Option Explicit

' ============================================================
' AKTARIM PANELI  v34  -  M04_Yenile
'
'   Yenile            : acik dosyalar, sayfalar, sutunlar, menuler;
'                       eski eslesmeleri geri koyar, profili uygular.
'   KaynakTara        : TEK GECISTE veri tipi + filtre adaylari
'                       (sutun basina en fazla FILTRE_TARAMA satir).
'   BlokMenuleriTazele: kalan hedef menuleri, renkler, harita, DURUM.
'   HedefDegisti      : Panel sayfasindaki hucre degisikliklerini
'                       yonlendirir (oneri onayi, SABIT kutusu,
'                       "Bu paneli aktar", dosya secimi).
'   SABIT DEGERLER    : SabitAlaniniDuzenle ve yardimcilari.
'   HEDEF HARITASI    : HaritayiYaz (renkler + [+] + YESIL BOSTA).
'   VERI TIPI         : TipSutunuYaz.
' ============================================================


' ============================================================
'  YENILE
' ============================================================

Public Sub Yenile()
    YenileCalistir False
End Sub


Public Sub YenileCalistir(ByVal sessiz As Boolean)

    Dim pnl As Worksheet
    Dim k As Long, adet As Long
    Dim ozet As String, adim As String
    Dim bas As Double

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub

    adet = BlokSayisi(pnl)
    bas = Timer

    EkranKilitle
    On Error GoTo Hata

    g_SonAcmaHatasi = ""
    Set g_SabitOnbellek = New Sozluk

    For k = 1 To adet
        If Not BlokGizliMi(pnl, k) Then
            adim = "PANEL " & k & " yenileniyor"
            DurumCubugu "YENILE: " & adim & " ..."
            ozet = ozet & BlokYenile(pnl, k) & vbNewLine
        End If
    Next k

    DurumCubugu ""
    EkranAc
    g_SonAktifBlok = 0
    AktifPanelGoster pnl, AktifBlok()

    If sessiz Then Exit Sub

    ozet = "YENILE TAMAMLANDI   (acik dosya: " & (Workbooks.Count - 1) & ", " & _
           Format$(Timer - bas, "0.0") & " sn)" & vbNewLine & vbNewLine & ozet
    If g_SonAcmaHatasi <> "" Then
        ozet = ozet & vbNewLine & "DOSYA ACMA HATASI:" & vbNewLine & g_SonAcmaHatasi & vbNewLine
    End If
    ozet = ozet & vbNewLine & "Ayrintilar her panelin DURUM satirinda."
    MsgBox Sigdir(ozet, MSG_SINIR), vbInformation, "Yenile"
    Exit Sub

Hata:
    DurumCubugu ""
    EkranAc
    MsgBox "YENILE SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Adim  : " & adim & vbNewLine & _
           "Hata  : " & Err.Number & " - " & Err.Description, vbCritical, "Hata"

End Sub


' Panel sayfasi var ve v34 duzeninde mi? Degilse kullaniciya ne
' yapacagini soyler.
Public Function PanelHazirMi(ByVal pnl As Worksheet) As Boolean

    PanelHazirMi = False
    If pnl Is Nothing Then
        MsgBox "Panel sayfasi bulunamadi." & vbNewLine & _
               "Alt+F8 > SifirdanKur calistir.", vbExclamation, "Panel"
        Exit Function
    End If
    If Not PanelDuzeniGuncel() Then
        MsgBox "Panel sayfasi eski (v33 veya oncesi) duzende." & vbNewLine & vbNewLine & _
               "Alt+F8 > SifirdanKur calistir: paneldeki tum ayarlar ve" & vbNewLine & _
               "eslestirmeler korunarak yeni duzene tasinir.", vbExclamation, "Panel"
        Exit Function
    End If
    If BlokSayisi(pnl) = 0 Then
        MsgBox "Panel blogu bulunamadi. Alt+F8 > SifirdanKur calistir.", vbExclamation, "Panel"
        Exit Function
    End If
    PanelHazirMi = True

End Function


Public Function PanelDuzeniGuncel() As Boolean
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(DUZEN_ADI).RefersTo
    On Error GoTo 0
    PanelDuzeniGuncel = (InStr(s, DUZEN_DEGERI) > 0)
End Function


' Tek bir blogu yeniler; Yenile ozeti icin kisa bir satir dondurur.
Public Function BlokYenile(ByVal pnl As Worksheet, ByVal k As Long) As String

    Dim lst As Worksheet
    Dim wb As Workbook, ws As Worksheet
    Dim kaynakWb As Workbook, hedefWb As Workbook
    Dim kaynakWs As Worksheet, hedefWs As Worksheet
    Dim c As Long, i As Long, n As Long, sonSut As Long, son As Long
    Dim eskiKaynak As Variant, eskiHedef As Variant, eskiFiltre As Variant, eskiIslem As Variant
    Dim eskiOneri() As String
    Dim dosyalar() As Variant, sayfalar() As Variant, basliklar As Variant, yaz() As Variant
    Dim profilAdi As String, profilBulundu As Boolean
    Dim pTamam As Long, pOneri As Long
    Dim pEksik As String, pYeni As String, pHedefYok As String
    Dim notlar As String, ayniKaynak As String, ayniHedef As String
    Dim durumMetni As String, otomatikSecilen As String

    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)

    ' --- Mevcut eslesmeleri ve onay bekleyen onerileri hafizaya al ---
    son = BlokSonSatir(pnl, k)
    If son >= ILK_SATIR Then
        eskiKaynak = IkiBoyutlu(pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_KAYNAK), pnl.Cells(son, c + OFS_KAYNAK)))
        eskiHedef = IkiBoyutlu(pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(son, c + OFS_HEDEF)))
        eskiFiltre = IkiBoyutlu(pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_FILTRE), pnl.Cells(son, c + OFS_FILTRE)))
        eskiIslem = IkiBoyutlu(pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_ISLEM), pnl.Cells(son, c + OFS_ISLEM)))
        ReDim eskiOneri(1 To son - ILK_SATIR + 1)
        For i = ILK_SATIR To son
            If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then
                eskiOneri(i - ILK_SATIR + 1) = OneriMetni(pnl.Cells(i, c + OFS_HEDEF))
            End If
        Next i
    End If

    lst.UsedRange.Clear

    ' --- Acik dosyalar (bu calisma kitabi haric) ---
    n = 0
    ReDim dosyalar(0 To Workbooks.Count)
    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            dosyalar(n) = wb.Name
            n = n + 1
        End If
    Next wb
    If n > 0 Then
        ReDim Preserve dosyalar(0 To n - 1)
        ListeYaz lst, LST_DOSYALAR, dosyalar, True
    End If
    MenuKurSerbest Ayar(pnl, k, SATIR_KAYNAK_DOSYA), lst, LST_DOSYALAR
    MenuKurSerbest Ayar(pnl, k, SATIR_HEDEF_DOSYA), lst, LST_DOSYALAR

    ' --- Sabit listeler (Evet/Hayir, yazim modu, ISLEM) ---
    SabitListeleriYaz lst
    MenuKur Ayar(pnl, k, SATIR_DAHIL), lst, LST_EVET_HAYIR
    MenuKur Ayar(pnl, k, SATIR_YAZIM), lst, LST_YAZIM
    If DahilMi(pnl, k) Then
        Ayar(pnl, k, SATIR_DAHIL).Value = EvetMetni()
    Else
        Ayar(pnl, k, SATIR_DAHIL).Value = HayirMetni()
    End If
    Ayar(pnl, k, SATIR_YAZIM).Value = YazimMetni(YazimModuOku(Ayar(pnl, k, SATIR_YAZIM).Value))
    DahilBoya pnl, k

    ' --- Kaynak / hedef dosyanin sayfalari ---
    Set kaynakWb = DosyaBul(Ayar(pnl, k, SATIR_KAYNAK_DOSYA).Value)
    SayfaListesiYaz pnl, k, lst, kaynakWb, True
    Set hedefWb = DosyaBul(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value)
    SayfaListesiYaz pnl, k, lst, hedefWb, False

    ' --- Profil listesi ---
    ProfilListesiYaz lst
    MenuKur Ayar(pnl, k, SATIR_PROFIL), lst, LST_PROFILLER

    ' --- Eslesme tablosunu temizle (notlar/oneriler dahil) ---
    AlanTemizle pnl.Range(pnl.Cells(ILK_SATIR, c), pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1))
    TabloCercevesi pnl, k, ILK_SATIR, SON_SATIR

    ' --- Kaynak sutunlari ---
    sonSut = 0
    Set kaynakWs = SayfaBul(kaynakWb, Ayar(pnl, k, SATIR_KAYNAK_SAYFA).Value)
    If Not kaynakWs Is Nothing Then
        sonSut = SonSutun(kaynakWs, EN_FAZLA_SUTUN)
        If sonSut > 0 Then
            basliklar = IkiBoyutlu(kaynakWs.Range(kaynakWs.Cells(1, 1), kaynakWs.Cells(1, sonSut)))
            ReDim yaz(1 To sonSut, 1 To 1)
            For i = 1 To sonSut
                yaz(i, 1) = HucreMetni(SutunEtiketi(i, GuvenliMetin(basliklar(1, i))))
            Next i
            With pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_KAYNAK), pnl.Cells(ILK_SATIR - 1 + sonSut, c + OFS_KAYNAK))
                .Value = yaz
                .Interior.Color = Palet(prKaynakZemin)
                .Font.Color = RGB(30, 41, 59)
            End With
            ayniKaynak = AyniAdlar(yaz)
            KaynakTara pnl, k, lst, kaynakWs, sonSut
        End If
    End If

    ' --- Hedef sutun listesi ve ayar menuleri ---
    Set hedefWs = SayfaBul(hedefWb, Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value)
    If Not hedefWs Is Nothing Then
        SutunlariYaz hedefWs, lst
        otomatikSecilen = KontrolSutunlariniBul(pnl, k, lst)
        MenuKur Ayar(pnl, k, SATIR_HASAR), lst, LST_HEDEF_SUTUNLAR
        MenuKur Ayar(pnl, k, SATIR_IHBAR), lst, LST_HEDEF_SUTUNLAR
        MenuKur Ayar(pnl, k, SATIR_TUTAR), lst, LST_HEDEF_SUTUNLAR
        If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" Then
            ayniHedef = AyniAdlar(IkiBoyutlu(lst.Range(lst.Cells(1, LST_HEDEF_SUTUNLAR), _
                        lst.Cells(lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row, LST_HEDEF_SUTUNLAR))))
        End If
    End If

    If Not hedefWs Is Nothing And sonSut > 0 Then
        pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(ILK_SATIR - 1 + sonSut, c + OFS_HEDEF)).Interior.Color = Palet(prHedefZemin)
        With pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_FILTRE), pnl.Cells(ILK_SATIR - 1 + sonSut, c + OFS_FILTRE))
            .Interior.Color = Palet(prFiltreZemin)
            .NumberFormat = "@"
        End With
        With pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_ISLEM), pnl.Cells(ILK_SATIR - 1 + sonSut, c + OFS_ISLEM))
            .Interior.Color = Palet(prIslemZemin)
            .HorizontalAlignment = xlCenter
        End With
        MenuKur pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_ISLEM), pnl.Cells(ILK_SATIR - 1 + sonSut, c + OFS_ISLEM)), lst, LST_ISLEM
    End If

    ' --- Sabit degerler kutusu ---
    SabitAlaniniDuzenle pnl, k, lst, hedefWs

    ' --- Eski eslesmeleri (ve onaylanmamis onerileri) geri koy ---
    If Not IsEmpty(eskiKaynak) Then
        EslesmeGeriYukle pnl, k, lst, eskiKaynak, eskiHedef, eskiFiltre, eskiIslem, eskiOneri
    End If

    ' --- Profil varsa uygula ---
    profilAdi = AyarMetni(pnl, k, SATIR_PROFIL)
    If profilAdi <> "" And sonSut > 0 And Not hedefWs Is Nothing Then
        profilBulundu = ProfilUygula(pnl, k, kaynakWs, hedefWs, profilAdi, _
                                     pTamam, pOneri, pEksik, pYeni, pHedefYok)
    End If

    ' --- Profil / ayni ad notlari (durum satirinda gorunur) ---
    notlar = ""
    If profilAdi <> "" And sonSut > 0 And Not hedefWs Is Nothing Then
        If Not profilBulundu Then
            notlar = "PROFIL '" & profilAdi & "' BULUNAMADI"
        Else
            notlar = "PROFIL: " & pTamam & " eslesti"
            If pOneri > 0 Then notlar = notlar & ", " & pOneri & " oneri"
            If pEksik <> "" Then notlar = notlar & "  |  PROFILDE VAR, DOSYADA YOK: " & pEksik
            If pYeni <> "" Then notlar = notlar & "  |  DOSYADA YENI: " & pYeni
            If pHedefYok <> "" Then notlar = notlar & "  |  ANA DOSYADA BULUNAMADI: " & pHedefYok
        End If
    End If
    If otomatikSecilen <> "" Then notlar = notlar & "  |  OTOMATIK SECILDI: " & otomatikSecilen
    If ayniKaynak <> "" Then notlar = notlar & "  |  AYNI ADLI KAYNAK SUTUN: " & ayniKaynak
    If ayniHedef <> "" Then notlar = notlar & "  |  ANA DOSYADA AYNI ADLI SUTUN: " & ayniHedef
    If Left$(notlar, 5) = "  |  " Then notlar = Mid$(notlar, 6)
    lst.Cells(LST_DURUM_NOT, LST_DURUM).Value = HucreMetni(Left$(notlar, 1500))

    ' --- Hedef menuleri, renkler, harita, tip, durum ---
    BlokMenuleriTazele pnl, k

    BlokBasligiYaz pnl, k, kaynakWb, kaynakWs
    durumMetni = DurumYaz(pnl, k)

    BlokYenile = "PANEL " & k & " (" & Left$(DosyaAdiGuvenli(kaynakWb), 28) & "): " & Left$(durumMetni, 110)

End Function


' Bir dosyanin sayfa adlarini Listeler'e yazar ve sayfa menusunu kurar
Public Sub SayfaListesiYaz(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                           ByVal wb As Workbook, ByVal kaynakMi As Boolean)

    Dim ws As Worksheet
    Dim sayfalar() As Variant
    Dim n As Long, lc As Long, satir As Long

    If kaynakMi Then
        lc = LST_KAYNAK_SAYFALAR
        satir = SATIR_KAYNAK_SAYFA
    Else
        lc = LST_HEDEF_SAYFALAR
        satir = SATIR_HEDEF_SAYFA
    End If

    lst.Columns(lc).ClearContents
    If wb Is Nothing Then
        On Error Resume Next
        Ayar(pnl, k, satir).Validation.Delete
        On Error GoTo 0
        Exit Sub
    End If

    ReDim sayfalar(0 To wb.Worksheets.Count - 1)
    For Each ws In wb.Worksheets
        sayfalar(n) = CStr(ws.Name)
        n = n + 1
    Next ws
    ' Sayi/tarih gibi gorunen sayfa adlari ("2026", "09.2026") metin kalsin
    ListeYaz lst, lc, sayfalar, True
    MenuKur Ayar(pnl, k, satir), lst, lc

End Sub


' Tutar / Hasar tarihi / Ihbar tarihi ayarlari BOSSA ana dosya
' basliklarindan TEK ve ACIK bir aday varsa onu secer. Kullanicinin, bir
' Profilin ya da Setin doldurdugu degere ASLA dokunmaz.
' Donus: secilenlerin kisa listesi (DURUM notu icin).
Public Function KontrolSutunlariniBul(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet) As String

    Dim metin As String, aday As String

    If AyarMetni(pnl, k, SATIR_TUTAR) = "" Then
        aday = SutunAdayi(lst, Array("tutar", "odenen", "muallak"), Empty, "tutar")
        If aday <> "" Then
            Ayar(pnl, k, SATIR_TUTAR).Value = HucreMetni(aday)
            metin = metin & "tutar=" & aday & ", "
        End If
    End If
    If AyarMetni(pnl, k, SATIR_HASAR) = "" Then
        aday = SutunAdayi(lst, Array("hasar"), "tarih", "")
        If aday <> "" Then
            Ayar(pnl, k, SATIR_HASAR).Value = HucreMetni(aday)
            metin = metin & "hasar=" & aday & ", "
        End If
    End If
    If AyarMetni(pnl, k, SATIR_IHBAR) = "" Then
        aday = SutunAdayi(lst, Array("ihbar"), "tarih", "")
        If aday <> "" Then
            Ayar(pnl, k, SATIR_IHBAR).Value = HucreMetni(aday)
            metin = metin & "ihbar=" & aday & ", "
        End If
    End If

    If Len(metin) > 2 Then metin = Left$(metin, Len(metin) - 2)
    KontrolSutunlariniBul = metin

End Function


' Sade adi (Normalize) anahtarlardan birini ve (verildiyse) "zorunlu"
' parcayi iceren hedef sutunu. Birden fazla aday varsa sade adi tam
' "tercih" olan secilir; o da yoksa belirsiz sayilir ve "" doner.
Private Function SutunAdayi(ByVal lst As Worksheet, ByVal anahtarlar As Variant, ByVal zorunlu As Variant, _
                            ByVal tercih As String) As String

    Dim son As Long, h As Long, t As Long, adet As Long
    Dim deger As String, ad As String, bulunan As String, tam As String
    Dim uyar As Boolean

    SutunAdayi = ""
    If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) = "" Then Exit Function
    son = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row

    For h = 1 To son
        deger = GuvenliMetin(lst.Cells(h, LST_HEDEF_SUTUNLAR).Value)
        ad = Normalize(SutunAdi(deger))
        If ad <> "" Then
            uyar = False
            For t = LBound(anahtarlar) To UBound(anahtarlar)
                If InStr(ad, anahtarlar(t)) > 0 Then uyar = True
            Next t
            If uyar And Not IsEmpty(zorunlu) Then uyar = (InStr(ad, CStr(zorunlu)) > 0)
            If uyar Then
                adet = adet + 1
                bulunan = deger
                If tercih <> "" And ad = tercih Then tam = deger
            End If
        End If
    Next h

    If adet = 1 Then
        SutunAdayi = bulunan
    ElseIf tam <> "" Then
        SutunAdayi = tam
    End If

End Function


' Ana dosya basliklarini "F - TUTAR" bicimiyle Listeler'e yazar
' (eski surumlerin gecici kontrol sutunu listeye girmez)
Public Sub SutunlariYaz(ByVal ws As Worksheet, ByVal lst As Worksheet)

    Dim sonSut As Long, i As Long, n As Long
    Dim v As Variant
    Dim yaz() As Variant

    lst.Columns(LST_HEDEF_SUTUNLAR).ClearContents
    sonSut = SonSutun(ws, EN_FAZLA_SUTUN)
    If sonSut < 1 Then Exit Sub

    v = IkiBoyutlu(ws.Range(ws.Cells(1, 1), ws.Cells(1, sonSut)))
    ReDim yaz(1 To sonSut, 1 To 1)
    For i = 1 To sonSut
        If Trim$(GuvenliMetin(v(1, i))) <> KONTROL_BASLIK Then
            n = n + 1
            yaz(n, 1) = HucreMetni(SutunEtiketi(i, GuvenliMetin(v(1, i))))
        End If
    Next i
    If n > 0 Then lst.Cells(1, LST_HEDEF_SUTUNLAR).Resize(sonSut, 1).Value = yaz

End Sub


' ============================================================
'  KAYNAK TARAMA  (tek gecis: veri tipi + filtre adaylari)
'
'  Her kaynak sutunu BIR KEZ ve en fazla FILTRE_TARAMA satir olarak
'  okunur. Ayni dizi hem veri tipi (ilk TIP_ORNEK dolu satir) hem de
'  filtre menusu (farkli degerler, FILTRE_SINIRI'na kadar) icin
'  kullanilir. Farkli deger cok fazlaysa (police no, tarih) menu
'  kurulmaz; filtre hucresine yine elle yazilabilir.
' ============================================================
Public Sub KaynakTara(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                      ByVal kaynakWs As Worksheet, ByVal sonSut As Long)

    Dim j As Long, i As Long, sonR As Long, okunan As Long
    Dim v As Variant
    Dim tipKisa() As Variant, tipUzun() As Variant
    Dim kisa As String, uzun As String
    Dim degerler As Sozluk
    Dim deger As String, anahtar As String
    Dim c As Long
    Dim asildi As Boolean

    c = BlokSutun(k)
    ReDim tipKisa(1 To sonSut, 1 To 1)
    ReDim tipUzun(1 To sonSut, 1 To 1)

    ' Listeler<k> BlokYenile basinda temizlendi; sutun sutun temizlenmez.
    For j = 1 To sonSut
        sonR = kaynakWs.Cells(kaynakWs.Rows.Count, j).End(xlUp).Row

        If sonR < 2 Then
            tipKisa(j, 1) = "BOS"
            tipUzun(j, 1) = "BOS"
        Else
            okunan = sonR
            If okunan > FILTRE_TARAMA + 1 Then okunan = FILTRE_TARAMA + 1
            v = SutunOku(kaynakWs, j, 2, okunan)

            TipAnalizDizi v, kisa, uzun
            tipKisa(j, 1) = kisa
            tipUzun(j, 1) = uzun

            ' --- filtre adaylari ---
            ' (Sozluk anahtari zaten buyuk/kucuk harf ayirmaz; filtre de
            ' BH ile karsilastirdigi icin ayrica BH gerekmez - hiz.)
            Set degerler = New Sozluk
            asildi = False
            For i = 1 To UBound(v, 1)
                deger = Trim$(GuvenliMetin(v(i, 1)))
                If deger <> "" Then
                    anahtar = deger
                    If Not degerler.Exists(anahtar) Then
                        If degerler.Count >= FILTRE_SINIRI Then
                            asildi = True
                            Exit For
                        End If
                        degerler.Item(anahtar) = deger
                    End If
                End If
            Next i

            If Not asildi And degerler.Count > 0 Then
                ListeYaz lst, LST_FILTRE_ILK + j - 1, degerler.Items, True
                If degerler.Count > 1 Then
                    lst.Range(lst.Cells(1, LST_FILTRE_ILK + j - 1), lst.Cells(degerler.Count, LST_FILTRE_ILK + j - 1)).Sort _
                        Key1:=lst.Cells(1, LST_FILTRE_ILK + j - 1), Order1:=xlAscending, Header:=xlNo
                End If
                MenuKurSerbest pnl.Cells(ILK_SATIR - 1 + j, c + OFS_FILTRE), lst, LST_FILTRE_ILK + j - 1
            End If
        End If
    Next j

    lst.Cells(1, LST_KAYNAK_TIP).Resize(sonSut, 1).Value = tipKisa
    lst.Cells(1, LST_KAYNAK_TIP_METIN).Resize(sonSut, 1).Value = tipUzun

End Sub


' Ekranda gosterilen kisa tip metni: "TARIH", "SAYI  %85" (azinlik
' varsa baskin tipin orani), "KARISIK", "BOS". Ornek deger tooltip
' yerine Listeler'de tutulmaz; hucre dar oldugu icin kisa tutulur.
Public Sub TipAnalizDizi(ByVal v As Variant, ByRef kisa As String, ByRef uzun As String)

    Dim tarihSay As Long, sayiSay As Long, metinSay As Long, toplam As Long, baskin As Long
    Dim ornek As String

    VeriTipiSay v, tarihSay, sayiSay, metinSay, ornek, TIP_ORNEK
    kisa = TipKarari(tarihSay, sayiSay, metinSay)
    toplam = tarihSay + sayiSay + metinSay

    If toplam = 0 Then
        uzun = "BOS"
        Exit Sub
    End If

    Select Case kisa
        Case "TARIH": baskin = tarihSay
        Case "SAYI":  baskin = sayiSay
        Case "METIN": baskin = metinSay
        Case Else:    baskin = toplam
    End Select

    uzun = kisa
    If baskin < toplam Then uzun = uzun & "  %" & Format$(baskin / toplam * 100, "0")

End Sub


' ============================================================
'  MENULER, RENKLER, HARITA, DURUM  (tek blok)
' ============================================================

Public Sub MenuleriTazele()

    Dim pnl As Worksheet

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    EkranKilitle
    On Error GoTo Cikis
    BlokMenuleriTazele pnl, AktifBlok()
Cikis:
    EkranAc

End Sub


' Kullanilmis hedef sutunlarini menuden duser; renkleri, haritayi,
' veri tipi sutununu ve DURUM satirini tazeler.
Public Sub BlokMenuleriTazele(ByVal pnl As Worksheet, ByVal k As Long)

    Dim lst As Worksheet
    Dim son As Long, hedefSon As Long, i As Long, n As Long, c As Long
    Dim kullanilan As Sozluk
    Dim deger As String
    Dim hedefWs As Worksheet
    Dim renkler() As Long, renkVar() As Boolean, sonSut As Long
    Dim hedefler As Variant, kalan() As Variant, panelHedef As Variant

    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)

    son = BlokSonSatir(pnl, k)
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    If son < ILK_SATIR Or Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) = "" Then
        TipSutunuYaz pnl, k, Nothing
        DurumYaz pnl, k
        Exit Sub
    End If

    ' Kullanilmis hedefler TAM DEGERLE ("F - TUTAR") tutulur; ana dosyada
    ' ayni adli iki sutun varsa biri secilince digeri menuden dusmez.
    Set kullanilan = New Sozluk
    panelHedef = IkiBoyutlu(pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(son, c + OFS_HEDEF)))
    For i = 1 To UBound(panelHedef, 1)
        deger = BH(Trim$(GuvenliMetin(panelHedef(i, 1))))
        If deger <> "" Then kullanilan.Item(deger) = True
    Next i

    hedefler = IkiBoyutlu(lst.Range(lst.Cells(1, LST_HEDEF_SUTUNLAR), lst.Cells(hedefSon, LST_HEDEF_SUTUNLAR)))
    ReDim kalan(0 To hedefSon)
    n = 0
    For i = 1 To hedefSon
        deger = BH(Trim$(GuvenliMetin(hedefler(i, 1))))
        If deger <> "" And Not kullanilan.Exists(deger) Then
            kalan(n) = GuvenliMetin(hedefler(i, 1))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        kalan(0) = "(kalan hedef sutun yok)"
        n = 1
    End If
    ReDim Preserve kalan(0 To n - 1)
    ListeYaz lst, LST_KALAN_HEDEF, kalan

    MenuKur pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(son, c + OFS_HEDEF)), lst, LST_KALAN_HEDEF

    Set hedefWs = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value), Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value)
    If Not hedefWs Is Nothing Then
        sonSut = BaslikRenkleriOnb(lst, hedefWs, renkler, renkVar)
        HedefRenkleriUygula pnl, k, renkler, renkVar, sonSut
        HaritayiYaz pnl, k, hedefWs, renkler, renkVar, sonSut
    End If

    TipSutunuYaz pnl, k, hedefWs
    DurumYaz pnl, k

End Sub


' Blogun DURUM satirini bastan yazar ve durum metnini dondurur.
Public Function DurumYaz(ByVal pnl As Worksheet, ByVal k As Long) As String

    Dim lst As Worksheet
    Dim c As Long, i As Long, son As Long
    Dim kaynakSay As Long, eslesen As Long, oneriSay As Long
    Dim yesilToplam As Long, yesilBos As Long, tipUyusmaz As Long
    Dim metin As String, notlar As String
    Dim renk As Long

    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)
    son = BlokSonSatir(pnl, k)

    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" Then
            kaynakSay = kaynakSay + 1
            If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) <> "" Then
                eslesen = eslesen + 1
                If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then oneriSay = oneriSay + 1
            End If
        End If
    Next i

    notlar = GuvenliMetin(lst.Cells(LST_DURUM_NOT, LST_DURUM).Value)
    yesilToplam = Val(GuvenliMetin(lst.Cells(LST_DURUM_YESIL_TOPLAM, LST_DURUM).Value))
    yesilBos = Val(GuvenliMetin(lst.Cells(LST_DURUM_YESIL_BOS, LST_DURUM).Value))
    tipUyusmaz = Val(GuvenliMetin(lst.Cells(LST_DURUM_TIP_UYUSMAZ, LST_DURUM).Value))

    renk = Palet(prDurumNotr)
    If kaynakSay = 0 Then
        metin = "EKSIK: kaynak dosya/sayfa secilmemis"
        renk = Palet(prDurumHata)
    ElseIf Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) = "" Then
        metin = "EKSIK: hedef dosya/sayfa secilmemis"
        renk = Palet(prDurumHata)
    ElseIf eslesen = 0 Then
        metin = "EKSIK: hic sutun eslestirilmemis"
        renk = Palet(prDurumUyari)
    Else
        metin = "HAZIR: " & eslesen & " / " & kaynakSay & " sutun eslesti"
        If DahilMi(pnl, k) Then
            metin = metin & "  |  AKTARIMA DAHIL"
            renk = Palet(prDurumHazir)
        Else
            metin = metin & "  |  aktarima dahil degil"
        End If
        If oneriSay > 0 Then
            metin = metin & "  |  ONAY BEKLEYEN ONERI: " & oneriSay
            renk = Palet(prDurumOneri)
        End If
        If yesilToplam > 0 Then
            metin = metin & "  |  YESIL: " & yesilToplam & " sutun, " & yesilBos & " tanesi BOSTA"
            If yesilBos > 0 Then renk = Palet(prDurumHata)
        End If
        If tipUyusmaz > 0 Then
            metin = metin & "  |  VERI TIPI UYUSMAZ: " & tipUyusmaz & " sutun"
            renk = Palet(prDurumHata)
        End If
    End If

    If notlar <> "" Then metin = metin & "  |  " & notlar

    pnl.Cells(SATIR_DURUM, c).Value = HucreMetni("DURUM: " & metin)
    pnl.Cells(SATIR_DURUM, c).Interior.Color = renk
    DurumYaz = metin

End Function


' Yenile sonrasi eski eslesmeleri geri koyar. Kaynak tarafi TAM degerle
' ("C - Pol No") eslenir. Hedef degeri yeni listede yoksa (ana dosyada
' sutun yer degistirmisse) ayni adli TEK sutun aranir; bulunamazsa
' eslesme bos birakilir - eski harf asla korunmaz.
Public Sub EslesmeGeriYukle(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                            ByVal eskiKaynak As Variant, ByVal eskiHedef As Variant, _
                            ByVal eskiFiltre As Variant, ByVal eskiIslem As Variant, _
                            eskiOneri() As String)

    Dim son As Long, i As Long, kk As Long, c As Long, h As Long, hedefSon As Long
    Dim mevcut As String, hedefDeger As String, yeniDeger As String, ad As String
    Dim listede As Sozluk, adla As Sozluk, eskiYer As Sozluk
    Dim islem As String

    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)

    Set listede = New Sozluk
    Set adla = New Sozluk
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" Then
        For h = 1 To hedefSon
            yeniDeger = GuvenliMetin(lst.Cells(h, LST_HEDEF_SUTUNLAR).Value)
            listede.Item(BH(Trim$(yeniDeger))) = yeniDeger
            ad = Normalize(SutunAdi(yeniDeger))
            If adla.Exists(ad) Then
                adla.Item(ad) = ""
            Else
                adla.Item(ad) = yeniDeger
            End If
        Next h
    End If

    ' Eski kaynak degerinden eski satira hizli erisim
    Set eskiYer = New Sozluk
    For kk = LBound(eskiKaynak, 1) To UBound(eskiKaynak, 1)
        mevcut = GuvenliMetin(eskiKaynak(kk, 1))
        If mevcut <> "" Then
            If Not eskiYer.Exists(mevcut) Then eskiYer.Item(mevcut) = kk
        End If
    Next kk

    For i = ILK_SATIR To son
        mevcut = GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)
        If mevcut <> "" Then
            If eskiYer.Exists(mevcut) Then
                kk = eskiYer.Item(mevcut)

                hedefDeger = GuvenliMetin(eskiHedef(kk, 1))
                If hedefDeger <> "" Then
                    yeniDeger = ""
                    If listede.Exists(BH(Trim$(hedefDeger))) Then
                        yeniDeger = listede.Item(BH(Trim$(hedefDeger)))
                    Else
                        ad = Normalize(SutunAdi(hedefDeger))
                        If adla.Exists(ad) Then yeniDeger = adla.Item(ad)
                    End If
                    If yeniDeger <> "" Then
                        pnl.Cells(i, c + OFS_HEDEF).Value = HucreMetni(yeniDeger)
                        If eskiOneri(kk) <> "" Then OneriIsaretle pnl.Cells(i, c + OFS_HEDEF), eskiOneri(kk)
                    End If
                End If

                If GuvenliMetin(eskiFiltre(kk, 1)) <> "" Then pnl.Cells(i, c + OFS_FILTRE).Value = eskiFiltre(kk, 1)
                islem = GuvenliMetin(eskiIslem(kk, 1))
                If islem <> "" Then
                    pnl.Cells(i, c + OFS_ISLEM).Value = IslemMetni(IslemSecili(islem), IslemFormul(islem))
                End If
            End If
        End If
    Next i

End Sub


' ============================================================
'  HUCRE DEGISIKLIGI YONLENDIRICI  (Panel sayfa kodundan)
'
'   - HEDEF hucresi elle secildi  -> o satirin onerisi onaylanir;
'                                    menuler/renkler/harita/durum tazelenir
'   - SABIT kutusu degisti        -> satirlar duzenlenir, DEGER menusu
'   - "Bu paneli aktar" degisti   -> renk ve DURUM
'   - Kaynak/hedef DOSYA secildi  -> o dosyanin sayfa menusu hemen kurulur
' ============================================================
Public Sub HedefDegisti(ByVal Target As Range)

    Dim pnl As Worksheet
    Dim alan As Range, parca As Range, hucre As Range, sutunParcasi As Range
    Dim bloklar As Sozluk, sabitBloklar As Sozluk
    Dim tamBloklar As Sozluk, satirlar As Sozluk, boyanacak As Sozluk
    Dim blok As Variant, parcalar() As String
    Dim ofs As Long, k As Long, col As Long, sonCol As Long, enSagCol As Long
    Dim hedefWs As Worksheet

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    If Not Target.Worksheet Is pnl Then Exit Sub
    If Not PanelDuzeniGuncel() Then Exit Sub
    If Target.Row > SON_SATIR Then Exit Sub

    EkranKilitle
    On Error GoTo Cikis

    Set bloklar = New Sozluk
    Set sabitBloklar = New Sozluk
    Set tamBloklar = New Sozluk
    Set satirlar = New Sozluk
    Set boyanacak = New Sozluk

    ' Sayfanin tamami silinse bile milyonlarca hucre dolasilmaz:
    ' yalnizca blok sutunlarina, SUTUN SUTUN bakilir.
    enSagCol = BlokSayisi(pnl) * BLOK_GENISLIK

    ' --- Eslesme tablosu: HEDEF hucresi elle secildi -> oneri onaylandi ---
    Set alan = Intersect(Target, pnl.Range(pnl.Rows(ILK_SATIR), pnl.Rows(SON_SATIR)))
    If Not alan Is Nothing Then
        For Each parca In alan.Areas
            sonCol = parca.Column + parca.Columns.Count - 1
            If sonCol > enSagCol Then sonCol = enSagCol
            For col = parca.Column To sonCol
                If ((col - 1) Mod BLOK_GENISLIK) = OFS_HEDEF Then
                    Set sutunParcasi = Intersect(parca, pnl.Columns(col))
                    If Not sutunParcasi Is Nothing Then
                        If sutunParcasi.Cells.Count <= 600 Then
                            For Each hucre In sutunParcasi.Cells
                                OneriKaldir hucre
                            Next hucre
                        End If
                    End If
                    bloklar.Item(CLng(Int((col - 1) / BLOK_GENISLIK) + 1)) = True
                End If
            Next col
        Next parca
    End If

    ' --- SABIT DEGERLER kutusu ---
    Set alan = Intersect(Target, pnl.Range(pnl.Rows(SABIT_ILK), pnl.Rows(SABIT_SON)))
    If Not alan Is Nothing Then
        For Each parca In alan.Areas
            sonCol = parca.Column + parca.Columns.Count - 1
            If sonCol > enSagCol Then sonCol = enSagCol
            For col = parca.Column To sonCol
                ofs = (col - 1) Mod BLOK_GENISLIK
                k = CLng(Int((col - 1) / BLOK_GENISLIK) + 1)
                If ofs = OFS_SABIT_SUTUN Or ofs = OFS_SABIT_DEGER Then
                    Set sutunParcasi = Intersect(parca, pnl.Columns(col))
                    If Not sutunParcasi Is Nothing Then
                        For Each hucre In sutunParcasi.Cells
                            SabitHucreDegisti pnl, k, hucre, ofs, tamBloklar, satirlar, boyanacak
                        Next hucre
                    End If
                End If
            Next col
        Next parca
    End If

    For Each blok In tamBloklar.Keys
        k = CLng(blok)
        Set hedefWs = BlokHedefSayfasi(pnl, k)
        SabitAlaniniDuzenle pnl, k, ListeSayfasi(k), hedefWs
        sabitBloklar.Item(k) = True
    Next blok

    For Each blok In satirlar.Keys
        parcalar = Split(CStr(blok), "|")
        k = CLng(parcalar(0))
        If Not tamBloklar.Exists(k) Then
            Set hedefWs = BlokHedefSayfasi(pnl, k)
            If SabitSatirGuncelle(pnl, k, ListeSayfasi(k), CLng(parcalar(1)), hedefWs) Then
                SabitAlaniniDuzenle pnl, k, ListeSayfasi(k), hedefWs
                tamBloklar.Item(k) = True
            End If
            sabitBloklar.Item(k) = True
        End If
    Next blok

    For Each blok In boyanacak.Keys
        parcalar = Split(CStr(blok), "|")
        k = CLng(parcalar(0))
        If Not tamBloklar.Exists(k) Then SabitSatirBoya pnl, k, CLng(parcalar(1))
    Next blok

    For Each blok In sabitBloklar.Keys
        If Not bloklar.Exists(blok) Then SabitSonrasiTazele pnl, CLng(blok)
    Next blok

    For Each blok In bloklar.Keys
        BlokMenuleriTazele pnl, CLng(blok)
    Next blok

    ' --- Ayar hucreleri (tek hucre degisikliklerinde) ---
    If Target.Cells.Count = 1 And Target.Row <= SATIR_AYAR_SON Then
        ofs = (Target.Column - 1) Mod BLOK_GENISLIK
        k = CLng(Int((Target.Column - 1) / BLOK_GENISLIK) + 1)
        If ofs = OFS_AYAR And k >= 1 And k <= BlokSayisi(pnl) Then
            Select Case Target.Row
                Case SATIR_DAHIL
                    DahilBoya pnl, k
                    DurumYaz pnl, k
                Case SATIR_KAYNAK_DOSYA
                    SayfaListesiYaz pnl, k, ListeSayfasi(k), DosyaBul(Target.Value), True
                    g_SonAktifBlok = 0
                    AktifPanelGoster pnl, k
                Case SATIR_HEDEF_DOSYA
                    SayfaListesiYaz pnl, k, ListeSayfasi(k), DosyaBul(Target.Value), False
            End Select
        End If
    End If

Cikis:
    EkranAc

End Sub


Public Function BlokHedefSayfasi(ByVal pnl As Worksheet, ByVal k As Long) As Worksheet
    Set BlokHedefSayfasi = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value), _
                                    Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value)
End Function


Public Function BlokKaynakSayfasi(ByVal pnl As Worksheet, ByVal k As Long) As Worksheet
    Set BlokKaynakSayfasi = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_KAYNAK_DOSYA).Value), _
                                     Ayar(pnl, k, SATIR_KAYNAK_SAYFA).Value)
End Function


' SEC isaretlerini kaldirir (FORMUL isaretleri kalir)
Public Sub IsaretleriTemizle()

    Dim pnl As Worksheet
    Dim k As Long, adet As Long

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    k = AktifBlok()

    EkranKilitle
    On Error GoTo Cikis
    adet = SecIsaretleriniKaldir(pnl, k)
Cikis:
    EkranAc

    MsgBox "PANEL " & k & ": " & adet & " SEC isareti kaldirildi." & vbNewLine & _
           "(FORMUL isaretleri korunur.)", vbInformation

End Sub


Public Function SecIsaretleriniKaldir(ByVal pnl As Worksheet, ByVal k As Long) As Long

    Dim c As Long, i As Long, son As Long, adet As Long
    Dim v As Variant

    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)
    For i = ILK_SATIR To son
        v = pnl.Cells(i, c + OFS_ISLEM).Value
        If IslemSecili(v) Then
            pnl.Cells(i, c + OFS_ISLEM).Value = IslemMetni(False, IslemFormul(v))
            adet = adet + 1
        End If
    Next i
    SecIsaretleriniKaldir = adet

End Function


' ============================================================
'  SABIT DEGERLER KUTUSU
'
'  Her satir: ana dosya sutunu (menu) + deger (menu: o sutunda ana
'  dosyada bulunan degerler; elle de yazilabilir). Doldurulan satirin
'  altinda bir bos satir acik durur; sutunu silinen satir kalkar ve
'  alttakiler yukari kayar. En fazla 15 satir.
' ============================================================

Public Sub SabitAlaniniDuzenle(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                               ByVal hedefWs As Worksheet)

    Dim c As Long, r As Long, t As Long, n As Long, h As Long, hedefSon As Long
    Dim sut(1 To SABIT_EN_FAZLA) As String, deg(1 To SABIT_EN_FAZLA) As Variant
    Dim s As String, ad As String, yeniDeger As String
    Dim listede As Sozluk, adla As Sozluk

    c = BlokSutun(k)

    ' --- Dolu satirlari sirasiyla topla (sutunu VEYA degeri dolu) ---
    n = 0
    For r = SABIT_ILK To SABIT_SON
        s = Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_SUTUN).Value))
        If s <> "" Or Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_DEGER).Value)) <> "" Then
            n = n + 1
            sut(n) = s
            deg(n) = pnl.Cells(r, c + OFS_SABIT_DEGER).Value
        End If
    Next r

    ' --- Ana dosyada sutun yer degistirmisse adla tekil ara ---
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" Then
        Set listede = New Sozluk
        Set adla = New Sozluk
        For h = 1 To hedefSon
            yeniDeger = GuvenliMetin(lst.Cells(h, LST_HEDEF_SUTUNLAR).Value)
            listede.Item(BH(Trim$(yeniDeger))) = yeniDeger
            ad = Normalize(SutunAdi(yeniDeger))
            If adla.Exists(ad) Then
                adla.Item(ad) = ""
            Else
                adla.Item(ad) = yeniDeger
            End If
        Next h
        For t = 1 To n
            If sut(t) = "" Then
                ' deger var, sutun henuz secilmemis
            ElseIf listede.Exists(BH(sut(t))) Then
                sut(t) = listede.Item(BH(sut(t)))
            Else
                ad = Normalize(SutunAdi(sut(t)))
                If adla.Exists(ad) Then
                    If adla.Item(ad) <> "" Then sut(t) = adla.Item(ad)
                End If
            End If
        Next t
    End If

    ' --- Alani yeniden yaz ---
    AlanTemizle pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_DEGER))

    For t = 1 To n
        r = SABIT_ILK + t - 1
        pnl.Cells(r, c + OFS_SABIT_SUTUN).Value = HucreMetni(sut(t))
        pnl.Cells(r, c + OFS_SABIT_DEGER).Value = deg(t)
        SabitSatirBoya pnl, k, r
    Next t

    If n < SABIT_EN_FAZLA Then SabitBosSatirBoya pnl, k, SABIT_ILK + n

    ' --- Menuler ---
    If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" Then
        If n < SABIT_EN_FAZLA Then
            MenuKur pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_ILK + n, c + OFS_SABIT_SUTUN)), lst, LST_HEDEF_SUTUNLAR
        Else
            MenuKur pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_SUTUN)), lst, LST_HEDEF_SUTUNLAR
        End If
    End If

    For t = 1 To n
        SabitDegerMenusu pnl, k, lst, SABIT_ILK + t - 1, hedefWs
    Next t

End Sub


' Sabit kutusunda tek hucre degisikligini siniflandirir:
'   sutun silindi             -> degeri de silinir, kutu tam duzenlenir
'   sutun secildi             -> sadece o satir guncellenir
'   deger yazildi (sutun var) -> sadece o satir boyanir
'   deger yazildi (sutun yok) -> kutu tam duzenlenir (yeni bos satir icin)
Public Sub SabitHucreDegisti(ByVal pnl As Worksheet, ByVal k As Long, ByVal hucre As Range, _
                             ByVal ofs As Long, ByVal tamBloklar As Sozluk, _
                             ByVal satirlar As Sozluk, ByVal boyanacak As Sozluk)

    Dim c As Long, r As Long

    c = BlokSutun(k)
    r = hucre.Row

    If ofs = OFS_SABIT_SUTUN Then
        If Trim$(GuvenliMetin(hucre.Value)) = "" Then
            pnl.Cells(r, c + OFS_SABIT_DEGER).ClearContents
            tamBloklar.Item(k) = True
        Else
            satirlar.Item(k & "|" & r) = True
        End If
    Else
        If Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_SUTUN).Value)) = "" Then
            tamBloklar.Item(k) = True
        Else
            boyanacak.Item(k & "|" & r) = True
        End If
    End If

End Sub


' Sutunu secilen tek bir sabit satirini gunceller. Ustunde bos satir
' varsa True dondurur: o zaman kutu tam duzenlenmelidir.
Public Function SabitSatirGuncelle(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                                   ByVal r As Long, ByVal hedefWs As Worksheet) As Boolean

    Dim c As Long, rr As Long

    c = BlokSutun(k)
    SabitSatirGuncelle = False

    For rr = SABIT_ILK To r - 1
        If Trim$(GuvenliMetin(pnl.Cells(rr, c + OFS_SABIT_SUTUN).Value)) = "" And _
           Trim$(GuvenliMetin(pnl.Cells(rr, c + OFS_SABIT_DEGER).Value)) = "" Then
            SabitSatirGuncelle = True
            Exit Function
        End If
    Next rr

    SabitSatirBoya pnl, k, r
    SabitDegerMenusu pnl, k, lst, r, hedefWs

    ' Altinda bir bos satir acik dursun
    If r < SABIT_SON Then
        If Trim$(GuvenliMetin(pnl.Cells(r + 1, c + OFS_SABIT_SUTUN).Value)) = "" And _
           Trim$(GuvenliMetin(pnl.Cells(r + 1, c + OFS_SABIT_DEGER).Value)) = "" Then
            SabitBosSatirBoya pnl, k, r + 1
            If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" Then
                MenuKur pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), _
                                  pnl.Cells(r + 1, c + OFS_SABIT_SUTUN)), lst, LST_HEDEF_SUTUNLAR
            End If
        End If
    End If

End Function


' Dolu bir sabit satirini boyar; eksik (bos) hucreyi kirmiziyla gosterir
Public Sub SabitSatirBoya(ByVal pnl As Worksheet, ByVal k As Long, ByVal r As Long)

    Dim c As Long

    c = BlokSutun(k)
    With pnl.Range(pnl.Cells(r, c + OFS_SABIT_SUTUN), pnl.Cells(r, c + OFS_SABIT_DEGER))
        .Interior.Color = Palet(prSabitDolu)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = Palet(prAyarKenar)
    End With
    If Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_DEGER).Value)) = "" Then
        pnl.Cells(r, c + OFS_SABIT_DEGER).Interior.Color = Palet(prSabitEksik)
    End If
    If Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_SUTUN).Value)) = "" Then
        pnl.Cells(r, c + OFS_SABIT_SUTUN).Interior.Color = Palet(prSabitEksik)
    End If

End Sub


Private Sub SabitBosSatirBoya(ByVal pnl As Worksheet, ByVal k As Long, ByVal r As Long)

    Dim c As Long

    c = BlokSutun(k)
    With pnl.Range(pnl.Cells(r, c + OFS_SABIT_SUTUN), pnl.Cells(r, c + OFS_SABIT_DEGER))
        .Interior.Color = Palet(prSabitBosSatir)
        .Borders(xlEdgeBottom).LineStyle = xlDash
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = Palet(prCerceve)
    End With

End Sub


' Sabit sutun degisince yalnizca hedef haritasi ve DURUM tazelenir
Public Sub SabitSonrasiTazele(ByVal pnl As Worksheet, ByVal k As Long)

    Dim lst As Worksheet, hedefWs As Worksheet
    Dim renkler() As Long, renkVar() As Boolean, sonSut As Long

    Set lst = ListeSayfasi(k)
    Set hedefWs = BlokHedefSayfasi(pnl, k)
    If Not hedefWs Is Nothing Then
        sonSut = BaslikRenkleriOnb(lst, hedefWs, renkler, renkVar)
        HaritayiYaz pnl, k, hedefWs, renkler, renkVar, sonSut
    End If
    DurumYaz pnl, k

End Sub


' Bir sabit satirinin DEGER menusu: ana dosyanin o sutununda son
' SABIT_TARAMA satirda bulunan farkli degerler. Onbellek (g_SabitOnbellek)
' sayesinde ayni sutun bir Yenile boyunca yeniden taranmaz.
Public Sub SabitDegerMenusu(ByVal pnl As Worksheet, ByVal k As Long, ByVal lst As Worksheet, _
                            ByVal satir As Long, ByVal hedefWs As Worksheet)

    Dim c As Long, hNo As Long, lc As Long
    Dim son As Long, ilk As Long, i As Long, adet As Long
    Dim v As Variant, degerler As Variant
    Dim sozluk As Sozluk
    Dim anahtar As String, onbellekAnahtari As String

    c = BlokSutun(k)
    lc = LST_SABIT_ILK + (satir - SABIT_ILK)
    lst.Columns(lc).ClearContents

    On Error Resume Next
    pnl.Cells(satir, c + OFS_SABIT_DEGER).Validation.Delete
    On Error GoTo 0

    If hedefWs Is Nothing Then Exit Sub
    hNo = SutunNo(pnl.Cells(satir, c + OFS_SABIT_SUTUN).Value)
    If hNo = 0 Then Exit Sub

    If g_SabitOnbellek Is Nothing Then Set g_SabitOnbellek = New Sozluk
    onbellekAnahtari = BH(hedefWs.Parent.Name & "|" & hedefWs.Name & "|" & hNo)

    If g_SabitOnbellek.Exists(onbellekAnahtari) Then
        degerler = g_SabitOnbellek.Item(onbellekAnahtari)
    Else
        son = hedefWs.Cells(hedefWs.Rows.Count, hNo).End(xlUp).Row
        If son < 2 Then Exit Sub
        ilk = son - SABIT_TARAMA + 1
        If ilk < 2 Then ilk = 2

        v = SutunOku(hedefWs, hNo, ilk, son)
        Set sozluk = New Sozluk
        For i = UBound(v, 1) To 1 Step -1
            anahtar = BH(Trim$(GuvenliMetin(v(i, 1))))
            If anahtar <> "" Then
                If Not sozluk.Exists(anahtar) Then
                    sozluk.Item(anahtar) = v(i, 1)
                    If sozluk.Count >= SABIT_MENU_SINIRI Then Exit For
                End If
            End If
        Next i
        If sozluk.Count = 0 Then Exit Sub
        ' Asil degerler (sayi sayi, tarih tarih olarak) saklanir
        degerler = sozluk.Items
        g_SabitOnbellek.Item(onbellekAnahtari) = degerler
    End If

    adet = UBound(degerler) - LBound(degerler) + 1
    If adet < 1 Then Exit Sub
    ListeYaz lst, lc, degerler
    If adet > 1 Then
        lst.Range(lst.Cells(1, lc), lst.Cells(adet, lc)).Sort _
            Key1:=lst.Cells(1, lc), Order1:=xlAscending, Header:=xlNo
    End If

    MenuKurSerbest pnl.Cells(satir, c + OFS_SABIT_DEGER), lst, lc

End Sub


' Paneldeki sabit degerlerin metin hali (Rapor ve Setler icin):
' her satir "F - DURUM" & TAB & deger, satirlar vbLf ile ayrilir
Public Function SabitlerMetni(ByVal pnl As Worksheet, ByVal k As Long) As String

    Dim c As Long, r As Long
    Dim metin As String

    c = BlokSutun(k)
    For r = SABIT_ILK To SABIT_SON
        If Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_SUTUN).Value)) <> "" Then
            If metin <> "" Then metin = metin & vbLf
            metin = metin & Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_SUTUN).Value)) & vbTab & _
                    GuvenliMetin(pnl.Cells(r, c + OFS_SABIT_DEGER).Value)
        End If
    Next r
    SabitlerMetni = metin

End Function


' SabitlerMetni bicimindeki metni kutuya yazar (Setler, gecis)
Public Sub SabitleriYaz(ByVal pnl As Worksheet, ByVal k As Long, ByVal sabitler As String)

    Dim c As Long, r As Long, t As Long
    Dim satirlar() As String, parca() As String

    c = BlokSutun(k)
    AlanTemizle pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_DEGER))
    If sabitler = "" Then Exit Sub

    satirlar = Split(sabitler, vbLf)
    r = SABIT_ILK
    For t = 0 To UBound(satirlar)
        parca = Split(satirlar(t), vbTab)
        If UBound(parca) >= 1 And r <= SABIT_SON Then
            pnl.Cells(r, c + OFS_SABIT_SUTUN).Value = HucreMetni(parca(0))
            pnl.Cells(r, c + OFS_SABIT_DEGER).Value = HucreMetni(parca(1))
            r = r + 1
        End If
    Next t

End Sub


' ============================================================
'  HEDEF SUTUN RENKLERI VE HARITA
'
'  Ana dosyadaki baslik hucrelerinin rengi panelde HEDEF hucresine ve
'  HEDEF HARITASI sutununa uygulanir. Yesil (zorunlu) olup eslesmemis
'  sutunlar "<< YESIL, BOSTA" ile isaretlenir.
' ============================================================

Public Sub HedefRenkleriUygula(ByVal pnl As Worksheet, ByVal k As Long, renkler() As Long, _
                               renkVar() As Boolean, ByVal sonSut As Long)

    Dim c As Long, son As Long, i As Long

    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" Then
            HedefHucreBoya pnl.Cells(i, c + OFS_HEDEF), renkler, renkVar, sonSut
        End If
    Next i

End Sub


Private Sub HedefHucreBoya(ByVal h As Range, renkler() As Long, renkVar() As Boolean, ByVal sonSut As Long)

    Dim hNo As Long

    If OneriMi(h) Then
        h.Interior.Color = Palet(prOneri)
        Exit Sub
    End If

    h.Interior.Color = Palet(prHedefZemin)
    hNo = SutunNo(h.Value)
    If hNo >= 1 And hNo <= sonSut Then
        If renkVar(hNo) Then h.Interior.Color = renkler(hNo)
    End If

End Sub


' Ana dosyanin baslik renklerini bir kere okur. Donus: sutun sayisi.
Public Function BaslikRenkleri(ByVal hedefWs As Worksheet, renkler() As Long, renkVar() As Boolean) As Long

    Dim sonSut As Long, i As Long, r As Long

    sonSut = SonSutun(hedefWs, EN_FAZLA_SUTUN)
    If sonSut < 1 Then sonSut = 1
    ReDim renkler(1 To sonSut)
    ReDim renkVar(1 To sonSut)

    For i = 1 To sonSut
        renkVar(i) = HucreRengi(hedefWs.Cells(1, i), r)
        renkler(i) = r
    Next i
    BaslikRenkleri = sonSut

End Function


' BaslikRenkleri'nin onbellekli hali (Listeler<k>, her Yenile'de sifirlanir)
Public Function BaslikRenkleriOnb(ByVal lst As Worksheet, ByVal hedefWs As Worksheet, _
                                  renkler() As Long, renkVar() As Boolean) As Long

    Dim n As Long, i As Long
    Dim v As Variant, w As Variant

    n = Val(GuvenliMetin(lst.Cells(1, LST_RENK + 2).Value))
    If n > 0 Then
        ReDim renkler(1 To n)
        ReDim renkVar(1 To n)
        v = IkiBoyutlu(lst.Range(lst.Cells(1, LST_RENK), lst.Cells(n, LST_RENK)))
        w = IkiBoyutlu(lst.Range(lst.Cells(1, LST_RENK + 1), lst.Cells(n, LST_RENK + 1)))
        For i = 1 To n
            renkler(i) = CLng(Val(GuvenliMetin(v(i, 1))))
            renkVar(i) = (GuvenliMetin(w(i, 1)) = "1")
        Next i
        BaslikRenkleriOnb = n
        Exit Function
    End If

    n = BaslikRenkleri(hedefWs, renkler, renkVar)
    If n > 0 Then
        ReDim v(1 To n, 1 To 1)
        ReDim w(1 To n, 1 To 1)
        For i = 1 To n
            v(i, 1) = renkler(i)
            w(i, 1) = IIf(renkVar(i), "1", "0")
        Next i
        lst.Cells(1, LST_RENK).Resize(n, 1).Value = v
        lst.Cells(1, LST_RENK + 1).Resize(n, 1).Value = w
        lst.Cells(1, LST_RENK + 2).Value = n
    End If
    BaslikRenkleriOnb = n

End Function


Public Sub HaritayiYaz(ByVal pnl As Worksheet, ByVal k As Long, ByVal hedefWs As Worksheet, _
                       renkler() As Long, renkVar() As Boolean, ByVal sonSut As Long)

    Dim c As Long, i As Long, son As Long, hNo As Long, r As Long
    Dim kullanilan As Sozluk
    Dim baslik As String
    Dim yesilBos As Long, yesilToplam As Long
    Dim hucre As Range
    Dim lst As Worksheet
    Dim basliklar As Variant

    If hedefWs Is Nothing Then Exit Sub

    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)
    Set lst = ListeSayfasi(k)

    ' Bu blokta hangi hedef sutunlari dolacak (sutun NUMARASIYLA)
    Set kullanilan = New Sozluk
    For i = ILK_SATIR To son
        hNo = SutunNo(pnl.Cells(i, c + OFS_HEDEF).Value)
        If hNo > 0 Then kullanilan.Item(hNo) = True
    Next i
    For i = SABIT_ILK To SABIT_SON
        hNo = SutunNo(pnl.Cells(i, c + OFS_SABIT_SUTUN).Value)
        If hNo > 0 Then kullanilan.Item(hNo) = True
    Next i

    With pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HARITA), pnl.Cells(SON_SATIR, c + OFS_HARITA))
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Font.ColorIndex = xlAutomatic
    End With

    If sonSut < 1 Then Exit Sub
    basliklar = IkiBoyutlu(hedefWs.Range(hedefWs.Cells(1, 1), hedefWs.Cells(1, sonSut)))

    r = ILK_SATIR - 1
    For i = 1 To sonSut
        baslik = GuvenliMetin(basliklar(1, i))
        If Trim$(baslik) <> KONTROL_BASLIK And r < SON_SATIR Then
            r = r + 1
            Set hucre = pnl.Cells(r, c + OFS_HARITA)

            If kullanilan.Exists(i) Then
                hucre.Value = HucreMetni("[+] " & SutunEtiketi(i, baslik))
            Else
                hucre.Value = HucreMetni(SutunEtiketi(i, baslik))
            End If

            If renkVar(i) Then
                hucre.Interior.Color = renkler(i)
                If YesilMi(renkler(i)) Then
                    yesilToplam = yesilToplam + 1
                    If Not kullanilan.Exists(i) Then
                        yesilBos = yesilBos + 1
                        hucre.Value = HucreMetni(SutunEtiketi(i, baslik) & "   << YESIL, BOSTA")
                        hucre.Font.Bold = True
                        hucre.Font.Color = RGB(255, 255, 255)
                        hucre.Interior.Color = Palet(prButonKirmizi)
                    End If
                End If
            End If
        End If
    Next i
    TabloCercevesi pnl, k, ILK_SATIR, SON_SATIR

    lst.Cells(LST_DURUM_YESIL_TOPLAM, LST_DURUM).Value = yesilToplam
    lst.Cells(LST_DURUM_YESIL_BOS, LST_DURUM).Value = yesilBos

End Sub


' Ana dosyada YESIL olup bu aktarimda dolmayacak sutunlar
Public Function YesilBosSutunlar(ByVal hedefWs As Worksheet, ByVal dolacak As Sozluk) As String

    Dim renkler() As Long, renkVar() As Boolean
    Dim sonSut As Long, i As Long
    Dim metin As String, baslik As String

    sonSut = BaslikRenkleri(hedefWs, renkler, renkVar)
    For i = 1 To sonSut
        If renkVar(i) Then
            If YesilMi(renkler(i)) And Not dolacak.Exists(i) Then
                baslik = GuvenliMetin(hedefWs.Cells(1, i).Value)
                If Trim$(baslik) <> KONTROL_BASLIK Then metin = metin & SutunHarfi(i) & " " & baslik & ", "
            End If
        End If
    Next i

    If Len(metin) > 2 Then metin = Left$(metin, Len(metin) - 2)
    YesilBosSutunlar = metin

End Function


' ============================================================
'  VERI TIPI SUTUNU
' ============================================================

Public Sub TipSutunuYaz(ByVal pnl As Worksheet, ByVal k As Long, ByVal hedefWs As Worksheet)

    Dim lst As Worksheet
    Dim c As Long, i As Long, r As Long, son As Long, hNo As Long, uyusmaz As Long
    Dim kTip As String, hTip As String, metin As String
    Dim h As Range

    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)
    son = BlokSonSatir(pnl, k)

    For i = ILK_SATIR To son
        Set h = pnl.Cells(i, c + OFS_TIP)
        r = i - ILK_SATIR + 1
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) = "" Then
            h.ClearContents
            h.Interior.ColorIndex = xlNone
        Else
            kTip = GuvenliMetin(lst.Cells(r, LST_KAYNAK_TIP).Value)
            metin = GuvenliMetin(lst.Cells(r, LST_KAYNAK_TIP_METIN).Value)
            h.Value = HucreMetni(metin)
            h.Font.Bold = False
            h.Font.Color = RGB(30, 41, 59)
            h.HorizontalAlignment = xlCenter
            h.Interior.Color = TipRengi(kTip, metin)

            hNo = SutunNo(pnl.Cells(i, c + OFS_HEDEF).Value)
            If hNo > 0 And Not hedefWs Is Nothing Then
                hTip = HedefTipOnbellek(lst, hedefWs, hNo)
                If TipCatisir(kTip, hTip) Then
                    h.Value = HucreMetni(kTip & " <> " & hTip)
                    h.Interior.Color = Palet(prTipUyusmaz)
                    h.Font.Color = Palet(prTipUyusmazYazi)
                    h.Font.Bold = True
                    uyusmaz = uyusmaz + 1
                End If
            End If
        End If
    Next i

    lst.Cells(LST_DURUM_TIP_UYUSMAZ, LST_DURUM).Value = uyusmaz

End Sub


Private Function TipRengi(ByVal tip As String, ByVal metin As String) As Long

    Select Case tip
        Case "TARIH":   TipRengi = RGB(219, 234, 254)
        Case "SAYI":    TipRengi = RGB(220, 252, 231)
        Case "METIN":   TipRengi = RGB(241, 245, 249)
        Case "KARISIK": TipRengi = RGB(254, 240, 138)
        Case Else:      TipRengi = RGB(255, 255, 255)
    End Select

    ' Baskin tip olsa da azinlik varsa (orn. %85 tarih) uyari tonu
    If InStr(metin, "%") > 0 And tip <> "KARISIK" Then TipRengi = RGB(254, 243, 199)

End Function


' Ana dosya sutununun tipi; Listeler'de saklanir, her Yenile'de sifirlanir
Public Function HedefTipOnbellek(ByVal lst As Worksheet, ByVal hedefWs As Worksheet, ByVal hNo As Long) As String

    Dim tip As String

    If hNo < 1 Or hNo > 16000 Then Exit Function
    tip = GuvenliMetin(lst.Cells(hNo, LST_HEDEF_TIP).Value)
    If tip = "" Then
        tip = VeriTipiHedef(hedefWs, hNo)
        If tip = "" Then tip = "?"
        lst.Cells(hNo, LST_HEDEF_TIP).Value = tip
    End If
    If tip = "?" Then tip = ""
    HedefTipOnbellek = tip

End Function


' ============================================================
'  CIFT TIKLAMA KISAYOLLARI  (Panel sayfa kodundan)
'
'   - Turuncu ONERI hucresine (HEDEF SUTUN) cift tik : oneri onaylanir
'   - ISLEM hucresine cift tik                      : SEC acilir/kapanir
'                                                     (FORMUL korunur)
'  Kisayol uygulandiysa Cancel = True (Excel hucre duzenlemeye girmez).
' ============================================================
Public Sub PanelCiftTik(ByVal Target As Range, ByRef Cancel As Boolean)

    Dim pnl As Worksheet
    Dim k As Long, ofs As Long, c As Long
    Dim v As Variant
    Dim kaynak As String, hedef As String

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    If Not Target.Worksheet Is pnl Then Exit Sub
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Row < ILK_SATIR Or Target.Row > SON_SATIR Then Exit Sub
    If Not PanelDuzeniGuncel() Then Exit Sub

    k = CLng(Int((Target.Column - 1) / BLOK_GENISLIK) + 1)
    If k < 1 Or k > BlokSayisi(pnl) Then Exit Sub
    ofs = (Target.Column - 1) Mod BLOK_GENISLIK
    c = BlokSutun(k)
    kaynak = Trim$(GuvenliMetin(pnl.Cells(Target.Row, c + OFS_KAYNAK).Value))
    If kaynak = "" Then Exit Sub

    If ofs = OFS_HEDEF Then
        If Not OneriMi(Target) Then Exit Sub
        Cancel = True
        hedef = SutunAdi(Target.Value)
        EkranKilitle
        On Error GoTo Cikis
        OneriKaldir Target
        BlokMenuleriTazele pnl, k
Cikis:
        EkranAc
        OlayYaz "ONERI ONAY", k, 0, 0, "Cift tikla onaylandi: " & SutunAdi(kaynak) & " -> " & hedef, "Onaylandi"
        DurumCubugu "PANEL " & k & ": oneri onaylandi  (" & SutunAdi(kaynak) & " -> " & hedef & ")"
        On Error Resume Next
        Application.OnTime Now + TimeValue("00:00:06"), "DurumCubuguTemizle"
        On Error GoTo 0

    ElseIf ofs = OFS_ISLEM Then
        Cancel = True
        v = Target.Value
        EkranKilitle
        Target.Value = IslemMetni(Not IslemSecili(v), IslemFormul(v))
        EkranAc
    End If

End Sub
