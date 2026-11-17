Attribute VB_Name = "M10_Denetim"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.0  -  M10_Denetim
'
' AKTUERYAL KONTROL NOKTALARI (K1..K14) ve YONETICI GORUNUMU
'
'   K1 DOSYA, BASLIK & HEDEF   dosya/sayfa erisimi, salt okunur kilit, baslik
'                              yeri/adi, iki kaynak ayni hedefe, hedef aralik
'   K2 MUKERRER AKTARIM        ayni kaynak bu ana dosyada zaten etkin mi
'   K3 SON SATIR & KUYRUK      cogunluk satiri, alt bolum (toplam) ayraci
'   K4 ZORUNLU (YESIL) SUTUN   eslesmeyen yesil sutun + yesil sutunda bos hucre
'   K5 VERI TIPI & SABIT DEGER kaynak/hedef tip catismasi, sabit deger tipi
'   K6 BICIM & KOD KORUMASI    bastaki sifir, 16+ hane, "=..." metin, ##### tasma
'   K7 TUTAR & MUTABAKAT       bozuk/bos/eksi tutar + 4'lu mutabakat (FARK 0,00)
'   K8 HASAR & IHBAR TARIHI    ihbar < hasar, bos / bozuk / metin tarih
'   K9 ANAHTAR TEKILLIGI       anahtar(+tutar+tarih) tekrari, birebir kopya, ana dosya ornegi
'   K10 DONEM TUTARLILIGI      donem sonundan sonra, makul aralik disi tarih
'   K11 MAKULLUK               z-skor (12 ay), buyuk hasar, Benford ilk hane
'   K12 SEMA & KOD KAYMASI     profil/son aktarim semasi, yeni/kaybolan kirilim kodu
'   K13 KATMANLI MUTABAKAT     kirilim x (adet, tutar sutunlari): kaynak = yazilan
'   K14 KAYNAK BUTUNLUGU & SOY On Kontrol -> Aktar kaynak izi, kaydedilmemis kaynak
'
' Durumlar (M01_Ayarlar.KD_*): "-" bekliyor, YOK uygulanmaz, OK, DIKKAT, ENGEL.
'
'   KarneCiz / KarneYaz / KarneHafif : Panel'deki canli KONTROL NOKTALARI
'                                      karnesi (her blokta c+5, satir 8-17)
'   HazirlikGuncelle / HazirlikGoster: seritteki HAZIRLIK gostergesi
'   DenetimOzeti                     : "Denetim" sayfasi (yonetici ozeti)
' ============================================================

Private Const LST_KARNE As Long = 344          ' Listeler<k>: 1..K_SAYI durum, +1 not, +2 kalite


' ============================================================
'  KONTROL NOKTASI ADLARI
' ============================================================

' Hucre/sekil icin Turkce ad
Public Function KontrolAdi(ByVal n As Long) As String
    Select Case n
        Case 1: KontrolAdi = TR("Dosya, ba#sl#ik & hedef")
        Case 2: KontrolAdi = TR("M#ukerrer aktar#im")
        Case 3: KontrolAdi = TR("Son sat#ir & kuyruk")
        Case 4: KontrolAdi = TR("Zorunlu (ye#sil) s#utun")
        Case 5: KontrolAdi = TR("Veri tipi & sabit de#ger")
        Case 6: KontrolAdi = TR("Bi#cim & kod korumas#i")
        Case 7: KontrolAdi = TR("Tutar & mutabakat")
        Case 8: KontrolAdi = TR("Hasar / ihbar tarihi")
        Case 9: KontrolAdi = TR("Anahtar tekilli#gi")
        Case 10: KontrolAdi = TR("D#onem tutarl#il#i#g#i")
        Case 11: KontrolAdi = TR("Makull#uk & ayk#ir#i de#ger")
        Case 12: KontrolAdi = TR("#Sema & kod kaymas#i")
        Case 13: KontrolAdi = TR("Katmanl#i mutabakat")
        Case 14: KontrolAdi = TR("Kaynak b#ut#unl#u#g#u & soy")
    End Select
End Function


' MsgBox / Rapor icin ASCII ad
Public Function KontrolAdiAscii(ByVal n As Long) As String
    Select Case n
        Case 1: KontrolAdiAscii = "DOSYA, BASLIK & HEDEF"
        Case 2: KontrolAdiAscii = "MUKERRER AKTARIM"
        Case 3: KontrolAdiAscii = "SON SATIR & KUYRUK"
        Case 4: KontrolAdiAscii = "ZORUNLU (YESIL) SUTUN"
        Case 5: KontrolAdiAscii = "VERI TIPI & SABIT DEGER"
        Case 6: KontrolAdiAscii = "BICIM & KOD KORUMASI"
        Case 7: KontrolAdiAscii = "TUTAR & MUTABAKAT"
        Case 8: KontrolAdiAscii = "HASAR & IHBAR TARIHI"
        Case 9: KontrolAdiAscii = "ANAHTAR TEKILLIGI"
        Case 10: KontrolAdiAscii = "DONEM TUTARLILIGI"
        Case 11: KontrolAdiAscii = "MAKULLUK & AYKIRI DEGER"
        Case 12: KontrolAdiAscii = "SEMA & KOD KAYMASI"
        Case 13: KontrolAdiAscii = "KATMANLI MUTABAKAT"
        Case 14: KontrolAdiAscii = "KAYNAK BUTUNLUGU & SOY"
    End Select
End Function


' Durumlarin tek satirlik ozeti (dialog icin):
'   "K1 OK  K2 OK  K3 DIKKAT  K4 OK  K5 DIKKAT  K6 OK  K7 OK  K8 -"
Public Function KarneSatiri(durum() As String) As String
    Dim n As Long
    Dim s As String
    For n = 1 To K_SAYI
        s = s & "K" & n & " " & durum(n) & IIf(n < K_SAYI, "  ", "")
        If n = 5 Or n = 10 Then s = s & vbNewLine
    Next n
    KarneSatiri = s
End Function


Public Function KSira(ByVal d As String) As Long
    Select Case d
        Case KD_YOK:    KSira = 1
        Case KD_OK:     KSira = 2
        Case KD_DIKKAT: KSira = 3
        Case KD_ENGEL:  KSira = 4
        Case Else:      KSira = 0
    End Select
End Function


' ============================================================
'  PANEL KARNESI
' ============================================================

Public Sub KarneCiz(ByVal pnl As Worksheet, ByVal k As Long)

    Dim c As Long, n As Long
    Dim durum() As String

    c = BlokSutun(k) + OFS_KARNE
    With pnl.Cells(KARNE_BASLIK, c)
        .Value = TR("KONTROL NOKTALARI")
        .Font.Bold = True
        .Font.Color = Palet(prZemin)
        .Interior.Color = Palet(prBlokBaslik)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
    On Error Resume Next
    pnl.Cells(KARNE_BASLIK, c).ClearComments
    pnl.Cells(KARNE_BASLIK, c).AddComment _
        "AKTUERYAL KONTROL NOKTALARI" & vbLf & _
        "K1 dosya, baslik, salt okunur, hedef aralik" & vbLf & _
        "K2 mukerrer aktarim" & vbLf & _
        "K3 son satir / toplam satiri" & vbLf & _
        "K4 zorunlu (yesil) sutun ve bos hucre" & vbLf & _
        "K5 veri tipi ve sabit deger tipi" & vbLf & _
        "K6 bastaki sifir, uzun numara, formul metni, #####" & vbLf & _
        "K7 tutar butunlugu ve 4'lu mutabakat" & vbLf & _
        "K8 hasar / ihbar tarih mantigi" & vbLf & _
        "K9 anahtar tekilligi, birebir kopya satir" & vbLf & _
        "K10 donem sonu ve makul tarih araligi" & vbLf & _
        "K11 z-skor, buyuk hasar, Benford" & vbLf & _
        "K12 sema ve kirilim kodu kaymasi" & vbLf & _
        "K13 kirilim bazinda adet ve tutar mutabakati" & vbLf & _
        "K14 On Kontrol -> Aktar kaynak izi ve soy" & vbLf & vbLf & _
        "K1, K4, K5: YENILE ve eslestirmede canli." & vbLf & _
        "Hepsi: 3 On Kontrol ve 4 AKTAR."
    pnl.Cells(KARNE_BASLIK, c).Comment.Visible = False
    pnl.Cells(KARNE_BASLIK, c).Comment.Shape.TextFrame.AutoSize = True
    On Error GoTo 0

    ReDim durum(1 To K_SAYI)
    For n = 1 To K_SAYI
        durum(n) = KD_BEKLIYOR
    Next n
    KarneHucreleri pnl, k, durum, TR("YEN#ILE ve #On Kontrol ile dolar"), -1

End Sub


' Karneyi yazar, Listeler'e saklar ve HAZIRLIK gostergesini tazeler
Public Sub KarneYaz(ByVal pnl As Worksheet, ByVal k As Long, durum() As String, ByVal notMetni As String, _
                    Optional ByVal kalite As Double = -1)

    Dim lst As Worksheet
    Dim n As Long

    Set lst = ListeSayfasi(k)
    For n = 1 To K_SAYI
        lst.Cells(n, LST_KARNE).Value = durum(n)
    Next n
    lst.Cells(K_SAYI + 1, LST_KARNE).Value = HucreMetni(notMetni)
    lst.Cells(K_SAYI + 2, LST_KARNE).Value = kalite
    KarneHucreleri pnl, k, durum, notMetni, kalite
    HazirlikGuncelle pnl

End Sub


Private Sub KarneHucreleri(ByVal pnl As Worksheet, ByVal k As Long, durum() As String, ByVal notMetni As String, _
                           ByVal kalite As Double)

    Dim c As Long, n As Long
    Dim isaret As String
    Dim zemin As Long, yazi As Long
    Dim dahil As Boolean
    Dim ok As Long, dikkat As Long, engel As Long, bekleyen As Long

    c = BlokSutun(k) + OFS_KARNE
    dahil = DahilMi(pnl, k)
    For n = 1 To K_SAYI
        If Not dahil Then
            isaret = ChrW(8211): zemin = Palet(prZemin): yazi = Palet(prPasifMetin)
        Else
            Select Case durum(n)
                Case KD_OK:     isaret = ChrW(10003): zemin = Palet(prZemin):       yazi = Palet(prOk)
                Case KD_DIKKAT: isaret = "!":         zemin = Palet(prDikkatZemin): yazi = Palet(prDikkat)
                Case KD_ENGEL:  isaret = ChrW(10005): zemin = Palet(prEngelZemin):  yazi = Palet(prEngel)
                Case KD_YOK:    isaret = ChrW(8211):  zemin = Palet(prZemin):       yazi = Palet(prPasifMetin)
                Case Else:      isaret = ChrW(9675):  zemin = Palet(prZemin):       yazi = Palet(prIkincilMetin)
            End Select
        End If
        Select Case durum(n)
            Case KD_OK:       ok = ok + 1
            Case KD_DIKKAT:   dikkat = dikkat + 1
            Case KD_ENGEL:    engel = engel + 1
            Case KD_BEKLIYOR: bekleyen = bekleyen + 1
        End Select
        With pnl.Cells(KARNE_ILK + n - 1, c)
            .Value = isaret & "   K" & n & "   " & KontrolAdi(n) & IIf(dahil, DurumEki(durum(n)), "")
            .Interior.Color = zemin
            .Font.Color = yazi
            .Font.Bold = (dahil And (durum(n) = KD_DIKKAT Or durum(n) = KD_ENGEL))
            .Font.Size = YZ_GOVDE
            .IndentLevel = 1
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlHairline
            .Borders(xlEdgeBottom).Color = Palet(prCizgi)
        End With
    Next n

    ' --- Baslik rozeti (PN-03): "11 OK . 2 ! . 0 X" (gliflerle) ---
    With pnl.Cells(KARNE_BASLIK, c)
        If dahil Then
            .Value = TR("KONTROL NOKTALARI   ") & ok & " " & ChrW(10003) & "  " & ChrW(183) & "  " & dikkat & " !  " & _
                     ChrW(183) & "  " & engel & " " & ChrW(10005) & IIf(bekleyen > 0, "  " & ChrW(183) & "  " & bekleyen & " " & ChrW(9675), "")
            If engel > 0 Then
                .Interior.Color = Palet(prEngelZemin): .Font.Color = Palet(prEngel)
            ElseIf dikkat > 0 Or bekleyen > 0 Then
                .Interior.Color = Palet(prDikkatZemin): .Font.Color = Palet(prDikkat)
            Else
                .Interior.Color = Palet(prOkZemin): .Font.Color = Palet(prOk)
            End If
        Else
            .Value = TR("KONTROL NOKTALARI   #-  aktar#ima dahil de#gil")
            .Interior.Color = Palet(prYuzey): .Font.Color = Palet(prPasifMetin)
        End If
        .Font.Bold = True
        .Font.Size = YZ_GOVDE
    End With

    ' --- Veri kalitesi skoru (DN-06) ---
    With pnl.Cells(KARNE_KALITE, c)
        .IndentLevel = 1
        .Font.Size = YZ_GOVDE
        .Font.Bold = True
        .Interior.Color = Palet(prZemin)
        If kalite < 0 Or Not dahil Then
            .Value = TR("Veri kalitesi   #-   (#On Kontrol ile hesaplan#ir)")
            .Font.Color = Palet(prPasifMetin)
            .Font.Bold = False
        Else
            .Value = TR("Veri kalitesi   ") & Format$(kalite, "0") & " / 100" & _
                     IIf(kalite >= Esik(ekKaliteHedef), "   " & ChrW(10003), TR("   hedef ") & Format$(Esik(ekKaliteHedef), "0"))
            .Font.Color = IIf(kalite >= Esik(ekKaliteHedef), Palet(prOk), Palet(prDikkat))
        End If
    End With

    With pnl.Cells(KARNE_NOT, c)
        .Value = HucreMetni(notMetni)
        .WrapText = True
        .EntireRow.RowHeight = 28
        .Font.Italic = True
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
        .IndentLevel = 1
    End With

End Sub


Private Function DurumEki(ByVal d As String) As String
    Select Case d
        Case KD_ENGEL:  DurumEki = TR("  #*  ENGEL")
        Case KD_YOK:    DurumEki = TR("  #*  yok")
        Case Else:      DurumEki = ""
    End Select
End Function


' Listeler'de saklanan son karne (Listeler bossa hepsi "-")
Public Function KarneKalite(ByVal k As Long) As Double
    Dim x As Variant
    x = ListeSayfasi(k).Cells(K_SAYI + 2, LST_KARNE).Value
    KarneKalite = -1
    If IsNumeric(x) And Not IsEmpty(x) Then KarneKalite = CDbl(x)
End Function


Public Function KarneOku(ByVal k As Long, durum() As String) As String

    Dim lst As Worksheet
    Dim n As Long

    Set lst = ListeSayfasi(k)
    ReDim durum(1 To K_SAYI)
    For n = 1 To K_SAYI
        durum(n) = Trim$(GuvenliMetin(lst.Cells(n, LST_KARNE).Value))
        If durum(n) = "" Then durum(n) = KD_BEKLIYOR
    Next n
    KarneOku = GuvenliMetin(lst.Cells(K_SAYI + 1, LST_KARNE).Value)

End Function


' YENILE ve eslestirme degisikliklerinde: aninda bilinen K1, K4, K5
' hesaplanir; aktarim verisi gerektirenler (K2, K3, K6, K7, K8) "-"
' (On Kontrol bekliyor) olur. Eski On Kontrol sonucu eslestirme degistigi
' icin artik gecerli degildir.
Public Sub KarneHafif(ByVal pnl As Worksheet, ByVal k As Long)

    Dim lst As Worksheet
    Dim durum() As String
    Dim c As Long, i As Long, son As Long, kaynakSay As Long, eslesen As Long, oneri As Long, n As Long

    Set lst = ListeSayfasi(k)
    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" Then
            kaynakSay = kaynakSay + 1
            If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) <> "" Then
                eslesen = eslesen + 1
                If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then oneri = oneri + 1
            End If
        End If
    Next i

    ReDim durum(1 To K_SAYI)
    For n = 1 To K_SAYI
        durum(n) = KD_BEKLIYOR
    Next n

    If kaynakSay = 0 Or Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) = "" Or eslesen = 0 Then
        durum(1) = KD_ENGEL
    ElseIf oneri > 0 Then
        durum(1) = KD_DIKKAT
    Else
        durum(1) = KD_OK
    End If

    If Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) <> "" And kaynakSay > 0 Then
        durum(4) = IIf(Val(GuvenliMetin(lst.Cells(LST_DURUM_YESIL_BOS, LST_DURUM).Value)) > 0, KD_DIKKAT, KD_OK)
        durum(5) = IIf(Val(GuvenliMetin(lst.Cells(LST_DURUM_TIP_UYUSMAZ, LST_DURUM).Value)) > 0, KD_DIKKAT, KD_OK)
    End If
    If AyarMetni(pnl, k, SATIR_TUTAR) = "" Then durum(7) = KD_YOK
    If AyarMetni(pnl, k, SATIR_HASAR) = "" And AyarMetni(pnl, k, SATIR_IHBAR) = "" Then
        durum(8) = KD_YOK
        durum(10) = KD_YOK
    End If
    If AyarMetni(pnl, k, SATIR_ANAHTAR) = "" Then durum(9) = KD_YOK
    If AyarMetni(pnl, k, SATIR_KIRILIM) = "" Then durum(13) = KD_YOK

    KarneYaz pnl, k, durum, TR("Canl#i: K1 K4 K5  #*  tam kontrol: 3 #On Kontrol  (") & Format$(Now, "hh:nn") & ")"

End Sub


' ============================================================
'  HAZIRLIK (serit gostergesi)
'
'  Aktarima dahil (Evet) her panel icin karneye bakar:
'    ENGEL varsa          -> engelli
'    K2/K3/K6/K7/K8 "-"   -> On Kontrol yapilmamis
'    DIKKAT varsa         -> dikkatli hazir
'    hepsi OK / YOK       -> hazir
' ============================================================
Public Sub HazirlikGuncelle(ByVal pnl As Worksheet)

    Dim shp As Object
    Dim k As Long, dahil As Long, hazir As Long, dikkat As Long, engel As Long, bekleyen As Long
    Dim metin As String, renk As Long

    On Error Resume Next
    Set shp = pnl.Shapes(HAZIRLIK_ADI)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    If Not PanelDuzeniGuncel() Then Exit Sub

    For k = 1 To BlokSayisi(pnl)
        If Not BlokGizliMi(pnl, k) And DahilMi(pnl, k) Then
            dahil = dahil + 1
            Select Case HazirlikDurumu(k)
                Case KD_ENGEL:    engel = engel + 1
                Case KD_BEKLIYOR: bekleyen = bekleyen + 1
                Case KD_DIKKAT:   dikkat = dikkat + 1
                Case Else:        hazir = hazir + 1
            End Select
        End If
    Next k

    ' BT-07: tek satir, durum zemini + ayni anlamda koyu metin
    If dahil = 0 Then
        metin = TR("HAZIRLIK   aktar#ima dahil panel yok  #*  'Bu paneli aktar' = Evet")
    Else
        metin = TR("HAZIRLIK   ") & (hazir + dikkat) & " / " & dahil & TR(" panel haz#ir")
        If dikkat > 0 Then metin = metin & "   " & ChrW(183) & "   " & dikkat & TR(" d#ikkatli (!)")
        If bekleyen > 0 Then metin = metin & "   " & ChrW(183) & "   " & bekleyen & TR(" #On Kontrol bekliyor")
        If engel > 0 Then metin = metin & "   " & ChrW(183) & "   " & engel & " ENGEL (" & ChrW(10005) & ")"
        If engel = 0 And bekleyen = 0 Then metin = metin & "   " & ChrW(183) & TR("   4 AKTAR'a haz#ir")
    End If

    On Error Resume Next
    AdimlariGuncelle pnl, (engel > 0 Or dahil = 0), (bekleyen > 0)
    shp.TextFrame2.TextRange.Text = metin
    If engel > 0 Then
        shp.Fill.ForeColor.RGB = Palet(prEngelZemin)
        shp.Line.ForeColor.RGB = Palet(prEngel)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Palet(prEngel)
    ElseIf bekleyen > 0 Or dikkat > 0 Then
        shp.Fill.ForeColor.RGB = Palet(prDikkatZemin)
        shp.Line.ForeColor.RGB = Palet(prDikkat)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Palet(prDikkat)
    ElseIf dahil > 0 Then
        shp.Fill.ForeColor.RGB = Palet(prOkZemin)
        shp.Line.ForeColor.RGB = Palet(prOk)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Palet(prOk)
    Else
        shp.Fill.ForeColor.RGB = Palet(prYuzey)
        shp.Line.ForeColor.RGB = Palet(prCizgi)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Palet(prIkincilMetin)
    End If
    On Error GoTo 0

End Sub


' Bir panelin hazirlik durumu: KD_ENGEL / KD_BEKLIYOR / KD_DIKKAT / KD_OK
Public Function HazirlikDurumu(ByVal k As Long) As String

    Dim durum() As String
    Dim n As Long, enKotu As Long

    KarneOku k, durum
    For n = 1 To K_SAYI
        If durum(n) = KD_ENGEL Then
            HazirlikDurumu = KD_ENGEL
            Exit Function
        End If
    Next n
    For n = 1 To K_SAYI
        If durum(n) = KD_BEKLIYOR Then
            HazirlikDurumu = KD_BEKLIYOR
            Exit Function
        End If
        If KSira(durum(n)) > enKotu Then enKotu = KSira(durum(n))
    Next n
    HazirlikDurumu = IIf(enKotu >= KSira(KD_DIKKAT), KD_DIKKAT, KD_OK)

End Function


' Aktar ve HAZIRLIK gostergesi icin tek panelin kisa etiketi
Public Function HazirlikEtiketi(ByVal k As Long) As String

    Dim durum() As String
    Dim n As Long
    Dim liste As String

    KarneOku k, durum
    Select Case HazirlikDurumu(k)
        Case KD_ENGEL
            For n = 1 To K_SAYI
                If durum(n) = KD_ENGEL Then liste = liste & "K" & n & " "
            Next n
            HazirlikEtiketi = "[ENGEL: " & Trim$(liste) & "]"
        Case KD_BEKLIYOR
            HazirlikEtiketi = "[ON KONTROL YAPILMADI]"
        Case KD_DIKKAT
            For n = 1 To K_SAYI
                If durum(n) = KD_DIKKAT Then liste = liste & "K" & n & " "
            Next n
            HazirlikEtiketi = "[HAZIR, DIKKAT: " & Trim$(liste) & "]"
        Case Else
            HazirlikEtiketi = "[HAZIR]"
    End Select

End Function


' HAZIRLIK gostergesine tiklaninca: panel panel durum
Public Sub HazirlikGoster()

    Dim pnl As Worksheet
    Dim k As Long
    Dim metin As String, notMetni As String
    Dim durum() As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub

    For k = 1 To BlokSayisi(pnl)
        If Not BlokGizliMi(pnl, k) Then
            notMetni = KarneOku(k, durum)
            metin = metin & "PANEL " & k & IIf(DahilMi(pnl, k), " (dahil)", " (dahil degil)") & "  " & _
                    Left$(AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA), 30) & vbNewLine & _
                    "   " & HazirlikEtiketi(k) & vbNewLine & _
                    "   " & Replace(KarneSatiri(durum), vbNewLine, vbNewLine & "   ") & vbNewLine
        End If
    Next k

    MsgBox "AKTARIM HAZIRLIGI" & vbNewLine & vbNewLine & Sigdir(metin, 800) & vbNewLine & _
           "OK = gecti, DIKKAT = okunup onaylanir, ENGEL = aktarim yapilamaz," & vbNewLine & _
           "- = On Kontrol bekliyor, YOK = bu panelde uygulanmaz.", vbInformation, "Hazirlik"

End Sub


' ============================================================
'  DENETIM SAYFASI  (yonetici / denetci ozeti)
'
'  Rapor ve Gecmis'ten uretilir; elle duzenlenmez. Her AKTAR ve Geri
'  Al'dan sonra kendiliginden, istenince "Denetim" butonuyla yenilenir.
'    - KPI kartlari: aktarim sayisi, etkin / geri alinan, aktarilan satir
'      ve tutar, mutabakat basari orani
'    - K1..K8 istatistigi: kac kez DIKKAT / ENGEL, en son ne zaman
'    - Son 15 aktarim: kaynak, hedef, satir, mutabakat, durum
' ============================================================
Public Sub DenetimOzeti()

    EkranKilitle
    On Error GoTo Cikis
    DenetimYaz
    EkranAc
    On Error Resume Next
    ThisWorkbook.Worksheets(SAYFA_DENETIM).Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    Exit Sub

Cikis:
    EkranAc
    MsgBox "Denetim sayfasi olusturulamadi: " & Err.Number & " - " & Err.Description, vbExclamation, "Denetim"

End Sub


' Sessiz yenileme (Aktar / Geri Al sonunda). Hata asil islemi etkilemez.
Public Sub DenetimSessiz()
    On Error Resume Next
    DenetimYaz
    On Error GoTo 0
End Sub


Private Sub DenetimYaz()

    Dim ws As Worksheet, rpr As Worksheet
    Dim v As Variant
    Dim aktifler As Sozluk, kDikkat As Sozluk, kEngel As Sozluk, kSon As Sozluk, gorulen As Sozluk
    Dim son As Long, i As Long, n As Long, r As Long, sayac As Long
    Dim toplamAktarim As Long, etkin As Long, geriAlinan As Long, tutuyor As Long, fark As Long
    Dim satir As Double, tutar As Double
    Dim no As String, kod As String, parcalar() As String, mut As String
    Dim kSay As Long
    Dim mutabakatOrani As Variant
    Dim aylikSatir(0 To DN_TREND_AY - 1) As Double, aylikTutar(0 To DN_TREND_AY - 1) As Double
    Dim ilkAy As Date, ay As Long

    ilkAy = DateSerial(Year(Date), Month(Date) - DN_TREND_AY + 1, 1)
    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    Set aktifler = AktifAktarimlar()
    Set kDikkat = New Sozluk
    Set kEngel = New Sozluk
    Set kSon = New Sozluk
    Set gorulen = New Sozluk

    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    If son >= 2 Then
        v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))
        For i = 1 To UBound(v, 1)
            If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" Then
                no = Trim$(GuvenliMetin(v(i, R_AKTARIM)))
                If Not gorulen.Exists(no) Then
                    gorulen.Item(no) = True
                    toplamAktarim = toplamAktarim + 1
                    If aktifler.Exists(no) Then
                        etkin = etkin + 1
                        satir = satir + GuvenliSayi(v(i, R_SATIR))
                        If IsDate(v(i, R_TARIH)) Then
                            ay = DateDiff("m", ilkAy, CDate(v(i, R_TARIH)))
                            If ay >= 0 And ay < DN_TREND_AY Then
                                aylikSatir(ay) = aylikSatir(ay) + GuvenliSayi(v(i, R_SATIR))
                                If IsNumeric(v(i, R_AYLIK_TOPLAM)) Then aylikTutar(ay) = aylikTutar(ay) + CDbl(v(i, R_AYLIK_TOPLAM))
                            End If
                        End If
                        If IsNumeric(v(i, R_AYLIK_TOPLAM)) And GuvenliMetin(v(i, R_AYLIK_TOPLAM)) <> "" Then
                            tutar = tutar + CDbl(v(i, R_AYLIK_TOPLAM))
                        End If
                    Else
                        geriAlinan = geriAlinan + 1
                    End If
                    mut = Left$(GuvenliMetin(v(i, R_MUTABAKAT)), 8)
                    If mut = "TUTUYOR:" Then tutuyor = tutuyor + 1
                    If mut = "FARK VAR" Then fark = fark + 1
                End If
            End If
            kod = Trim$(GuvenliMetin(v(i, R_KONTROL)))
            If kod <> "" Then
                parcalar = Split(kod, " ")
                If UBound(parcalar) >= 1 Then
                    If parcalar(1) = KD_ENGEL Then
                        kEngel.Artir parcalar(0)
                    ElseIf parcalar(1) = KD_DIKKAT Then
                        kDikkat.Artir parcalar(0)
                    End If
                    kSon.Item(parcalar(0)) = TarihMetni(v(i, R_TARIH))
                End If
            End If
        Next i
    End If

    Set ws = SayfaOlustur(SAYFA_DENETIM)
    On Error Resume Next
    ws.Cells.ClearOutline
    On Error GoTo 0
    ws.UsedRange.Clear
    ws.Cells.Font.Name = PANEL_YAZI_TIPI
    ws.Cells.Font.Size = YZ_GOVDE
    ws.Columns(1).ColumnWidth = 2
    ws.Range(ws.Columns(2), ws.Columns(13)).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 26
    ws.Columns(4).ColumnWidth = 22

    With ws.Cells(2, 2)
        .Value = TR("DENET#IM #OZET#I  #-  Aktar#im Paneli ") & SURUM
        .Font.Size = YZ_BASLIK
        .Font.Bold = True
        .Font.Color = Palet(prMarka)
    End With
    ws.Cells(3, 2).Value = TR("Rapor ve Ge#cmi#s kay#itlar#indan #uretildi. Elle d#uzenlemeyin.")
    If toplamAktarim = 0 Then ws.Cells(3, 2).Value = TR("Hen#uz aktar#im yok.")
    ws.Cells(3, 2).Font.Color = Palet(prIkincilMetin)

    ' --- KPI kartlari ---
    KpiKart ws, 5, 2, TR("Aktar#im"), toplamAktarim, TR("toplam kay#it")
    KpiKart ws, 5, 4, TR("Etkin"), etkin, TR("geri al#inmam#i#s")
    KpiKart ws, 5, 6, TR("Geri al#inan"), geriAlinan, TR("aktar#im")
    KpiKart ws, 5, 8, TR("Aktar#ilan sat#ir"), satir, TR("etkin aktar#imlar")
    KpiKart ws, 5, 10, TR("Aktar#ilan tutar"), tutar, TR("etkin, kaynak toplam#i")
    mutabakatOrani = Empty
    If tutuyor + fark > 0 Then mutabakatOrani = tutuyor / (tutuyor + fark)
    KpiKart ws, 5, 12, "Mutabakat", mutabakatOrani, _
            tutuyor & " TUTUYOR  " & ChrW(183) & "  " & fark & " FARK"
    ws.Cells(6, 12).NumberFormat = "0%"
    ws.Cells(6, 10).NumberFormat = "#,##0.00"
    If fark > 0 Then ws.Cells(6, 12).Font.Color = Palet(prButonKirmizi)

    ' --- Kontrol noktalari ---
    r = 10
    BaslikSatiri ws, r, Array("KONTROL NOKTASI", "", TR("D#IKKAT (adet)"), "ENGEL (adet)", TR("Son g#or#ulme"))
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 5)).HorizontalAlignment = xlRight
    For n = 1 To K_SAYI
        r = r + 1
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
        ws.Cells(r, 2).Value = "K" & n & "   " & KontrolAdi(n)
        ws.Cells(r, 4).Value = CLng(Val(GuvenliMetin(kDikkat.Item("K" & n))))
        ws.Cells(r, 5).Value = CLng(Val(GuvenliMetin(kEngel.Item("K" & n))))
        ws.Cells(r, 6).Value = HucreMetni(GuvenliMetin(kSon.Item("K" & n)))
        If ws.Cells(r, 5).Value > 0 Then ws.Cells(r, 5).Interior.Color = Palet(prDurumHata)
        If ws.Cells(r, 4).Value > 0 Then ws.Cells(r, 4).Interior.Color = Palet(prDurumUyari)
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Borders(xlEdgeBottom).Color = Palet(prAyarKenar)
    Next n

    ' --- Son 15 aktarim ---
    r = r + 3
    BaslikSatiri ws, r, Array("TARIH", TR("KAYNAK (ayl#ik)"), "HEDEF (ana dosya)", "PANEL", "PARTI", _
                              TR("AKTARIM NO"), TR("SATIR"), "AYLIK TOPLAM", "MUTABAKAT", "DURUM", TR("KULLANICI"))
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 9)).HorizontalAlignment = xlRight   ' sayi sutunlari
    If son >= 2 Then
        For i = UBound(v, 1) To 1 Step -1
            If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And sayac < 15 Then
                sayac = sayac + 1
                r = r + 1
                no = Trim$(GuvenliMetin(v(i, R_AKTARIM)))
                ws.Cells(r, 2).Value = TarihMetni(v(i, R_TARIH))
                ws.Cells(r, 3).Value = HucreMetni(GuvenliMetin(v(i, R_KAYNAK_DOSYA)) & " / " & GuvenliMetin(v(i, R_KAYNAK_SAYFA)))
                ws.Cells(r, 4).Value = HucreMetni(GuvenliMetin(v(i, R_HEDEF_DOSYA)))
                ws.Cells(r, 5).Value = v(i, R_PANEL)
                ws.Cells(r, 6).Value = v(i, R_PARTI)
                ws.Cells(r, 7).Value = v(i, R_AKTARIM)
                ws.Cells(r, 8).Value = v(i, R_SATIR)
                ws.Cells(r, 9).Value = v(i, R_AYLIK_TOPLAM)
                ws.Cells(r, 9).NumberFormat = "#,##0.00"
                mut = Left$(GuvenliMetin(v(i, R_MUTABAKAT)), 8)
                ws.Cells(r, 10).Value = IIf(mut = "TUTUYOR:", "TUTUYOR", IIf(mut = "FARK VAR", "FARK VAR", ChrW(8211)))
                If mut = "TUTUYOR:" Then ws.Cells(r, 10).Interior.Color = Palet(prDurumHazir)
                If mut = "FARK VAR" Then ws.Cells(r, 10).Interior.Color = Palet(prDurumHata)
                ws.Cells(r, 11).Value = IIf(aktifler.Exists(no), "ETKIN", "GERI ALINDI")
                If Not aktifler.Exists(no) Then ws.Range(ws.Cells(r, 2), ws.Cells(r, 12)).Font.Color = Palet(prIkincilMetin)
                ws.Cells(r, 12).Value = HucreMetni(GuvenliMetin(v(i, R_KULLANICI)))
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 12)).Borders(xlEdgeBottom).Color = Palet(prAyarKenar)
            End If
        Next i
    End If
    If sayac = 0 Then ws.Cells(r + 1, 2).Value = TR("Hen#uz aktar#im yok.")

    DenetimTrend ws, ilkAy, aylikSatir, aylikTutar
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = ws.Range(ws.Cells(2, 2), ws.Cells(r + 1, 13)).Address
        .LeftHeader = "&F"
        .RightHeader = "&D"
        .CenterFooter = "&P / &N"
    End With
    ws.Tab.Color = Palet(prMarka)

End Sub


Private Sub KpiKart(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal baslik As String, _
                    ByVal deger As Variant, ByVal alt As String)

    Dim rng As Range

    Set rng = ws.Range(ws.Cells(r, c), ws.Cells(r + 2, c + 1))
    rng.Interior.Color = Palet(prYuzey)
    KartCercevesi rng
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
        .Merge
        .Value = baslik
        .Font.Size = YZ_DIPNOT
        .Font.Bold = True
        .Font.Color = Palet(prIkincilMetin)
        .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + 1))
        .Merge
        .Value2 = deger
        .NumberFormat = "#,##0"
        .Font.Size = YZ_GOSTERGE
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
        .IndentLevel = 1
    End With
    ws.Rows(r + 1).RowHeight = 26
    With ws.Range(ws.Cells(r + 2, c), ws.Cells(r + 2, c + 1))
        .Merge
        .Value = HucreMetni(alt)
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
        .IndentLevel = 1
    End With

End Sub


Private Sub BaslikSatiri(ByVal ws As Worksheet, ByVal r As Long, ByVal basliklar As Variant)
    Dim i As Long
    For i = 0 To UBound(basliklar)
        ws.Cells(r, 2 + i).Value = basliklar(i)
    Next i
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 2 + UBound(basliklar)))
        .Font.Bold = True
        .Font.Color = Palet(prZemin)
        .Interior.Color = Palet(prBlokBaslik)
    End With
End Sub

Private Sub DenetimTrend(ByVal ws As Worksheet, ByVal ilkAy As Date, satir() As Double, tutar() As Double)
    Dim i As Long, r As Long, c As Long, grafik As ChartObject
    r = DN_TREND_SATIR: c = DN_TREND_SUTUN
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = DN_TREND_GRAFIK Then ws.ChartObjects(i).Delete
    Next i
    ws.Cells(r, c).Value = "AY"
    ws.Cells(r, c + 1).Value = "SATIR"
    ws.Cells(r, c + 2).Value = "TUTAR"
    For i = 0 To DN_TREND_AY - 1
        ws.Cells(r + i + 1, c).Value = DateAdd("m", i, ilkAy)
        ws.Cells(r + i + 1, c).NumberFormat = "mmm yy"
        ws.Cells(r + i + 1, c + 1).Value2 = satir(i)
        ws.Cells(r + i + 1, c + 2).Value2 = tutar(i)
    Next i
    Set grafik = ws.ChartObjects.Add(ws.Cells(r, DN_TREND_GRAFIK_SUTUN).Left, ws.Cells(r, DN_TREND_GRAFIK_SUTUN).Top, _
                  ws.Range(ws.Cells(r, DN_TREND_GRAFIK_SUTUN), ws.Cells(r, 13)).Width, ws.Rows(r).Height * 9)
    grafik.Name = DN_TREND_GRAFIK
    With grafik.Chart
        .ChartType = xlColumnClustered
        .SetSourceData ws.Range(ws.Cells(r, c), ws.Cells(r + DN_TREND_AY, c + 2))
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).AxisGroup = xlSecondary
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = Palet(prMarka)
        .SeriesCollection(2).Format.Line.ForeColor.RGB = Palet(prOneriMetin)
        .HasTitle = True
        .ChartTitle.Text = TR("Son 12 ay #* etkin aktar#imlar")
        .ChartArea.Font.Name = PANEL_YAZI_TIPI
        .ChartArea.Font.Size = YZ_DIPNOT
        .PlotVisibleOnly = False
        .Legend.Position = xlLegendPositionBottom
    End With
    ws.Range(ws.Columns(c), ws.Columns(c + 2)).Hidden = True
End Sub
