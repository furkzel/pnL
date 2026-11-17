Attribute VB_Name = "M15_Esik"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M15_Esik  (AYARLAR SAYFASI)
'
'   Esik(ek...)          : duzenlenebilir esigin etkin degeri. Ayarlar
'                          sayfasindaki deger bos / sayi degil / negatifse
'                          M01'deki varsayilan (ESV_*) kullanilir.
'   AyarlarSayfasiHazirla: sayfayi kurar; kullanicinin girdigi degerler
'                          KOD ile eslenerek korunur.
'   ZorunluMeta          : ZORUNLU SUTUNLAR tablosu (M10). Bir ana dosya
'                          sayfasi icin satir varsa zorunluluk BURADAN
'                          okunur; yoksa baslik rengi (yesil) varsayilan
'                          oneri olarak kalir.
'   EsikFarklari         : varsayilandan farkli esikler (Denetim'de gosterilir)
' ============================================================


Public Function EsikKodu(ByVal ek As EsikKod) As String
    Select Case ek
        Case ekTarihAltYil:       EsikKodu = "K10_TARIH_ALT_YIL"
        Case ekTarihUstAy:        EsikKodu = "K10_TARIH_UST_AY"
        Case ekZSkor:             EsikKodu = "K11_Z"
        Case ekZMinAy:            EsikKodu = "K11_Z_MIN_AY"
        Case ekBuyukHasar:        EsikKodu = "K11_BUYUK_HASAR"
        Case ekBenfordMad:        EsikKodu = "K11_BENFORD_MAD"
        Case ekBenfordMinN:       EsikKodu = "K11_BENFORD_MIN_N"
        Case ekK9Pencere:         EsikKodu = "K9_PENCERE"
        Case ekK9Ornek:           EsikKodu = "K9_ORNEK"
        Case ekK9KopyaEngel:      EsikKodu = "K9_KOPYA_ENGEL"
        Case ekK12Pencere:        EsikKodu = "K12_PENCERE"
        Case ekTamlikHedef:       EsikKodu = "KALITE_TAMLIK"
        Case ekKaliteHedef:       EsikKodu = "KALITE_HEDEF"
        Case ekYedekParti:        EsikKodu = "YEDEK_PARTI"
        Case ekKilitDakika:       EsikKodu = "KILIT_DAKIKA"
        Case ekMutabakatSutun:    EsikKodu = "K7_SUTUN"
        Case ekMutabakatTolerans: EsikKodu = "K7_TOLERANS"
        Case ekDosyaMB:           EsikKodu = "DOSYA_MB"
        Case ekAnalizGrup:        EsikKodu = "ANALIZ_GRUP"
        Case ekAnalizSatir:       EsikKodu = "ANALIZ_SATIR"
    End Select
End Function


Public Function EsikVarsayilan(ByVal ek As EsikKod) As Double
    Select Case ek
        Case ekTarihAltYil:       EsikVarsayilan = ESV_TARIH_ALT_YIL
        Case ekTarihUstAy:        EsikVarsayilan = ESV_TARIH_UST_AY
        Case ekZSkor:             EsikVarsayilan = ESV_Z_SKOR
        Case ekZMinAy:            EsikVarsayilan = ESV_Z_MIN_AY
        Case ekBuyukHasar:        EsikVarsayilan = ESV_BUYUK_HASAR
        Case ekBenfordMad:        EsikVarsayilan = ESV_BENFORD_MAD
        Case ekBenfordMinN:       EsikVarsayilan = ESV_BENFORD_MIN_N
        Case ekK9Pencere:         EsikVarsayilan = ESV_K9_PENCERE
        Case ekK9Ornek:           EsikVarsayilan = ESV_K9_ORNEK
        Case ekK9KopyaEngel:      EsikVarsayilan = ESV_K9_KOPYA_ENGEL
        Case ekK12Pencere:        EsikVarsayilan = ESV_K12_PENCERE
        Case ekTamlikHedef:       EsikVarsayilan = ESV_TAMLIK_HEDEF
        Case ekKaliteHedef:       EsikVarsayilan = ESV_KALITE_HEDEF
        Case ekYedekParti:        EsikVarsayilan = ESV_YEDEK_PARTI
        Case ekKilitDakika:       EsikVarsayilan = ESV_KILIT_DAKIKA
        Case ekMutabakatSutun:    EsikVarsayilan = ESV_MUTABAKAT_SUTUN
        Case ekMutabakatTolerans: EsikVarsayilan = ESV_MUTABAKAT_TOLERANS
        Case ekDosyaMB:           EsikVarsayilan = ESV_DOSYA_MB
        Case ekAnalizGrup:        EsikVarsayilan = ESV_ANALIZ_GRUP
        Case ekAnalizSatir:       EsikVarsayilan = ESV_ANALIZ_SATIR
    End Select
End Function


' Kullaniciya gorunen ad ve birim (hucre metni, TR)
Private Function EsikAdi(ByVal ek As EsikKod) As String
    Select Case ek
        Case ekTarihAltYil:       EsikAdi = TR("Makul tarih alt s#in#ir#i (y#il)")
        Case ekTarihUstAy:        EsikAdi = TR("Makul tarih #ust s#in#ir#i: d#onem sonu + ay")
        Case ekZSkor:             EsikAdi = TR("Z-skor e#si#gi (sat#ir ve tutar)")
        Case ekZMinAy:            EsikAdi = TR("Z-skor i#cin en az ge#cmi#s ay")
        Case ekBuyukHasar:        EsikAdi = TR("B#uy#uk hasar tutar#i (0 = kapal#i)")
        Case ekBenfordMad:        EsikAdi = TR("Benford ilk hane MAD e#si#gi (0 = kapal#i)")
        Case ekBenfordMinN:       EsikAdi = TR("Benford i#cin en az sat#ir")
        Case ekK9Pencere:         EsikAdi = TR("K9: ana dosyada taranan son sat#ir")
        Case ekK9Ornek:           EsikAdi = TR("K9: ana dosyada aranan #ornek sat#ir")
        Case ekK9KopyaEngel:      EsikAdi = TR("K9: birebir kopya sat#ir ENGEL (1) / D#IKKAT (0)")
        Case ekK12Pencere:        EsikAdi = TR("K12: kod kaymas#i i#cin taranan son sat#ir")
        Case ekTamlikHedef:       EsikAdi = TR("Taml#ik hedefi (%)")
        Case ekKaliteHedef:       EsikAdi = TR("Veri kalitesi skoru hedefi (0-100)")
        Case ekYedekParti:        EsikAdi = TR("Yede#gi saklanan son parti")
        Case ekKilitDakika:       EsikAdi = TR("Ana dosya kilidi bayatl#ik s#uresi (dk)")
        Case ekMutabakatSutun:    EsikAdi = TR("Mutabakat yap#ilan en fazla tutar s#utunu")
        Case ekMutabakatTolerans: EsikAdi = TR("Mutabakat tolerans#i (para birimi)")
        Case ekDosyaMB:           EsikAdi = TR("Dosya boyutu uyar#is#i (MB)")
        Case ekAnalizGrup:        EsikAdi = TR("Analiz: karma gruplama s#in#ir#i")
        Case ekAnalizSatir:       EsikAdi = TR("Analiz: tabloda en fazla sat#ir")
    End Select
End Function


Private Function EsikAciklama(ByVal ek As EsikKod) As String
    Select Case ek
        Case ekTarihAltYil, ekTarihUstAy
            EsikAciklama = TR("Bu aral#i#g#in d#i#s#indaki hasar/ihbar tarihleri K10'da D#IKKAT.")
        Case ekZSkor, ekZMinAy
            EsikAciklama = TR("Bu ay#in sat#ir/tutar#i ayn#i hedefin son 12 ay#ina g#ore |z| > e#sik ise K11 D#IKKAT.")
        Case ekBuyukHasar
            EsikAciklama = TR("Bu tutar#in #ust#undeki sat#irlar K11'de listelenir. Varsay#im: 10.000.000.")
        Case ekBenfordMad, ekBenfordMinN
            EsikAciklama = TR("Nigrini ilk hane uyumsuzluk s#in#ir#i 0,015.")
        Case ekK9Pencere, ekK9Ornek, ekK9KopyaEngel
            EsikAciklama = TR("Anahtar tekilli#gi ve m#ukerrer sat#ir.")
        Case ekK12Pencere
            EsikAciklama = TR("K#ir#il#im s#utununda ana dosyada g#or#ulmemi#s kodlar.")
        Case ekYedekParti
            EsikAciklama = TR("Daha eski partilerin yede#gi silinir; o partiler geri al#inamaz.")
        Case ekKilitDakika
            EsikAciklama = TR("Bu s#ureden eski kilit dosyas#i bayat say#il#ir.")
        Case ekMutabakatSutun, ekMutabakatTolerans
            EsikAciklama = TR("Ana tutar + say#isal e#slenen s#utunlar; Currency ile kuru#s hassasiyetinde.")
        Case Else
            EsikAciklama = ""
    End Select
End Function


Private Function AyarlarSayfasi() As Worksheet
    On Error Resume Next
    Set AyarlarSayfasi = ThisWorkbook.Worksheets(SAYFA_AYARLAR)
    On Error GoTo 0
End Function


' Etkin esik degeri (sayfa yoksa / deger gecersizse varsayilan)
Public Function Esik(ByVal ek As EsikKod) As Double

    Dim ws As Worksheet
    Dim x As Variant
    Dim d As Double

    Esik = EsikVarsayilan(ek)
    Set ws = AyarlarSayfasi()
    If ws Is Nothing Then Exit Function
    On Error GoTo Cikis
    If GuvenliMetin(ws.Cells(AYR_ILK + ek - 1, AYR_KOD).Value) <> EsikKodu(ek) Then Exit Function
    x = ws.Cells(AYR_ILK + ek - 1, AYR_DEGER).Value
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If Not SayiDegeri(x, d) Then Exit Function
    If d < 0 Then Exit Function
    Esik = d
Cikis:

End Function


Public Sub AyarlarSayfasiHazirla()

    Dim ws As Worksheet
    Dim eski As Sozluk
    Dim ek As Long, r As Long, sonR As Long
    Dim kod As String

    Set ws = SayfaOlustur(SAYFA_AYARLAR)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' --- Kullanicinin degerlerini KOD ile hatirla ---
    Set eski = New Sozluk
    sonR = ws.Cells(ws.Rows.Count, AYR_KOD).End(xlUp).Row
    For r = AYR_ILK To sonR
        If r >= AYR_Z_BASLIK Then Exit For
        kod = Trim$(GuvenliMetin(ws.Cells(r, AYR_KOD).Value))
        If kod <> "" And Not IsEmpty(ws.Cells(r, AYR_DEGER).Value) Then eski.Item(kod) = ws.Cells(r, AYR_DEGER).Value
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(AYR_Z_BASLIK - 1, AYR_ACIKLAMA + 1)).Clear
    ws.Cells.Font.Name = PANEL_YAZI_TIPI
    ws.Cells.Font.Size = YZ_GOVDE
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(AYR_KOD).ColumnWidth = 22
    ws.Columns(AYR_AD).ColumnWidth = 48
    ws.Columns(AYR_DEGER).ColumnWidth = 14
    ws.Columns(AYR_VARSAYILAN).ColumnWidth = 14
    ws.Columns(AYR_ACIKLAMA).ColumnWidth = 70

    With ws.Cells(2, AYR_KOD)
        .Value = TR("AYARLAR  #*  e#sikler ve zorunlu s#utunlar")
        .Font.Size = YZ_BASLIK
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
    End With
    With ws.Cells(3, AYR_KOD)
        .Value = TR("Yaln#iz DE#GER s#utunu d#uzenlenir. Bo#s b#irak#ilan de#ger varsay#ilana d#u#ser. " & _
                    "De#gi#siklikler Denetim'de 'varsay#ilandan farkl#i e#sikler' olarak listelenir.")
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
    End With

    ws.Cells(AYR_BASLIK_SATIR, AYR_KOD).Value = "KOD"
    ws.Cells(AYR_BASLIK_SATIR, AYR_AD).Value = TR("E#S#IK")
    ws.Cells(AYR_BASLIK_SATIR, AYR_DEGER).Value = TR("DE#GER")
    ws.Cells(AYR_BASLIK_SATIR, AYR_VARSAYILAN).Value = TR("VARSAYILAN")
    ws.Cells(AYR_BASLIK_SATIR, AYR_ACIKLAMA).Value = TR("A#CIKLAMA")
    TabloBasligiBicim ws.Range(ws.Cells(AYR_BASLIK_SATIR, AYR_KOD), ws.Cells(AYR_BASLIK_SATIR, AYR_ACIKLAMA))

    For ek = 1 To ESIK_SAYI
        r = AYR_ILK + ek - 1
        kod = EsikKodu(ek)
        ws.Cells(r, AYR_KOD).Value = kod
        ws.Cells(r, AYR_AD).Value = EsikAdi(ek)
        If eski.Exists(kod) Then
            ws.Cells(r, AYR_DEGER).Value = eski.Item(kod)
        End If
        ws.Cells(r, AYR_VARSAYILAN).Value = EsikVarsayilan(ek)
        ws.Cells(r, AYR_ACIKLAMA).Value = EsikAciklama(ek)
        ws.Range(ws.Cells(r, AYR_KOD), ws.Cells(r, AYR_ACIKLAMA)).Borders(xlEdgeBottom).Color = Palet(prCizgi)
        ws.Cells(r, AYR_DEGER).Interior.Color = Palet(prZemin)
        ws.Cells(r, AYR_DEGER).Borders.Color = Palet(prCizgi)
        ws.Cells(r, AYR_DEGER).Locked = False
    Next ek
    ws.Range(ws.Cells(AYR_ILK, AYR_KOD), ws.Cells(AYR_ILK + ESIK_SAYI - 1, AYR_KOD)).Font.Color = Palet(prIkincilMetin)
    ws.Range(ws.Cells(AYR_ILK, AYR_VARSAYILAN), ws.Cells(AYR_ILK + ESIK_SAYI - 1, AYR_VARSAYILAN)).Font.Color = Palet(prIkincilMetin)
    ws.Range(ws.Cells(AYR_ILK, AYR_ACIKLAMA), ws.Cells(AYR_ILK + ESIK_SAYI - 1, AYR_ACIKLAMA)).Font.Size = YZ_DIPNOT
    ws.Range(ws.Cells(AYR_ILK, AYR_DEGER), ws.Cells(AYR_ILK + ESIK_SAYI - 1, AYR_VARSAYILAN)).NumberFormat = "#,##0.###"

    ' --- ZORUNLU SUTUNLAR (meta veri; icerik korunur) ---
    With ws.Cells(AYR_Z_BASLIK, AYR_KOD)
        .Value = TR("ZORUNLU S#UTUNLAR")
        .Font.Size = YZ_BOLUM
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
    End With
    With ws.Cells(AYR_Z_BASLIK, AYR_AD)
        .Value = TR("Bir ana dosya sayfas#i i#cin sat#ir girilirse zorunluluk buradan okunur (K4, harita); " & _
                    "girilmezse ba#sl#ik rengi (ye#sil) varsay#ilan kabul edilir.")
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
    End With
    ws.Cells(AYR_Z_ILK - 1, AYR_Z_SAYFA).Value = TR("ANA DOSYA SAYFASI (bo#s = hepsi)")
    ws.Cells(AYR_Z_ILK - 1, AYR_Z_SUTUN).Value = TR("S#UTUN BA#SLI#GI")
    TabloBasligiBicim ws.Range(ws.Cells(AYR_Z_ILK - 1, AYR_Z_SAYFA), ws.Cells(AYR_Z_ILK - 1, AYR_Z_SUTUN))
    With ws.Range(ws.Cells(AYR_Z_ILK, AYR_Z_SAYFA), ws.Cells(AYR_Z_SON, AYR_Z_SUTUN))
        .Locked = False
        .NumberFormat = "@"
    End With

    ws.Tab.Color = Palet(prIkincilMetin)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = ws.Range(ws.Cells(2, AYR_KOD), ws.Cells(AYR_Z_ILK + 20, AYR_ACIKLAMA)).Address
    End With

End Sub


Public Sub TabloBasligiBicim(ByVal rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
        .Interior.Color = Palet(prTabloBaslik)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = Palet(prCerceve)
        .VerticalAlignment = xlCenter
    End With
End Sub


' Bir ana dosya sayfasi icin ZORUNLU sutun basliklari (Normalize edilmis).
' Bu sayfa icin satir yoksa Nothing (renk kurali gecerli).
Public Function ZorunluMeta(ByVal hedefWs As Worksheet) As Sozluk

    Dim ws As Worksheet
    Dim v As Variant
    Dim i As Long
    Dim sayfa As String, sutun As String
    Dim sonuc As Sozluk

    Set ZorunluMeta = Nothing
    If hedefWs Is Nothing Then Exit Function
    Set ws = AyarlarSayfasi()
    If ws Is Nothing Then Exit Function
    v = IkiBoyutlu(ws.Range(ws.Cells(AYR_Z_ILK, AYR_Z_SAYFA), ws.Cells(AYR_Z_SON, AYR_Z_SUTUN)))
    Set sonuc = New Sozluk
    For i = 1 To UBound(v, 1)
        sayfa = Trim$(GuvenliMetin(v(i, 1)))
        sutun = Normalize(GuvenliMetin(v(i, 2)))
        If sutun <> "" Then
            If sayfa = "" Or BH(sayfa) = BH(CStr(hedefWs.Name)) Then sonuc.Item(sutun) = True
        End If
    Next i
    If sonuc.Count > 0 Then Set ZorunluMeta = sonuc

End Function


' "KOD = deger (varsayilan v)" listesi; hepsi varsayilansa ""
Public Function EsikFarklari() As String
    Dim ek As Long
    Dim metin As String
    For ek = 1 To ESIK_SAYI
        If Abs(Esik(ek) - EsikVarsayilan(ek)) > 0.0000001 Then
            metin = metin & EsikKodu(ek) & " = " & Format$(Esik(ek), "#,##0.###") & _
                    " (" & Format$(EsikVarsayilan(ek), "#,##0.###") & "); "
        End If
    Next ek
    EsikFarklari = metin
End Function
