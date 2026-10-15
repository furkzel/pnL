Attribute VB_Name = "M03_Panel"
Option Explicit

' ============================================================
' AKTARIM PANELI  v34  -  M03_Panel
'
' Panel sayfasinin GEOMETRISI ve GORUNUMU:
'   - blok konumlari (BlokSutun, AktifBlok, Ayar ...)
'   - ust buton seridi (3 is akisi asamasi + gezinme)
'   - blok iskeletinin cizimi (BlokCiz)
'   - Panel Ekle / Panel Kaldir (gizli blok yeniden kullanimi)
'   - yapiskan gezinme (ButonlariEkranaTasi, PanelOnceki/Sonraki)
'
' BLOK DUZENI (c = blogun ilk sutunu, 6 sutun):
'   c+0 KAYNAK SUTUN   c+1 VERI TIPI   c+2 HEDEF SUTUN
'   c+3 HEDEF HARITASI c+4 FILTRE      c+5 ISLEM (SEC / FORMUL)
'   Ayar etiketleri c+0, ayar degerleri c+2 (Ayar fonksiyonu).
'   SABIT DEGERLER kutusu c+3 (ana dosya sutunu) / c+4 (deger).
'
' BUTON ADLARI: tum serit sekilleri "AP_" ile baslar; ButonlariEkranaTasi
' yalnizca bunlari tasir.
' ============================================================

Private Const ONEK_SEKIL As String = "AP_"
Private Const GOSTERGE_ADI As String = "AP_GOSTERGE"

' Office sabitleri (Mac / eski surumlerde ad cozumlemesi sorun olmasin)
Private Const SEKIL_YUVARLAK As Long = 5     ' msoShapeRoundedRectangle
Private Const METIN_YATAY As Long = 1        ' msoTextOrientationHorizontal
Private Const DIKEY_ORTA As Long = 3         ' msoAnchorMiddle
Private Const HIZA_ORTA As Long = 2          ' msoAlignCenter
Private Const HIZA_SOL As Long = 1           ' msoAlignLeft

' Buton stilleri
Private Const STIL_ANA As Long = 1
Private Const STIL_YESIL As Long = 2
Private Const STIL_KIRMIZI As Long = 3
Private Const STIL_IKINCIL As Long = 4
Private Const STIL_KUCUK As Long = 5         ' seyrek ve geri donussuz islem (Set Sil)


' ============================================================
'  GEOMETRI
' ============================================================

Public Function PanelSayfasi() As Worksheet
    On Error Resume Next
    Set PanelSayfasi = ThisWorkbook.Worksheets(SAYFA_PANEL)
    On Error GoTo 0
End Function


' k. blogun ilk sutun numarasi (1. blok = A, 2. blok = G ...)
Public Function BlokSutun(ByVal k As Long) As Long
    BlokSutun = 1 + (k - 1) * BLOK_GENISLIK
End Function


' Sayfadaki blok sayisi (baslik satirindaki "PANEL n" yazilarina gore;
' gizli bloklar dahil)
Public Function BlokSayisi(ByVal pnl As Worksheet) As Long

    Dim k As Long

    For k = 1 To MAX_BLOK
        If Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)).Value))), 5) <> "PANEL" Then
            BlokSayisi = k - 1
            Exit Function
        End If
    Next k
    BlokSayisi = MAX_BLOK

End Function


' Kaldirilmis (gizlenmis) blok mu?
Public Function BlokGizliMi(ByVal pnl As Worksheet, ByVal k As Long) As Boolean
    BlokGizliMi = pnl.Columns(BlokSutun(k)).Hidden
End Function


Public Function GorunurBlokSayisi(ByVal pnl As Worksheet) As Long
    Dim k As Long, n As Long
    For k = 1 To BlokSayisi(pnl)
        If Not BlokGizliMi(pnl, k) Then n = n + 1
    Next k
    GorunurBlokSayisi = n
End Function


' Imlecin bulundugu blok (panel disindaysa ilk gorunur blok)
Public Function AktifBlok() As Long

    Dim pnl As Worksheet
    Dim k As Long, adet As Long

    AktifBlok = 1
    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Function
    adet = BlokSayisi(pnl)
    If adet < 1 Then Exit Function

    k = 0
    On Error Resume Next
    If ActiveSheet.Name = pnl.Name Then k = Int((ActiveCell.Column - 1) / BLOK_GENISLIK) + 1
    On Error GoTo 0

    If k < 1 Then k = 1
    If k > adet Then k = adet
    If BlokGizliMi(pnl, k) Then k = KomsuBlok(pnl, k, 1)
    AktifBlok = k

End Function


' k'dan baslayarak yon (+1 / -1) dogrultusunda ilk gorunur blok;
' bulamazsa ilk gorunur blok.
Public Function KomsuBlok(ByVal pnl As Worksheet, ByVal k As Long, ByVal yon As Long) As Long

    Dim adet As Long, i As Long, aday As Long

    adet = BlokSayisi(pnl)
    KomsuBlok = k
    If adet < 1 Then Exit Function

    aday = k
    For i = 1 To adet
        aday = aday + yon
        If aday > adet Then aday = 1
        If aday < 1 Then aday = adet
        If Not BlokGizliMi(pnl, aday) Then
            KomsuBlok = aday
            Exit Function
        End If
    Next i

End Function


' Bir blogun ayar (deger) hucresi
Public Function Ayar(ByVal pnl As Worksheet, ByVal k As Long, ByVal satir As Long) As Range
    Set Ayar = pnl.Cells(satir, BlokSutun(k) + OFS_AYAR)
End Function


Public Function AyarMetni(ByVal pnl As Worksheet, ByVal k As Long, ByVal satir As Long) As String
    AyarMetni = Trim$(GuvenliMetin(Ayar(pnl, k, satir).Value))
End Function


' Eslesme tablosundaki son dolu satir (KAYNAK sutununa gore)
Public Function BlokSonSatir(ByVal pnl As Worksheet, ByVal k As Long) As Long

    Dim c As Long, r As Long

    c = BlokSutun(k) + OFS_KAYNAK
    r = pnl.Cells(SON_SATIR + 1, c).End(xlUp).Row
    If r < ILK_SATIR Then r = ILK_SATIR - 1
    If r > SON_SATIR Then r = SON_SATIR
    BlokSonSatir = r

End Function


' Her blogun kendi gizli liste sayfasi (menuler birbirine karismasin)
Public Function ListeSayfasi(ByVal k As Long) As Worksheet

    Dim ws As Worksheet

    Set ws = SayfaOlustur(SAYFA_LISTE_ONEK & k)
    If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
    Set ListeSayfasi = ws

End Function


' ------------------------------------------------------------
'  Hucre degerlerinin anlamlari (Turkce harfli menu metinleri BH ile
'  ASCII'ye katlanarak okunur; eski v33 degerleri de taninir)
' ------------------------------------------------------------

Public Function EvetMetni() As String
    EvetMetni = "Evet"
End Function


Public Function HayirMetni() As String
    HayirMetni = TR("Hay#ir")
End Function


Public Function DahilMi(ByVal pnl As Worksheet, ByVal k As Long) As Boolean
    DahilMi = (BH(AyarMetni(pnl, k, SATIR_DAHIL)) = "EVET")
End Function


Public Function YazimModuOku(ByVal deger As Variant) As YazimModu

    Dim v As String

    v = BH(Trim$(GuvenliMetin(deger)))
    If InStr(v, "GERI ALINAN") > 0 Then
        YazimModuOku = ymGeriAlinan
    ElseIf InStr(v, "SUTUN") > 0 Then
        YazimModuOku = ymSutunSonu
    Else
        YazimModuOku = ymSayfaSonu
    End If

End Function


Public Function YazimMetni(ByVal m As YazimModu) As String
    Select Case m
        Case ymSutunSonu:  YazimMetni = TR("S#utunun kendi sonu")
        Case ymGeriAlinan: YazimMetni = TR("Geri al#inan aral#i#ga")
        Case Else:         YazimMetni = TR("Sayfan#in en alt#i")
    End Select
End Function


' ISLEM sutunu: "SEC" (secili sutun yamasi / Geri Al kapsami),
' "FORMUL" (formul olarak tasi) ya da ikisi. v33'teki "Evet"/"X"
' isaretleri SEC sayilir.
Public Function IslemSecili(ByVal deger As Variant) As Boolean
    Dim v As String
    v = BH(Trim$(GuvenliMetin(deger)))
    IslemSecili = (InStr(v, "SEC") > 0 Or v = "X" Or v = "EVET")
End Function


Public Function IslemFormul(ByVal deger As Variant) As Boolean
    IslemFormul = (InStr(BH(Trim$(GuvenliMetin(deger))), "FORMUL") > 0)
End Function


Public Function IslemMetni(ByVal sec As Boolean, ByVal formul As Boolean) As String
    If sec And formul Then
        IslemMetni = TR("SE#C + FORM#UL")
    ElseIf sec Then
        IslemMetni = TR("SE#C")
    ElseIf formul Then
        IslemMetni = TR("FORM#UL")
    Else
        IslemMetni = ""
    End If
End Function


' Bir menu hucresine (Data Validation listesi) sabit secenekleri yazar
Public Sub SabitListeleriYaz(ByVal lst As Worksheet)

    ListeYaz lst, LST_EVET_HAYIR, Array(HayirMetni(), EvetMetni())
    ListeYaz lst, LST_YAZIM, Array(YazimMetni(ymSayfaSonu), YazimMetni(ymSutunSonu), YazimMetni(ymGeriAlinan))
    ListeYaz lst, LST_ISLEM, Array(IslemMetni(True, False), IslemMetni(False, True), IslemMetni(True, True))

End Sub


' ============================================================
'  BLOK CIZIMI
' ============================================================

Public Sub PanelSatirYukseklikleri(ByVal pnl As Worksheet)

    pnl.Rows(1).RowHeight = 6
    pnl.Rows(2).RowHeight = 15
    pnl.Rows(3).RowHeight = 28
    pnl.Rows(4).RowHeight = 28
    pnl.Rows(5).RowHeight = 8
    pnl.Rows(6).RowHeight = 8
    pnl.Rows(SATIR_BLOK_BASLIK).RowHeight = 26
    pnl.Range(pnl.Rows(SATIR_DAHIL), pnl.Rows(SABIT_NOT)).RowHeight = 18
    pnl.Rows(SATIR_KAYNAK_BOLUM).RowHeight = 20
    pnl.Rows(SATIR_HEDEF_BOLUM).RowHeight = 20
    pnl.Rows(SATIR_KONTROL_BOLUM).RowHeight = 20
    pnl.Rows(SABIT_NOT + 1).RowHeight = 6
    pnl.Rows(SATIR_DURUM).RowHeight = 24
    pnl.Rows(SATIR_DURUM + 1).RowHeight = 8
    pnl.Rows(BASLIK_SATIR).RowHeight = 30

End Sub


' Bir blogun iskeletini cizer (etiketler, basliklar, renkler, varsayilanlar).
' Icerigi (eslesme tablosu) silmez; cagiran gerekiyorsa once temizler.
Public Sub BlokCiz(ByVal pnl As Worksheet, ByVal k As Long)

    Dim c As Long, r As Long
    Dim ayarSatirlari As Variant, i As Long

    c = BlokSutun(k)

    ' --- Sutun genislikleri ---
    pnl.Columns(c + OFS_KAYNAK).ColumnWidth = 30
    pnl.Columns(c + OFS_TIP).ColumnWidth = 13
    pnl.Columns(c + OFS_HEDEF).ColumnWidth = 30
    pnl.Columns(c + OFS_HARITA).ColumnWidth = 30
    pnl.Columns(c + OFS_FILTRE).ColumnWidth = 18
    pnl.Columns(c + OFS_ISLEM).ColumnWidth = 12

    ' --- Blok basligi ---
    With pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c), pnl.Cells(SATIR_BLOK_BASLIK, c + BLOK_GENISLIK - 1))
        .UnMerge
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = Palet(prBlokBaslik)
    End With
    pnl.Cells(SATIR_BLOK_BASLIK, c).Value = "PANEL " & k

    ' --- Etiketler ---
    EtiketYaz pnl, SATIR_DAHIL, c, TR("Bu paneli aktar"), True
    BolumYaz pnl, SATIR_KAYNAK_BOLUM, c, TR("KAYNAK  #*  ayl#ik dosya")
    EtiketYaz pnl, SATIR_KAYNAK_DOSYA, c, "Dosya", False
    EtiketYaz pnl, SATIR_KAYNAK_SAYFA, c, "Sayfa", False
    EtiketYaz pnl, SATIR_SON_SATIR, c, TR("Son sat#ir  (bo#s = otomatik)"), False
    BolumYaz pnl, SATIR_HEDEF_BOLUM, c, TR("HEDEF  #*  ana dosya")
    EtiketYaz pnl, SATIR_HEDEF_DOSYA, c, "Dosya", False
    EtiketYaz pnl, SATIR_HEDEF_SAYFA, c, "Sayfa", False
    EtiketYaz pnl, SATIR_YAZIM, c, TR("Yaz#im ba#slang#ic#i"), False
    BolumYaz pnl, SATIR_KONTROL_BOLUM, c, TR("E#SLE#ST#IRME  #*  KONTROL")
    EtiketYaz pnl, SATIR_PROFIL, c, TR("Profil (#sirket)"), False
    EtiketYaz pnl, SATIR_TUTAR, c, TR("Tutar s#utunu (mutabakat)"), False
    EtiketYaz pnl, SATIR_HASAR, c, TR("Hasar tarihi s#utunu"), False
    EtiketYaz pnl, SATIR_IHBAR, c, TR("#Ihbar tarihi s#utunu"), False

    ' --- Ayar (deger) hucreleri ---
    ayarSatirlari = Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                          SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, _
                          SATIR_PROFIL, SATIR_TUTAR, SATIR_HASAR, SATIR_IHBAR)
    For i = 0 To UBound(ayarSatirlari)
        r = ayarSatirlari(i)
        With pnl.Cells(r, c + OFS_AYAR)
            .Interior.Color = Palet(prAyarZemin)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = Palet(prAyarKenar)
            .Font.Color = RGB(15, 23, 42)
        End With
    Next i
    pnl.Cells(SATIR_SON_SATIR, c + OFS_AYAR).HorizontalAlignment = xlLeft
    pnl.Cells(SATIR_SON_SATIR, c + OFS_AYAR).NumberFormat = "0"
    ' Dosya ve sayfa adlari metin kalsin ("2026", "09.2026" adli sayfalar
    ' menuden secildiginde sayiya/tarihe donusmesin)
    pnl.Cells(SATIR_KAYNAK_DOSYA, c + OFS_AYAR).NumberFormat = "@"
    pnl.Cells(SATIR_KAYNAK_SAYFA, c + OFS_AYAR).NumberFormat = "@"
    pnl.Cells(SATIR_HEDEF_DOSYA, c + OFS_AYAR).NumberFormat = "@"
    pnl.Cells(SATIR_HEDEF_SAYFA, c + OFS_AYAR).NumberFormat = "@"

    If AyarMetni(pnl, k, SATIR_DAHIL) = "" Then Ayar(pnl, k, SATIR_DAHIL).Value = HayirMetni()
    If AyarMetni(pnl, k, SATIR_YAZIM) = "" Then Ayar(pnl, k, SATIR_YAZIM).Value = YazimMetni(ymSayfaSonu)
    DahilBoya pnl, k

    ' --- SABIT DEGERLER kutusu ---
    SabitBasligiCiz pnl, k

    ' --- DURUM satiri ---
    With pnl.Range(pnl.Cells(SATIR_DURUM, c), pnl.Cells(SATIR_DURUM, c + BLOK_GENISLIK - 1))
        .UnMerge
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Interior.Color = Palet(prDurumNotr)
        .Borders(xlEdgeTop).Color = Palet(prAyarKenar)
        .Borders(xlEdgeBottom).Color = Palet(prAyarKenar)
    End With
    If Trim$(GuvenliMetin(pnl.Cells(SATIR_DURUM, c).Value)) = "" Then
        pnl.Cells(SATIR_DURUM, c).Value = TR("DURUM: hen#uz yenilenmedi  #-  dosyalar#i se#cip YEN#ILE'ye bas#in")
    End If

    ' --- Tablo basligi ---
    TabloBasligiCiz pnl, k
    TabloCercevesi pnl, k, ILK_SATIR, SON_SATIR

End Sub


Private Sub EtiketYaz(ByVal pnl As Worksheet, ByVal r As Long, ByVal c As Long, _
                      ByVal metin As String, ByVal kalin As Boolean)
    With pnl.Cells(r, c + OFS_ETIKET)
        .Value = metin
        .Font.Bold = kalin
        .Font.Color = Palet(prEtiketYazi)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
End Sub


Private Sub BolumYaz(ByVal pnl As Worksheet, ByVal r As Long, ByVal c As Long, ByVal metin As String)
    With pnl.Cells(r, c + OFS_ETIKET)
        .Value = metin
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = Palet(prBolumYazi)
        .VerticalAlignment = xlBottom
    End With
    With pnl.Range(pnl.Cells(r, c), pnl.Cells(r, c + OFS_AYAR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = Palet(prBlokBaslikAktif)
        .Weight = xlThin
    End With
End Sub


Public Sub SabitBasligiCiz(ByVal pnl As Worksheet, ByVal k As Long)

    Dim c As Long

    c = BlokSutun(k)
    pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_SUTUN).Value = TR("SAB#IT DE#GERLER")
    pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_DEGER).Value = TR("DE#GER")
    BaslikNotu pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_SUTUN), _
        "SABIT DEGERLER: aylik dosyada olmayan bilgiler" & vbLf & _
        "(RAPOR_DONEMI, KAYNAK, DOSYA_TURU ...)." & vbLf & _
        "Soldaki hucre: ana dosya sutunu (menu)." & vbLf & _
        "Sagdaki hucre: her satira yazilacak deger" & vbLf & _
        "(menu ana dosyadaki degerleri onerir; elle de yazilir)."
    With pnl.Range(pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_DEGER))
        .Font.Bold = True
        .Font.Color = RGB(6, 78, 59)
        .Interior.Color = Palet(prSabitBaslik)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With

    With pnl.Cells(SABIT_NOT, c + OFS_SABIT_SUTUN)
        .Value = TR("Sat#ir#i doldurunca yenisi a#c#il#ir (en fazla 15). S#utunu silince sat#ir kalkar.")
        .Font.Italic = True
        .Font.Size = 8
        .Font.Color = RGB(100, 116, 139)
    End With

End Sub


Public Sub TabloBasligiCiz(ByVal pnl As Worksheet, ByVal k As Long)

    Dim c As Long

    c = BlokSutun(k)
    pnl.Cells(BASLIK_SATIR, c + OFS_KAYNAK).Value = TR("KAYNAK S#UTUN  (ayl#ik)")
    pnl.Cells(BASLIK_SATIR, c + OFS_TIP).Value = TR("VER#I T#IP#I")
    pnl.Cells(BASLIK_SATIR, c + OFS_HEDEF).Value = TR("HEDEF S#UTUN  (ana dosya)")
    pnl.Cells(BASLIK_SATIR, c + OFS_HARITA).Value = TR("HEDEF HAR#ITASI")
    pnl.Cells(BASLIK_SATIR, c + OFS_FILTRE).Value = TR("F#ILTRE")
    pnl.Cells(BASLIK_SATIR, c + OFS_ISLEM).Value = TR("#I#SLEM")

    With pnl.Range(pnl.Cells(BASLIK_SATIR, c), pnl.Cells(BASLIK_SATIR, c + BLOK_GENISLIK - 1))
        .Font.Bold = True
        .Font.Color = RGB(30, 41, 59)
        .Interior.Color = Palet(prTabloBaslik)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = Palet(prCerceve)
    End With

    BaslikNotu pnl.Cells(BASLIK_SATIR, c + OFS_FILTRE), _
        "FILTRE ornekleri:" & vbLf & _
        "  AXA          tam deger" & vbLf & _
        "  AXA, ALLIANZ birden fazla deger" & vbLf & _
        "  AX*  ?X      joker (* ve ?)" & vbLf & _
        "  >100  <=500  sayi / tarih karsilastirma" & vbLf & _
        "  <>IPTAL      esit olmayan" & vbLf & _
        "  BOS / DOLU   bos ya da dolu hucre"
    BaslikNotu pnl.Cells(BASLIK_SATIR, c + OFS_ISLEM), _
        "ISLEM:" & vbLf & _
        "  SEC    : 'Secili Sutun Yamasi' bu sutunu tek basina aktarir;" & vbLf & _
        "           'Geri Al' yalnizca SEC'li sutunlari geri alir." & vbLf & _
        "  FORMUL : sutun formul olarak tasinir (satir referanslari kaydirilir)."
    BaslikNotu pnl.Cells(BASLIK_SATIR, c + OFS_TIP), _
        "Aylik sutunun veri tipi (ilk 300 satir)." & vbLf & _
        "Kirmizi: ana dosyadaki sutunun tipiyle celisiyor."
    BaslikNotu pnl.Cells(BASLIK_SATIR, c + OFS_HARITA), _
        "Ana dosyanin butun sutunlari, gercek baslik renkleriyle." & vbLf & _
        "[+] : bu aktarimda dolacak" & vbLf & _
        "<< YESIL, BOSTA : zorunlu (yesil) sutun eslesmemis"

End Sub


Private Sub BaslikNotu(ByVal h As Range, ByVal metin As String)
    On Error Resume Next
    h.ClearComments
    h.AddComment metin
    h.Comment.Visible = False
    h.Comment.Shape.TextFrame.AutoSize = True
    On Error GoTo 0
End Sub


' Tablo alanina ince izgara ve blok sonuna kalin ayrac cizer
Public Sub TabloCercevesi(ByVal pnl As Worksheet, ByVal k As Long, ByVal ilk As Long, ByVal son As Long)

    Dim c As Long

    c = BlokSutun(k)
    With pnl.Range(pnl.Cells(ilk, c), pnl.Cells(son, c + BLOK_GENISLIK - 1)).Borders
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = Palet(prAyarKenar)
    End With
    With pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c + BLOK_GENISLIK - 1), _
                   pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = Palet(prBlokBaslik)
    End With

End Sub


' "Bu paneli aktar" hucresinin rengi: Evet yesil, Hayir acik kirmizi
Public Sub DahilBoya(ByVal pnl As Worksheet, ByVal k As Long)
    With Ayar(pnl, k, SATIR_DAHIL)
        .Font.Bold = True
        If DahilMi(pnl, k) Then
            .Interior.Color = Palet(prDahilEvet)
        Else
            .Interior.Color = Palet(prDahilHayir)
        End If
    End With
End Sub


' Blok basligina kaynak dosya adini yazar. "PANEL n" on eki KALMALI:
' bloklar bu ek ile taniniyor (BlokSayisi).
Public Sub BlokBasligiYaz(ByVal pnl As Worksheet, ByVal k As Long, ByVal kaynakWb As Workbook, _
                          ByVal kaynakWs As Worksheet)

    Dim metin As String

    metin = "PANEL " & k
    If Not kaynakWs Is Nothing Then
        metin = metin & "   " & ChrW(8212) & "   " & kaynakWb.Name & "  /  " & CStr(kaynakWs.Name)
    ElseIf Not kaynakWb Is Nothing Then
        metin = metin & "   " & ChrW(8212) & "   " & kaynakWb.Name
    End If
    pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)).Value = HucreMetni(metin)

End Sub


' ============================================================
'  BUTON SERIDI
'
'  Satir 1-5 koyu bir serit olur ve dondurulur (asagi kaydirinca
'  kaybolmaz). Saga kaydirinca ButonlariEkranaTasi butonlari gorunen
'  alanin soluna getirir (Panel sayfasinda hucre secildikce).
'
'  1 HAZIRLA            : YENILE, + Panel, - Panel, Set Yukle/Kaydet/Sil
'  2 ESLESTIR & DOGRULA : Otomatik Eslestir, On Kontrol, Oneri Onayla
'                         (akilli), Profil Kaydet
'  3 AKTAR & GERI AL    : AKTAR, Secili Sutun Yamasi, Geri Al, Tutar Mutabakati
'  PANEL                : < Aktif panel gostergesi >, Kilavuz, Tani
'  Numarali 4 ana buton (1 YENILE, 2 Otomatik Eslestir, 3 On Kontrol,
'  4 AKTAR) yeni kullaniciya yolu gosterir.
' ============================================================

Public Sub ButonlariYenile()

    Dim pnl As Worksheet
    Dim hata As String

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then
        MsgBox "Panel sayfasi yok. Once 'SifirdanKur' calistir.", vbExclamation
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Cikis
    ButonlariKur pnl
Cikis:
    If Err.Number <> 0 Then hata = Err.Number & " - " & Err.Description
    EkranAc
    If hata <> "" Then
        MsgBox "Butonlar kurulamadi: " & hata, vbExclamation
    Else
        MsgBox "Butonlar yeniden olusturuldu.", vbInformation
    End If

End Sub


Public Sub ButonlariKur(ByVal pnl As Worksheet)

    Dim x As Double, x0 As Double, yA As Double, yB As Double, yE As Double
    Dim h As Double, ara As Double, grupAra As Double

    ' --- Eski serit (v33 form dugmeleri ve v34 sekilleri) ---
    On Error Resume Next
    pnl.Buttons.Delete
    On Error GoTo 0
    SekilleriSil pnl

    ' --- Serit zemini (hucre dolgusu: yatay kaydirmada da gorunur) ---
    PanelSatirYukseklikleri pnl
    With pnl.Range(pnl.Cells(1, 1), pnl.Cells(SERIT_SON_SATIR, TEMIZLIK_SON_SUTUN))
        .Interior.Color = Palet(prSeritZemin)
    End With

    yE = pnl.Rows(2).Top + 1
    yA = pnl.Rows(3).Top + 2
    yB = pnl.Rows(4).Top + 2
    h = 24
    ara = 6
    grupAra = 22

    ' --- Baslik ---
    x0 = 10
    MetinKutusu pnl, "AP_BASLIK", TR("AKTARIM PANEL#I"), x0, yA - 2, 128, 22, 13, Palet(prSeritYazi), True
    MetinKutusu pnl, "AP_SURUM", TR("e#sleme ile veri aktar#im#i  #*  ") & SURUM, x0, yB, 128, 18, 8, _
                Palet(prSeritGrupYazi), False

    ' 1-2-3-4 is akisi: dort ANA buton (YENILE, Otomatik Eslestir, On Kontrol,
    ' AKTAR) numarali, daha genis ve daha buyuk yazili. Seyrek kullanilan
    ' makrolar seritte yok ama Alt+F8 ile calisir: MenuleriTazele,
    ' IsaretleriTemizle, ProfilSil, ButonlariYenile, PaneliYenidenKur.

    ' --- 1. HAZIRLA ---
    x0 = x0 + 128 + grupAra
    GrupEtiketi pnl, 1, TR("1  #*  HAZIRLA"), x0, yE
    x = x0
    ButonEkle pnl, "Yenile", TR("1   YEN#ILE"), x, yA, 116, h, STIL_ANA: x = x + 116 + ara
    ButonEkle pnl, "YeniPanelEkle", "+ Panel", x, yA, 58, h, STIL_IKINCIL: x = x + 58 + ara
    ButonEkle pnl, "PanelKaldir", ChrW(8722) & " Panel", x, yA, 58, h, STIL_IKINCIL
    x = x0
    ButonEkle pnl, "SetYukle", TR("Set Y#ukle"), x, yB, 90, h, STIL_IKINCIL: x = x + 90 + ara
    ButonEkle pnl, "SetKaydet", "Set Kaydet", x, yB, 90, h, STIL_IKINCIL: x = x + 90 + ara
    ButonEkle pnl, "SetSil", "Set Sil", x, yB, 52, h, STIL_KUCUK
    x0 = x0 + 244 + grupAra
    AyracCiz pnl, 1, x0 - grupAra / 2, yE, yB + h

    ' --- 2. ESLESTIR & DOGRULA ---
    GrupEtiketi pnl, 2, TR("2  #*  E#SLE#ST#IR & DO#GRULA"), x0, yE
    x = x0
    ButonEkle pnl, "OtomatikEslestir", TR("2   Otomatik E#sle#stir"), x, yA, 150, h, STIL_ANA: x = x + 150 + ara
    ButonEkle pnl, "OnKontrol", TR("3   #On Kontrol"), x, yA, 124, h, STIL_ANA
    x = x0
    ButonEkle pnl, "OneriOnayla", ChrW(10003) & TR(" #Oneri Onayla"), x, yB, 150, h, STIL_IKINCIL: x = x + 150 + ara
    ButonEkle pnl, "ProfilKaydet", "Profil Kaydet", x, yB, 124, h, STIL_IKINCIL
    x0 = x0 + 280 + grupAra
    AyracCiz pnl, 2, x0 - grupAra / 2, yE, yB + h

    ' --- 3. AKTAR & GERI AL ---
    GrupEtiketi pnl, 3, TR("3  #*  AKTAR & GER#I AL"), x0, yE
    x = x0
    ButonEkle pnl, "Aktar", "4   " & ChrW(9654) & "  AKTAR", x, yA, 130, h, STIL_YESIL: x = x + 130 + ara
    ButonEkle pnl, "AktarSecili", TR("Se#cili S#utun Yamas#i"), x, yA, 140, h, STIL_IKINCIL
    x = x0
    ButonEkle pnl, "GeriAl", ChrW(8630) & TR("  Geri Al"), x, yB, 130, h, STIL_KIRMIZI: x = x + 130 + ara
    ButonEkle pnl, "TutarMutabakatiSon", TR("Tutar Mutabakat#i"), x, yB, 140, h, STIL_IKINCIL
    x0 = x0 + 276 + grupAra
    AyracCiz pnl, 3, x0 - grupAra / 2, yE, yB + h

    ' --- PANEL / YARDIM ---
    GrupEtiketi pnl, 4, TR("PANEL  #*  YARDIM"), x0, yE
    x = x0
    ButonEkle pnl, "PanelOnceki", ChrW(9664), x, yA, 26, h, STIL_IKINCIL: x = x + 26 + 4
    MetinKutusu pnl, GOSTERGE_ADI, "PANEL 1", x, yA, 170, h, 9, Palet(prSeritYazi), True
    With pnl.Shapes(GOSTERGE_ADI)
        .Fill.Visible = -1
        .Fill.ForeColor.RGB = Palet(prBlokBaslikAktif)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = HIZA_ORTA
        .TextFrame2.VerticalAnchor = DIKEY_ORTA
    End With
    x = x + 170 + 4
    ButonEkle pnl, "PanelSonraki", ChrW(9654), x, yA, 26, h, STIL_IKINCIL
    x = x0
    ButonEkle pnl, "Yardim", TR("K#ILAVUZ"), x, yB, 112, h, STIL_IKINCIL: x = x + 112 + ara
    ButonEkle pnl, "Tani", TR("Tan#i"), x, yB, 112, h, STIL_IKINCIL

    g_SonAktifBlok = 0
    AktifPanelGoster pnl, AktifBlok()

End Sub


' Seritteki (AP_ ile baslayan) butun sekilleri siler
Public Sub SekilleriSil(ByVal pnl As Worksheet)

    Dim i As Long

    On Error Resume Next
    For i = pnl.Shapes.Count To 1 Step -1
        If Left$(pnl.Shapes(i).Name, Len(ONEK_SEKIL)) = ONEK_SEKIL Then pnl.Shapes(i).Delete
    Next i
    On Error GoTo 0

End Sub


Private Sub ButonEkle(ByVal pnl As Worksheet, ByVal makro As String, ByVal yazi As String, _
                      ByVal sol As Double, ByVal ust As Double, ByVal g As Double, ByVal y As Double, _
                      ByVal stil As Long)

    Dim shp As Shape
    Dim zemin As Long, yaziRenk As Long, kenar As Long

    Select Case stil
        Case STIL_ANA:     zemin = Palet(prButonAna)
        Case STIL_YESIL:   zemin = Palet(prButonYesil)
        Case STIL_KIRMIZI: zemin = Palet(prButonKirmizi)
        Case STIL_KUCUK:   zemin = Palet(prSeritZemin)
        Case Else:         zemin = Palet(prButonIkincil)
    End Select
    yaziRenk = RGB(255, 255, 255)
    If stil = STIL_KUCUK Then yaziRenk = Palet(prSeritGrupYazi)
    kenar = Palet(prButonKenar)

    Set shp = pnl.Shapes.AddShape(SEKIL_YUVARLAK, sol, ust, g, y)
    shp.Name = ONEK_SEKIL & "BTN_" & makro
    shp.OnAction = makro
    On Error Resume Next
    shp.Adjustments.Item(1) = 0.22
    shp.Shadow.Visible = 0
    On Error GoTo 0
    shp.Fill.ForeColor.RGB = zemin
    If stil = STIL_IKINCIL Or stil = STIL_KUCUK Then
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = kenar
        shp.Line.Weight = 0.75
    Else
        shp.Line.Visible = 0
    End If

    With shp.TextFrame2
        .MarginLeft = 3
        .MarginRight = 3
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = 0
        .VerticalAnchor = DIKEY_ORTA
        .TextRange.Text = yazi
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        Select Case stil
            Case STIL_IKINCIL: .TextRange.Font.Size = 9
            Case STIL_KUCUK:   .TextRange.Font.Size = 8
            Case Else:         .TextRange.Font.Size = 10      ' 1-2-3-4 ana adimlar
        End Select
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = yaziRenk
        .TextRange.ParagraphFormat.Alignment = HIZA_ORTA
    End With
    shp.Placement = xlFreeFloating

End Sub


Private Sub MetinKutusu(ByVal pnl As Worksheet, ByVal ad As String, ByVal metin As String, _
                        ByVal sol As Double, ByVal ust As Double, ByVal g As Double, ByVal y As Double, _
                        ByVal boyut As Long, ByVal renk As Long, ByVal kalin As Boolean)

    Dim shp As Shape

    Set shp = pnl.Shapes.AddTextbox(METIN_YATAY, sol, ust, g, y)
    shp.Name = ad
    shp.Fill.Visible = 0
    shp.Line.Visible = 0
    With shp.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = 0
        .VerticalAnchor = DIKEY_ORTA
        .TextRange.Text = metin
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = boyut
        .TextRange.Font.Bold = IIf(kalin, -1, 0)
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = HIZA_SOL
    End With
    shp.Placement = xlFreeFloating

End Sub


Private Sub GrupEtiketi(ByVal pnl As Worksheet, ByVal n As Long, ByVal metin As String, _
                        ByVal sol As Double, ByVal ust As Double)
    MetinKutusu pnl, ONEK_SEKIL & "GRUP_" & n, metin, sol, ust, 300, 14, 8, Palet(prSeritGrupYazi), True
End Sub


Private Sub AyracCiz(ByVal pnl As Worksheet, ByVal n As Long, ByVal x As Double, _
                     ByVal ust As Double, ByVal alt As Double)

    Dim shp As Shape

    Set shp = pnl.Shapes.AddLine(x, ust + 2, x, alt)
    shp.Name = ONEK_SEKIL & "AYRAC_" & n
    shp.Line.ForeColor.RGB = Palet(prButonIkincil)
    shp.Line.Weight = 1
    shp.Placement = xlFreeFloating

End Sub


' Butonlari, gorunen alanin sol kenarina tasir (yalnizca yatay).
' Panel sayfasinin SelectionChange / Activate olayindan cagrilir.
Public Sub ButonlariEkranaTasi()

    Dim pnl As Worksheet
    Dim shp As Shape
    Dim hedefSol As Double, fark As Double, enSol As Double
    Dim ilk As Boolean

    On Error GoTo Cikis

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    If Not ActiveSheet Is pnl Then Exit Sub

    hedefSol = pnl.Columns(ActiveWindow.ScrollColumn).Left + 10

    ilk = True
    For Each shp In pnl.Shapes
        If Left$(shp.Name, Len(ONEK_SEKIL)) = ONEK_SEKIL Then
            If ilk Then
                enSol = shp.Left
                ilk = False
            ElseIf shp.Left < enSol Then
                enSol = shp.Left
            End If
        End If
    Next shp
    If ilk Then Exit Sub

    fark = hedefSol - enSol
    If Abs(fark) < 5 Then GoTo Gosterge

    Application.ScreenUpdating = False
    For Each shp In pnl.Shapes
        If Left$(shp.Name, Len(ONEK_SEKIL)) = ONEK_SEKIL Then shp.Left = shp.Left + fark
    Next shp
    Application.ScreenUpdating = True

Gosterge:
    AktifPanelGoster pnl, AktifBlok()

Cikis:
    On Error GoTo 0

End Sub


' Seritteki "AKTIF PANEL" gostergesini ve blok basliklarinin rengini
' gunceller. Blok degismediyse hicbir sey yapmaz (hizli).
Public Sub AktifPanelGoster(ByVal pnl As Worksheet, ByVal k As Long)

    Dim metin As String, kaynak As String

    On Error Resume Next
    If k = g_SonAktifBlok Then Exit Sub

    If g_SonAktifBlok >= 1 And g_SonAktifBlok <= BlokSayisi(pnl) Then
        pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(g_SonAktifBlok)).Interior.Color = Palet(prBlokBaslik)
    End If
    If k >= 1 Then pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)).Interior.Color = Palet(prBlokBaslikAktif)

    kaynak = AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA)
    metin = TR("AKT#IF:  PANEL ") & k
    If kaynak <> "" Then metin = metin & "  " & ChrW(183) & "  " & Left$(UzantisiSil(kaynak), 16)
    pnl.Shapes(GOSTERGE_ADI).TextFrame2.TextRange.Text = metin

    g_SonAktifBlok = k
    On Error GoTo 0

End Sub


Public Sub PanelOnceki()
    PaneleGit KomsuBlok(PanelSayfasi(), AktifBlok(), -1)
End Sub


Public Sub PanelSonraki()
    PaneleGit KomsuBlok(PanelSayfasi(), AktifBlok(), 1)
End Sub


Public Sub PaneleGit(ByVal k As Long)

    Dim pnl As Worksheet
    Dim olay As Boolean

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    If k < 1 Or k > BlokSayisi(pnl) Then Exit Sub

    On Error Resume Next
    olay = Application.EnableEvents
    pnl.Activate
    ActiveWindow.ScrollColumn = BlokSutun(k)
    Application.EnableEvents = False
    pnl.Cells(SATIR_KAYNAK_DOSYA, BlokSutun(k) + OFS_AYAR).Select
    Application.EnableEvents = olay
    ButonlariEkranaTasi
    On Error GoTo 0

End Sub


' Panel sayfasinin ust seridini dondurur (yalnizca satirlar; saga
' serbest kaydirma korunur).
Public Sub SeritDondur(ByVal pnl As Worksheet)

    Dim eski As Object

    On Error Resume Next
    Set eski = ActiveSheet
    pnl.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = SERIT_SON_SATIR
        .FreezePanes = True
        .DisplayGridlines = False
        .ScrollRow = 1
    End With
    If Not eski Is Nothing Then
        If Not eski Is pnl Then eski.Activate
    End If
    On Error GoTo 0

End Sub


' ============================================================
'  PANEL EKLE / KALDIR
'
'  Kaldirilan blok SILINMEZ: icerigi temizlenir ve sutunlari gizlenir;
'  boylece sagdaki panellerin numarasi ve setlerdeki panel numaralari
'  degismez. Panel Ekle once gizli bir blogu yeniden kullanir.
' ============================================================

Public Sub YeniPanelEkle()

    Dim pnl As Worksheet
    Dim k As Long, i As Long, ornek As Long

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then
        MsgBox "Panel sayfasi yok. Once 'SifirdanKur' calistir.", vbExclamation
        Exit Sub
    End If

    k = 0
    For i = 1 To BlokSayisi(pnl)
        If BlokGizliMi(pnl, i) Then
            k = i
            Exit For
        End If
    Next i

    If k = 0 Then
        k = BlokSayisi(pnl) + 1
        If k > MAX_BLOK Then
            MsgBox "En fazla " & MAX_BLOK & " panel blogu desteklenir.", vbExclamation
            Exit Sub
        End If
    End If

    ' Hedef ayarlari ilk gorunur bloktan kopyalanir (genelde ayni ana dosya)
    ornek = 0
    For i = 1 To BlokSayisi(pnl)
        If i <> k And Not BlokGizliMi(pnl, i) Then
            ornek = i
            Exit For
        End If
    Next i

    EkranKilitle
    On Error GoTo Hata

    pnl.Range(pnl.Columns(BlokSutun(k)), pnl.Columns(BlokSutun(k) + BLOK_GENISLIK - 1)).Hidden = False
    AlanTemizle pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)), _
                          pnl.Cells(SON_SATIR, BlokSutun(k) + BLOK_GENISLIK - 1))
    BlokCiz pnl, k

    If ornek > 0 Then
        Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value = Ayar(pnl, ornek, SATIR_HEDEF_DOSYA).Value
        Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value = Ayar(pnl, ornek, SATIR_HEDEF_SAYFA).Value
        Ayar(pnl, k, SATIR_YAZIM).Value = Ayar(pnl, ornek, SATIR_YAZIM).Value
        Ayar(pnl, k, SATIR_HASAR).Value = Ayar(pnl, ornek, SATIR_HASAR).Value
        Ayar(pnl, k, SATIR_IHBAR).Value = Ayar(pnl, ornek, SATIR_IHBAR).Value
        Ayar(pnl, k, SATIR_TUTAR).Value = Ayar(pnl, ornek, SATIR_TUTAR).Value
    End If

    ListeSayfasi(k).UsedRange.Clear
    ButonlariKur pnl
    EkranAc

    YenileCalistir True
    PaneleGit k

    OlayYaz "PANEL EKLE", k, 0, 0, "PANEL " & k & " eklendi", ""

    MsgBox "PANEL " & k & " eklendi." & vbNewLine & vbNewLine & _
           IIf(ornek > 0, "Hedef ayarlari PANEL " & ornek & "'den kopyalandi." & vbNewLine, "") & _
           "Kaynak dosyayi ve sayfayi secip 'YENILE' calistir.", vbInformation
    Exit Sub

Hata:
    EkranAc
    MsgBox "PANEL EKLENEMEDI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Panel Ekle"

End Sub


Public Sub PanelKaldir()

    Dim pnl As Worksheet
    Dim adet As Long, k As Long, c As Long
    Dim cevap As String, kaynakAdi As String

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    adet = BlokSayisi(pnl)

    If GorunurBlokSayisi(pnl) <= 1 Then
        MsgBox "Tek panel kaldi, kaldirilamaz." & vbNewLine & _
               "Icerigini temizlemek istiyorsan 'PanelKur' calistir.", vbExclamation
        Exit Sub
    End If

    cevap = Trim$(InputBox("Hangi panel kaldirilsin?" & vbNewLine & _
                           "1 - " & adet & " arasi bir numara yaz." & vbNewLine & vbNewLine & _
                           "Panelin tum ayarlari ve eslestirmeleri silinir." & vbNewLine & _
                           "Diger panellerin numarasi DEGISMEZ.", _
                           "Panel Kaldir", CStr(AktifBlok())))
    If cevap = "" Then Exit Sub
    If Not IsNumeric(cevap) Then
        MsgBox "Gecerli bir panel numarasi yaz.", vbExclamation
        Exit Sub
    End If

    k = CLng(cevap)
    If k < 1 Or k > adet Then
        MsgBox "1 - " & adet & " arasi bir numara olmali.", vbExclamation
        Exit Sub
    End If
    If BlokGizliMi(pnl, k) Then
        MsgBox "PANEL " & k & " zaten kaldirilmis.", vbInformation
        Exit Sub
    End If

    kaynakAdi = AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA)

    If MsgBox("PANEL " & k & " kaldirilacak." & vbNewLine & vbNewLine & _
              "Kaynak: " & kaynakAdi & vbNewLine & _
              "Hedef : " & AyarMetni(pnl, k, SATIR_HEDEF_DOSYA) & vbNewLine & vbNewLine & _
              "Ayarlari ve eslestirmeleri silinir, panel gizlenir." & vbNewLine & _
              "Diger panellerin numarasi degismez." & vbNewLine & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbExclamation + vbDefaultButton2, "Panel Kaldir") <> vbYes Then
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Hata

    c = BlokSutun(k)
    AlanTemizle pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c), pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1))

    ' Baslik "PANEL k" olarak kalir: blok sayimi ve numaralar bozulmasin
    pnl.Cells(SATIR_BLOK_BASLIK, c).Value = "PANEL " & k
    pnl.Range(pnl.Columns(c), pnl.Columns(c + BLOK_GENISLIK - 1)).Hidden = True
    ListeSayfasi(k).UsedRange.Clear

    g_SonAktifBlok = 0
    EkranAc

    OlayYaz "PANEL KALDIR", k, 0, 0, "PANEL " & k & " kaldirildi (gizlendi). Kaynak: " & kaynakAdi, ""
    PaneleGit KomsuBlok(pnl, k, -1)

    MsgBox "PANEL " & k & " kaldirildi." & vbNewLine & vbNewLine & _
           "Diger panellerin numarasi degismedi." & vbNewLine & _
           "'+ Panel' bu yeri yeniden kullanir.", vbInformation
    Exit Sub

Hata:
    EkranAc
    MsgBox "PANEL KALDIRMA SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Hata"

End Sub
