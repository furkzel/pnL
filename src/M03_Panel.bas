Attribute VB_Name = "M03_Panel"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.0  -  M03_Panel
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
Public Const HAZIRLIK_ADI As String = "AP_HAZIRLIK"

' Office sabitleri (Mac / eski surumlerde ad cozumlemesi sorun olmasin)
Private Const SEKIL_YUVARLAK As Long = 5     ' msoShapeRoundedRectangle
Private Const METIN_YATAY As Long = 1        ' msoTextOrientationHorizontal
Private Const DIKEY_ORTA As Long = 3         ' msoAnchorMiddle
Private Const HIZA_ORTA As Long = 2          ' msoAlignCenter
Private Const HIZA_SOL As Long = 1           ' msoAlignLeft

' Buton stilleri (BT-01: birincil / ikincil / ucuncul hiyerarsisi)
Private Const STIL_ADIM As Long = 1          ' 1-4: durumu AdimlariGuncelle verir
Private Const STIL_IKINCIL As Long = 2       ' kenarlikli
Private Const STIL_MENU As Long = 3          ' ucuncul: "Diger" menusu satiri
Private Const STIL_TEHLIKE As Long = 4       ' Geri Al: kirmizi hayalet
Private Const STIL_PIL As Long = 5           ' HAZIRLIK
Private Const STIL_BILDIRIM As Long = 6      ' bildirim seridi
Private Const SEKIL_DIKDORTGEN As Long = 1   ' msoShapeRectangle


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
    If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
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
    pnl.Rows(3).RowHeight = 32
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
    pnl.Columns(c + OFS_FILTRE).ColumnWidth = 24      ' ayni zamanda SABIT sutun adi
    pnl.Columns(c + OFS_ISLEM).ColumnWidth = 16       ' ayni zamanda SABIT deger
    pnl.Columns(c + OFS_HARITA).ColumnWidth = 30      ' ayni zamanda KONTROL karnesi

    ' --- Blok basligi ---
    With pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c), pnl.Cells(SATIR_BLOK_BASLIK, c + BLOK_GENISLIK - 1))
        .UnMerge
        .Merge
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Bold = True
        .Font.Size = YZ_BASLIK
        .Font.Color = Palet(prZemin)
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
    EtiketYaz pnl, SATIR_ANAHTAR, c, TR("Anahtar s#utunu (hasar no)"), False
    EtiketYaz pnl, SATIR_KIRILIM, c, TR("K#ir#il#im s#utunu (bran#s)"), False
    EtiketYaz pnl, SATIR_DONEM, c, TR("Rapor d#onemi sonu (bo#s = ge#cen ay)"), False

    ' --- Ayar (deger) hucreleri ---
    ayarSatirlari = Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                          SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, _
                          SATIR_PROFIL, SATIR_TUTAR, SATIR_HASAR, SATIR_IHBAR, _
                          SATIR_ANAHTAR, SATIR_KIRILIM, SATIR_DONEM)
    For i = 0 To UBound(ayarSatirlari)
        r = ayarSatirlari(i)
        With pnl.Cells(r, c + OFS_AYAR)
            .Interior.Color = Palet(prAyarZemin)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = Palet(prAyarKenar)
            .Font.Color = Palet(prMetin)
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
    pnl.Cells(SATIR_DONEM, c + OFS_AYAR).NumberFormat = "dd.mm.yyyy"
    pnl.Cells(SATIR_DONEM, c + OFS_AYAR).HorizontalAlignment = xlLeft

    If AyarMetni(pnl, k, SATIR_DAHIL) = "" Then Ayar(pnl, k, SATIR_DAHIL).Value = HayirMetni()
    If AyarMetni(pnl, k, SATIR_YAZIM) = "" Then Ayar(pnl, k, SATIR_YAZIM).Value = YazimMetni(ymSayfaSonu)
    DahilBoya pnl, k

    ' --- SABIT DEGERLER kutusu ve KONTROL NOKTALARI karnesi ---
    SabitBasligiCiz pnl, k
    KarneCiz pnl, k

    ' --- Kart cerceveleri: ayarlar | sabit degerler | kontrol karnesi ---
    KartCercevesi pnl.Range(pnl.Cells(SATIR_DAHIL, c), pnl.Cells(SATIR_AYAR_SON, c + OFS_AYAR))
    KartCercevesi pnl.Range(pnl.Cells(SABIT_BASLIK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_NOT, c + OFS_SABIT_DEGER))
    KartCercevesi pnl.Range(pnl.Cells(KARNE_BASLIK, c + OFS_KARNE), pnl.Cells(KARNE_NOT, c + OFS_KARNE))

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
    For i = LBound(ayarSatirlari) To UBound(ayarSatirlari)
        pnl.Cells(CLng(ayarSatirlari(i)), c + OFS_AYAR).Locked = False
    Next i
    pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_DEGER)).Locked = False
    pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(SON_SATIR, c + OFS_ISLEM)).Locked = False

End Sub


' Ayar kartinin cevresine ince, yumusak cerceve
Public Sub KartCercevesi(ByVal rng As Range)
    Dim kenar As Variant
    For Each kenar In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With rng.Borders(kenar)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = Palet(prCerceve)
        End With
    Next kenar
End Sub


' Etiket c..c+1 BIRLESIK hucreye yazilir (deger c+2): etiket dar
' VERI TIPI sutununa tasip kopuk gorunmez.
Private Sub EtiketYaz(ByVal pnl As Worksheet, ByVal r As Long, ByVal c As Long, _
                      ByVal metin As String, ByVal kalin As Boolean)
    With pnl.Range(pnl.Cells(r, c + OFS_ETIKET), pnl.Cells(r, c + OFS_AYAR - 1))
        .UnMerge
        .Merge
    End With
    With pnl.Cells(r, c + OFS_ETIKET)
        .Value = metin
        .Font.Bold = kalin
        .Font.Color = Palet(prEtiketYazi)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
End Sub


Private Sub BolumYaz(ByVal pnl As Worksheet, ByVal r As Long, ByVal c As Long, ByVal metin As String)
    With pnl.Range(pnl.Cells(r, c + OFS_ETIKET), pnl.Cells(r, c + OFS_AYAR - 1))
        .UnMerge
        .Merge
    End With
    With pnl.Cells(r, c + OFS_ETIKET)
        .Value = metin
        .Font.Bold = True
        .Font.Size = YZ_DIPNOT
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
        .Font.Color = Palet(prMetin)
        .Interior.Color = Palet(prSabitBaslik)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With

    With pnl.Cells(SABIT_NOT, c + OFS_SABIT_SUTUN)
        .Value = TR("Sat#ir#i doldurunca yenisi a#c#il#ir (en fazla 15). S#utunu silince sat#ir kalkar.")
        .Font.Italic = True
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
    End With

End Sub


Public Sub TabloBasligiCiz(ByVal pnl As Worksheet, ByVal k As Long)

    Dim c As Long

    c = BlokSutun(k)
    pnl.Cells(BASLIK_SATIR, c + OFS_KAYNAK).Value = TR("KAYNAK S#UTUN  (ayl#ik)")
    pnl.Cells(BASLIK_SATIR, c + OFS_TIP).Value = TR("VER#I T#IP#I")
    ' PN-06: cift tiklanabilen sutunlarin basliginda glif (on ek)
    pnl.Cells(BASLIK_SATIR, c + OFS_HEDEF).Value = ChrW(10003) & TR("  HEDEF S#UTUN  (ana dosya)")
    pnl.Cells(BASLIK_SATIR, c + OFS_HARITA).Value = ChrW(8646) & TR("  HEDEF HAR#ITASI")
    pnl.Cells(BASLIK_SATIR, c + OFS_FILTRE).Value = TR("F#ILTRE")
    pnl.Cells(BASLIK_SATIR, c + OFS_ISLEM).Value = ChrW(8646) & TR("  #I#SLEM")

    With pnl.Range(pnl.Cells(BASLIK_SATIR, c), pnl.Cells(BASLIK_SATIR, c + BLOK_GENISLIK - 1))
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
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

    Dim doluSon As Long
    c = BlokSutun(k)
    ' TS-06: yalnizca dolu (kaynak sutunu olan) satirlara ince alt cizgi;
    ' bos satirlarda kenarlik yok. Dolu satirlar ILK_SATIR'dan bitisiktir:
    ' hucre hucre degil, tek aralik olarak cizilir (v1.0: 502 hucre okumasi).
    pnl.Range(pnl.Cells(ilk, c), pnl.Cells(son, c + BLOK_GENISLIK - 1)).Borders.LineStyle = xlNone
    doluSon = BlokSonSatir(pnl, k)
    If doluSon > son Then doluSon = son
    If doluSon >= ilk Then
        With pnl.Range(pnl.Cells(ilk, c), pnl.Cells(doluSon, c + BLOK_GENISLIK - 1))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlHairline
            .Borders(xlEdgeBottom).Color = Palet(prCizgi)
            If doluSon > ilk Then
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).Weight = xlHairline
                .Borders(xlInsideHorizontal).Color = Palet(prCizgi)
            End If
        End With
    End If
    With pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c + BLOK_GENISLIK - 1), _
                   pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = Palet(prBlokBaslik)
    End With

End Sub


' "Bu paneli aktar": anahtar glifi (TS-03). Deger "Evet"/"Hayir" olarak
' KALIR (menu ve kod onu okur); glif yalnizca sayi bicimidir:
' Evet = "(dolu nokta) Evet" marka rengi, Hayir = "(bos daire) Hayir" soluk.
Public Sub DahilBoya(ByVal pnl As Worksheet, ByVal k As Long)
    HazirlikGuncelle pnl
    With Ayar(pnl, k, SATIR_DAHIL)
        .Font.Bold = True
        .Interior.Color = Palet(prDahilEvet)
        If DahilMi(pnl, k) Then
            .NumberFormat = ";;;""" & ChrW(9679) & "  ""@"
            .Font.Color = Palet(prMarka)
        Else
            .NumberFormat = ";;;""" & ChrW(9675) & "  ""@"
            .Font.Color = Palet(prPasifMetin)
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
'  PANEL                : < Aktif panel gostergesi >, Kilavuz, Denetim, Tani
'  HAZIRLIK             : aktarima dahil panellerin K1..K8 karnesinden canli ozet
'  Numarali 4 ana buton (1 YENILE, 2 Otomatik Eslestir, 3 On Kontrol,
'  4 AKTAR) yeni kullaniciya yolu gosterir.
' ============================================================

Public Sub ButonlariYenile()

    Dim pnl As Worksheet
    Dim hata As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub

    EkranKilitle
    On Error GoTo Cikis
    ButonlariKur pnl
Cikis:
    If Err.Number <> 0 Then hata = Err.Number & " - " & Err.Description
    EkranAc
    If hata <> "" Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "Butonlar kurulamadi: " & hata
    Else
        Bildir BS_OK, "Aktarim Paneli", "Butonlar yeniden olusturuldu."
    End If

End Sub


Public Sub ButonlariKur(ByVal pnl As Worksheet)

    Dim i As Long
    Dim x As Double, y As Double, x2 As Double, menuX As Double
    Dim shp As Shape
    Dim makrolar As Variant, adlar As Variant

    On Error Resume Next
    pnl.Buttons.Delete
    On Error GoTo 0
    SekilleriSil pnl
    PanelSatirYukseklikleri pnl
    pnl.Range(pnl.Cells(1, 1), pnl.Cells(SERIT_SON_SATIR, TEMIZLIK_SON_SUTUN)).Interior.Color = Palet(prYuzey)
    With pnl.Range(pnl.Cells(SERIT_SON_SATIR, 1), pnl.Cells(SERIT_SON_SATIR, TEMIZLIK_SON_SUTUN)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = Palet(prCizgi)
    End With

    ' --- Satir 2: urun adi + aktif panel gostergesi ---
    MetinKutusu pnl, "AP_BASLIK", TR("AKTARIM PANEL#I  ") & SURUM, SERIT_SOL, pnl.Rows(2).Top, _
                 SERIT_ADIM_GENISLIK * 2, pnl.Rows(2).Height, YZ_BASLIK, Palet(prMetin), True
    MetinKutusu pnl, GOSTERGE_ADI, "PANEL 1", SERIT_SOL + SERIT_ADIM_GENISLIK * 2, pnl.Rows(2).Top, _
                 SERIT_ADIM_GENISLIK * 3, pnl.Rows(2).Height, YZ_GOVDE, Palet(prIkincilMetin), True

    ' --- Satir 3: BIRINCIL (4 adim) | IKINCIL | TEHLIKELI ---
    y = pnl.Rows(3).Top + (pnl.Rows(3).Height - SERIT_BUTON_YUKSEKLIK) / 2
    makrolar = Array("Yenile", "OtomatikEslestir", "OnKontrol", "Aktar")
    For i = 0 To UBound(makrolar)
        x = SERIT_SOL + i * (SERIT_ADIM_GENISLIK + SERIT_BUTON_ARALIK)
        ButonEkle pnl, CStr(makrolar(i)), AdimAdi(i + 1), x, y, SERIT_ADIM_GENISLIK, SERIT_BUTON_YUKSEKLIK, STIL_ADIM
    Next i
    x2 = SERIT_SOL + SERIT_ADIM_SAYI * (SERIT_ADIM_GENISLIK + SERIT_BUTON_ARALIK) - SERIT_BUTON_ARALIK + SERIT_GRUP_ARALIK
    y = pnl.Rows(3).Top + (pnl.Rows(3).Height - SERIT_IKINCIL_YUKSEKLIK) / 2
    makrolar = Array("AnalizAc", "DenetimOzeti", "DigerMenu")
    adlar = Array("Analiz", "Denetim", TR("Di#ger  ") & ChrW(9662))
    For i = 0 To UBound(makrolar)
        x = x2 + i * (SERIT_IKINCIL_GENISLIK + SERIT_BUTON_ARALIK)
        ButonEkle pnl, CStr(makrolar(i)), CStr(adlar(i)), x, y, SERIT_IKINCIL_GENISLIK, SERIT_IKINCIL_YUKSEKLIK, STIL_IKINCIL
        If CStr(makrolar(i)) = "DigerMenu" Then menuX = x
    Next i
    x = x2 + 3 * (SERIT_IKINCIL_GENISLIK + SERIT_BUTON_ARALIK) - SERIT_BUTON_ARALIK + SERIT_GRUP_ARALIK
    AyracCiz pnl, 1, x - SERIT_GRUP_ARALIK / 2, pnl.Rows(3).Top + 4, pnl.Rows(3).Top + pnl.Rows(3).Height - 4
    ButonEkle pnl, "GeriAl", ChrW(8630) & TR("  Geri Al"), x, y, SERIT_IKINCIL_GENISLIK, SERIT_IKINCIL_YUKSEKLIK, STIL_TEHLIKE

    ' --- Satir 4: HAZIRLIK pili (adimlarin altinda) + bildirim seridi ---
    y = pnl.Rows(4).Top + (pnl.Rows(4).Height - SERIT_IKINCIL_YUKSEKLIK) / 2
    ButonEkle pnl, "HazirlikGoster", TR("HAZIRLIK"), SERIT_SOL, y, _
              SERIT_ADIM_SAYI * (SERIT_ADIM_GENISLIK + SERIT_BUTON_ARALIK) - SERIT_BUTON_ARALIK, SERIT_IKINCIL_YUKSEKLIK, STIL_PIL
    pnl.Shapes(ONEK_SEKIL & "BTN_HazirlikGoster").Name = HAZIRLIK_ADI
    ButonEkle pnl, "BildirimListesi", TR("Bildirim yok"), x2, y, _
              x + SERIT_IKINCIL_GENISLIK - x2, SERIT_IKINCIL_YUKSEKLIK, STIL_BILDIRIM
    pnl.Shapes(ONEK_SEKIL & "BTN_BildirimListesi").Name = BILDIRIM_ADI

    ' --- UCUNCUL: "Diger" menusu (gizli; DigerMenu acar/kapatir) ---
    makrolar = Array("YeniPanelEkle", "PanelKaldir", "PanelOnceki", "PanelSonraki", "SetYukle", "SetKaydet", "SetSil", _
                     "OneriOnayla", "ProfilKaydet", "AktarSecili", "TutarMutabakatiSon", "AyarlarAc", "Yardim", "Tani")
    adlar = Array(TR("+  Panel ekle"), TR("#-  Panel kald#ir"), TR("#Onceki panel"), "Sonraki panel", TR("Set y#ukle"), _
                 "Set kaydet", "Set sil", TR("#Oneri onayla"), "Profil kaydet", TR("Se#cili s#utun yamas#i"), _
                 TR("Tutar mutabakat#i (son parti)"), TR("Ayarlar ve e#sikler"), TR("K#ilavuz"), TR("Tan#i"))
    For i = 0 To UBound(makrolar)
        ButonEkle pnl, CStr(makrolar(i)), CStr(adlar(i)), menuX, _
                  pnl.Rows(SERIT_SON_SATIR + 1).Top + i * SERIT_MENU_YUKSEKLIK, SERIT_MENU_GENISLIK, SERIT_MENU_YUKSEKLIK, STIL_MENU
        Set shp = pnl.Shapes(ONEK_SEKIL & "BTN_" & CStr(makrolar(i)))
        shp.Name = SERIT_MENU_ONEK & CStr(makrolar(i))
        shp.Visible = False
    Next i

    g_SonAktifBlok = 0
    AktifPanelGoster pnl, AktifBlok()
    HazirlikGuncelle pnl
    BildirimSeridiniYaz

End Sub


Private Function AdimAdi(ByVal n As Long) As String
    Select Case n
        Case 1: AdimAdi = TR("1   Yenile")
        Case 2: AdimAdi = TR("2   E#sle#stir")
        Case 3: AdimAdi = TR("3   #On Kontrol")
        Case Else: AdimAdi = "4   AKTAR"
    End Select
End Function


Public Sub DigerMenu()
    Dim pnl As Worksheet, shp As Shape, ac As Boolean
    EkranKilitle
    On Error GoTo Cikis
    Set pnl = PanelSayfasi()
    ac = Not pnl.Shapes(SERIT_MENU_ONEK & "YeniPanelEkle").Visible
    For Each shp In pnl.Shapes
        If Left$(shp.Name, Len(SERIT_MENU_ONEK)) = SERIT_MENU_ONEK Then
            shp.Visible = ac
            If ac Then shp.ZOrder 0
        End If
    Next shp
Cikis:
    EkranAc
End Sub


Public Sub AyarlarAc()
    On Error Resume Next
    If Not SayfaVarMi(SAYFA_AYARLAR) Then AyarlarSayfasiHazirla
    ThisWorkbook.Worksheets(SAYFA_AYARLAR).Activate
    ActiveWindow.ScrollRow = 1
    On Error GoTo 0
End Sub


' BT-02/03: adim durumu. Tamamlanan "isaretli", SIRADAKI dolgulu (marka),
' sonrakiler pasif. ENGEL varken AKTAR pasif gorunur (tiklanirsa engeli
' anlatir; kontrol atlanmaz).
Public Sub AdimlariGuncelle(ByVal pnl As Worksheet, ByVal engelli As Boolean, ByVal bekliyor As Boolean)

    Dim shp As Shape
    Dim i As Long, sonraki As Long, k As Long, c As Long, r As Long, eslesen As Long
    Dim makrolar As Variant

    k = AktifBlok(): c = BlokSutun(k)
    sonraki = 1
    If Len(GuvenliMetin(pnl.Cells(ILK_SATIR, c + OFS_KAYNAK).Value)) > 0 Then
        sonraki = 2
        For r = ILK_SATIR To BlokSonSatir(pnl, k)
            If Len(GuvenliMetin(pnl.Cells(r, c + OFS_HEDEF).Value)) > 0 Then eslesen = eslesen + 1
        Next r
        If eslesen > 0 Then sonraki = 3
        If Not engelli And Not bekliyor And eslesen > 0 Then sonraki = 4
    End If

    makrolar = Array("Yenile", "OtomatikEslestir", "OnKontrol", "Aktar")
    For i = 0 To UBound(makrolar)
        Set shp = pnl.Shapes(ONEK_SEKIL & "BTN_" & CStr(makrolar(i)))
        shp.AlternativeText = ""
        If i + 1 = sonraki Then
            AdimBoya shp, Palet(prMarka), Palet(prMarka), Palet(prZemin), AdimAdi(i + 1)
        ElseIf i + 1 < sonraki Then
            AdimBoya shp, Palet(prZemin), Palet(prCerceve), Palet(prIkincilMetin), ChrW(10003) & "  " & AdimAdi(i + 1)
        Else
            AdimBoya shp, Palet(prZemin), Palet(prCizgi), Palet(prPasifMetin), AdimAdi(i + 1)
        End If
        If i = 3 And engelli Then
            AdimBoya shp, Palet(prYuzey), Palet(prCizgi), Palet(prPasifMetin), AdimAdi(4) & "   " & ChrW(10005)
            shp.AlternativeText = TR("ENGEL var: aktar#im yap#ilamaz. Karnedeki k#irm#iz#i maddeyi giderip #On Kontrol'#u " & _
                                     "yeniden #cal#i#st#ir#in.")
        End If
    Next i

End Sub


Private Sub AdimBoya(ByVal shp As Shape, ByVal zemin As Long, ByVal kenar As Long, ByVal yazi As Long, ByVal metin As String)
    shp.Fill.ForeColor.RGB = zemin
    shp.Line.Visible = -1
    shp.Line.ForeColor.RGB = kenar
    shp.Line.Weight = 0.75
    shp.TextFrame2.TextRange.Text = metin
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = yazi
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
    Dim zemin As Long, yaziRenk As Long, kenar As Long, boyut As Double, hiza As Long
    Dim kenarVar As Boolean, kalin As Boolean

    kenarVar = True
    kalin = True
    hiza = HIZA_ORTA
    Select Case stil
        Case STIL_ADIM
            zemin = Palet(prZemin): kenar = Palet(prCizgi): yaziRenk = Palet(prPasifMetin): boyut = YZ_BOLUM
        Case STIL_TEHLIKE
            zemin = Palet(prZemin): kenar = Palet(prEngel): yaziRenk = Palet(prEngel): boyut = YZ_GOVDE
        Case STIL_PIL
            zemin = Palet(prYuzey): kenar = Palet(prCizgi): yaziRenk = Palet(prMetin): boyut = YZ_GOVDE
            hiza = HIZA_SOL
        Case STIL_BILDIRIM
            zemin = Palet(prZemin): kenar = Palet(prCizgi): yaziRenk = Palet(prIkincilMetin): boyut = YZ_DIPNOT
            hiza = HIZA_SOL
            kalin = False
        Case STIL_MENU
            zemin = Palet(prZemin): kenar = Palet(prCizgi): yaziRenk = Palet(prMetin): boyut = YZ_GOVDE
            hiza = HIZA_SOL
            kalin = False
        Case Else
            zemin = Palet(prZemin): kenar = Palet(prCerceve): yaziRenk = Palet(prMetin): boyut = YZ_GOVDE
    End Select

    Set shp = pnl.Shapes.AddShape(IIf(stil = STIL_MENU, SEKIL_DIKDORTGEN, SEKIL_YUVARLAK), sol, ust, g, y)
    shp.Name = ONEK_SEKIL & "BTN_" & makro
    shp.OnAction = makro
    On Error Resume Next
    shp.Adjustments.Item(1) = 0.18
    shp.Shadow.Visible = 0
    On Error GoTo 0
    shp.Fill.ForeColor.RGB = zemin
    shp.Line.Visible = IIf(kenarVar, -1, 0)
    shp.Line.ForeColor.RGB = kenar
    shp.Line.Weight = 0.75

    With shp.TextFrame2
        .MarginLeft = IIf(hiza = HIZA_SOL, 8, 3)
        .MarginRight = 3
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = 0
        .VerticalAnchor = DIKEY_ORTA
        .TextRange.Text = yazi
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = boyut
        .TextRange.Font.Bold = IIf(kalin, -1, 0)
        .TextRange.Font.Fill.ForeColor.RGB = yaziRenk
        .TextRange.ParagraphFormat.Alignment = hiza
    End With
    shp.Placement = xlFreeFloating

End Sub


Private Sub MetinKutusu(ByVal pnl As Worksheet, ByVal ad As String, ByVal metin As String, _
                        ByVal sol As Double, ByVal ust As Double, ByVal g As Double, ByVal y As Double, _
                        ByVal boyut As Double, ByVal renk As Long, ByVal kalin As Boolean)

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
    MetinKutusu pnl, ONEK_SEKIL & "GRUP_" & n, metin, sol, ust, 300, 14, YZ_DIPNOT, Palet(prSeritGrupYazi), True
End Sub


Private Sub AyracCiz(ByVal pnl As Worksheet, ByVal n As Long, ByVal x As Double, _
                     ByVal ust As Double, ByVal alt As Double)

    Dim shp As Shape

    Set shp = pnl.Shapes.AddLine(x, ust + 2, x, alt)
    shp.Name = ONEK_SEKIL & "AYRAC_" & n
    shp.Line.ForeColor.RGB = Palet(prCizgi)
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

    For Each shp In pnl.Shapes
        If Left$(shp.Name, Len(SERIT_MENU_ONEK)) = SERIT_MENU_ONEK Then shp.Visible = False
    Next shp
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
    HazirlikGuncelle pnl
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
    If Not PanelHazirMi(pnl) Then Exit Sub

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
            Bildir BS_DIKKAT, "Aktarim Paneli", "En fazla " & MAX_BLOK & " panel blogu desteklenir."
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

    Bildir BS_OK, "Aktarim Paneli", "PANEL " & k & " eklendi." & vbNewLine & vbNewLine & _
           IIf(ornek > 0, "Hedef ayarlari PANEL " & ornek & "'den kopyalandi." & vbNewLine, "") & _
           "Kaynak dosyayi ve sayfayi secip 'YENILE' calistir."
    Exit Sub

Hata:
    EkranAc
    Bildir BS_ENGEL, "Panel Ekle", "PANEL EKLENEMEDI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description

End Sub


Public Sub PanelKaldir()

    Dim pnl As Worksheet
    Dim adet As Long, k As Long, c As Long
    Dim cevap As String, kaynakAdi As String

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    adet = BlokSayisi(pnl)

    If GorunurBlokSayisi(pnl) <= 1 Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "Tek panel kaldi, kaldirilamaz." & vbNewLine & _
               "Icerigini temizlemek istiyorsan 'PanelKur' calistir."
        Exit Sub
    End If

    cevap = Trim$(InputBox("Hangi panel kaldirilsin?" & vbNewLine & _
                           "1 - " & adet & " arasi bir numara yaz." & vbNewLine & vbNewLine & _
                           "Panelin tum ayarlari ve eslestirmeleri silinir." & vbNewLine & _
                           "Diger panellerin numarasi DEGISMEZ.", _
                           "Panel Kaldir", CStr(AktifBlok())))
    If cevap = "" Then Exit Sub
    If Not IsNumeric(cevap) Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "Gecerli bir panel numarasi yaz."
        Exit Sub
    End If

    k = CLng(cevap)
    If k < 1 Or k > adet Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "1 - " & adet & " arasi bir numara olmali."
        Exit Sub
    End If
    If BlokGizliMi(pnl, k) Then
        Bildir BS_BILGI, "Aktarim Paneli", "PANEL " & k & " zaten kaldirilmis."
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

    Bildir BS_OK, "Aktarim Paneli", "PANEL " & k & " kaldirildi." & vbNewLine & vbNewLine & _
           "Diger panellerin numarasi degismedi." & vbNewLine & _
           "'+ Panel' bu yeri yeniden kullanir."
    Exit Sub

Hata:
    EkranAc
    Bildir BS_ENGEL, "Hata", "PANEL KALDIRMA SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description

End Sub
