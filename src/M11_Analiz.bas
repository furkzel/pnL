Attribute VB_Name = "M11_Analiz"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M11_Analiz  (ANALIZ SAYFASI)
'
'   AnalizAc           : sayfayi acar (yoksa kurar)
'   AnalizCalistir     : istek kartini okur, YALNIZCA istenen sutunlari
'                        okur (satir, sutun, olcu, filtre, donem), gruplar,
'                        tablo + yerel grafik + 3 maddelik yorum yazar
'   AnalizDegisti      : Sablon secilince kart dolar; Kaynak degisince
'                        sutun menuleri yenilenir (ThisWorkbook olayindan)
'   AnalizSablonKaydet / AnalizSablonSil : kullanici sablonlari
'
' Gruplama: Collection karmasi (grup <= Ayarlar ANALIZ_GRUP); daha cok
' farkli deger varsa siralama tabanli gruplamaya gecilir. Filtre, Panel
' FILTRE sutunuyla AYNI motordur (M14_Filtre). Tarih tipli satir/sutun
' boyutu ay ay (yyyy-aa) gruplanir.
'
' Sablonlar: A1 veri profili, A2 ay-ay (z-skor), A3 kirilim x sirket,
' A4 hasar gelisim ucgeni (zincir merdiven faktorleri), A5 ihbar gecikmesi,
' A6 buyuk hasar Top-N, A7 aktarim etkisi, A8 kontrol noktasi trendi,
' A9 frekans / siddet ayrisimi.
' ============================================================

Private Const OLCU_ADET As String = "Adet"
Private Const OLCU_TOPLAM As String = "Toplam"
Private Const OLCU_ORT As String = "Ortalama"
Private Const OLCU_MIN As String = "Min"
Private Const OLCU_MAKS As String = "Maks"
Private Const OLCU_TEKIL As String = "Tekil"
Private Const OLCU_P50 As String = "P50"
Private Const OLCU_P95 As String = "P95"
Private Const BOS_GRUP As String = "(bos)"
Private Const DIGER_GRUP As String = "Diger"
Private Const AYRAC As String = "|"
Private Const GRAFIK_SATIR As Long = 15
Private Const TOPN As Long = 20

Private Const LS_SABLON As Long = 0          ' AN_LST_SUTUN + ofs
Private Const LS_KAYNAK As Long = 1
Private Const LS_SUTUNLAR As Long = 2
Private Const LS_OLCU As Long = 3

' --- Bir calistirmanin gruplari (modul duzeyi; her calistirmada sifirlanir) ---
Private gSay As Long
Private gSatir() As String
Private gSutun() As String
Private gN() As Double
Private gSayisal() As Double
Private gT() As Double
Private gMin() As Double
Private gMax() As Double
Private gOzel() As Double            ' Tekil / P50 / P95


' ============================================================
'  GIRIS
' ============================================================
Public Sub AnalizAc()
    On Error Resume Next
    If Not SayfaVarMi(SAYFA_ANALIZ) Then AnalizSayfasiHazirla
    AnalizMenuleri
    ThisWorkbook.Worksheets(SAYFA_ANALIZ).Activate
    ActiveWindow.ScrollRow = 1
    On Error GoTo 0
End Sub


Public Sub AnalizDegisti(ByVal Target As Range)
    Dim ws As Worksheet
    On Error GoTo Cikis
    Set ws = Target.Worksheet
    If Target.Cells.Count <> 1 Or Target.Column <> AN_DEGER_SUTUN Then Exit Sub
    EkranKilitle
    Select Case Target.Row
        Case AN_SATIR_SABLON
            AnalizSablonUygula ws, GuvenliMetin(Target.Value)
            AnalizMenuleri
        Case AN_SATIR_KAYNAK
            AnalizMenuleri
    End Select
    EkranAc
    Exit Sub
Cikis:
    EkranAc
End Sub


' ============================================================
'  SAYFA
' ============================================================
Public Sub AnalizSayfasiHazirla()

    Dim ws As Worksheet
    Dim r As Long
    Dim etiketler As Variant
    Dim shp As Shape
    Dim x As Double

    Set ws = SayfaOlustur(SAYFA_ANALIZ)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    ws.Cells.Font.Name = PANEL_YAZI_TIPI
    ws.Cells.Font.Size = YZ_GOVDE
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(AN_ETIKET_SUTUN).ColumnWidth = 24
    ws.Columns(AN_ETIKET_SUTUN + 1).ColumnWidth = 2
    ws.Columns(AN_DEGER_SUTUN).ColumnWidth = 34
    ws.Range(ws.Columns(AN_DEGER_SUTUN + 1), ws.Columns(AN_DEGER_SUTUN + 12)).ColumnWidth = 13
    ws.Range(ws.Columns(AN_LST_SUTUN), ws.Columns(AN_LST_SUTUN + LS_OLCU)).Hidden = True

    With ws.Cells(2, AN_ETIKET_SUTUN)
        .Value = TR("ANAL#IZ")
        .Font.Size = YZ_BASLIK
        .Font.Bold = True
        .Font.Color = Palet(prMetin)
    End With
    With ws.Cells(3, AN_ETIKET_SUTUN)
        .Value = TR("#Istek kart#in#i doldurun, #Cal#i#st#ir'a bas#in. Yaln#iz istenen s#utunlar okunur; " & _
                    "filtre Panel F#ILTRE s#utunuyla ayn#i s#oz dizimidir (AXA, >100, <>IPTAL, BOS ...).")
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
    End With

    ws.Cells(AN_ILK_ALAN - 1, AN_ETIKET_SUTUN).Value = TR("#ISTEK KARTI")
    TabloBasligiBicim ws.Range(ws.Cells(AN_ILK_ALAN - 1, AN_ETIKET_SUTUN), ws.Cells(AN_ILK_ALAN - 1, AN_DEGER_SUTUN))
    etiketler = Array(TR("#Sablon"), "Kaynak", TR("Sat#ir boyutu"), TR("S#utun boyutu (iste#ge ba#gl#i)"), TR("#Ol#c#u"), _
                      TR("#Ol#c#u s#utunu"), TR("Filtre s#utunu"), TR("Filtre (Panel s#oz dizimi)"), TR("D#onem s#utunu"), _
                      TR("D#onem ba#slang#ic#i"), TR("D#onem sonu"))
    For r = AN_ILK_ALAN To AN_SON_ALAN
        With ws.Cells(r, AN_ETIKET_SUTUN)
            .Value = etiketler(r - AN_ILK_ALAN)
            .Font.Color = Palet(prIkincilMetin)
            .IndentLevel = 1
        End With
        With ws.Cells(r, AN_DEGER_SUTUN)
            .Interior.Color = Palet(prZemin)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = Palet(prCizgi)
            .Font.Color = Palet(prMetin)
            .Locked = False
        End With
        ws.Range(ws.Cells(r, AN_ETIKET_SUTUN), ws.Cells(r, AN_DEGER_SUTUN)).Borders(xlEdgeBottom).Color = Palet(prCizgi)
    Next r
    ws.Range(ws.Cells(AN_SATIR_DONEM_BAS, AN_DEGER_SUTUN), ws.Cells(AN_SATIR_DONEM_SON, AN_DEGER_SUTUN)).NumberFormat = "dd.mm.yyyy"
    ws.Range(ws.Cells(AN_SATIR_FILTRE, AN_DEGER_SUTUN), ws.Cells(AN_SATIR_FILTRE, AN_DEGER_SUTUN)).NumberFormat = "@"
    KartCercevesi ws.Range(ws.Cells(AN_ILK_ALAN - 1, AN_ETIKET_SUTUN), ws.Cells(AN_SON_ALAN, AN_DEGER_SUTUN))

    ' --- Butonlar: birincil Calistir + ikincil sablon islemleri ---
    For Each shp In ws.Shapes
        If Left$(shp.Name, 6) = "AP_AN_" Then shp.Delete
    Next shp
    x = ws.Cells(AN_ILK_ALAN, AN_DEGER_SUTUN + 1).Left + 12
    AnalizButonu ws, "AP_AN_Calistir", ChrW(9654) & TR("   #Cal#i#st#ir"), "AnalizCalistir", x, ws.Cells(AN_ILK_ALAN, 1).Top, True
    AnalizButonu ws, "AP_AN_Yenile", TR("Se#cenekleri yenile"), "AnalizMenuleriTazele", x, ws.Cells(AN_ILK_ALAN + 2, 1).Top, False
    AnalizButonu ws, "AP_AN_Kaydet", TR("#Sablon kaydet"), "AnalizSablonKaydet", x, ws.Cells(AN_ILK_ALAN + 4, 1).Top, False
    AnalizButonu ws, "AP_AN_Sil", TR("#Sablon sil"), "AnalizSablonSil", x, ws.Cells(AN_ILK_ALAN + 6, 1).Top, False

    ws.Cells(AN_YORUM_ILK - 1, AN_ETIKET_SUTUN).Value = "YORUM"
    TabloBasligiBicim ws.Range(ws.Cells(AN_YORUM_ILK - 1, AN_ETIKET_SUTUN), ws.Cells(AN_YORUM_ILK - 1, AN_DEGER_SUTUN + 8))
    For r = AN_YORUM_ILK To AN_YORUM_ILK + 2
        With ws.Range(ws.Cells(r, AN_ETIKET_SUTUN), ws.Cells(r, AN_DEGER_SUTUN + 8))
            .Merge
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        ws.Rows(r).RowHeight = 28
    Next r
    ws.Cells(AN_YORUM_ILK, AN_ETIKET_SUTUN).Value = TR("Hen#uz analiz yok. Bir #sablon se#cin (#orne#gin A3) ve #Cal#i#st#ir'a bas#in.")
    ws.Cells(AN_YORUM_ILK, AN_ETIKET_SUTUN).Font.Color = Palet(prIkincilMetin)

    ws.Cells(AN_SATIR_OLCU, AN_DEGER_SUTUN).Value = OLCU_TOPLAM
    ws.Tab.Color = Palet(prMarka)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    AnalizMenuleri
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0

End Sub


Private Sub AnalizButonu(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal makro As String, _
                         ByVal x As Double, ByVal y As Double, ByVal birincil As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(5, x, y + 2, 150, SERIT_IKINCIL_YUKSEKLIK + IIf(birincil, 6, 0))
    shp.Name = ad
    shp.OnAction = makro
    shp.Shadow.Visible = 0
    shp.Fill.ForeColor.RGB = IIf(birincil, Palet(prMarka), Palet(prZemin))
    shp.Line.ForeColor.RGB = IIf(birincil, Palet(prMarka), Palet(prCerceve))
    shp.Line.Weight = 0.75
    With shp.TextFrame2
        .MarginLeft = 4
        .MarginRight = 4
        .VerticalAnchor = 3
        .TextRange.Text = metin
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = IIf(birincil, YZ_BOLUM, YZ_GOVDE)
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = IIf(birincil, Palet(prZemin), Palet(prMetin))
        .TextRange.ParagraphFormat.Alignment = 2
    End With
    shp.Placement = xlFreeFloating
End Sub


Private Function AnalizSayfasi() As Worksheet
    If Not SayfaVarMi(SAYFA_ANALIZ) Then AnalizSayfasiHazirla
    Set AnalizSayfasi = ThisWorkbook.Worksheets(SAYFA_ANALIZ)
End Function


Private Function Kart(ByVal ws As Worksheet, ByVal satir As Long) As String
    Kart = Trim$(GuvenliMetin(ws.Cells(satir, AN_DEGER_SUTUN).Value))
End Function


Public Sub AnalizMenuleriTazele()
    EkranKilitle
    On Error Resume Next
    AnalizMenuleri
    On Error GoTo 0
    EkranAc
End Sub


' Sablon, kaynak, sutun ve olcu menuleri (gizli liste sutunlari)
Public Sub AnalizMenuleri()

    Dim ws As Worksheet, kws As Worksheet, pnl As Worksheet, sb As Worksheet
    Dim liste() As Variant
    Dim n As Long, k As Long, i As Long, sonSut As Long, son As Long, ilk As Long
    Dim basliklar As Variant
    Dim aciklama As String

    Set ws = AnalizSayfasi()
    Set pnl = PanelSayfasi()

    ' --- sablonlar ---
    ReDim liste(0 To 99)
    For i = 1 To 9
        liste(n) = SablonAdi(i)
        n = n + 1
    Next i
    If SayfaVarMi(SAYFA_SABLON) Then
        Set sb = ThisWorkbook.Worksheets(SAYFA_SABLON)
        son = sb.Cells(sb.Rows.Count, 1).End(xlUp).Row
        For i = 2 To son
            If n <= UBound(liste) And Trim$(GuvenliMetin(sb.Cells(i, 1).Value)) <> "" Then
                liste(n) = GuvenliMetin(sb.Cells(i, 1).Value)
                n = n + 1
            End If
        Next i
    End If
    ReDim Preserve liste(0 To n - 1)
    ListeYaz ws, AN_LST_SUTUN + LS_SABLON, liste, True
    MenuKurSerbest ws.Cells(AN_SATIR_SABLON, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SABLON

    ' --- kaynaklar ---
    n = 0
    ReDim liste(0 To 3 * MAX_BLOK)
    If Not pnl Is Nothing Then
        If PanelDuzeniGuncel() Then
            For k = 1 To BlokSayisi(pnl)
                If Not BlokGizliMi(pnl, k) Then
                    If AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA) <> "" Then
                        liste(n) = KaynakMetni(k, 1): n = n + 1
                    End If
                    If AyarMetni(pnl, k, SATIR_HEDEF_DOSYA) <> "" Then
                        liste(n) = KaynakMetni(k, 2): n = n + 1
                        liste(n) = KaynakMetni(k, 3): n = n + 1
                    End If
                End If
            Next k
        End If
    End If
    If n = 0 Then
        liste(0) = TR("(Panel'de dosya se#cilmemi#s)")
        n = 1
    End If
    ReDim Preserve liste(0 To n - 1)
    ListeYaz ws, AN_LST_SUTUN + LS_KAYNAK, liste, True
    MenuKur ws.Cells(AN_SATIR_KAYNAK, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_KAYNAK

    ' --- secili kaynagin sutunlari ---
    ws.Columns(AN_LST_SUTUN + LS_SUTUNLAR).ClearContents
    If KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        sonSut = SonSutun(kws, EN_FAZLA_SUTUN)
        If sonSut > 0 Then
            basliklar = IkiBoyutlu(kws.Range(kws.Cells(1, 1), kws.Cells(1, sonSut)))
            ReDim liste(0 To sonSut - 1)
            For i = 1 To sonSut
                liste(i - 1) = SutunEtiketi(i, GuvenliMetin(basliklar(1, i)))
            Next i
            ListeYaz ws, AN_LST_SUTUN + LS_SUTUNLAR, liste, True
        End If
    End If
    MenuKurSerbest ws.Cells(AN_SATIR_SATIR, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SUTUNLAR
    MenuKurSerbest ws.Cells(AN_SATIR_SUTUN, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SUTUNLAR
    MenuKurSerbest ws.Cells(AN_SATIR_OLCU_SUTUN, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SUTUNLAR
    MenuKurSerbest ws.Cells(AN_SATIR_FILTRE_SUTUN, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SUTUNLAR
    MenuKurSerbest ws.Cells(AN_SATIR_DONEM_SUTUN, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_SUTUNLAR

    ' --- olculer ---
    ListeYaz ws, AN_LST_SUTUN + LS_OLCU, Array(OLCU_ADET, OLCU_TOPLAM, OLCU_ORT, OLCU_MIN, OLCU_MAKS, OLCU_TEKIL, OLCU_P50, OLCU_P95), True
    MenuKur ws.Cells(AN_SATIR_OLCU, AN_DEGER_SUTUN), ws, AN_LST_SUTUN + LS_OLCU

End Sub


Private Function SablonAdi(ByVal n As Long) As String
    Select Case n
        Case 1: SablonAdi = TR("A1  Veri profili")
        Case 2: SablonAdi = TR("A2  Ay-ay kar#s#ila#st#irma (z-skor)")
        Case 3: SablonAdi = TR("A3  K#ir#il#im x #sirket")
        Case 4: SablonAdi = TR("A4  Hasar geli#sim #u#cgeni (zincir merdiven)")
        Case 5: SablonAdi = TR("A5  #Ihbar gecikmesi da#g#il#im#i")
        Case 6: SablonAdi = TR("A6  B#uy#uk hasar Top-20")
        Case 7: SablonAdi = TR("A7  Aktar#im etkisi (son parti)")
        Case 8: SablonAdi = TR("A8  Kontrol noktas#i trendi (12 ay)")
        Case 9: SablonAdi = TR("A9  Frekans / #siddet ayr#i#s#im#i")
    End Select
End Function


Private Function SablonKodu(ByVal ad As String) As Long
    Dim t As String
    t = UCase$(Left$(Trim$(ad), 2))
    If Len(t) = 2 And Left$(t, 1) = "A" And Mid$(t, 2, 1) >= "1" And Mid$(t, 2, 1) <= "9" Then SablonKodu = CLng(Mid$(t, 2, 1))
End Function


' tur: 1 aylik dosya, 2 ana dosya, 3 son parti
Private Function KaynakMetni(ByVal k As Long, ByVal tur As Long) As String
    Select Case tur
        Case 1: KaynakMetni = "PANEL " & k & TR("  #*  Ayl#ik dosya")
        Case 2: KaynakMetni = "PANEL " & k & TR("  #*  Ana dosya")
        Case Else: KaynakMetni = "PANEL " & k & TR("  #*  Son parti (etkin aktar#im)")
    End Select
End Function


' Kaynak metnini sayfa + satir araligina cevirir
Private Function KaynakCoz(ByVal metin As String, ByRef ws As Worksheet, ByRef ilk As Long, ByRef son As Long, _
                           ByRef aciklama As String) As Boolean

    Dim pnl As Worksheet
    Dim k As Long, tur As Long, aktarimNo As Long
    Dim b As String
    Dim kaynakEslesti As Boolean

    KaynakCoz = False
    Set ws = Nothing
    b = BH(metin)
    If Left$(b, 6) <> "PANEL " Then Exit Function
    k = Val(Mid$(b, 7))
    Set pnl = PanelSayfasi()
    If pnl Is Nothing Or k < 1 Then Exit Function
    If k > BlokSayisi(pnl) Then Exit Function
    If InStr(b, "AYLIK") > 0 Then
        tur = 1
    ElseIf InStr(b, "SON PARTI") > 0 Then
        tur = 3
    Else
        tur = 2
    End If
    If tur = 1 Then
        Set ws = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_KAYNAK_DOSYA).Value), Ayar(pnl, k, SATIR_KAYNAK_SAYFA).Value)
    Else
        Set ws = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value), Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value)
    End If
    If ws Is Nothing Then Exit Function
    ilk = 2
    son = SonDoluSatir(ws)
    If tur = 3 Then
        If Not SonAktifAralik(AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA), AyarMetni(pnl, k, SATIR_KAYNAK_SAYFA), _
                              ws.Parent.Name, CStr(ws.Name), aktarimNo, ilk, son, kaynakEslesti) Then Exit Function
        aciklama = ws.Parent.Name & " / " & ws.Name & TR(", aktar#im ") & aktarimNo & " (" & ilk & "-" & son & ")"
    Else
        aciklama = ws.Parent.Name & " / " & ws.Name
    End If
    If son < ilk Then Exit Function
    KaynakCoz = True

End Function


' ============================================================
'  SABLONLAR: kart doldurma
' ============================================================
Private Sub AnalizSablonUygula(ByVal ws As Worksheet, ByVal ad As String)

    Dim kod As Long, k As Long, tur As Long
    Dim pnl As Worksheet, sb As Worksheet
    Dim i As Long, son As Long, j As Long

    kod = SablonKodu(ad)
    If kod = 0 Then
        ' kullanici sablonu
        If Not SayfaVarMi(SAYFA_SABLON) Then Exit Sub
        Set sb = ThisWorkbook.Worksheets(SAYFA_SABLON)
        son = sb.Cells(sb.Rows.Count, 1).End(xlUp).Row
        For i = 2 To son
            If BH(Trim$(GuvenliMetin(sb.Cells(i, 1).Value))) = BH(Trim$(ad)) Then
                For j = AN_SATIR_KAYNAK To AN_SON_ALAN
                    ws.Cells(j, AN_DEGER_SUTUN).Value = sb.Cells(i, 2 + j - AN_SATIR_KAYNAK).Value
                Next j
                Exit Sub
            End If
        Next i
        Exit Sub
    End If

    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    k = Val(Mid$(BH(Kart(ws, AN_SATIR_KAYNAK)), 7))
    If k < 1 Or k > BlokSayisi(pnl) Then k = AktifBlok()
    Select Case kod
        Case 1, 5, 6: tur = 1
        Case Else: tur = 2
    End Select
    ws.Cells(AN_SATIR_KAYNAK, AN_DEGER_SUTUN).Value = KaynakMetni(k, tur)
    For j = AN_SATIR_SATIR To AN_SON_ALAN
        ws.Cells(j, AN_DEGER_SUTUN).ClearContents
    Next j
    ws.Cells(AN_SATIR_OLCU, AN_DEGER_SUTUN).Value = OLCU_TOPLAM

    Select Case kod
        Case 2, 9
            KartSutun ws, AN_SATIR_SATIR, pnl, k, tur, SATIR_HASAR
            KartSutun ws, AN_SATIR_OLCU_SUTUN, pnl, k, tur, SATIR_TUTAR
        Case 3
            KartSutun ws, AN_SATIR_SATIR, pnl, k, tur, SATIR_KIRILIM
            ws.Cells(AN_SATIR_SUTUN, AN_DEGER_SUTUN).Value = AdayBaslik(pnl, k, tur, Array("sirket"))
            KartSutun ws, AN_SATIR_OLCU_SUTUN, pnl, k, tur, SATIR_TUTAR
        Case 4
            KartSutun ws, AN_SATIR_SATIR, pnl, k, tur, SATIR_HASAR
            ws.Cells(AN_SATIR_DONEM_SUTUN, AN_DEGER_SUTUN).Value = AdayBaslik(pnl, k, tur, Array("odemetarih", "raporlamatarih", "islemtarih", "raportarih", "donem"))
            KartSutun ws, AN_SATIR_OLCU_SUTUN, pnl, k, tur, SATIR_TUTAR
        Case 5
            KartSutun ws, AN_SATIR_SATIR, pnl, k, tur, SATIR_HASAR
            KartSutun ws, AN_SATIR_SUTUN, pnl, k, tur, SATIR_IHBAR
            ws.Cells(AN_SATIR_OLCU, AN_DEGER_SUTUN).Value = OLCU_ADET
        Case 6
            KartSutun ws, AN_SATIR_SATIR, pnl, k, tur, SATIR_ANAHTAR
            KartSutun ws, AN_SATIR_SUTUN, pnl, k, tur, SATIR_KIRILIM
            KartSutun ws, AN_SATIR_OLCU_SUTUN, pnl, k, tur, SATIR_TUTAR
            ws.Cells(AN_SATIR_OLCU, AN_DEGER_SUTUN).Value = OLCU_MAKS
    End Select

End Sub


' Panel ayarindaki (ana dosya) sutunu karta yazar; aylik dosyada o
' sutunu dolduran KAYNAK sutunu kullanilir.
Private Sub KartSutun(ByVal ws As Worksheet, ByVal satir As Long, ByVal pnl As Worksheet, ByVal k As Long, _
                      ByVal tur As Long, ByVal ayarSatiri As Long)
    Dim hedef As String
    Dim c As Long, r As Long
    hedef = AyarMetni(pnl, k, ayarSatiri)
    If hedef = "" Then Exit Sub
    If tur <> 1 Then
        ws.Cells(satir, AN_DEGER_SUTUN).Value = HucreMetni(hedef)
        Exit Sub
    End If
    c = BlokSutun(k)
    For r = ILK_SATIR To BlokSonSatir(pnl, k)
        If BH(Trim$(GuvenliMetin(pnl.Cells(r, c + OFS_HEDEF).Value))) = BH(hedef) Then
            ws.Cells(satir, AN_DEGER_SUTUN).Value = HucreMetni(GuvenliMetin(pnl.Cells(r, c + OFS_KAYNAK).Value))
            Exit Sub
        End If
    Next r
End Sub


' Kaynak/hedef basliklarinda sade adi verilen parcalardan birini iceren ilk sutun
Private Function AdayBaslik(ByVal pnl As Worksheet, ByVal k As Long, ByVal tur As Long, ByVal parcalar As Variant) As String
    Dim ws As Worksheet
    Dim sonSut As Long, i As Long, t As Long
    Dim v As Variant
    Dim ad As String
    If tur = 1 Then
        Set ws = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_KAYNAK_DOSYA).Value), Ayar(pnl, k, SATIR_KAYNAK_SAYFA).Value)
    Else
        Set ws = SayfaBul(DosyaBul(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value), Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value)
    End If
    If ws Is Nothing Then Exit Function
    sonSut = SonSutun(ws, EN_FAZLA_SUTUN)
    If sonSut < 1 Then Exit Function
    v = IkiBoyutlu(ws.Range(ws.Cells(1, 1), ws.Cells(1, sonSut)))
    For t = LBound(parcalar) To UBound(parcalar)
        For i = 1 To sonSut
            ad = Normalize(GuvenliMetin(v(1, i)))
            If InStr(ad, CStr(parcalar(t))) > 0 Then
                AdayBaslik = SutunEtiketi(i, GuvenliMetin(v(1, i)))
                Exit Function
            End If
        Next i
    Next t
End Function


Public Sub AnalizSablonKaydet()

    Dim ws As Worksheet, sb As Worksheet
    Dim ad As String
    Dim r As Long, i As Long, son As Long, j As Long

    Set ws = AnalizSayfasi()
    ad = Trim$(InputBox("Analiz sablonu hangi adla kaydedilsin?" & vbNewLine & _
                        "(A1-A9 hazir sablon adlari kullanilamaz; ayni ad varsa uzerine yazilir)", "Sablon kaydet"))
    If ad = "" Then Exit Sub
    If SablonKodu(ad) > 0 Then
        Bildir BS_DIKKAT, "Analiz", "Sablon adi 'A' + rakamla baslayamaz (hazir sablonlar icin ayrildi): " & ad
        Exit Sub
    End If
    EkranKilitle
    On Error GoTo Cikis
    Set sb = SayfaOlustur(SAYFA_SABLON)
    sb.Visible = xlSheetVeryHidden
    If Trim$(GuvenliMetin(sb.Cells(1, 1).Value)) = "" Then
        sb.Range("A1:L1").Value = Array("AD", "KAYNAK", "SATIR", "SUTUN", "OLCU", "OLCU SUTUNU", _
                                        "FILTRE SUTUNU", "FILTRE", "DONEM SUTUNU", "DONEM BAS", "DONEM SON", "KAYIT")
    End If
    son = sb.Cells(sb.Rows.Count, 1).End(xlUp).Row
    r = son + 1
    For i = 2 To son
        If BH(Trim$(GuvenliMetin(sb.Cells(i, 1).Value))) = BH(ad) Then r = i
    Next i
    sb.Cells(r, 1).Value = HucreMetni(ad)
    For j = AN_SATIR_KAYNAK To AN_SON_ALAN
        sb.Cells(r, 2 + j - AN_SATIR_KAYNAK).NumberFormat = "@"
        sb.Cells(r, 2 + j - AN_SATIR_KAYNAK).Value = ws.Cells(j, AN_DEGER_SUTUN).Value
    Next j
    sb.Cells(r, 12).Value = Now
    AnalizMenuleri
    ws.Cells(AN_SATIR_SABLON, AN_DEGER_SUTUN).Value = HucreMetni(ad)
    EkranAc
    Bildir BS_OK, "Analiz", "'" & ad & "' sablonu kaydedildi."
    Exit Sub
Cikis:
    EkranAc
    Bildir BS_ENGEL, "Analiz", "Sablon kaydedilemedi: " & Err.Description

End Sub


Public Sub AnalizSablonSil()

    Dim ws As Worksheet, sb As Worksheet
    Dim ad As String
    Dim i As Long, son As Long

    Set ws = AnalizSayfasi()
    If Not SayfaVarMi(SAYFA_SABLON) Then
        Bildir BS_BILGI, "Analiz", "Kayitli kullanici sablonu yok."
        Exit Sub
    End If
    ad = Trim$(InputBox("Silinecek sablonun adini yaz (onay icin):", "Sablon sil", Kart(ws, AN_SATIR_SABLON)))
    If ad = "" Then Exit Sub
    EkranKilitle
    On Error GoTo Cikis
    Set sb = ThisWorkbook.Worksheets(SAYFA_SABLON)
    son = sb.Cells(sb.Rows.Count, 1).End(xlUp).Row
    For i = son To 2 Step -1
        If BH(Trim$(GuvenliMetin(sb.Cells(i, 1).Value))) = BH(ad) Then sb.Rows(i).Delete
    Next i
    AnalizMenuleri
    EkranAc
    Bildir BS_OK, "Analiz", "'" & ad & "' sablonu silindi."
    Exit Sub
Cikis:
    EkranAc

End Sub


' ============================================================
'  CALISTIR
' ============================================================
Public Sub AnalizCalistir()

    Dim ws As Worksheet
    Dim kod As Long
    Dim bas As Double
    Dim sonuc As String, hata As String

    Set ws = AnalizSayfasi()
    kod = SablonKodu(Kart(ws, AN_SATIR_SABLON))
    bas = Timer
    EkranKilitle
    On Error GoTo Hata
    DurumCubugu "ANALIZ calisiyor ..."
    CiktiTemizle ws
    Select Case kod
        Case 1: sonuc = A1Profil(ws)
        Case 4: sonuc = A4Ucgen(ws)
        Case 5: sonuc = A5Gecikme(ws)
        Case 6: sonuc = A6BuyukHasar(ws)
        Case 7: sonuc = A7AktarimEtkisi(ws)
        Case 8: sonuc = A8KontrolTrendi(ws)
        Case Else: sonuc = GenelAnaliz(ws, kod)
    End Select
    GoTo Cikis
Hata:
    hata = Err.Number & " - " & Err.Description
    Resume Cikis
Cikis:
    On Error Resume Next
    DurumCubugu ""
    EkranAc
    ws.Activate
    ActiveWindow.ScrollRow = 1
    If hata <> "" Then
        Bildir BS_ENGEL, "Analiz", "Analiz tamamlanamadi: " & hata
    ElseIf Left$(sonuc, 1) = "!" Then
        Bildir BS_DIKKAT, "Analiz", Mid$(sonuc, 2)
    Else
        Bildir BS_OK, "Analiz", sonuc & "  (" & Format$(SureSn(bas), "0.0") & " sn)"
    End If

End Sub


Private Function SureSn(ByVal bas As Double) As Double
    SureSn = Timer - bas
    If SureSn < 0 Then SureSn = SureSn + 86400
End Function


Private Sub CiktiTemizle(ByVal ws As Worksheet)
    Dim i As Long, r As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = AN_GRAFIK Then ws.ChartObjects(i).Delete
    Next i
    For r = AN_YORUM_ILK To AN_YORUM_ILK + 2
        ws.Cells(r, AN_ETIKET_SUTUN).Value = ""
        ws.Cells(r, AN_ETIKET_SUTUN).Font.Color = Palet(prMetin)
    Next r
    With ws.Range(ws.Cells(AN_TABLO_ILK - 1, 1), ws.Cells(AN_TABLO_ILK + 400, AN_LST_SUTUN - 1))
        .Clear
        .Font.Name = PANEL_YAZI_TIPI
        .Font.Size = YZ_GOVDE
    End With
End Sub


Private Sub Yorum(ByVal ws As Worksheet, ByVal n As Long, ByVal metin As String)
    If metin = "" Then Exit Sub
    ws.Cells(AN_YORUM_ILK + n - 1, AN_ETIKET_SUTUN).Value = HucreMetni(ChrW(8226) & "  " & metin)
End Sub


Private Sub SonucBasligi(ByVal ws As Worksheet, ByVal metin As String)
    With ws.Cells(AN_TABLO_ILK - 1, AN_TABLO_SUTUN)
        .Value = HucreMetni(metin)
        .Font.Size = YZ_DIPNOT
        .Font.Color = Palet(prIkincilMetin)
    End With
End Sub


' ------------------------------------------------------------
'  Okuma ve satir secimi (yalniz istenen sutunlar)
' ------------------------------------------------------------
Private Function SatirlariSec(ByVal ws As Worksheet, ByVal kws As Worksheet, ByVal ilk As Long, ByVal son As Long, _
                              ByRef secili() As Long, ByRef aciklama As String) As Long

    Dim cF As Long, cD As Long, i As Long, n As Long, m As Long
    Dim vF As Variant, vD As Variant
    Dim f As FiltreSutun
    Dim filtreVar As Boolean, donemVar As Boolean
    Dim basT As Double, sonT As Double, d As Double

    n = son - ilk + 1
    ReDim secili(1 To n)
    cF = SutunNo(Kart(ws, AN_SATIR_FILTRE_SUTUN))
    If cF > 0 And Kart(ws, AN_SATIR_FILTRE) <> "" Then
        filtreVar = True
        vF = SutunOku(kws, cF, ilk, son)
        f = FiltreAyristir(Kart(ws, AN_SATIR_FILTRE))
        FiltreBelirsizligiGider f, vF
        aciklama = aciklama & "; filtre " & SutunAdi(Kart(ws, AN_SATIR_FILTRE_SUTUN)) & " = " & Kart(ws, AN_SATIR_FILTRE)
    End If
    cD = SutunNo(Kart(ws, AN_SATIR_DONEM_SUTUN))
    basT = -1
    sonT = 3000000
    If cD > 0 Then
        If IsDate(ws.Cells(AN_SATIR_DONEM_BAS, AN_DEGER_SUTUN).Value) Then basT = CDbl(CDate(ws.Cells(AN_SATIR_DONEM_BAS, AN_DEGER_SUTUN).Value))
        If IsDate(ws.Cells(AN_SATIR_DONEM_SON, AN_DEGER_SUTUN).Value) Then sonT = CDbl(CDate(ws.Cells(AN_SATIR_DONEM_SON, AN_DEGER_SUTUN).Value)) + 0.99999
        If basT >= 0 Or sonT < 3000000 Then
            donemVar = True
            vD = SutunOku(kws, cD, ilk, son)
            aciklama = aciklama & "; donem " & SutunAdi(Kart(ws, AN_SATIR_DONEM_SUTUN)) & " " & _
                       IIf(basT >= 0, Format$(basT, "dd.mm.yyyy"), "...") & " - " & IIf(sonT < 3000000, Format$(Int(sonT), "dd.mm.yyyy"), "...")
        End If
    End If
    For i = 1 To n
        If filtreVar Then
            If Not FiltreUyar(f, vF(i, 1)) Then GoTo Sonraki
        End If
        If donemVar Then
            If Not TarihSayisi(vD(i, 1), d) Then GoTo Sonraki
            If d < basT Or d > sonT Then GoTo Sonraki
        End If
        m = m + 1
        secili(m) = i
Sonraki:
    Next i
    SatirlariSec = m

End Function


' Tarih tipli sutun mu (tabakali ornekle)?
Private Function TarihSutunuMu(ByVal kws As Worksheet, ByVal c As Long, ByVal ilk As Long, ByVal son As Long) As Boolean
    TarihSutunuMu = (VeriTipiDizi(TabakaliOrnek(kws, c, ilk, son)) = "TARIH")
End Function


Private Function GrupAnahtari(ByVal x As Variant, ByVal ayMi As Boolean) As String
    Dim d As Double
    If ayMi Then
        If TarihSayisi(x, d) Then
            GrupAnahtari = CStr(Year(d)) & "-" & Right$("0" & CStr(Month(d)), 2)
            Exit Function
        End If
    End If
    GrupAnahtari = Trim$(GuvenliMetin(x))
    If GrupAnahtari = "" Then GrupAnahtari = BOS_GRUP
End Function


' ------------------------------------------------------------
'  GENEL GRUPLAMA MOTORU
' ------------------------------------------------------------
Private Function GenelAnaliz(ByVal ws As Worksheet, ByVal kod As Long) As String

    Dim kws As Worksheet
    Dim ilk As Long, son As Long, cS As Long, cC As Long, cO As Long, m As Long, i As Long, r As Long
    Dim secili() As Long, grup() As Long
    Dim vS As Variant, vC As Variant, vO As Variant
    Dim olcu As String, aciklama As String, anahtar As String
    Dim satirAy As Boolean, sutunAy As Boolean, olcuTarih As Boolean
    Dim anahtarlar() As String
    Dim degerler() As Double, sayisal() As Boolean
    Dim d As Double
    Dim bosBoyut As Long, sayiDegil As Long

    If Not KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        GenelAnaliz = "!Kaynak bulunamadi: Panel'de dosyalari secip YENILE calistir, sonra Kaynak menusunden sec."
        Exit Function
    End If
    cS = SutunNo(Kart(ws, AN_SATIR_SATIR))
    cC = SutunNo(Kart(ws, AN_SATIR_SUTUN))
    cO = SutunNo(Kart(ws, AN_SATIR_OLCU_SUTUN))
    olcu = Kart(ws, AN_SATIR_OLCU)
    If olcu = "" Then olcu = OLCU_ADET
    If cS = 0 Then
        GenelAnaliz = "!Satir boyutu secilmemis."
        Exit Function
    End If
    If olcu <> OLCU_ADET And cO = 0 Then
        GenelAnaliz = "!'" & olcu & "' icin olcu sutunu secilmemis."
        Exit Function
    End If

    m = SatirlariSec(ws, kws, ilk, son, secili, aciklama)
    If m = 0 Then
        GenelAnaliz = "!Filtreye / doneme uyan satir yok."
        Exit Function
    End If
    satirAy = TarihSutunuMu(kws, cS, ilk, son)
    vS = SutunOku(kws, cS, ilk, son)
    If cC > 0 Then
        sutunAy = TarihSutunuMu(kws, cC, ilk, son)
        vC = SutunOku(kws, cC, ilk, son)
    End If
    If cO > 0 Then
        vO = SutunOku(kws, cO, ilk, son)
        olcuTarih = TarihSutunuMu(kws, cO, ilk, son)
    End If

    ' --- anahtarlar ve olcu degerleri (secili satirlar) ---
    ReDim anahtarlar(1 To m)
    ReDim degerler(1 To m)
    ReDim sayisal(1 To m)
    For i = 1 To m
        r = secili(i)
        anahtar = GrupAnahtari(vS(r, 1), satirAy)
        If anahtar = BOS_GRUP Then bosBoyut = bosBoyut + 1
        If cC > 0 Then anahtar = anahtar & AYRAC & GrupAnahtari(vC(r, 1), sutunAy)
        anahtarlar(i) = anahtar
        If cO > 0 Then
            If olcu = OLCU_TEKIL Then
                sayisal(i) = True
            ElseIf SayiDegeri(vO(r, 1), d) Then
                degerler(i) = d
                sayisal(i) = True
            ElseIf Trim$(GuvenliMetin(vO(r, 1))) <> "" Then
                sayiDegil = sayiDegil + 1
            End If
        End If
    Next i

    Gruplandir anahtarlar, degerler, sayisal, m, grup
    If olcu = OLCU_TEKIL Or olcu = OLCU_P50 Or olcu = OLCU_P95 Then OzelOlcu grup, degerler, sayisal, vO, secili, m, olcu

    TabloYaz ws, olcu, (cC > 0), satirAy, olcuTarih, Kart(ws, AN_SATIR_SATIR), Kart(ws, AN_SATIR_SUTUN), Kart(ws, AN_SATIR_OLCU_SUTUN)
    SonucBasligi ws, TR("SONU#C  #*  ") & aciklama & "  " & ChrW(183) & "  " & SayiMetni(m) & " / " & SayiMetni(son - ilk + 1) & _
                     TR(" sat#ir  #*  ") & gSay & " grup"
    GenelYorum ws, olcu, satirAy, (cC > 0), m, son - ilk + 1, bosBoyut, sayiDegil, kod
    GenelAnaliz = olcu & " / " & SutunAdi(Kart(ws, AN_SATIR_SATIR)) & ": " & gSay & " grup, " & SayiMetni(m) & " satir"

End Function


' Anahtar -> grup. Karma (Collection) once; grup sayisi Ayarlar
' ANALIZ_GRUP esigini asarsa siralama tabanli gruplamaya gecilir.
Private Sub Gruplandir(anahtarlar() As String, degerler() As Double, sayisal() As Boolean, ByVal m As Long, grup() As Long)

    Dim karma As Collection
    Dim i As Long, g As Long, sinir As Long, j As Long
    Dim sira() As Long
    Dim tasti As Boolean

    sinir = CLng(Esik(ekAnalizGrup))
    If sinir < 1 Then sinir = 1
    ReDim grup(1 To m)
    GruplariSifirla sinir
    Set karma = New Collection
    For i = 1 To m
        g = 0
        On Error Resume Next
        g = karma.Item("k" & anahtarlar(i))
        On Error GoTo 0
        If g = 0 Then
            If gSay >= sinir Then
                tasti = True
                Exit For
            End If
            gSay = gSay + 1
            g = gSay
            karma.Add g, "k" & anahtarlar(i)
            GrupAdiAta g, anahtarlar(i)
        End If
        grup(i) = g
    Next i

    If tasti Then
        ' --- yuksek kardinalite: siralama tabanli ---
        ReDim sira(1 To m)
        For i = 1 To m
            sira(i) = i
        Next i
        HizliSiralaIndeks anahtarlar, sira, 1, m
        GruplariSifirla m
        For i = 1 To m
            j = sira(i)
            If i = 1 Then
                gSay = 1
                GrupAdiAta 1, anahtarlar(j)
            ElseIf anahtarlar(j) <> anahtarlar(sira(i - 1)) Then
                gSay = gSay + 1
                GrupAdiAta gSay, anahtarlar(j)
            End If
            grup(j) = gSay
        Next i
    End If

    For i = 1 To m
        g = grup(i)
        gN(g) = gN(g) + 1
        If sayisal(i) Then
            If gSayisal(g) = 0 Then
                gMin(g) = degerler(i)
                gMax(g) = degerler(i)
            Else
                If degerler(i) < gMin(g) Then gMin(g) = degerler(i)
                If degerler(i) > gMax(g) Then gMax(g) = degerler(i)
            End If
            gSayisal(g) = gSayisal(g) + 1
            gT(g) = gT(g) + degerler(i)
        End If
    Next i

End Sub


Private Sub GruplariSifirla(ByVal n As Long)
    gSay = 0
    ReDim gSatir(1 To n)
    ReDim gSutun(1 To n)
    ReDim gN(1 To n)
    ReDim gSayisal(1 To n)
    ReDim gT(1 To n)
    ReDim gMin(1 To n)
    ReDim gMax(1 To n)
    ReDim gOzel(1 To n)
End Sub


Private Sub GrupAdiAta(ByVal g As Long, ByVal anahtar As String)
    Dim p As Long
    p = InStr(anahtar, AYRAC)
    If p > 0 Then
        gSatir(g) = Left$(anahtar, p - 1)
        gSutun(g) = Mid$(anahtar, p + 1)
    Else
        gSatir(g) = anahtar
        gSutun(g) = ""
    End If
End Sub


' Tekil / P50 / P95: gruplar sayma siralamasiyla bitisik dilimlere
' konur, her dilim kendi icinde siralanir.
Private Sub OzelOlcu(grup() As Long, degerler() As Double, sayisal() As Boolean, ByVal vO As Variant, _
                     secili() As Long, ByVal m As Long, ByVal olcu As String)

    Dim bas() As Long, dolu() As Long
    Dim i As Long, g As Long, a As Long, b As Long, k As Long, toplam As Long
    Dim dilim() As Double, metinler() As String
    Dim tekil As Long

    ReDim bas(1 To gSay + 1)
    ReDim dolu(1 To gSay)
    toplam = 1
    For g = 1 To gSay
        bas(g) = toplam
        toplam = toplam + gSayisal(g)
    Next g
    bas(gSay + 1) = toplam
    If olcu = OLCU_TEKIL Then
        ReDim metinler(1 To IIf(toplam > 1, toplam - 1, 1))
    Else
        ReDim dilim(1 To IIf(toplam > 1, toplam - 1, 1))
    End If
    For i = 1 To m
        If sayisal(i) Then
            g = grup(i)
            k = bas(g) + dolu(g)
            dolu(g) = dolu(g) + 1
            If olcu = OLCU_TEKIL Then
                metinler(k) = Trim$(GuvenliMetin(vO(secili(i), 1)))
            Else
                dilim(k) = degerler(i)
            End If
        End If
    Next i
    For g = 1 To gSay
        a = bas(g)
        b = bas(g + 1) - 1
        If b >= a Then
            If olcu = OLCU_TEKIL Then
                HizliSiralaMetin metinler, a, b
                tekil = 0
                For k = a To b
                    If metinler(k) <> "" Then
                        If k = a Then
                            tekil = tekil + 1
                        ElseIf metinler(k) <> metinler(k - 1) Then
                            tekil = tekil + 1
                        End If
                    End If
                Next k
                gOzel(g) = tekil
            Else
                HizliSiralaSayi dilim, a, b
                If olcu = OLCU_P50 Then
                    gOzel(g) = dilim(a + YuzdelikSira(b - a + 1, 0.5) - 1)
                Else
                    gOzel(g) = dilim(a + YuzdelikSira(b - a + 1, 0.95) - 1)
                End If
            End If
        End If
    Next g

End Sub


' En yakin sira yontemi (nearest-rank): ceil(p * n)
Private Function YuzdelikSira(ByVal n As Long, ByVal p As Double) As Long
    YuzdelikSira = -Int(-p * n)
    If YuzdelikSira < 1 Then YuzdelikSira = 1
    If YuzdelikSira > n Then YuzdelikSira = n
End Function


Private Function GrupDegeri(ByVal g As Long, ByVal olcu As String) As Variant
    Select Case olcu
        Case OLCU_ADET:   GrupDegeri = gN(g)
        Case OLCU_TOPLAM: GrupDegeri = gT(g)
        Case OLCU_ORT
            If gSayisal(g) > 0 Then GrupDegeri = gT(g) / gSayisal(g) Else GrupDegeri = Empty
        Case OLCU_MIN
            If gSayisal(g) > 0 Then GrupDegeri = gMin(g) Else GrupDegeri = Empty
        Case OLCU_MAKS
            If gSayisal(g) > 0 Then GrupDegeri = gMax(g) Else GrupDegeri = Empty
        Case Else
            If gSayisal(g) > 0 Then GrupDegeri = gOzel(g) Else GrupDegeri = Empty
    End Select
End Function


Private Function Toplanabilir(ByVal olcu As String) As Boolean
    Toplanabilir = (olcu = OLCU_ADET Or olcu = OLCU_TOPLAM)
End Function


' ------------------------------------------------------------
'  TABLO + GRAFIK
' ------------------------------------------------------------
Private Sub TabloYaz(ByVal ws As Worksheet, ByVal olcu As String, ByVal pivot As Boolean, ByVal satirAy As Boolean, _
                     ByVal olcuTarih As Boolean, ByVal satirAd As String, ByVal sutunAd As String, ByVal olcuAd As String)

    Dim satirlar As Collection, sutunlar As Collection
    Dim sAd() As String, cAd() As String
    Dim sDeger() As Double, cDeger() As Double, sN() As Double
    Dim sSay As Long, cSay As Long, g As Long, i As Long, j As Long, r As Long, enFazla As Long
    Dim sSira() As Long, cSira() As Long, sIdx() As Long, cIdx() As Long
    Dim cikti() As Variant
    Dim toplamDeger As Double, toplamN As Double, x As Variant
    Dim bicim As String, idx As Long

    enFazla = CLng(Esik(ekAnalizSatir))
    If enFazla < 1 Then enFazla = 1
    bicim = IIf(olcu = OLCU_ADET Or olcu = OLCU_TEKIL, "#,##0", IIf(olcuTarih And olcu <> OLCU_TOPLAM And olcu <> OLCU_ORT, "dd.mm.yyyy", "#,##0.00"))

    ' --- satir ve sutun degerleri ---
    Set satirlar = New Collection
    Set sutunlar = New Collection
    ReDim sAd(1 To gSay)
    ReDim sDeger(1 To gSay)
    ReDim sN(1 To gSay)
    ReDim cAd(1 To gSay)
    ReDim cDeger(1 To gSay)
    ReDim sIdx(1 To gSay)
    ReDim cIdx(1 To gSay)
    For g = 1 To gSay
        idx = 0
        On Error Resume Next
        idx = satirlar.Item("k" & gSatir(g))
        On Error GoTo 0
        If idx = 0 Then
            sSay = sSay + 1
            idx = sSay
            satirlar.Add idx, "k" & gSatir(g)
            sAd(idx) = gSatir(g)
        End If
        sIdx(g) = idx
        sN(idx) = sN(idx) + gN(g)
        x = GrupDegeri(g, olcu)
        If Not IsEmpty(x) Then sDeger(idx) = sDeger(idx) + CDbl(x)
        If pivot Then
            idx = 0
            On Error Resume Next
            idx = sutunlar.Item("k" & gSutun(g))
            On Error GoTo 0
            If idx = 0 Then
                cSay = cSay + 1
                idx = cSay
                sutunlar.Add idx, "k" & gSutun(g)
                cAd(idx) = gSutun(g)
            End If
            cIdx(g) = idx
            cDeger(idx) = cDeger(idx) + gN(g)
        End If
    Next g
    If Not pivot Then
        ' tek boyutta grup = satir: toplanamayan olcu dogrudan grup degeridir
        For g = 1 To gSay
            x = GrupDegeri(g, olcu)
            If IsEmpty(x) Then sDeger(sIdx(g)) = 0 Else sDeger(sIdx(g)) = CDbl(x)
        Next g
    End If

    ' --- siralama: tarih (ay) boyutu kronolojik, digerleri olcuye gore azalan ---
    ReDim sSira(1 To sSay)
    For i = 1 To sSay
        sSira(i) = i
    Next i
    If satirAy Then
        HizliSiralaIndeks sAd, sSira, 1, sSay
    Else
        SiralaAzalan sDeger, sSira, sSay
    End If
    If pivot Then
        ReDim cSira(1 To cSay)
        For i = 1 To cSay
            cSira(i) = i
        Next i
        SiralaAzalan cDeger, cSira, cSay
        If cSay > AN_EN_FAZLA_SUTUN Then cSay = AN_EN_FAZLA_SUTUN
    End If

    ' --- tablo ---
    r = AN_TABLO_ILK
    If pivot Then
        ReDim cikti(1 To IIf(sSay > enFazla, enFazla + 1, sSay) + 1, 1 To cSay + 2)
        cikti(1, 1) = SutunAdi(satirAd) & "  \  " & SutunAdi(sutunAd)
        For j = 1 To cSay
            cikti(1, j + 1) = cAd(cSira(j))
        Next j
        cikti(1, cSay + 2) = TR("Toplam")
        For i = 1 To sSay
            If i <= enFazla Then
                cikti(i + 1, 1) = sAd(sSira(i))
            Else
                cikti(enFazla + 2, 1) = TR("Di#ger (") & (sSay - enFazla) & ")"
            End If
        Next i
        For g = 1 To gSay
            i = PozisyonBul(sSira, sSay, sIdx(g))
            j = PozisyonBul(cSira, cSay, cIdx(g))
            If i > enFazla Then i = enFazla + 1
            x = GrupDegeri(g, olcu)
            If Not IsEmpty(x) And (i <= enFazla Or Toplanabilir(olcu)) Then
                If j > 0 Then cikti(i + 1, j + 1) = SayiTopla(cikti(i + 1, j + 1), x, olcu)
                If Toplanabilir(olcu) Then cikti(i + 1, cSay + 2) = SayiTopla(cikti(i + 1, cSay + 2), x, olcu)
            End If
        Next g
        ws.Cells(r, AN_TABLO_SUTUN).Resize(UBound(cikti, 1), UBound(cikti, 2)).Value = cikti
        ws.Cells(r + 1, AN_TABLO_SUTUN + 1).Resize(UBound(cikti, 1) - 1, UBound(cikti, 2) - 1).NumberFormat = bicim
        TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, UBound(cikti, 2))
        TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(UBound(cikti, 1) - 1, UBound(cikti, 2))
        GrafikCiz ws, ws.Cells(r, AN_TABLO_SUTUN).Resize(IIf(UBound(cikti, 1) > GRAFIK_SATIR + 1, GRAFIK_SATIR + 1, UBound(cikti, 1)), _
                  IIf(cSay > 5, 6, cSay + 1)), satirAy, r + UBound(cikti, 1) + 2, True
    Else
        ReDim cikti(1 To IIf(sSay > enFazla, enFazla + 1, sSay) + 2, 1 To 4)
        cikti(1, 1) = SutunAdi(satirAd)
        cikti(1, 2) = TR("Sat#ir")
        cikti(1, 3) = olcu & IIf(olcu <> OLCU_ADET And olcuAd <> "", " " & SutunAdi(olcuAd), "")
        cikti(1, 4) = TR("Pay")
        For i = 1 To sSay
            toplamN = toplamN + sN(i)
            toplamDeger = toplamDeger + sDeger(i)
        Next i
        For i = 1 To sSay
            If i <= enFazla Then
                cikti(i + 1, 1) = sAd(sSira(i))
                cikti(i + 1, 2) = sN(sSira(i))
                cikti(i + 1, 3) = sDeger(sSira(i))
                If Toplanabilir(olcu) And toplamDeger <> 0 Then cikti(i + 1, 4) = sDeger(sSira(i)) / toplamDeger
            Else
                cikti(enFazla + 2, 1) = TR("Di#ger (") & (sSay - enFazla) & ")"
                cikti(enFazla + 2, 2) = SayiTopla(cikti(enFazla + 2, 2), sN(sSira(i)), OLCU_ADET)
                If Toplanabilir(olcu) Then
                    cikti(enFazla + 2, 3) = SayiTopla(cikti(enFazla + 2, 3), sDeger(sSira(i)), olcu)
                    If toplamDeger <> 0 Then cikti(enFazla + 2, 4) = cikti(enFazla + 2, 3) / toplamDeger
                End If
            End If
        Next i
        cikti(UBound(cikti, 1), 1) = TR("Toplam")
        cikti(UBound(cikti, 1), 2) = toplamN
        If Toplanabilir(olcu) Then
            cikti(UBound(cikti, 1), 3) = toplamDeger
            cikti(UBound(cikti, 1), 4) = 1
        End If
        ws.Cells(r, AN_TABLO_SUTUN).Resize(UBound(cikti, 1), 4).Value = cikti
        ws.Cells(r + 1, AN_TABLO_SUTUN + 1).Resize(UBound(cikti, 1) - 1, 1).NumberFormat = "#,##0"
        ws.Cells(r + 1, AN_TABLO_SUTUN + 2).Resize(UBound(cikti, 1) - 1, 1).NumberFormat = bicim
        ws.Cells(r + 1, AN_TABLO_SUTUN + 3).Resize(UBound(cikti, 1) - 1, 1).NumberFormat = "0.0%"
        TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 4)
        TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(UBound(cikti, 1) - 1, 4)
        ws.Cells(r + UBound(cikti, 1) - 1, AN_TABLO_SUTUN).Resize(1, 4).Font.Bold = True
        GrafikCiz ws, Union(ws.Cells(r, AN_TABLO_SUTUN).Resize(IIf(UBound(cikti, 1) - 1 > GRAFIK_SATIR + 1, GRAFIK_SATIR + 1, UBound(cikti, 1) - 1), 1), _
                            ws.Cells(r, AN_TABLO_SUTUN + 2).Resize(IIf(UBound(cikti, 1) - 1 > GRAFIK_SATIR + 1, GRAFIK_SATIR + 1, UBound(cikti, 1) - 1), 1)), _
                  satirAy, r, False
    End If

End Sub


Private Function SayiTopla(ByVal mevcut As Variant, ByVal ek As Variant, ByVal olcu As String) As Variant
    If IsEmpty(mevcut) Then
        SayiTopla = CDbl(ek)
    Else
        SayiTopla = CDbl(mevcut) + CDbl(ek)
    End If
End Function


Private Function PozisyonBul(sira() As Long, ByVal n As Long, ByVal idx As Long) As Long
    Dim i As Long
    For i = 1 To n
        If sira(i) = idx Then
            PozisyonBul = i
            Exit Function
        End If
    Next i
    PozisyonBul = UBound(sira) + 1
End Function


Private Sub TabloCizgileri(ByVal rng As Range)
    With rng
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = Palet(prCizgi)
        If .Rows.Count > 1 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = Palet(prCizgi)
        End If
    End With
End Sub


' Grafik: ay boyutu cizgi, kategori cubuk; marka rengi + notr
Private Sub GrafikCiz(ByVal ws As Worksheet, ByVal kaynak As Range, ByVal zaman As Boolean, ByVal ustSatir As Long, _
                      ByVal altta As Boolean)

    Dim grafik As ChartObject
    Dim sol As Double, ust As Double
    Dim i As Long

    If kaynak.Rows.Count < 2 Then Exit Sub
    If altta Then
        sol = ws.Cells(ustSatir, AN_TABLO_SUTUN).Left
        ust = ws.Cells(ustSatir, AN_TABLO_SUTUN).Top
    Else
        sol = ws.Cells(ustSatir, AN_TABLO_SUTUN + 5).Left
        ust = ws.Cells(ustSatir, 1).Top
    End If
    Set grafik = ws.ChartObjects.Add(sol, ust, 460, 250)
    grafik.Name = AN_GRAFIK
    With grafik.Chart
        .SetSourceData kaynak
        If zaman Then
            .ChartType = xlLineMarkers
        ElseIf altta Then
            .ChartType = xlColumnClustered
        Else
            .ChartType = xlBarClustered
            .Axes(1).ReversePlotOrder = True
        End If
        For i = 1 To .SeriesCollection.Count
            If i = 1 Then
                .SeriesCollection(i).Format.Fill.ForeColor.RGB = Palet(prMarka)
                .SeriesCollection(i).Format.Line.ForeColor.RGB = Palet(prMarka)
            Else
                .SeriesCollection(i).Format.Fill.ForeColor.RGB = Palet(prIkincilMetin)
                .SeriesCollection(i).Format.Line.ForeColor.RGB = Palet(prIkincilMetin)
            End If
        Next i
        .HasTitle = False
        .HasLegend = altta
        .ChartArea.Font.Name = PANEL_YAZI_TIPI
        .ChartArea.Font.Size = YZ_DIPNOT
        .ChartArea.Format.Line.ForeColor.RGB = Palet(prCizgi)
    End With

End Sub


' ------------------------------------------------------------
'  OTOMATIK YORUM (3 madde)
' ------------------------------------------------------------
Private Sub GenelYorum(ByVal ws As Worksheet, ByVal olcu As String, ByVal satirAy As Boolean, ByVal pivot As Boolean, _
                       ByVal m As Long, ByVal n As Long, ByVal bosBoyut As Long, ByVal sayiDegil As Long, ByVal kod As Long)

    Dim r As Long, son As Long, i As Long, k As Long
    Dim ilk As Double, toplam As Double, ilk3 As Double, ort As Double, sapma As Double, z As Double, x As Double
    Dim ad As String, metin As String
    Dim ozet As String

    r = AN_TABLO_ILK + 1
    son = ws.Cells(ws.Rows.Count, AN_TABLO_SUTUN).End(xlUp).Row - 1
    If pivot Then
        ' pivot: satir toplamlari son sutunda
        Yorum ws, 1, "Satir x sutun kirilimi; satirlar ve sutunlar " & olcu & "a gore buyukten kucuge, en fazla " & _
                     CLng(Esik(ekAnalizSatir)) & " satir ve " & AN_EN_FAZLA_SUTUN & " sutun (gerisi 'Diger')."
    ElseIf satirAy Then
        ' zaman serisi: son ay, onceki aylarin ortalamasi ve z-skor
        k = son - r + 1
        If k >= 3 Then
            For i = r To son - 1
                ort = ort + CDbl(ws.Cells(i, AN_TABLO_SUTUN + 2).Value)
            Next i
            ort = ort / (k - 1)
            For i = r To son - 1
                sapma = sapma + (CDbl(ws.Cells(i, AN_TABLO_SUTUN + 2).Value) - ort) ^ 2
            Next i
            sapma = Sqr(sapma / (k - 2))
            x = CDbl(ws.Cells(son, AN_TABLO_SUTUN + 2).Value)
            ad = GuvenliMetin(ws.Cells(son, AN_TABLO_SUTUN).Value)
            If sapma > 0 Then
                z = (x - ort) / sapma
                Yorum ws, 1, "Son donem (" & ad & ") " & olcu & " " & Format$(x, "#,##0") & "; onceki " & (k - 1) & _
                             " donem ortalamasi " & Format$(ort, "#,##0") & " (z = " & Format$(z, "+0.0;-0.0") & ")" & _
                             IIf(Abs(z) > Esik(ekZSkor), "  ->  esik " & Format$(Esik(ekZSkor), "0.0") & " asildi, kontrol edin.", ", olagan aralikta.")
            Else
                Yorum ws, 1, "Son donem (" & ad & ") " & olcu & " " & Format$(x, "#,##0") & "; onceki donemler sabit."
            End If
        Else
            Yorum ws, 1, "Zaman serisi icin en az 3 donem gerekir; " & k & " donem var."
        End If
    Else
        For i = r To son
            If IsNumeric(ws.Cells(i, AN_TABLO_SUTUN + 2).Value) Then toplam = toplam + CDbl(ws.Cells(i, AN_TABLO_SUTUN + 2).Value)
        Next i
        If son >= r And toplam <> 0 And Toplanabilir(olcu) Then
            ilk = CDbl(ws.Cells(r, AN_TABLO_SUTUN + 2).Value)
            For i = r To r + 2
                If i <= son Then ilk3 = ilk3 + CDbl(ws.Cells(i, AN_TABLO_SUTUN + 2).Value)
            Next i
            Yorum ws, 1, "En buyuk grup " & GuvenliMetin(ws.Cells(r, AN_TABLO_SUTUN).Value) & ": " & olcu & "in %" & _
                         Format$(ilk / toplam * 100, "0") & "'i; ilk 3 grup %" & Format$(ilk3 / toplam * 100, "0") & _
                         IIf(ilk3 / toplam > 0.8, " (yogunlasma yuksek).", ".")
        ElseIf son >= r Then
            Yorum ws, 1, "En yuksek " & olcu & ": " & GuvenliMetin(ws.Cells(r, AN_TABLO_SUTUN).Value) & " (" & _
                         Format$(ws.Cells(r, AN_TABLO_SUTUN + 2).Value, "#,##0.00") & "); en dusuk: " & _
                         GuvenliMetin(ws.Cells(son, AN_TABLO_SUTUN).Value) & " (" & Format$(ws.Cells(son, AN_TABLO_SUTUN + 2).Value, "#,##0.00") & ")."
        End If
    End If
    Yorum ws, 2, gSay & " grup; satirlarin %" & Format$(m / n * 100, "0.0") & "'i filtre ve doneme uydu (" & _
                 SayiMetni(m) & " / " & SayiMetni(n) & ")." & IIf(kod = 9, " Frekans = satir sayisi, siddet = ortalama tutar.", "")
    ozet = ""
    If bosBoyut > 0 Then ozet = ozet & SayiMetni(bosBoyut) & " satirda boyut bos ('" & BOS_GRUP & "'); "
    If sayiDegil > 0 Then ozet = ozet & SayiMetni(sayiDegil) & " satirda olcu sayi degil (hesaba girmedi); "
    If ozet = "" Then ozet = "Veri notu yok: boyut dolu, olcu sutunu tamamen sayisal."
    Yorum ws, 3, ozet
    If kod = 9 Then A9Ayrisim ws

End Sub


' ------------------------------------------------------------
'  A1  VERI PROFILI
' ------------------------------------------------------------
Private Function A1Profil(ByVal ws As Worksheet) As String

    Dim kws As Worksheet
    Dim ilk As Long, son As Long, sonSut As Long, c As Long, i As Long, r As Long, m As Long, pencere As Long
    Dim v As Variant, basliklar As Variant
    Dim secili() As Long
    Dim cikti() As Variant
    Dim dolu As Long, sayi As Long, tarih As Long, enDusukDolu As Double, enDusukAd As String
    Dim degerler() As Double, k As Long
    Dim tekil As Collection
    Dim tekilSay As Long, tekilTasti As Boolean
    Dim aciklama As String, x As Variant, d As Double
    Dim zorunluAlti As Long

    If Not KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        A1Profil = "!Kaynak bulunamadi."
        Exit Function
    End If
    pencere = CLng(Esik(ekK12Pencere))
    If son - ilk + 1 > pencere And InStr(BH(Kart(ws, AN_SATIR_KAYNAK)), "ANA DOSYA") > 0 Then
        ilk = son - pencere + 1
        aciklama = aciklama & "; son " & SayiMetni(pencere) & " satir"
    End If
    m = SatirlariSec(ws, kws, ilk, son, secili, aciklama)
    If m = 0 Then
        A1Profil = "!Filtreye / doneme uyan satir yok."
        Exit Function
    End If
    sonSut = SonSutun(kws, EN_FAZLA_SUTUN)
    basliklar = IkiBoyutlu(kws.Range(kws.Cells(1, 1), kws.Cells(1, sonSut)))
    ReDim cikti(1 To sonSut + 1, 1 To 9)
    cikti(1, 1) = TR("S#UTUN")
    cikti(1, 2) = TR("T#IP")
    cikti(1, 3) = TR("DOLU %")
    cikti(1, 4) = TR("TEK#IL")
    cikti(1, 5) = "MIN"
    cikti(1, 6) = "P50"
    cikti(1, 7) = "P95"
    cikti(1, 8) = "MAKS"
    cikti(1, 9) = TR("TOPLAM")
    enDusukDolu = 2
    For c = 1 To sonSut
        DurumCubugu "ANALIZ A1: " & c & "/" & sonSut
        v = SutunOku(kws, c, ilk, son)
        dolu = 0: sayi = 0: tarih = 0: k = 0: tekilSay = 0: tekilTasti = False
        ReDim degerler(1 To m)
        Set tekil = New Collection
        For i = 1 To m
            x = v(secili(i), 1)
            If Trim$(GuvenliMetin(x)) <> "" Then
                dolu = dolu + 1
                If VarType(x) = vbDate Then tarih = tarih + 1
                If SayiDegeri(x, d) Then
                    sayi = sayi + 1
                    k = k + 1
                    degerler(k) = d
                End If
                If Not tekilTasti Then
                    On Error Resume Next
                    tekil.Add 1, "k" & Trim$(GuvenliMetin(x))
                    If Err.Number = 0 Then tekilSay = tekilSay + 1
                    Err.Clear
                    On Error GoTo 0
                    If tekilSay > 100000 Then tekilTasti = True
                End If
            End If
        Next i
        r = c + 1
        cikti(r, 1) = SutunEtiketi(c, GuvenliMetin(basliklar(1, c)))
        If dolu = 0 Then
            cikti(r, 2) = "BOS"
        ElseIf tarih / dolu >= 0.6 Then
            cikti(r, 2) = "TARIH"
        ElseIf sayi / dolu >= 0.6 Then
            cikti(r, 2) = "SAYI"
        Else
            cikti(r, 2) = "METIN"
        End If
        cikti(r, 3) = dolu / m
        cikti(r, 4) = IIf(tekilTasti, ">100.000", tekilSay)
        If k > 0 And cikti(r, 2) <> "METIN" Then
            HizliSiralaSayi degerler, 1, k
            cikti(r, 5) = degerler(1)
            cikti(r, 6) = degerler(YuzdelikSira(k, 0.5))
            cikti(r, 7) = degerler(YuzdelikSira(k, 0.95))
            cikti(r, 8) = degerler(k)
            If cikti(r, 2) = "SAYI" Then
                For i = 1 To k
                    d = d + degerler(i)
                Next i
                cikti(r, 9) = d
                d = 0
            End If
        End If
        If dolu / m < enDusukDolu Then
            enDusukDolu = dolu / m
            enDusukAd = cikti(r, 1)
        End If
        If dolu / m * 100 < Esik(ekTamlikHedef) Then zorunluAlti = zorunluAlti + 1
    Next c
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(sonSut + 1, 9).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 9)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(sonSut, 9)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 2).Resize(sonSut, 1).NumberFormat = "0.0%"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 3).Resize(sonSut, 1).NumberFormat = "#,##0"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 4).Resize(sonSut, 5).NumberFormat = "#,##0.00"
    For c = 1 To sonSut
        If cikti(c + 1, 2) = "TARIH" Then ws.Cells(r + c, AN_TABLO_SUTUN + 4).Resize(1, 4).NumberFormat = "dd.mm.yyyy"
        If IsNumeric(cikti(c + 1, 3)) Then
            If cikti(c + 1, 3) * 100 < Esik(ekTamlikHedef) Then
                ws.Cells(r + c, AN_TABLO_SUTUN + 2).Interior.Color = Palet(prDikkatZemin)
                ws.Cells(r + c, AN_TABLO_SUTUN + 2).Font.Color = Palet(prDikkat)
            End If
        End If
    Next c
    SonucBasligi ws, TR("VER#I PROF#IL#I  #*  ") & aciklama & "  " & ChrW(183) & "  " & SayiMetni(m) & TR(" sat#ir")
    Yorum ws, 1, sonSut & " sutun; tamlik hedefinin (%" & Format$(Esik(ekTamlikHedef), "0.0") & ") altinda " & zorunluAlti & " sutun."
    Yorum ws, 2, "En dusuk doluluk: " & enDusukAd & " (%" & Format$(enDusukDolu * 100, "0.0") & ")."
    Yorum ws, 3, "Tip, satirlarin %60 kuralina gore; TEKIL 100.000 farkli degerde durur. Yuzdelikler en yakin sira yontemiyle."
    A1Profil = "A1 veri profili: " & sonSut & " sutun, " & SayiMetni(m) & " satir"

End Function


' ------------------------------------------------------------
'  A4  HASAR GELISIM UCGENI (zincir merdiven faktorleri)
' ------------------------------------------------------------
Private Function A4Ucgen(ByVal ws As Worksheet) As String

    Dim kws As Worksheet
    Dim ilk As Long, son As Long, cK As Long, cG As Long, cT As Long, m As Long, i As Long, r As Long
    Dim secili() As Long
    Dim vK As Variant, vG As Variant, vT As Variant
    Dim yilMin As Long, yilMax As Long, ky As Long, gy As Long, n As Long, j As Long, gecersiz As Long
    Dim ucgen() As Double, kum() As Double, faktor() As Double
    Dim cikti() As Variant
    Dim dK As Double, dG As Double, t As Double, pay As Double, payda As Double
    Dim aciklama As String

    If Not KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        A4Ucgen = "!Kaynak bulunamadi."
        Exit Function
    End If
    cK = SutunNo(Kart(ws, AN_SATIR_SATIR))
    cG = SutunNo(Kart(ws, AN_SATIR_DONEM_SUTUN))
    cT = SutunNo(Kart(ws, AN_SATIR_OLCU_SUTUN))
    If cK = 0 Or cG = 0 Or cT = 0 Then
        A4Ucgen = "!A4 icin: Satir boyutu = hasar (kaza) tarihi, Donem sutunu = odeme/rapor tarihi, Olcu sutunu = tutar gerekli."
        Exit Function
    End If
    m = SatirlariSec(ws, kws, ilk, son, secili, aciklama)
    If m = 0 Then
        A4Ucgen = "!Filtreye / doneme uyan satir yok."
        Exit Function
    End If
    vK = SutunOku(kws, cK, ilk, son)
    vG = SutunOku(kws, cG, ilk, son)
    vT = SutunOku(kws, cT, ilk, son)
    yilMin = 9999
    For i = 1 To m
        r = secili(i)
        If TarihSayisi(vK(r, 1), dK) Then
            If Year(dK) < yilMin Then yilMin = Year(dK)
            If Year(dK) > yilMax Then yilMax = Year(dK)
        End If
    Next i
    If yilMax = 0 Then
        A4Ucgen = "!Kaza tarihi sutununda gecerli tarih yok."
        Exit Function
    End If
    If yilMax - yilMin > 19 Then yilMin = yilMax - 19
    n = yilMax - yilMin + 1
    ReDim ucgen(1 To n, 1 To n)
    For i = 1 To m
        r = secili(i)
        If TarihSayisi(vK(r, 1), dK) And TarihSayisi(vG(r, 1), dG) And SayiDegeri(vT(r, 1), t) Then
            ky = Year(dK) - yilMin + 1
            gy = Year(dG) - Year(dK) + 1
            If ky >= 1 And ky <= n And gy >= 1 And gy <= n - ky + 1 Then
                ucgen(ky, gy) = ucgen(ky, gy) + t
            Else
                gecersiz = gecersiz + 1
            End If
        Else
            gecersiz = gecersiz + 1
        End If
    Next i
    ReDim kum(1 To n, 1 To n)
    For ky = 1 To n
        For gy = 1 To n - ky + 1
            kum(ky, gy) = ucgen(ky, gy)
            If gy > 1 Then kum(ky, gy) = kum(ky, gy) + kum(ky, gy - 1)
        Next gy
    Next ky
    ' Hacim agirlikli yasa-yas faktorleri: f_j = sum C(i,j+1) / sum C(i,j)
    ReDim faktor(1 To n)
    For gy = 1 To n - 1
        pay = 0
        payda = 0
        For ky = 1 To n - gy
            pay = pay + kum(ky, gy + 1)
            payda = payda + kum(ky, gy)
        Next ky
        If payda <> 0 Then faktor(gy) = pay / payda
    Next gy
    ReDim cikti(1 To n + 3, 1 To n + 1)
    cikti(1, 1) = TR("KAZA YILI \ GEL#I#S#IM")
    For gy = 1 To n
        cikti(1, gy + 1) = gy
    Next gy
    For ky = 1 To n
        cikti(ky + 1, 1) = yilMin + ky - 1
        For gy = 1 To n - ky + 1
            cikti(ky + 1, gy + 1) = kum(ky, gy)
        Next gy
    Next ky
    cikti(n + 2, 1) = TR("Fakt#or (j #> j+1)")
    For gy = 1 To n - 1
        If faktor(gy) <> 0 Then cikti(n + 2, gy + 1) = faktor(gy)
    Next gy
    cikti(n + 3, 1) = TR("K#um#ulatif fakt#or (ultimate'a)")
    t = 1
    For gy = n - 1 To 1 Step -1
        If faktor(gy) <> 0 Then t = t * faktor(gy)
        cikti(n + 3, gy + 1) = t
    Next gy
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(n + 3, n + 1).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, n + 1)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(n + 2, n + 1)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 1).Resize(n, n).NumberFormat = "#,##0"
    ws.Cells(r + n + 1, AN_TABLO_SUTUN + 1).Resize(2, n).NumberFormat = "0.000"
    ws.Cells(r + n + 1, AN_TABLO_SUTUN).Resize(2, n + 1).Font.Bold = True
    For ky = 2 To n
        ws.Cells(r + ky, AN_TABLO_SUTUN + 1 + n - ky + 1).Resize(1, ky - 1).Interior.Color = Palet(prYuzey)
    Next ky
    GrafikCiz ws, ws.Cells(r + n + 1, AN_TABLO_SUTUN).Resize(1, n), False, r + n + 5, True
    SonucBasligi ws, TR("HASAR GEL#I#S#IM #U#CGEN#I (k#um#ulatif, y#ill#ik)  #*  ") & aciklama & "  " & ChrW(183) & "  " & _
                     SayiMetni(m) & TR(" sat#ir")
    Yorum ws, 1, n & " kaza yili (" & yilMin & "-" & yilMax & "); gelisim yili = donem yili - kaza yili + 1. " & _
                 "Faktorler hacim agirlikli yasa-yas (chain-ladder) ortalamasidir."
    If n >= 2 Then Yorum ws, 2, "1 -> 2 faktoru " & Format$(faktor(1), "0.000") & "; kumulatif (ultimate'a) " & _
                                 Format$(cikti(n + 3, 2), "0.000") & ". Kuyruk faktoru varsayilmadi (1,000)."
    Yorum ws, 3, IIf(gecersiz > 0, SayiMetni(gecersiz) & " satir ucgen disinda (gecersiz tarih/tutar veya donem < kaza).", _
                     "Tum satirlar ucgene girdi.") & " Bu bir veri kontrolu araci; rezerv tahmini aktuerin yargisini gerektirir."
    A4Ucgen = "A4 gelisim ucgeni: " & n & " kaza yili"

End Function


' ------------------------------------------------------------
'  A5  IHBAR GECIKMESI
' ------------------------------------------------------------
Private Function A5Gecikme(ByVal ws As Worksheet) As String

    Dim kws As Worksheet
    Dim ilk As Long, son As Long, cH As Long, cI As Long, m As Long, i As Long, r As Long, k As Long, b As Long
    Dim secili() As Long
    Dim vH As Variant, vI As Variant
    Dim gun() As Double
    Dim dH As Double, dI As Double, toplam As Double
    Dim sinirlar As Variant, adlar As Variant
    Dim say(1 To 8) As Double
    Dim cikti() As Variant
    Dim aciklama As String

    If Not KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        A5Gecikme = "!Kaynak bulunamadi."
        Exit Function
    End If
    cH = SutunNo(Kart(ws, AN_SATIR_SATIR))
    cI = SutunNo(Kart(ws, AN_SATIR_SUTUN))
    If cH = 0 Or cI = 0 Then
        A5Gecikme = "!A5 icin: Satir boyutu = hasar tarihi, Sutun boyutu = ihbar tarihi gerekli."
        Exit Function
    End If
    m = SatirlariSec(ws, kws, ilk, son, secili, aciklama)
    If m = 0 Then
        A5Gecikme = "!Filtreye / doneme uyan satir yok."
        Exit Function
    End If
    vH = SutunOku(kws, cH, ilk, son)
    vI = SutunOku(kws, cI, ilk, son)
    ReDim gun(1 To m)
    sinirlar = Array(-1, 30, 90, 180, 365, 730, 1E+30)
    adlar = Array(TR("negatif (ihbar < hasar)"), TR("0 - 30 g#un"), TR("31 - 90 g#un"), TR("91 - 180 g#un"), _
                  TR("181 - 365 g#un"), TR("366 - 730 g#un"), TR("730 g#un #ust#u"), TR("tarih okunamad#i"))
    For i = 1 To m
        r = secili(i)
        If TarihSayisi(vH(r, 1), dH) And TarihSayisi(vI(r, 1), dI) Then
            k = k + 1
            gun(k) = Int(dI) - Int(dH)
            toplam = toplam + gun(k)
            For b = 0 To 6
                If gun(k) <= sinirlar(b) Then
                    say(b + 1) = say(b + 1) + 1
                    Exit For
                End If
            Next b
        Else
            say(8) = say(8) + 1
        End If
    Next i
    ReDim cikti(1 To 10, 1 To 3)
    cikti(1, 1) = TR("GEC#IKME")
    cikti(1, 2) = TR("Sat#ir")
    cikti(1, 3) = "Pay"
    For b = 1 To 8
        cikti(b + 1, 1) = adlar(b - 1)
        cikti(b + 1, 2) = say(b)
        cikti(b + 1, 3) = say(b) / m
    Next b
    cikti(10, 1) = TR("Toplam")
    cikti(10, 2) = m
    cikti(10, 3) = 1
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(10, 3).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 3)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(9, 3)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 1).Resize(9, 1).NumberFormat = "#,##0"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 2).Resize(9, 1).NumberFormat = "0.0%"
    ws.Cells(r + 9, AN_TABLO_SUTUN).Resize(1, 3).Font.Bold = True
    GrafikCiz ws, ws.Cells(r, AN_TABLO_SUTUN).Resize(8, 2), False, r, False
    SonucBasligi ws, TR("#IHBAR GEC#IKMES#I (ihbar - hasar, g#un)  #*  ") & aciklama
    If k > 0 Then
        HizliSiralaSayi gun, 1, k
        Yorum ws, 1, "Ortalama gecikme " & Format$(toplam / k, "0.0") & " gun; P50 " & Format$(gun(YuzdelikSira(k, 0.5)), "0") & _
                     ", P90 " & Format$(gun(YuzdelikSira(k, 0.9)), "0") & ", en uzun " & Format$(gun(k), "#,##0") & " gun."
    End If
    Yorum ws, 2, IIf(say(1) > 0, SayiMetni(say(1)) & " satirda ihbar hasardan ONCE (K8 ile ayni bulgu).", "Negatif gecikme yok.")
    Yorum ws, 3, "P90'in aylik trendi IBNR varsayimini dogrudan etkiler; uzun kuyruk (" & SayiMetni(say(7)) & _
                 " satir > 730 gun) ayrica incelenmeli."
    A5Gecikme = "A5 ihbar gecikmesi: " & SayiMetni(k) & " satir"

End Function


' ------------------------------------------------------------
'  A6  BUYUK HASAR TOP-N
' ------------------------------------------------------------
Private Function A6BuyukHasar(ByVal ws As Worksheet) As String

    Dim kws As Worksheet
    Dim ilk As Long, son As Long, cA As Long, cK As Long, cT As Long, m As Long, i As Long, r As Long, n As Long
    Dim secili() As Long, sira() As Long
    Dim vA As Variant, vK As Variant, vT As Variant
    Dim tutar() As Double
    Dim toplam As Double, ust As Double, d As Double
    Dim cikti() As Variant
    Dim aciklama As String

    If Not KaynakCoz(Kart(ws, AN_SATIR_KAYNAK), kws, ilk, son, aciklama) Then
        A6BuyukHasar = "!Kaynak bulunamadi."
        Exit Function
    End If
    cA = SutunNo(Kart(ws, AN_SATIR_SATIR))
    cK = SutunNo(Kart(ws, AN_SATIR_SUTUN))
    cT = SutunNo(Kart(ws, AN_SATIR_OLCU_SUTUN))
    If cT = 0 Then
        A6BuyukHasar = "!A6 icin Olcu sutunu (tutar) gerekli."
        Exit Function
    End If
    m = SatirlariSec(ws, kws, ilk, son, secili, aciklama)
    If m = 0 Then
        A6BuyukHasar = "!Filtreye / doneme uyan satir yok."
        Exit Function
    End If
    vT = SutunOku(kws, cT, ilk, son)
    If cA > 0 Then vA = SutunOku(kws, cA, ilk, son)
    If cK > 0 Then vK = SutunOku(kws, cK, ilk, son)
    ReDim tutar(1 To m)
    ReDim sira(1 To m)
    For i = 1 To m
        If SayiDegeri(vT(secili(i), 1), d) Then tutar(i) = d
        toplam = toplam + tutar(i)
        sira(i) = i
    Next i
    SiralaAzalan tutar, sira, m
    n = TOPN
    If n > m Then n = m
    ReDim cikti(1 To n + 1, 1 To 6)
    cikti(1, 1) = TR("SIRA")
    cikti(1, 2) = TR("SATIR (kaynak)")
    cikti(1, 3) = IIf(cA > 0, SutunAdi(Kart(ws, AN_SATIR_SATIR)), TR("ANAHTAR"))
    cikti(1, 4) = IIf(cK > 0, SutunAdi(Kart(ws, AN_SATIR_SUTUN)), TR("KIRILIM"))
    cikti(1, 5) = SutunAdi(Kart(ws, AN_SATIR_OLCU_SUTUN))
    cikti(1, 6) = TR("Pay")
    For i = 1 To n
        r = secili(sira(i))
        cikti(i + 1, 1) = i
        cikti(i + 1, 2) = ilk + r - 1
        If cA > 0 Then cikti(i + 1, 3) = HucreMetni(GuvenliMetin(vA(r, 1)))
        If cK > 0 Then cikti(i + 1, 4) = HucreMetni(GuvenliMetin(vK(r, 1)))
        cikti(i + 1, 5) = tutar(sira(i))
        If toplam <> 0 Then cikti(i + 1, 6) = tutar(sira(i)) / toplam
        ust = ust + tutar(sira(i))
    Next i
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(n + 1, 6).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 6)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(n, 6)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 4).Resize(n, 1).NumberFormat = "#,##0.00"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 5).Resize(n, 1).NumberFormat = "0.0%"
    SonucBasligi ws, TR("B#UY#UK HASAR TOP-") & n & TR("  #*  ") & aciklama
    If toplam <> 0 Then Yorum ws, 1, "En buyuk " & n & " satir toplamin %" & Format$(ust / toplam * 100, "0.0") & "'i; en buyuk tek tutar " & _
                                     Format$(tutar(sira(1)), "#,##0.00") & " (%" & Format$(tutar(sira(1)) / toplam * 100, "0.0") & ")."
    Yorum ws, 2, "Buyuk hasar esigi (Ayarlar K11_BUYUK_HASAR): " & Format$(Esik(ekBuyukHasar), "#,##0") & "; esigi asan satirlar K11'de DIKKAT olur."
    Yorum ws, 3, "SATIR sutunu kaynaktaki satir numarasidir; kaydi dogrudan bulmak icin kullanin."
    A6BuyukHasar = "A6 Top-" & n

End Function


' ------------------------------------------------------------
'  A7  AKTARIM ETKISI (Rapor: son etkin parti)
' ------------------------------------------------------------
Private Function A7AktarimEtkisi(ByVal ws As Worksheet) As String

    Dim rpr As Worksheet
    Dim v As Variant
    Dim son As Long, i As Long, sonParti As Long, n As Long, r As Long
    Dim aktifler As Sozluk
    Dim cikti() As Variant
    Dim toplamYazilan As Double, fark As Long

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    If son < 2 Then
        A7AktarimEtkisi = "!Rapor'da aktarim yok."
        Exit Function
    End If
    Set aktifler = AktifAktarimlar()
    v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))
    For i = 1 To UBound(v, 1)
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And aktifler.Exists(Trim$(GuvenliMetin(v(i, R_AKTARIM)))) Then
            If GuvenliSayi(v(i, R_PARTI)) > sonParti Then sonParti = GuvenliSayi(v(i, R_PARTI))
        End If
    Next i
    If sonParti = 0 Then
        A7AktarimEtkisi = "!Etkin aktarim yok."
        Exit Function
    End If
    ReDim cikti(1 To 2 * MAX_BLOK + 1, 1 To 9)
    cikti(1, 1) = "PANEL"
    cikti(1, 2) = TR("AKTARIM")
    cikti(1, 3) = TR("ANA DOSYA")
    cikti(1, 4) = TR("SATIR")
    cikti(1, 5) = TR("KAYNAK TOPLAM")
    cikti(1, 6) = TR("HEDEF #ONCE")
    cikti(1, 7) = TR("HEDEF SONRA")
    cikti(1, 8) = TR("ARTI#S %")
    cikti(1, 9) = "MUTABAKAT"
    n = 1
    For i = 1 To UBound(v, 1)
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And GuvenliSayi(v(i, R_PARTI)) = sonParti And n < UBound(cikti, 1) Then
            n = n + 1
            cikti(n, 1) = v(i, R_PANEL)
            cikti(n, 2) = v(i, R_AKTARIM)
            cikti(n, 3) = HucreMetni(GuvenliMetin(v(i, R_HEDEF_DOSYA)) & " / " & GuvenliMetin(v(i, R_HEDEF_SAYFA)))
            cikti(n, 4) = v(i, R_SATIR)
            cikti(n, 5) = v(i, R_AYLIK_TOPLAM)
            cikti(n, 6) = v(i, R_HEDEF_ONCEKI)
            cikti(n, 7) = v(i, R_HEDEF_YENI)
            If IsNumeric(v(i, R_HEDEF_ONCEKI)) And IsNumeric(v(i, R_HEDEF_YENI)) Then
                If CDbl(v(i, R_HEDEF_ONCEKI)) <> 0 Then cikti(n, 8) = CDbl(v(i, R_HEDEF_YENI)) / CDbl(v(i, R_HEDEF_ONCEKI)) - 1
            End If
            cikti(n, 9) = Left$(GuvenliMetin(v(i, R_MUTABAKAT)), 8)
            If cikti(n, 9) = "FARK VAR" Then fark = fark + 1
            If IsNumeric(v(i, R_AYLIK_TOPLAM)) Then toplamYazilan = toplamYazilan + CDbl(v(i, R_AYLIK_TOPLAM))
        End If
    Next i
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(n, 9).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 9)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(n - 1, 9)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 3).Resize(n - 1, 1).NumberFormat = "#,##0"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 4).Resize(n - 1, 3).NumberFormat = "#,##0.00"
    ws.Cells(r + 1, AN_TABLO_SUTUN + 7).Resize(n - 1, 1).NumberFormat = "0.00%"
    SonucBasligi ws, TR("AKTARIM ETK#IS#I  #*  son etkin parti ") & sonParti
    Yorum ws, 1, (n - 1) & " aktarim; kaynaktan toplam " & Format$(toplamYazilan, "#,##0.00") & " yazildi."
    Yorum ws, 2, IIf(fark = 0, "Tum aktarimlarda mutabakat TUTUYOR.", fark & " aktarimda FARK VAR: Tutar Mutabakati ile yeniden kontrol edin.")
    Yorum ws, 3, "ARTIS %, ana dosya tutar sutununun bu aktarimla buyumesidir; ayni ay icin olagan disi buyume K11 z-skoruyla da izlenir."
    A7AktarimEtkisi = "A7 aktarim etkisi: parti " & sonParti

End Function


' ------------------------------------------------------------
'  A8  KONTROL NOKTASI TRENDI (Rapor, 12 ay)
' ------------------------------------------------------------
Private Function A8KontrolTrendi(ByVal ws As Worksheet) As String

    Dim rpr As Worksheet
    Dim v As Variant
    Dim son As Long, i As Long, ay As Long, kNo As Long, r As Long, enCok As Long
    Dim tablo(1 To 14, 0 To 11) As Double
    Dim cikti() As Variant
    Dim ilkAy As Date
    Dim kod As String, parcalar() As String
    Dim toplam(1 To 14) As Double

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    ilkAy = DateSerial(Year(Date), Month(Date) - 11, 1)
    If son >= 2 Then
        v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))
        For i = 1 To UBound(v, 1)
            kod = Trim$(GuvenliMetin(v(i, R_KONTROL)))
            If kod <> "" And IsDate(v(i, R_TARIH)) Then
                parcalar = Split(kod, " ")
                kNo = Val(Mid$(parcalar(0), 2))
                ay = DateDiff("m", ilkAy, CDate(v(i, R_TARIH)))
                If kNo >= 1 And kNo <= 14 And ay >= 0 And ay <= 11 Then
                    tablo(kNo, ay) = tablo(kNo, ay) + 1
                    toplam(kNo) = toplam(kNo) + 1
                End If
            End If
        Next i
    End If
    ReDim cikti(1 To 15, 1 To 13)
    cikti(1, 1) = TR("KONTROL NOKTASI")
    For ay = 0 To 11
        cikti(1, ay + 2) = Format$(DateAdd("m", ay, ilkAy), "yyyy-mm")
    Next ay
    For kNo = 1 To 14
        cikti(kNo + 1, 1) = "K" & kNo & "  " & KontrolAdi(kNo)
        For ay = 0 To 11
            If tablo(kNo, ay) > 0 Then cikti(kNo + 1, ay + 2) = tablo(kNo, ay)
        Next ay
        If toplam(kNo) > toplam(IIf(enCok = 0, 1, enCok)) Or enCok = 0 Then enCok = kNo
    Next kNo
    r = AN_TABLO_ILK
    ws.Cells(r, AN_TABLO_SUTUN).Resize(15, 13).Value = cikti
    TabloBasligiBicim ws.Cells(r, AN_TABLO_SUTUN).Resize(1, 13)
    TabloCizgileri ws.Cells(r + 1, AN_TABLO_SUTUN).Resize(14, 13)
    ws.Cells(r + 1, AN_TABLO_SUTUN + 1).Resize(14, 12).NumberFormat = "#,##0"
    SonucBasligi ws, TR("KONTROL NOKTASI TREND#I  #*  D#IKKAT + ENGEL bulgu say#is#i, son 12 ay (Rapor)")
    Yorum ws, 1, "En sik bulgu: K" & enCok & " " & KontrolAdiAscii(enCok) & " (" & toplam(enCok) & " madde)."
    Yorum ws, 2, "Bir kontrol noktasinda aydan aya artan bulgu, kaynak sistemde kalici bir veri sorununa isaret eder."
    Yorum ws, 3, "Ayrinti icin Rapor sayfasinda KONTROL NOKTASI sutununu filtreleyin."
    A8KontrolTrendi = "A8 kontrol noktasi trendi"

End Function


' ------------------------------------------------------------
'  A9  FREKANS / SIDDET AYRISIMI (GenelAnaliz'den sonra)
'  Son iki donem arasinda toplam degisim = frekans etkisi + siddet etkisi
'  + etkilesim: dT = (n2-n1) s1 + n1 (s2-s1) + (n2-n1)(s2-s1)
' ------------------------------------------------------------
Private Sub A9Ayrisim(ByVal ws As Worksheet)

    Dim r As Long, son As Long
    Dim n1 As Double, n2 As Double, t1 As Double, t2 As Double, s1 As Double, s2 As Double
    Dim frekansEtki As Double, siddetEtki As Double, etkilesimEtki As Double, dT As Double

    r = AN_TABLO_ILK + 1
    son = ws.Cells(ws.Rows.Count, AN_TABLO_SUTUN).End(xlUp).Row - 1
    If son - r < 1 Then Exit Sub
    n1 = CDbl(ws.Cells(son - 1, AN_TABLO_SUTUN + 1).Value)
    n2 = CDbl(ws.Cells(son, AN_TABLO_SUTUN + 1).Value)
    t1 = CDbl(ws.Cells(son - 1, AN_TABLO_SUTUN + 2).Value)
    t2 = CDbl(ws.Cells(son, AN_TABLO_SUTUN + 2).Value)
    If n1 = 0 Or n2 = 0 Then Exit Sub
    s1 = t1 / n1
    s2 = t2 / n2
    frekansEtki = (n2 - n1) * s1
    siddetEtki = n1 * (s2 - s1)
    etkilesimEtki = (n2 - n1) * (s2 - s1)
    dT = t2 - t1
    Yorum ws, 1, GuvenliMetin(ws.Cells(son, AN_TABLO_SUTUN).Value) & " / " & GuvenliMetin(ws.Cells(son - 1, AN_TABLO_SUTUN).Value) & _
                 ": toplam degisim " & Format$(dT, "#,##0") & " = frekans " & Format$(frekansEtki, "#,##0") & _
                 " + siddet " & Format$(siddetEtki, "#,##0") & " + etkilesim " & Format$(etkilesimEtki, "#,##0") & _
                 "; ortalama tutar " & Format$(s1, "#,##0") & " -> " & Format$(s2, "#,##0") & "."

End Sub


' ============================================================
'  SIRALAMA
' ============================================================
Private Sub HizliSiralaSayi(a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim p As Double, t As Double
    If hi <= lo Then Exit Sub
    i = lo
    j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p
            i = i + 1
        Loop
        Do While a(j) > p
            j = j - 1
        Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then HizliSiralaSayi a, lo, j
    If i < hi Then HizliSiralaSayi a, i, hi
End Sub


Private Sub HizliSiralaMetin(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim p As String, t As String
    If hi <= lo Then Exit Sub
    i = lo
    j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p
            i = i + 1
        Loop
        Do While a(j) > p
            j = j - 1
        Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then HizliSiralaMetin a, lo, j
    If i < hi Then HizliSiralaMetin a, i, hi
End Sub


' sira() indekslerini anahtar(sira(i)) artan olacak bicimde siralar
Private Sub HizliSiralaIndeks(anahtar() As String, sira() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, t As Long
    Dim p As String
    If hi <= lo Then Exit Sub
    i = lo
    j = hi
    p = anahtar(sira((lo + hi) \ 2))
    Do While i <= j
        Do While anahtar(sira(i)) < p
            i = i + 1
        Loop
        Do While anahtar(sira(j)) > p
            j = j - 1
        Loop
        If i <= j Then
            t = sira(i): sira(i) = sira(j): sira(j) = t
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then HizliSiralaIndeks anahtar, sira, lo, j
    If i < hi Then HizliSiralaIndeks anahtar, sira, i, hi
End Sub


' sira(1..n) indekslerini deger(sira(i)) azalan olacak bicimde siralar
Private Sub SiralaAzalan(deger() As Double, sira() As Long, ByVal n As Long)
    SiralaAzalanIc deger, sira, 1, n
End Sub


Private Sub SiralaAzalanIc(deger() As Double, sira() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, t As Long
    Dim p As Double
    If hi <= lo Then Exit Sub
    i = lo
    j = hi
    p = deger(sira((lo + hi) \ 2))
    Do While i <= j
        Do While deger(sira(i)) > p
            i = i + 1
        Loop
        Do While deger(sira(j)) < p
            j = j - 1
        Loop
        If i <= j Then
            t = sira(i): sira(i) = sira(j): sira(j) = t
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then SiralaAzalanIc deger, sira, lo, j
    If i < hi Then SiralaAzalanIc deger, sira, i, hi
End Sub
