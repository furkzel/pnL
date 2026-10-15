Attribute VB_Name = "M09_Kilavuz"
Option Explicit

' ============================================================
' AKTARIM PANELI  v34  -  M09_Kilavuz  (KILAVUZ ve GELISTIRICI)
'
'   Kilavuz     : bolum bolum acilip kapanan kullanim kilavuzu
'   Gelistirici : araci devralacak kisi / yapay zeka icin baglam +
'                 projedeki butun kod
'
' TURKCE KARAKTER: metinler ASCII yazilir, "#" kisaltmalari TR()
' ile Turkce harfe cevrilir (bkz. M02_Yardimci.TR):
'   #i #I #s #S #g #G #u #U #o #O #c #C   #a = a(sapkali)  #- uzun tire
'   #> ok  #* orta nokta  ## = #
'
' KILAVUZ SAYFASI YERLESIMI:
'   C sutunu metin; H (gizli) satir turu: B baslik, G govde; I (gizli)
'   bolum no. Govdeler Excel gruplamasiyla gruplanir (makro kapaliyken
'   soldaki +/- ile acilir). Basliklar tiklanabilir sekillerdir
'   (OnAction = KilavuzTikla).
'
' METNI DEGISTIRMEK: K.. (kilavuz) ve GP.. (gelistirici) prosedurlerini
' duzenleyin. Satir turleri: P paragraf, A alt baslik, M madde, N adim,
' D dikkat kutusu, I ipucu kutusu, T "terim | aciklama", K kod.
' ============================================================

Private Const K_METIN As Long = 3
Private Const K_TUR As Long = 8
Private Const K_NO As Long = 9
Private Const K_EN_FAZLA_BOLUM As Long = 30

Private g_KWs As Worksheet
Private g_KSatir As Long
Private g_KBolum As Long

Private g_GSatirlar() As String
Private g_GSay As Long


' ============================================================
'  KILAVUZ: kurma
' ============================================================
Public Sub KilavuzKur()

    Dim ws As Worksheet
    Dim eskiSayfa As Object
    Dim i As Long

    On Error GoTo Hata
    Set eskiSayfa = ActiveSheet

    Set ws = KilavuzSayfasi()
    Set g_KWs = ws

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error Resume Next
    ws.Cells.ClearOutline
    On Error GoTo Hata
    ws.UsedRange.Clear
    ws.Rows.Hidden = False
    ws.Cells.RowHeight = 15
    ws.Cells.Font.Name = PANEL_YAZI_TIPI
    ws.Cells.Font.Size = 10
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 4
    ws.Columns(K_METIN).ColumnWidth = 125
    ws.Columns(4).ColumnWidth = 2
    ws.Columns(K_TUR).Hidden = True
    ws.Columns(K_NO).Hidden = True
    ws.Outline.SummaryRow = xlAbove

    KBaslikAlani ws
    g_KSatir = 6
    KBolumleriYaz
    KBitir ws

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ws.Range("A1").Select
    On Error Resume Next
    If Not eskiSayfa Is Nothing Then eskiSayfa.Activate
    On Error GoTo 0
    Exit Sub

Hata:
    MsgBox "KILAVUZ KURULAMADI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbExclamation, "Kilavuz"

End Sub


' Kilavuz sayfasini bulur; yoksa bos "Sheet1/Sayfa1" varsa onu, o da
' yoksa yeni bir sayfayi en basa ekleyip adlandirir.
Private Function KilavuzSayfasi() As Worksheet

    Dim ws As Worksheet

    Set ws = KVarOlanSayfa()
    If Not ws Is Nothing Then
        Set KilavuzSayfasi = ws
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        Select Case BH(Replace(ws.Name, " ", ""))
            Case "SHEET1", "SAYFA1", "FEUIL1", "TABELLE1", "KILAVUZ"
                If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
                    ws.Name = KilavuzAdi()
                    Set KilavuzSayfasi = ws
                    Exit Function
                End If
        End Select
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = KilavuzAdi()
    Set KilavuzSayfasi = ws

End Function


Private Sub KBaslikAlani(ByVal ws As Worksheet)

    ws.Rows(1).RowHeight = 10

    With ws.Cells(2, K_METIN)
        .Value = TR("AKTARIM PANEL#I  #-  KULLANIM KILAVUZU")
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(29, 78, 216)
    End With
    ws.Rows(2).RowHeight = 34

    With ws.Cells(3, K_METIN)
        .Value = TR("Ayl#ik hasar dosyalar#in#i (#odenen / muallak) birikimli ana dosyalara e#sleme ile aktaran ara#c" & _
                    "   #*   S#ur#um " & SURUM)
        .Font.Italic = True
        .Font.Color = RGB(71, 85, 105)
    End With

    With ws.Cells(4, K_METIN)
        .Value = TR("Nas#il okunur: bir ba#sl#i#ga t#iklay#in; o b#ol#um a#c#il#ir, di#gerleri kapan#ir. Ayn#i ba#sl#i#ga " & _
                    "tekrar t#iklay#inca kapan#ir. Makrolar kapal#iysa sol kenardaki + i#saretleriyle a#cabilirsiniz.")
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Color = RGB(51, 65, 85)
    End With
    ws.Rows(4).RowHeight = 32
    ws.Rows(5).RowHeight = 30
    ws.Rows(6).RowHeight = 8

    KButon ws, TR("T#um#un#u a#c"), "KilavuzTumunuAc", ws.Cells(5, K_METIN).Left, ws.Cells(5, K_METIN).Top + 3, 120
    KButon ws, TR("T#um#un#u kapat"), "KilavuzTumunuKapat", ws.Cells(5, K_METIN).Left + 130, ws.Cells(5, K_METIN).Top + 3, 120
    KButon ws, TR("Panele d#on"), "KilavuzPaneleDon", ws.Cells(5, K_METIN).Left + 260, ws.Cells(5, K_METIN).Top + 3, 120

End Sub


Private Sub KButon(ByVal ws As Worksheet, ByVal metin As String, ByVal makro As String, _
                   ByVal sol As Double, ByVal ust As Double, ByVal genislik As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(5, sol, ust, genislik, 24)
    shp.Name = "KLV_DUGME_" & makro
    shp.OnAction = makro
    shp.Fill.ForeColor.RGB = RGB(219, 234, 254)
    shp.Line.ForeColor.RGB = RGB(147, 197, 253)
    With shp.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(29, 78, 216)
        .TextRange.ParagraphFormat.Alignment = 2
        .VerticalAnchor = 3
    End With
    shp.Placement = xlFreeFloating

End Sub


Private Sub KBolum(ByVal no As Long, ByVal baslik As String)

    g_KSatir = g_KSatir + 1
    g_KWs.Rows(g_KSatir).RowHeight = 6

    g_KSatir = g_KSatir + 1
    g_KBolum = no
    g_KWs.Cells(g_KSatir, K_TUR).Value = "B"
    g_KWs.Cells(g_KSatir, K_NO).Value = no
    g_KWs.Rows(g_KSatir).RowHeight = 30

    With g_KWs.Range(g_KWs.Cells(g_KSatir, 2), g_KWs.Cells(g_KSatir, K_METIN))
        .Interior.Color = RGB(30, 41, 59)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .VerticalAlignment = xlCenter
    End With
    g_KWs.Cells(g_KSatir, K_METIN).Value = HucreMetni(no & ".  " & TR(baslik))

    KSatir "B", ""

End Sub


Private Sub KSatir(ByVal tur As String, ByVal metin As String)

    Dim h As Range
    Dim p As Long
    Dim terim As String, aciklama As String

    g_KSatir = g_KSatir + 1
    Set h = g_KWs.Cells(g_KSatir, K_METIN)
    g_KWs.Cells(g_KSatir, K_TUR).Value = "G"
    g_KWs.Cells(g_KSatir, K_NO).Value = g_KBolum

    h.WrapText = True
    h.VerticalAlignment = xlTop
    metin = TR(metin)

    Select Case tur
        Case "P"
            h.Value = HucreMetni(metin)
        Case "A"
            h.Value = HucreMetni(metin)
            h.Font.Bold = True
            h.Font.Size = 11
            h.Font.Color = RGB(29, 78, 216)
        Case "M"
            h.Value = HucreMetni(ChrW(8226) & "  " & metin)
            h.IndentLevel = 1
        Case "N"
            h.Value = HucreMetni(metin)
            h.IndentLevel = 1
        Case "D"
            h.Value = HucreMetni(metin)
            h.IndentLevel = 1
            h.Interior.Color = RGB(254, 243, 199)
            h.Borders(xlEdgeLeft).Color = RGB(234, 88, 12)
            h.Borders(xlEdgeLeft).Weight = xlThick
        Case "I"
            h.Value = HucreMetni(metin)
            h.IndentLevel = 1
            h.Interior.Color = RGB(220, 252, 231)
            h.Borders(xlEdgeLeft).Color = RGB(22, 163, 74)
            h.Borders(xlEdgeLeft).Weight = xlThick
        Case "T"
            p = InStr(metin, "|")
            If p > 0 Then
                terim = Trim$(Left$(metin, p - 1))
                aciklama = Trim$(Mid$(metin, p + 1))
            Else
                terim = metin
                aciklama = ""
            End If
            h.Value = HucreMetni(terim & "   " & ChrW(8212) & "   " & aciklama)
            h.IndentLevel = 1
            On Error Resume Next
            h.Characters(1, Len(terim)).Font.Bold = True
            h.Characters(1, Len(terim)).Font.Color = RGB(29, 78, 216)
            On Error GoTo 0
        Case "K"
            h.Value = HucreMetni(metin)
            h.IndentLevel = 1
            h.Font.Name = "Consolas"
            h.Interior.Color = RGB(241, 245, 249)
        Case Else
            h.Value = ""
    End Select

End Sub


' Satir yukseklikleri, gruplar, baslik sekilleri; hepsi kapali baslar
Private Sub KBitir(ByVal ws As Worksheet)

    Dim bas() As Long, ilk() As Long, son() As Long
    Dim n As Long, r As Long, adet As Long

    adet = KBolumleriBul(ws, bas, ilk, son)

    For n = 1 To adet
        If ilk(n) > 0 Then
            ws.Range(ws.Rows(ilk(n)), ws.Rows(son(n))).AutoFit
            For r = ilk(n) To son(n)
                If Trim$(GuvenliMetin(ws.Cells(r, K_METIN).Value)) = "" Then
                    ws.Rows(r).RowHeight = 6
                ElseIf ws.Rows(r).RowHeight < 400 Then
                    ws.Rows(r).RowHeight = ws.Rows(r).RowHeight + 4
                End If
            Next r
            ws.Rows(son(n)).RowHeight = ws.Rows(son(n)).RowHeight + 6
            ws.Range(ws.Rows(ilk(n)), ws.Rows(son(n))).Group
        End If
    Next n

    For n = 1 To adet
        If bas(n) > 0 Then KBaslikSekli ws, n, bas(n)
    Next n

    KHepsi ws, False

End Sub


Private Sub KBaslikSekli(ByVal ws As Worksheet, ByVal n As Long, ByVal satir As Long)

    Dim shp As Shape
    Dim sol As Double, genislik As Double

    sol = ws.Cells(satir, 2).Left
    genislik = ws.Cells(satir, 2).Width + ws.Cells(satir, K_METIN).Width

    Set shp = ws.Shapes.AddShape(5, sol, ws.Rows(satir).Top + 2, genislik, ws.Rows(satir).Height - 4)
    shp.Name = "KLV_B_" & n
    shp.OnAction = "KilavuzTikla"
    shp.Fill.ForeColor.RGB = RGB(30, 41, 59)
    shp.Line.Visible = 0
    With shp.TextFrame2
        .TextRange.Text = ChrW(9654) & "    " & GuvenliMetin(ws.Cells(satir, K_METIN).Value)
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = 12
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 1
        .VerticalAnchor = 3
        .MarginLeft = 12
        .WordWrap = 0
    End With
    shp.Placement = xlMove

End Sub


' Bolumlerin baslik ve govde satirlarini bulur. Donus: bolum sayisi.
Private Function KBolumleriBul(ByVal ws As Worksheet, bas() As Long, ilk() As Long, son() As Long) As Long

    Dim sonSatir As Long, r As Long, n As Long, enBuyuk As Long
    Dim tur As String

    ReDim bas(1 To K_EN_FAZLA_BOLUM)
    ReDim ilk(1 To K_EN_FAZLA_BOLUM)
    ReDim son(1 To K_EN_FAZLA_BOLUM)

    ' End(xlUp) gizli satirlari atlar; kapali son bolum bulunamazdi
    sonSatir = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To sonSatir
        tur = GuvenliMetin(ws.Cells(r, K_TUR).Value)
        If tur = "B" Or tur = "G" Then
            n = GuvenliSayi(ws.Cells(r, K_NO).Value)
            If n >= 1 And n <= K_EN_FAZLA_BOLUM Then
                If tur = "B" Then
                    bas(n) = r
                Else
                    If ilk(n) = 0 Then ilk(n) = r
                    son(n) = r
                End If
                If n > enBuyuk Then enBuyuk = n
            End If
        End If
    Next r

    KBolumleriBul = enBuyuk

End Function


' ============================================================
'  KILAVUZ: acma / kapama
' ============================================================
Public Sub KilavuzTikla()
    Dim ad As String
    On Error Resume Next
    ad = CStr(Application.Caller)
    On Error GoTo 0
    If Left$(ad, 6) = "KLV_B_" Then KBolumAcKapa CLng(Mid$(ad, 7))
End Sub


Public Sub KilavuzTumunuAc()
    Dim ws As Worksheet
    Set ws = KVarOlanSayfa()
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    KHepsi ws, True
    Application.ScreenUpdating = True
End Sub


Public Sub KilavuzTumunuKapat()
    Dim ws As Worksheet
    Set ws = KVarOlanSayfa()
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    KHepsi ws, False
    Application.ScreenUpdating = True
    On Error Resume Next
    ActiveWindow.ScrollRow = 1
    On Error GoTo 0
End Sub


Public Sub KilavuzPaneleDon()
    PaneleGit AktifBlok()
End Sub


' KILAVUZ butonundan: sayfa yoksa kurar, sonra acar
Public Sub KilavuzAc()

    Dim ws As Worksheet

    Set ws = KVarOlanSayfa()
    If ws Is Nothing Then
        KilavuzKur
        Set ws = KVarOlanSayfa()
    ElseIf Trim$(GuvenliMetin(ws.Cells(2, K_METIN).Value)) = "" Then
        KilavuzKur
    End If

    If Not ws Is Nothing Then
        ws.Activate
        On Error Resume Next
        ActiveWindow.ScrollRow = 1
        On Error GoTo 0
    End If

End Sub


Private Function KVarOlanSayfa() As Worksheet
    On Error Resume Next
    Set KVarOlanSayfa = ThisWorkbook.Worksheets(KilavuzAdi())
    On Error GoTo 0
End Function


Private Sub KBolumAcKapa(ByVal n As Long)

    Dim ws As Worksheet
    Dim bas() As Long, ilk() As Long, son() As Long
    Dim adet As Long, m As Long
    Dim acikti As Boolean

    Set ws = KVarOlanSayfa()
    If ws Is Nothing Then Exit Sub

    adet = KBolumleriBul(ws, bas, ilk, son)
    If n < 1 Or n > adet Then Exit Sub
    If ilk(n) = 0 Then Exit Sub

    acikti = Not ws.Rows(ilk(n)).Hidden
    Application.ScreenUpdating = False

    For m = 1 To adet
        If ilk(m) > 0 And m <> n Then
            ws.Range(ws.Rows(ilk(m)), ws.Rows(son(m))).Hidden = True
            KIsaretle ws, m, False
        End If
    Next m

    ws.Range(ws.Rows(ilk(n)), ws.Rows(son(n))).Hidden = acikti
    KIsaretle ws, n, Not acikti
    KSekilleriHizala ws, bas, adet
    Application.ScreenUpdating = True

    If Not acikti Then
        On Error Resume Next
        If ActiveSheet Is ws Then
            If bas(n) > 2 Then
                ActiveWindow.ScrollRow = bas(n) - 1
            Else
                ActiveWindow.ScrollRow = 1
            End If
        End If
        On Error GoTo 0
    End If

End Sub


Private Sub KHepsi(ByVal ws As Worksheet, ByVal ac As Boolean)

    Dim bas() As Long, ilk() As Long, son() As Long
    Dim adet As Long, n As Long

    adet = KBolumleriBul(ws, bas, ilk, son)
    For n = 1 To adet
        If ilk(n) > 0 Then
            ws.Range(ws.Rows(ilk(n)), ws.Rows(son(n))).Hidden = Not ac
            KIsaretle ws, n, ac
        End If
    Next n
    KSekilleriHizala ws, bas, adet

End Sub


Private Sub KIsaretle(ByVal ws As Worksheet, ByVal n As Long, ByVal acik As Boolean)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes("KLV_B_" & n)
    If shp Is Nothing Then Exit Sub
    If acik Then
        shp.TextFrame2.TextRange.Characters(1, 1).Text = ChrW(9660)
        shp.Fill.ForeColor.RGB = RGB(29, 78, 216)
    Else
        shp.TextFrame2.TextRange.Characters(1, 1).Text = ChrW(9654)
        shp.Fill.ForeColor.RGB = RGB(30, 41, 59)
    End If
    On Error GoTo 0

End Sub


Private Sub KSekilleriHizala(ByVal ws As Worksheet, bas() As Long, ByVal adet As Long)

    Dim shp As Shape
    Dim n As Long

    On Error Resume Next
    For n = 1 To adet
        If bas(n) > 0 Then
            Set shp = Nothing
            Set shp = ws.Shapes("KLV_B_" & n)
            If Not shp Is Nothing Then
                shp.Top = ws.Rows(bas(n)).Top + 2
                shp.Height = ws.Rows(bas(n)).Height - 4
                shp.Left = ws.Cells(bas(n), 2).Left
            End If
        End If
    Next n
    On Error GoTo 0

End Sub


' ============================================================
'  KILAVUZ METNI  (bolum basina bir prosedur)
' ============================================================
Private Sub KBolumleriYaz()
    K01
    K02
    K03
    K04
    K05
    K06
    K07
    K08
    K09
    K10
    K11
    K12
    K13
    K14
    K15
    K16
End Sub


Private Sub K01()
    KBolum 1, "Bu ara#c ne yapar?"
    KSatir "P", "Kapan#i#s d#oneminde her ay gelen hasar dosyalar#i (#odenen, muallak ...) birikimli ""ana dosya""lara eklenir. Ayl#ik dosyalarla ana dosyan#in s#utun adlar#i ve s#iralar#i birebir tutmaz; bu ara#c her ayl#ik s#utunu ana dosyadaki do#gru s#utuna E#SLEYEREK kopyalar."
    KSatir "M", "S#utunlar#i ad#iyla e#sle#stirir (birebir ad, benzer ad, #o#grenilmi#s eski adlar)."
    KSatir "M", "Aktarmadan #once her #seyi denetler: eksik s#utun, ayn#i hedefe iki kaynak, veri tipi uyu#smazl#i#g#i, bo#s kalacak zorunlu (ye#sil) s#utun, alttaki toplam sat#ir#i, tarih sorunlar#i."
    KSatir "M", "Yazd#i#g#i her #seyi Ge#cmi#s'e i#sler; Geri Al ile tek t#iklamayla ve g#uvenle geri al#inir."
    KSatir "M", "Aktar#ilan tutar#in ana dosyaya eksiksiz yaz#ild#i#g#in#i mutabakatla kan#itlar."
    KSatir "A", "Ne yapmaz?"
    KSatir "M", "Ana dosyadaki veriyi analiz etmez, #ozet / pivot / grafik #uretmez (v33'teki #Ozet sayfalar#i v34'te kald#ir#ild#i)."
    KSatir "M", "Kontrol i#cin ana dosyaya filtre ya da ge#cici s#utun yazmaz; t#um kontroller aktar#imdan #ONCE ve kaynak #uzerinde yap#il#ir."
    KSatir "M", "Kay#itlar#i poli#ce/hasar numaras#ina g#ore kar#s#ila#st#irmaz (m#ukerrer uyar#is#i dosya d#uzeyindedir)."
End Sub


Private Sub K02()
    KBolum 2, "Kurulum ve ilk a#c#il#i#s"
    KSatir "A", "Gerekenler"
    KSatir "M", "Excel (Windows #onerilir; Mac'te de #cal#i#s#ir). Makrolar#in etkin olmas#i (""#I#ceri#gi Etkinle#stir"")."
    KSatir "M", "Ek bir kurulum, eklenti ya da ActiveX gerekmez. Ara#c Windows'a ait ""Scripting"" kitapl#i#g#in#i bile kullanmaz; Citrix ve kilitli bilgisayarlarda #cal#i#s#ir."
    KSatir "A", "#Ilk kurulum / v33'ten y#ukseltme"
    KSatir "N", "1)  Dosyan#in bir YEDE#G#IN#I al#in."
    KSatir "N", "2)  Kod mod#ullerini y#ukleyin: geli#stirici tools\build.py ile otomatik y#ukler; elle y#uklemek i#cin README.md'deki ad#imlar."
    KSatir "N", "3)  Alt+F8 > SifirdanKur > #Cal#i#st#ir. Rapor, Ge#cmi#s, Profiller ve Setler KORUNUR; eski #Ozet / Dosya Durumu / Pivot Kontrol sayfalar#i silinir; Panel yeni d#uzene ta#s#in#ir (ayarlar, sabit de#gerler ve e#sle#stirmeler korunur)."
    KSatir "N", "4)  Dosyay#i kaydedin."
    KSatir "I", "#IPUCU: Panel ekran#i bir #sekilde bozulursa Alt+F8 > PaneliYenidenKur. #I#cerik silinmeden ekran yeniden #cizilir."
    KSatir "A", "Dosyalar#i a#cman#in iki yolu"
    KSatir "M", "Excel'de a#c#ip ""D#uzenlemeyi Etkinle#stir""e bas#in, sonra panelde men#uden se#cin."
    KSatir "M", "Ya da dosyan#in TAM YOLUNU (\\sunucu\klasor\dosya.xlsx) Dosya h#ucresine yap#i#st#ir#in; ara#c dosyay#i kendisi a#car (Korumal#i G#or#un#um'e tak#ilmaz)."
End Sub


Private Sub K03()
    KBolum 3, "Panel ekran#in#i tan#iyal#im"
    KSatir "A", "#Ust #serit: #u#c a#samal#i i#s ak#i#s#i"
    KSatir "P", "NUMARALI d#ort buton yolu g#osterir: 1 YEN#ILE  #>  2 Otomatik E#sle#stir  #>  3 #On Kontrol  #>  4 AKTAR."
    KSatir "T", "1 #* HAZIRLA | YEN#ILE (dosyalar#i, sayfalar#i, s#utunlar#i okur), + Panel / " & ChrW(8722) & " Panel, Set Y#ukle / Kaydet, k#u#c#uk Set Sil."
    KSatir "T", "2 #* E#SLE#ST#IR & DO#GRULA | Otomatik E#sle#stir, #On Kontrol (hi#cbir #sey yazmadan aktar#im#i dener), #Oneri Onayla (ak#ill#i: se#cili sat#irlarda #oneri varsa onlar#i, yoksa paneldeki hepsini), Profil Kaydet."
    KSatir "T", "3 #* AKTAR & GER#I AL | AKTAR (se#cili panelleri s#irayla), Se#cili S#utun Yamas#i, Geri Al, Tutar Mutabakat#i."
    KSatir "T", "PANEL #* YARDIM | " & ChrW(9664) & " " & ChrW(9654) & " paneller aras#inda gezinir; ortadaki g#osterge ""AKT#IF"" paneli g#osterir (butonlar hep bu panel i#cin #cal#i#s#ir). K#ILAVUZ, Tan#i."
    KSatir "T", "#Seritte olmayanlar | Alt+F8 ile: ProfilSil, MenuleriTazele (zaten otomatik), IsaretleriTemizle (SE#C i#saretlerini siler), ButonlariYenile, PaneliYenidenKur."
    KSatir "A", "#Cift t#iklama k#isayollar#i"
    KSatir "M", "Turuncu #ONER#I h#ucresine #cift t#iklay#in: #oneri hemen onaylan#ir."
    KSatir "M", "#I#SLEM h#ucresine #cift t#iklay#in: SE#C a#c#il#ir / kapan#ir (FORM#UL i#saretine dokunmaz)."
    KSatir "I", "#IPUCU: #Serit dondurulmu#stur; a#sa#g#i kayd#ir#inca kaybolmaz. Sa#ga kayd#ir#inca bir h#ucreye t#iklay#in, butonlar ekran#in soluna gelir."
    KSatir "A", "Panel blo#gu (her ayl#ik dosya i#cin bir blok, en fazla 10)"
    KSatir "T", "Bu paneli aktar | Evet olan paneller AKTAR'da s#irayla aktar#il#ir (ye#sil = Evet)."
    KSatir "T", "KAYNAK | Ayl#ik dosya, sayfa ve son sat#ir (bo#s = otomatik; alttaki toplam sat#ir#in#i d#i#sar#ida b#irakmak i#cin elle yaz#ilabilir)."
    KSatir "T", "HEDEF | Ana dosya, sayfa ve yaz#im ba#slang#ic#i (normal ayda ""Sayfan#in en alt#i"")."
    KSatir "T", "E#SLE#ST#IRME #* KONTROL | Profil, tutar s#utunu (mutabakat i#cin), hasar ve ihbar tarihi s#utunlar#i (tarih kontrol#u i#cin). Bo#ssa YEN#ILE ana dosyada TEK ve a#c#ik bir aday bulursa (TUTAR / HASAR+TAR#IH / #IHBAR+TAR#IH) kendisi se#cer; dolu de#gere dokunmaz, DURUM'da ""OTOMAT#IK SE#C#ILD#I"" yazar."
    KSatir "T", "SAB#IT DE#GERLER | Ayl#ik dosyada olmayan sabit bilgiler (RAPOR_DONEMI = 2026-09, KAYNAK = AXA ...)."
    KSatir "T", "DURUM | Panelin tek sat#irl#ik #ozeti: ka#c s#utun e#sle#sti, onay bekleyen #oneri, bo#s kalan ye#sil s#utun, tip uyu#smazl#i#g#i, profil notlar#i."
    KSatir "A", "E#sle#stirme tablosu (6 s#utun)"
    KSatir "T", "KAYNAK S#UTUN | Ayl#ik dosyan#in s#utunlar#i (""C - Pol No"")."
    KSatir "T", "VER#I T#IP#I | Ayl#ik s#utunun tipi (TAR#IH, SAYI, MET#IN, KARI#SIK, BO#S; az#inl#ik varsa ""%85""). K#irm#iz#i ""TARIH <> SAYI"": ana dosyadaki s#utunla #celi#siyor."
    KSatir "T", "HEDEF S#UTUN | Yaz#ilaca#g#i ana dosya s#utunu (men#uden). Renk, ana dosyadaki ba#sl#ik rengidir; turuncu = onay bekleyen #oneri."
    KSatir "T", "HEDEF HAR#ITASI | Ana dosyan#in b#ut#un s#utunlar#i kendi renkleriyle. [+] dolacak; ""<< YES#IL, BO#STA"" zorunlu s#utun e#sle#smemi#s."
    KSatir "T", "F#ILTRE | Sadece bu ko#sula uyan sat#irlar aktar#il#ir (b#ol#um 6)."
    KSatir "T", "#I#SLEM | SE#C: Se#cili S#utun Yamas#i ve Geri Al kapsam#i. FORM#UL: s#utun form#ul olarak ta#s#in#ir."
    KSatir "I", "#IPUCU: Ba#sl#iklar#in (F#ILTRE, #I#SLEM, VER#I T#IP#I, HAR#ITA) #uzerine gelince k#isa a#c#iklama #c#ikar."
End Sub


Private Sub K04()
    KBolum 4, "Ayl#ik i#s ak#i#s#i ad#im ad#im"
    KSatir "P", "#Ornek: HAYAT ana dosyas#ina o ay#in ayl#ik dosyalar#in#in aktar#ilmas#i."
    KSatir "A", "1 #* HAZIRLA"
    KSatir "N", "1)  Ana dosyay#i ve o ay#in ayl#ik dosyalar#in#i a#c#in."
    KSatir "N", "2)  Set Y#ukle > HAYAT. Paneller hedef dosya, profil, sabit de#gerler ve di#ger ayarlar#iyla gelir."
    KSatir "N", "3)  Her panelde KAYNAK dosya ve sayfay#i se#cin (dosyay#i se#cince sayfa men#us#u hemen dolar)."
    KSatir "N", "4)  YEN#ILE. E#sle#stirmeler profilden gelir; her panelin DURUM sat#ir#ina bak#in."
    KSatir "A", "2 #* E#SLE#ST#IR & DO#GRULA"
    KSatir "N", "5)  Bo#s kalan s#utun varsa Otomatik E#sle#stir. Turuncu #onerinin #uzerine gelip en iyi 3 aday#i okuyun; do#gruysa sat#ir#i se#cip #Oneri Onayla, yanl#i#ssa HEDEF men#us#unden do#grusunu se#cin."
    KSatir "N", "6)  HEDEF HAR#ITASI'nda ""YES#IL, BO#STA"" kalmad#i#g#indan emin olun (gerekirse SAB#IT DE#GER ile doldurun)."
    KSatir "N", "7)  E#sle#stirmeyi de#gi#stirdiyseniz Profil Kaydet: profil bu ay#in s#utun adlar#in#i #o#grenir."
    KSatir "N", "8)  #On Kontrol: hi#cbir #sey yazmadan ka#c sat#ir aktar#ilaca#g#in#i, ENGEL ve D#IKKAT maddelerini, tarih sorunlar#in#i listeler."
    KSatir "A", "3 #* AKTAR & GER#I AL"
    KSatir "N", "9)  Aktar#ilacak panellerde ""Bu paneli aktar"" = Evet."
    KSatir "N", "10)  AKTAR. Her panelin onay ekran#in#i okuyun; biti#ste mutabakat #ozeti g#or#un#ur."
    KSatir "N", "11)  #Once ANA DOSYAYI, sonra PANEL DOSYASINI kaydedin."
    KSatir "D", "D#IKKAT: Ana dosya ile panel dosyas#i birlikte kaydedilmelidir. Sadece ana dosya kaydedilirse Ge#cmi#s kayd#i kaybolur ve aktar#im Geri Al ile geri al#inamaz."
    KSatir "I", "#IPUCU: Ayn#i turdaki b#ut#un aktar#imlar tek bir ""parti"" numaras#i al#ir ve Geri Al ile birlikte geri al#inabilir."
End Sub


Private Sub K05()
    KBolum 5, "E#sle#stirme, #oneriler ve profiller"
    KSatir "A", "Otomatik E#sle#stir"
    KSatir "M", "1. tur: ad#i ayn#i olanlar (b#uy#uk/k#u#c#uk harf, T#urk#ce harf, bo#sluk ve noktalama fark#i yok say#il#ir). Ayn#i adl#i birden fazla s#utun varsa s#iras#i ayn#i olan e#sle#sir (2. ""Tutar"" #> 2. ""Tutar"")."
    KSatir "M", "Ad ayn#i ama veri tipi farkl#iysa (MET#IN #> SAYI) #ONER#I olur."
    KSatir "M", "2. tur: ad#i en az %70 benzeyenler #ONER#I olur; veri tipi kesin #celi#sen aday elenir. Notta en iyi 3 aday, benzerlik oran#i ve veri tipi yazar."
    KSatir "A", "#Oneri nas#il onaylan#ir?"
    KSatir "M", "Sat#ir(lar)#i se#cip #Oneri Onayla: yaln#izca se#cili #oneriler onaylan#ir."
    KSatir "M", "Hi#cbir #onerili sat#ir se#cili de#gilse #Oneri Onayla paneldeki b#ut#un #onerileri listeler ve onay ister."
    KSatir "M", "HEDEF h#ucresini men#uden yeniden se#cmek de onayd#ir."
    KSatir "A", "Profil (#sirket bazl#i e#sle#stirme haf#izas#i)"
    KSatir "M", "Profil Kaydet ayl#ik dosyan#in B#UT#UN s#utunlar#in#i saklar; bilerek e#sle#stirilmeyen s#utunlar da saklan#ir, b#oylece her ay sorulmaz."
    KSatir "M", "S#utunun ad#i de#gi#sir ve siz yine ayn#i hedefe e#slerseniz eski ad D#IGER ADLAR'a eklenir; gelecek ay iki ad da tan#in#ir."
    KSatir "M", "Bu ay dosyada olmayan eski e#sle#smeler silinmez, korunur."
    KSatir "M", "Profil uygulan#irken ad (veya di#ger ad) + s#ira birebir tutan s#utunlar do#grudan, ad#i de#gi#smi#s olabilecekler (benzer ad, ayn#i yer, ayn#i tip) #ONER#I olarak gelir."
End Sub


Private Sub K06()
    KBolum 6, "F#ILTRE yaz#im#i"
    KSatir "P", "F#ILTRE h#ucresi doluysa, o kaynak s#utunu ko#sula uyan sat#irlar aktar#il#ir. Birden fazla s#utunda filtre varsa HEPS#I sa#glanmal#id#ir. Men#u s#utundaki de#gerleri #onerir; elle de yaz#ilabilir."
    KSatir "K", "AXA                tam de#ger (b#uy#uk/k#u#c#uk harf ve T#urk#ce harf fark#i yok)"
    KSatir "K", "AXA, ALLIANZ       birden fazla de#ger (biri tutarsa yeter)"
    KSatir "K", "AX*   A?A          joker: * herhangi bir metin, ? tek karakter"
    KSatir "K", ">100   <=500   =0  say#i kar#s#ila#st#irma"
    KSatir "K", ">=01.01.2026       tarih kar#s#ila#st#irma"
    KSatir "K", "<>IPTAL            e#sit olmayan (birden fazla <> varsa hi#cbiri tutmamal#i)"
    KSatir "K", "BOS   DOLU         bo#s ya da dolu h#ucre"
    KSatir "I", "#IPUCU: De#gerin kendisi virg#ul i#ceriyorsa (""ALLIANZ, INC"") oldu#gu gibi yaz#in; s#utunda aynen varsa tek de#ger say#il#ir. Gerekirse ay#ira#c olarak ; kullan#in."
End Sub


Private Sub K07()
    KBolum 7, "Sabit de#gerler"
    KSatir "P", "Ayl#ik dosyada olmayan ama ana dosyada doldurulmas#i gereken s#utunlar i#cin (RAPOR_DONEMI, KAYNAK, DOSYA_TURU ...). Her aktar#ilan sat#ira ayn#i de#ger yaz#il#ir."
    KSatir "M", "Ba#sta tek sat#ir g#or#un#ur; s#utun ve de#ger dolunca alt#inda yeni sat#ir a#c#il#ir (en fazla 15). S#utunu silince sat#ir kalkar, alttakiler yukar#i kayar."
    KSatir "M", "DE#GER men#us#u ana dosyan#in o s#utununun son 50.000 sat#ir#indaki farkl#i de#gerleri listeler (say#ilar say#i, tarihler tarih olarak). Men#ude yoksa elle yaz#in."
    KSatir "M", "Sabit de#ger ile e#sle#stirilmi#s bir s#utun ayn#i hedefe gidemez (ENGEL). De#geri bo#s sabit sat#ir#i ENGEL'dir."
    KSatir "M", "Setler sabit de#gerleri de saklar."
End Sub


Private Sub K08()
    KBolum 8, "Aktar#im #oncesi kontroller: ENGEL, D#IKKAT, #On Kontrol"
    KSatir "A", "ENGEL: aktar#im yap#ilamaz, hi#cbir #sey yaz#ilmaz"
    KSatir "M", "Ana dosya salt okunur (ba#ska biri a#c#ik tutuyor)."
    KSatir "M", "E#sle#stirilmi#s bir s#utun dosyada yok ya da ayn#i adla birden fazla var."
    KSatir "M", "#Iki kaynak (ya da kaynak + sabit) ayn#i hedef s#utuna yaz#iyor."
    KSatir "M", "Sabit de#ger sat#ir#i eksik; Excel'in 1.048.576 sat#ir s#in#ir#i a#s#il#iyor."
    KSatir "A", "Altta toplam sat#ir#i var m#i?"
    KSatir "P", "S#utunlar#in #co#gu ayn#i sat#irda bitip birka#c#i daha a#sa#g#i iniyorsa ya da en altta bo#s sat#irla ayr#ilm#i#s k#u#c#uk bir b#ol#um varsa (genel toplam / not), ara#c o sat#irlar#in i#ceri#gini g#osterip sorar: EVET = temiz veriye kadar oku, HAYIR = hepsini oku, #IPTAL."
    KSatir "A", "D#IKKAT maddeleri (okuyup onaylars#in#iz, varsay#ilan d#u#gme Hay#ir)"
    KSatir "M", "Onaylanmam#i#s #oneri, veri tipi uyu#smazl#i#g#i, bo#s kalacak ye#sil s#utun, profille fark."
    KSatir "M", "Tarih bi#cimli h#ucrede ########## (ta#sma): o s#utunun tarihleri say#i olarak ta#s#in#ir."
    KSatir "M", "B#IC#IM KORUMASI: ba#staki s#if#ir, uzun numara gibi de#gerler metin olarak yaz#ilacak (b#ol#um 9)."
    KSatir "M", "KAYNAK TAR#IH KONTROL#U: hasar ve ihbar tarihi s#utunlar#i se#cildiyse, aktar#ilacak sat#irlarda ihbar tarihi hasardan #once olan, bo#s ya da tarihe #cevrilemeyen de#gerler (#ornek sat#ir numaralar#iyla)."
    KSatir "M", "M#ukerrer aktar#im: ayn#i ayl#ik dosya ayn#i ana dosyaya daha #once aktar#ilm#i#s ve geri al#inmam#i#s (#On Kontrol'de de g#or#un#ur)."
    KSatir "M", "ZORUNLU (YES#IL) SUTUNDA BO#S HUCRE: ye#sil s#utun e#sle#smi#s ama aktar#ilacak baz#i sat#irlarda de#ger bo#s (#ornek sat#ir numaralar#iyla)."
    KSatir "M", "TUTAR SUTUNU: say#i olmayan de#ger (""1.200 TL"", ""YOK"", hata), bo#s h#ucre ve eksi tutar say#is#i. Bunlar toplama girmedi#gi i#cin mutabakat ""TUTUYOR"" dese bile veri eksik olabilir."
    KSatir "M", "SAB#IT DE#GER T#IP#I UYU#SMAZ: #or. tarih s#utununa ""EYLUL"" gibi bir metin yaz#ilm#i#s."
    KSatir "A", "#On Kontrol"
    KSatir "P", "Bulundu#gunuz panel i#cin b#ut#un bu kontrolleri #cal#i#st#ir#ir ve tek ekranda raporlar; ana dosyaya da kaynak dosyaya da H#I#CB#IR #SEY yazmaz. v33'teki tarih kontrol d#u#gmelerinin yerini al#ir. Alttaki #s#upheli ""toplam"" sat#irlar#in#in i#ceri#gi de g#osterilir."
    KSatir "I", "#IPUCU: Ekrana s#i#gmayan maddeler ""... ve X madde daha"" olarak #ozetlenir; HER madde Rapor sayfas#inda ayr#i bir KONTROL sat#ir#ina k#isalt#ilmadan yaz#il#ir ve Rapor tek t#iklamayla a#c#ilabilir."
End Sub


Private Sub K09()
    KBolum 9, "Bi#cim korumas#i ve form#ul aktar#im#i"
    KSatir "P", "Excel h#ucreye yaz#ilan metni elle yaz#ilm#i#s gibi yorumlar: ""0004821"" 4821 olur, 16 haneli numara 1,23E+18 olur, ""=..."" form#ul olur. v34 bu metinleri oldu#gu gibi (metin olarak) yazar:"
    KSatir "M", "Ba#s#inda s#if#ir olan say#i gibi kodlar, 16 ve daha uzun haneli numaralar."
    KSatir "M", "= ile ba#slayan (ya da +, -, @ ile ba#slay#ip say#i olmayan) metinler."
    KSatir "M", "Ana dosyada MET#IN olan s#utuna giden say#i/tarih gibi metinler."
    KSatir "P", "Di#ger her #sey v33'teki gibi yaz#il#ir: say#i say#i, tarih tarih. Yeni sat#irlar, #ustteki sat#ir#in say#i bi#cimini (#or. 1.234,56 ya da gg.aa.yyyy) al#ir."
    KSatir "A", "Form#ul aktar#im#i (#I#SLEM = FORM#UL)"
    KSatir "P", "Form#ul#un g#oreli SATIR referanslar#i yeni sat#ira kayd#ir#il#ir, s#utun harfleri ayn#i kal#ir: kaynakta 5. sat#irdaki =B5*C5, ana dosyada 870005. sat#ira =B870005*C870005 olarak yaz#il#ir."
End Sub


Private Sub K10()
    KBolum 10, "Se#cili S#utun Yamas#i ve d#uzeltme"
    KSatir "P", "Bir aktar#imdan sonra yaln#izca bir ya da birka#c s#utunun yanl#i#s gitti#gini fark ettiyseniz t#um aktar#im#i geri alman#iz gerekmez."
    KSatir "N", "1)  E#sle#stirmeyi d#uzeltin; d#uzeltilecek s#utun(lar)da #I#SLEM = SE#C yap#in."
    KSatir "N", "2)  Se#cili S#utun Yamas#i. Yaz#im ba#slang#ic#i ""Sayfan#in en alt#i"" iken, se#cili s#utunlar bu ayl#ik dosyan#in SON AKT#IF aktar#im#in#in sat#irlar#ina UZER#INE yaz#il#ir."
    KSatir "N", "3)  #Uzerine yaz#ilan eski de#gerler gizli Yedek sayfas#ina al#in#ir; yamay#i Geri Al ile geri al#irsan#iz eski de#gerler geri gelir."
    KSatir "D", "D#IKKAT: Yama ayn#i filtre ve ayn#i son sat#irla okunan sat#ir say#is#i, son aktar#imdaki sat#ir say#is#iyla ayn#i olmal#id#ir; de#gilse ENGEL verir (sat#irlar kayard#i)."
    KSatir "A", "Di#ger d#uzeltme yollar#i"
    KSatir "M", "Geri Al ile aktar#im#i (ya da SE#C'li s#utunlar#i) geri al#in, yaz#im ba#slang#ic#in#i ""Geri al#inan aral#i#ga"" yap#in ve yeniden aktar#in: veri bo#salan yere yaz#il#ir. Yeni veri daha uzunsa ara#c sat#ir ekler ve Ge#cmi#s'i kayd#ir#ir."
    KSatir "M", """S#utunun kendi sonu"": her s#utun kendi son dolu h#ucresinin alt#ina yaz#il#ir (tek s#utun eklemek i#cin)."
End Sub


Private Sub K11()
    KBolum 11, "Geri Al"
    KSatir "P", "Son aktar#im#i, son partiyi (ayn#i AKTAR turundaki b#ut#un paneller) ya da yaln#izca SE#C i#saretli s#utunlar#i geri al#ir."
    KSatir "A", "#Iki a#samal#i ve g#uvenli"
    KSatir "M", "1. a#sama: geri al#inacak HER kay#it denetlenir: ana dosya a#c#ik m#i, salt okunur mu, s#utun ba#sl#i#g#i aktar#imdaki gibi mi. Tek bir sorun varsa hi#cbir h#ucre silinmez."
    KSatir "M", "PARMAK #IZ#I: her aktar#imda yaz#ilan aral#i#g#in k#isa bir #ozeti (dolu h#ucre say#is#i, toplam, ilk ve son de#ger) saklan#ir. Ana dosyaya sonradan sat#ir eklendiyse/silindiyse ya da h#ucreler de#gi#stiyse Geri Al bunu fark eder ve sizi uyar#ir."
    KSatir "M", "2. a#sama: h#ucreler temizlenir; #uzerine yaz#ilm#i#s olanlar (yama) yedekten eski haline d#oner. Ge#cmi#s'te kay#itlar ""GER#I ALINDI"" olur."
    KSatir "D", "D#IKKAT: Aktar#im ile geri alma aras#inda ana dosyada elle sat#ir veya s#utun ekleyip silmeyin. Parmak izi uyar#is#i #c#ikarsa ana dosyay#i kontrol etmeden devam etmeyin."
End Sub


Private Sub K12()
    KBolum 12, "Tutar mutabakat#i"
    KSatir "P", "Panelde ""Tutar s#utunu"" se#cildiyse her aktar#imdan sonra d#ort say#i kar#s#ila#st#ir#il#ir ve Rapor'a yaz#il#ir:"
    KSatir "T", "Kaynaktan okunan | Ayl#ik dosyada AKTARILAN sat#irlar#in toplam#i (metin olarak yaz#ilm#i#s say#ilar dahil)."
    KSatir "T", "Hedefe yaz#ilan | Ana dosyada yaz#ilan aral#i#g#in toplam#i."
    KSatir "T", "Hedef #onceki / yeni toplam | Ana dosyadaki tutar s#utununun aktar#imdan #onceki ve sonraki toplam#i. Art#i#s, yaz#ilan tutar kadar olmal#id#ir."
    KSatir "P", "FARK = 0,00 ise TUTUYOR. Fark varsa ara#c uyar#ir; aktar#im yap#ilm#i#st#ir, gerekirse Geri Al."
    KSatir "M", "Tutar Mutabakat#i butonu son AKT#IF PART#IN#IN (ayn#i AKTAR turundaki) HER panelini #SIMD#I yeniden kontrol eder ve tek tabloda g#osterir; bulundu#gunuz panel "">>"" ile i#saretlenir. Aktar#imdan sonra yap#ilan elle de#gi#siklikleri yakalar."
End Sub


Private Sub K13()
    KBolum 13, "Setler, Rapor ve di#ger sayfalar"
    KSatir "T", "Setler | Birden fazla panelin ayarlar#i tek isimle (#or. HAYAT): hedef dosya/sayfa, profil, sabit de#gerler, tutar/hasar/ihbar s#utunlar#i, yaz#im modu. Kaynak dosya ve son sat#ir saklanmaz (her ay de#gi#sir)."
    KSatir "T", "Rapor | #Insanlar i#cin olay g#unl#u#g#u: AKTARIM, D#IKKAT/UYARI kararlar#i, ENGEL, GER#I ALMA, KONTROL, MUTABAKAT, PROF#IL ve SET i#slemleri."
    KSatir "T", "Ge#cmi#s | Geri Al'#in defteri: her aktar#ilan s#utun i#cin bir sat#ir (parti, aktar#im no, aral#ik, durum, parmak izi, yedek). ELLE DE#G#I#ST#IRMEY#IN."
    KSatir "T", "Profiller | E#sle#stirme haf#izas#i."
    KSatir "T", "Geli#stirici | Arac#i devralacak ki#si i#cin ba#glam ve t#um kod."
    KSatir "T", "Listeler1..10, Yedek | Gizli: a#c#il#ir men#u i#cerikleri ve #uzerine yaz#ilan eski de#gerler."
End Sub


Private Sub K14()
    KBolum 14, "Sorun giderme"
    KSatir "I", "#IPUCU: Her sorunda #once Tan#i'ya bas#in (Excel'i de normale d#ond#ur#ur), sonra Rapor'daki son sat#irlara bak#in."
    KSatir "T", "Butonlar g#or#unm#uyor | Bir h#ucreye t#iklay#in. H#al#a yoksa Alt+F8 > ButonlariYenile."
    KSatir "T", "Hi#cbir buton #cal#i#sm#iyor | Makrolar kapal#i olabilir (""#I#ceri#gi Etkinle#stir"")."
    KSatir "T", """Panel eski d#uzende"" uyar#is#i | Alt+F8 > SifirdanKur (i#cerik korunur)."
    KSatir "T", "Men#ude dosyam yok | Dosya a#c#ik m#i, Korumal#i G#or#un#um'de mi? D#uzenlemeyi etkinle#stirip YEN#ILE."
    KSatir "T", """Kaynak/hedef bulunamad#i"" | Dosya kapanm#i#s veya ad#i de#gi#smi#s. Men#uden yeniden se#cin, YEN#ILE."
    KSatir "T", "DURUM: PROF#IL BULUNAMADI | Profil silinmi#s veya ad#i de#gi#smi#s. Profil men#us#unden yeniden se#cin."
    KSatir "T", "Aktar#im ENGEL veriyor | #On Kontrol ile b#ut#un maddeleri g#or#un; #co#gu zaman YEN#ILE yeterlidir."
    KSatir "T", "Geri Al ""ba#sl#i#g#i de#gi#smi#s"" diyor | Ana dosyaya s#utun eklenmi#s veya silinmi#s. Hi#cbir #sey silinmedi; ana dosyay#i eski d#uzenine getirin."
    KSatir "T", "Geri Al ""parmak izi"" uyar#is#i | Aktar#imdan sonra ana dosyada sat#ir eklenmi#s/silinmi#s ya da h#ucreler de#gi#smi#s. Kontrol etmeden ""Evet"" demeyin."
    KSatir "T", "Ana dosyada eski Kontrol filtresi kalm#i#s | v33'#un tarih kontrolleri ana dosyaya filtre b#irak#iyordu: Alt+F8 > KontrolTemizle."
    KSatir "T", "Excel ""Yan#it vermiyor"" | B#uy#uk dosyada normaldir, bekleyin. Durum #cubu#gu ilerlemeyi g#osterir."
End Sub


Private Sub K15()
    KBolum 15, "Alt#in kurallar ve s#in#irl#il#iklar"
    KSatir "N", "1)  Aktar#imdan sonra #once ana dosyay#i, sonra panel dosyas#in#i kaydedin."
    KSatir "N", "2)  Ge#cmi#s, Rapor ve Profiller sayfalar#in#i elle d#uzenlemeyin."
    KSatir "N", "3)  Normal ayl#ik aktar#imda yaz#im ba#slang#ic#i ""Sayfan#in en alt#i"" olmal#id#ir."
    KSatir "N", "4)  Onay ekran#in#i okumadan ""Evet"" demeyin. D#IKKAT maddeleri aktar#im#i durdurmaz; siz durdurursunuz."
    KSatir "N", "5)  Ana dosya ba#ska birinde a#c#ikken aktar#im yapmay#in."
    KSatir "A", "S#in#irl#il#iklar"
    KSatir "M", "Ba#sl#iklar 1. sat#irda, veri 2. sat#irdan itibaren olmal#id#ir."
    KSatir "M", "En fazla 10 panel, panel ba#s#ina 300 s#utun ve 15 sabit de#ger."
    KSatir "M", "Veri tipi tahmini ilk/son 300 sat#ira, filtre men#us#u ilk 10.000 sat#ira bakar (filtre yine b#ut#un sat#irlara uygulan#ir)."
    KSatir "M", "Ye#sil s#utun tespiti renk tonuna bakar (G > R+12 ve G > B+12)."
    KSatir "M", "Ayn#i anda tek ki#si #cal#i#smal#id#ir; kilitleme yoktur. Excel'in Ctrl+Z'si makrodan sonra #cal#i#smaz (Geri Al kullan#in)."
    KSatir "M", "Rapor resmi bir denetim kayd#i de#gil, #cal#i#sma tutana#g#id#ir."
End Sub


Private Sub K16()
    KBolum 16, "Arac#i geli#stirmek isteyenler i#cin"
    KSatir "P", "Kod VS Code'da d#uzenlenir: src\ klas#or#undeki .bas/.cls dosyalar#i tek do#gru kaynakt#ir. tools\build.py kodu #cal#i#sma kitab#ina y#ukler, derleme hatas#in#i yakalar ve test eder; tools\export.py Excel'de yap#ilan de#gi#siklikleri src\'ye geri al#ir."
    KSatir "K", "M01_Ayarlar     sabitler, enum'lar, renk paleti, global durum"
    KSatir "K", "M02_Yardimci    metin/T#urk#ce, dosya bulma, okuma, veri tipi, ekran kilidi"
    KSatir "K", "M03_Panel       panel geometrisi, buton #seridi, gezinme, panel ekle/kald#ir"
    KSatir "K", "M04_Yenile      Yenile, men#uler, DURUM, sabit de#gerler, harita, veri tipi"
    KSatir "K", "M05_Eslestirme  otomatik e#sle#stirme, #oneriler, profiller, setler"
    KSatir "K", "M06_Aktarim     aktar#im motoru, filtre, bi#cim korumas#i, #On Kontrol"
    KSatir "K", "M07_Kayit       Geri Al, Ge#cmi#s, Rapor, yedek, parmak izi"
    KSatir "K", "M08_Kurulum     SifirdanKur, PaneliYenidenKur, v33 ge#ci#si, Tan#i"
    KSatir "K", "M09_Kilavuz     bu k#ilavuz ve Geli#stirici sayfas#i"
    KSatir "K", "Sozluk (s#in#if)  Scripting.Dictionary yerine yerel s#ozl#uk"
    KSatir "A", "Kurallar"
    KSatir "M", "Kod ve MsgBox metinleri ASCII. H#ucreye T#urk#ce harf gerekiyorsa TR(""#s#i"") kullan#il#ir."
    KSatir "M", "Kar#s#ila#st#irmada UCase/LCase de#gil BH() / Normalize()."
    KSatir "M", "Yerle#sim de#gi#sikli#gi yaln#izca M01_Ayarlar'daki sabitlerden yap#il#ir."
    KSatir "M", "Her giri#s makrosu EkranKilitle / EkranAc #cifti ile #cal#i#s#ir; hata yolunda da EkranAc #ca#gr#il#ir."
    KSatir "M", "Ayr#int#i: Geli#stirici sayfas#i ve README.md."
End Sub


' ============================================================
'  GELISTIRICI SAYFASI
'
'  A sutununa once baglam (prompt), sonra projedeki butun kod yazilir.
'  Kodu okuyabilmek icin "VBA proje nesne modeline erisime guven" acik
'  olmalidir; kapaliysa elle ekleme adimlari yazilir.
' ============================================================
Public Sub GelistiriciGuncelle()

    Dim ws As Worksheet
    Dim eskiSayfa As Object
    Dim proje As Object, bilesen As Object
    Dim i As Long, adet As Long
    Dim erisim As Boolean
    Dim dizi() As Variant

    On Error GoTo Hata
    Set eskiSayfa = ActiveSheet

    Set ws = SayfaOlustur(SAYFA_GELISTIRICI)

    g_GSay = 0
    ReDim g_GSatirlar(1 To 4000)

    GEkle "Bu sayfa " & Format$(Now, "dd.mm.yyyy hh:nn") & " tarihinde GelistiriciGuncelle ile uretildi (" & SURUM & ")."
    GEkle ""
    GPromptYaz

    erisim = False
    On Error Resume Next
    Set proje = ThisWorkbook.VBProject
    adet = proje.VBComponents.Count
    erisim = (Err.Number = 0 And adet > 0)
    Err.Clear
    On Error GoTo Hata

    If erisim Then
        For Each bilesen In proje.VBComponents
            If bilesen.CodeModule.CountOfLines > 3 Then
                GEkle ""
                GEkle "===== KOD: " & BilesenAdi(bilesen) & " ====="
                For i = 1 To bilesen.CodeModule.CountOfLines
                    GEkle bilesen.CodeModule.Lines(i, 1)
                Next i
                GEkle "===== " & bilesen.Name & " SONU ====="
            End If
        Next bilesen
    Else
        GEkle ""
        GElleKodYaz
    End If

    ws.UsedRange.Clear
    ws.Columns(1).ColumnWidth = 150
    ws.Columns(1).NumberFormat = "@"
    ws.Columns(1).Font.Name = "Consolas"
    ws.Columns(1).Font.Size = 10

    ReDim dizi(1 To g_GSay, 1 To 1)
    For i = 1 To g_GSay
        dizi(i, 1) = " " & g_GSatirlar(i)
    Next i
    ws.Range("A1").Resize(g_GSay, 1).Value = dizi

    On Error Resume Next
    ws.Tab.Color = RGB(100, 116, 139)
    If Not eskiSayfa Is Nothing Then eskiSayfa.Activate
    On Error GoTo 0
    Exit Sub

Hata:
    MsgBox "GELISTIRICI SAYFASI GUNCELLENEMEDI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbExclamation, "Gelistirici"

End Sub


Private Function BilesenAdi(ByVal bilesen As Object) As String

    Dim ws As Worksheet

    BilesenAdi = bilesen.Name
    Select Case bilesen.Type
        Case 100
            For Each ws In ThisWorkbook.Worksheets
                If ws.CodeName = bilesen.Name Then
                    BilesenAdi = bilesen.Name & "  (""" & ws.Name & """ sayfasinin kod sayfasi = src\Panel.cls)"
                    Exit Function
                End If
            Next ws
            BilesenAdi = bilesen.Name & "  (belge modulu)"
        Case 2
            BilesenAdi = bilesen.Name & "  (sinif modulu, src\" & bilesen.Name & ".cls)"
        Case Else
            BilesenAdi = bilesen.Name & "  (standart modul, src\" & bilesen.Name & ".bas)"
    End Select

End Function


Private Sub GEkle(ByVal s As String)
    g_GSay = g_GSay + 1
    If g_GSay > UBound(g_GSatirlar) Then ReDim Preserve g_GSatirlar(1 To UBound(g_GSatirlar) * 2)
    g_GSatirlar(g_GSay) = s
End Sub


Private Sub GMetin(ByVal s As String)
    GEkle TR(s)
End Sub


Private Sub GPromptYaz()
    GP01
    GP02
    GP03
End Sub


Private Sub GP01()
    GMetin "================================================================"
    GMetin "AKTARIM PANEL#I (Excel VBA) " & SURUM & " #- GEL#I#ST#IRME BA#GLAMI"
    GMetin "v35: #On Kontrol'de m#ukerrer uyar#is#i ve kuyruk #onizlemesi, ye#sil s#utunda bo#s"
    GMetin "     h#ucre, tutar b#ut#unl#u#g#u, sabit de#ger tipi, parti bazl#i mutabakat, sade"
    GMetin "     #serit, kontrol s#utunlar#in#i otomatik bulma, #cift t#iklama k#isayollar#i."
    GMetin "================================================================"
    GMetin "Bu metin, arac#i devralacak ki#sinin (ya da onun kulland#i#g#i yapay zek#a"
    GMetin "asistan#in#in) arac#i h#izla anlamas#i i#cindir. B#ut#un kod a#sa#g#idad#ir."
    GMetin ""
    GMetin "#CALI#SMA B#I#C#IM#I"
    GMetin "- Kod VS Code'da, src\ klas#or#unde tutulur (tek do#gru kaynak)."
    GMetin "- tools\build.py: src\'yi #cal#i#sma kitab#ina y#ukler, derleme hatas#in#i"
    GMetin "  yakalar, istenirse u#ctan uca test eder ve .xlsm #uretir."
    GMetin "- tools\export.py: Excel'de VBA edit#or#unde yap#ilan de#gi#siklikleri src\'ye"
    GMetin "  geri yazar (sonra git diff ile g#ozden ge#cirin)."
    GMetin "- Ger#cek dosyalar gizlidir: yapay zek#aya dosya de#gil, s#utun adlar#i ve"
    GMetin "  birka#c #ornek sat#ir tarif edilir."
    GMetin ""
    GMetin "1. #I#S #IHT#IYACI"
    GMetin "Akt#uerya departman#inda her ay gelen hasar dosyalar#indaki (#odenen, muallak)"
    GMetin "veriler birikimli ana dosyalara (870.000+ sat#ir) s#utun e#slemesiyle eklenir."
    GMetin "Ayl#ik dosyalarla ana dosyan#in s#utun adlar#i/s#iralar#i birebir tutmaz."
    GMetin "Ama#c: h#izl#i, veri kay#ips#iz, denetlenebilir ""e#slemeli kopyala-yap#i#st#ir""."
    GMetin ""
    GMetin "2. ORTAM KISITLARI"
    GMetin "- Python/PyPI yok; Scripting.Dictionary, RegExp, ActiveX YOK (Citrix/Mac)."
    GMetin "  Yerine: Sozluk s#in#if#i (Collection tabanl#i), Shape butonlar."
    GMetin "- Ana dosyalar #cok b#uy#uk: sayfa tek seferde okunmaz, SutunOku ile s#utun"
    GMetin "  s#utun okunur. Application.Transpose kullan#ilmaz."
    GMetin "- Kullan#ic#ilar kod yazmaz; men#u ve butonlarla #cal#i#s#ir."
    GMetin ""
    GMetin "3. MOD#ULLER"
    GMetin "M01_Ayarlar     sabitler (yerle#sim, Listeler/Rapor/Gecmis/Profiller/Setler"
    GMetin "                s#utunlar#i), enum'lar, renk paleti, global durum"
    GMetin "M02_Yardimci    BH/Normalize/TR, SutunOku (ta#sma korumal#i), DosyaBul,"
    GMetin "                SayfaBul, veri tipi, BaslikDogrula, MenuKur, AlanTemizle,"
    GMetin "                EkranKilitle/EkranAc/HesapDurdur/HesapAc"
    GMetin "M03_Panel       blok geometrisi (BlokSutun, Ayar, AktifBlok), serit,"
    GMetin "                ButonlariEkranaTasi, PaneleGit, BlokCiz, Panel Ekle/Kald#ir"
    GMetin "M04_Yenile      Yenile/BlokYenile, KaynakTara (tek ge#ci#s tip+filtre),"
    GMetin "                BlokMenuleriTazele, DurumYaz, HedefDegisti (olay y#onlendirici),"
    GMetin "                SABIT kutusu, HaritayiYaz, TipSutunuYaz"
    GMetin "M05_Eslestirme  OtomatikEslestir, OneriOnayla (ak#ill#i), profiller, setler"
    GMetin "M06_Aktarim     Aktar/AktarSecili/OnKontrol; AktarimPlani tipiyle a#samal#i"
    GMetin "                motor; filtre motoru; KorumaGerekir; FormulSatirKaydir"
    GMetin "M07_Kayit       GeriAl (2 a#sama + parmak izi + yedek), Gecmis, Rapor,"
    GMetin "                AralikYedekle, AralikImzasi, TutarMutabakatiSon"
    GMetin "M08_Kurulum     SifirdanKur, PaneliYenidenKur (v33 ge#ci#si), PanelKur, Tani"
    GMetin "M09_Kilavuz     K#ilavuz ve bu sayfa"
    GMetin "Panel (sayfa)   Worksheet_Change -> HedefDegisti; SelectionChange ->"
    GMetin "                ButonlariEkranaTasi"
    GMetin "ThisWorkbook    Workbook_Open -> UygulamaDurumunuOnar"
End Sub


Private Sub GP02()
    GMetin ""
    GMetin "4. PANEL YERLE#S#IM#I (M01_Ayarlar)"
    GMetin "Yan yana en fazla 10 blok, her biri 6 s#utun (BLOK_GENISLIK)."
    GMetin "Sat#ir 1-5 #serit (dondurulmu#s), 7 blok ba#sl#i#g#i ""PANEL k - dosya / sayfa"","
    GMetin "8-21 ayarlar (etiket c, de#ger c+2), 8-24 SAB#IT kutusu (c+3 s#utun, c+4"
    GMetin "de#ger), 26 DURUM, 28 tablo ba#sl#i#g#i, 29-530 tablo:"
    GMetin "c KAYNAK | c+1 VER#I T#IP#I | c+2 HEDEF | c+3 HAR#ITA | c+4 F#ILTRE | c+5 #I#SLEM."
    GMetin "H#ucre de#gerleri ""F - SUTUN ADI"" bi#cimindedir (SutunNo/SutunAdi)."
    GMetin "Kald#ir#ilan blok silinmez, gizlenir (numaralar ve setler bozulmas#in)."
    GMetin "Panel d#uzeni ad#i: AP_DUZEN = ""v34"" (yoksa SifirdanKur ge#ci#s yapar)."
    GMetin ""
    GMetin "5. AKTARIM MOTORU (M06)"
    GMetin "PlanHazirla (ENGEL) -> M#ukerrer -> SonSatirBelirle (#co#gunluk + alt"
    GMetin "b#ol#um ayrac#i) -> VeriOku + SatirSec (filtre) -> HedefAralikBelirle (yaz#im"
    GMetin "modu / yama / sat#ir s#in#ir#i) -> EkDikkat -> OnayAl -> Yaz -> mutabakat."
    GMetin "Yaz: gerekiyorsa sat#ir ekler (GecmisSatirlariKaydir), dolu h#ucrenin"
    GMetin "#uzerine yaz#ilacaksa AralikYedekle, BicimDevral, KorumaGerekir ile"
    GMetin "metin korumas#i, yazma, AralikImzasi, LogYaz (Gecmis)."
    GMetin "Mutabakat: kaynak toplam#i, yaz#ilan aral#ik, hedef s#utun #onceki/yeni"
    GMetin "toplam#i (AGGREGATE(9,6) ile hatal#i h#ucreler yok say#il#ir)."
    GMetin ""
    GMetin "6. GER#I AL (M07)"
    GMetin "1. a#sama: her kay#it denetlenir (dosya a#c#ik, salt okunur de#gil, ba#sl#ik"
    GMetin "ayn#i, yedek var); parmak izi tutmazsa ayr#ica uyar#ilir. 2. a#sama: ters"
    GMetin "s#irayla temizle / yedekten geri y#ukle, Gecmis'te GERI ALINDI."
End Sub


Private Sub GP03()
    GMetin ""
    GMetin "7. GE#CM#I#STEN DERSLER (DE#G#I#SMEZ KURALLAR)"
    GMetin "- Kaynak kod %100 ASCII; T#urk#ce harf h#ucreye TR() ile yaz#il#ir. MsgBox ASCII."
    GMetin "- T#urk#ce Windows'ta UCase(""i"") noktal#i #I verir: kar#s#ila#st#irma BH()/Normalize()."
    GMetin "- Tek sat#irl#ik aral#ikta Range.Value skaler d#oner: IkiBoyutlu/SutunOku kullan."
    GMetin "- .Value hata 6 (##########) verirse Value2 ile okunur ve kullan#ic#i uyar#il#ir."
    GMetin "- Men#ul#u sayfada ws.Cells.Clear XML'i bozabilir: AlanTemizle kullan."
    GMetin "- Mod#ul d#uzeyi bildirimler ilk prosed#urden #once; IIf iki dal#i da hesaplar;"
    GMetin "  bir ifadede en fazla 24 sat#ir devam#i; MsgBox ~1024 karakter (Sigdir)."
    GMetin "- ""On Error"" sat#irlar#i Err'i s#if#irlar: hatay#i EkranAc'tan #once kaydet."
    GMetin "- = + - @ ile ba#slayan metin form#ul san#il#ir: HucreMetni / KorumaGerekir."
    GMetin "- Sayfa adlar#i say#i/tarih olabilir: CStr(ws.Name), men#u h#ucreleri metin bi#cimli."
    GMetin ""
    GMetin "8. B#IL#INEN SINIRLILIKLAR"
    GMetin "- Kay#it baz#inda m#ukerrer kontrol#u yok (dosya/sayfa d#uzeyinde)."
    GMetin "- Ba#sl#iklar 1. sat#irda olmal#i. Kilitleme yok (tek kullan#ic#i)."
    GMetin "- Ye#sil tespiti renk tonuna dayan#ir."
    GMetin ""
    GMetin "9. KOD (a#sa#g#ida; sat#ir ba#s#indaki bir bo#sluk h#ucreye g#uvenli yazmak i#cindir)"
End Sub


Private Sub GElleKodYaz()
    GMetin "KOD OTOMAT#IK OKUNAMADI"
    GMetin "Excel'in ""VBA proje nesne modeline eri#sime g#uven"" ayar#i kapal#i oldu#gu i#cin"
    GMetin "kod bu sayfaya yaz#ilamad#i. Kodun tamam#i src\ klas#or#undedir (README.md)."
    GMetin "(Ayar: Dosya > Se#cenekler > G#uven Merkezi > Makro Ayarlar#i; BT kontrol#undedir.)"
End Sub
