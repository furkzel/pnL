Attribute VB_Name = "M16_Bellek"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M16_Bellek  (DURUM, BILDIRIM, KILIT)
'
'   BellekYaz / BellekOku : gizli "Bellek" sayfasinda anahtar/deger
'                           (Yenile'de silinmeyen durum: On Kontrol kaynak
'                           izi (K14), son aktarim semasi (K12) ...)
'   Bildir                : bilgilendirme mesajlari MODAL DEGIL. Seritteki
'                           bildirim seridine (son mesaj) ve Bellek'e
'                           (son 200, zaman damgali, Rapor satiriyla) yazilir.
'                           Karar gerektiren sorular MsgBox olarak kalir.
'   BildirimListesi       : seritteki bildirime tiklaninca son 5 mesaj
'   KilitAl / KilitBirak  : ana dosya yanina ".aplock" (kullanici, makine,
'                           zaman). Bayat kilit (Ayarlar: KILIT_DAKIKA)
'                           devralinir. Bulut / kaydedilmemis dosyada kilit
'                           kurulamaz: aktarim durmaz, soy notuna yazilir.
' ============================================================

Public Const BILDIRIM_ADI As String = "AP_BILDIRIM"
Public Const BILDIRIM_LISTE_ADI As String = "AP_BILDIRIM_LISTE"

Public Const BS_OK As String = "OK"
Public Const BS_BILGI As String = "BILGI"
Public Const BS_DIKKAT As String = "DIKKAT"
Public Const BS_ENGEL As String = "ENGEL"


Private Function BellekSayfasi() As Worksheet
    Dim ws As Worksheet
    Set ws = SayfaOlustur(SAYFA_BELLEK)
    If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
    If Trim$(GuvenliMetin(ws.Cells(1, BLK_ANAHTAR).Value)) = "" Then
        ws.Cells(1, BLK_ANAHTAR).Value = "ANAHTAR"
        ws.Cells(1, BLK_DEGER).Value = "DEGER"
        ws.Cells(1, BLK_ZAMAN).Value = "ZAMAN"
        ws.Cells(1, BLD_ZAMAN).Value = "BILDIRIM ZAMANI"
        ws.Cells(1, BLD_SEVIYE).Value = "SEVIYE"
        ws.Cells(1, BLD_BASLIK).Value = "BASLIK"
        ws.Cells(1, BLD_METIN).Value = "METIN"
        ws.Cells(1, BLD_RAPOR).Value = "RAPOR SATIRI"
    End If
    Set BellekSayfasi = ws
End Function


' ------------------------------------------------------------
'  ANAHTAR / DEGER
' ------------------------------------------------------------
Private Function BellekSatiri(ByVal ws As Worksheet, ByVal anahtar As String) As Long
    Dim son As Long, i As Long
    Dim v As Variant
    son = ws.Cells(ws.Rows.Count, BLK_ANAHTAR).End(xlUp).Row
    If son < 2 Then Exit Function
    v = SutunOku(ws, BLK_ANAHTAR, 2, son)
    For i = 1 To UBound(v, 1)
        If GuvenliMetin(v(i, 1)) = anahtar Then
            BellekSatiri = i + 1
            Exit Function
        End If
    Next i
End Function


Public Sub BellekYaz(ByVal anahtar As String, ByVal deger As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = BellekSayfasi()
    r = BellekSatiri(ws, anahtar)
    If r = 0 Then r = ws.Cells(ws.Rows.Count, BLK_ANAHTAR).End(xlUp).Row + 1
    ws.Cells(r, BLK_ANAHTAR).Value = "'" & anahtar
    ws.Cells(r, BLK_DEGER).Value = HucreMetni(Left$(deger, 30000))
    ws.Cells(r, BLK_ZAMAN).Value = Now
End Sub


Public Function BellekOku(ByVal anahtar As String) As String
    Dim ws As Worksheet
    Dim r As Long
    If Not SayfaVarMi(SAYFA_BELLEK) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SAYFA_BELLEK)
    r = BellekSatiri(ws, anahtar)
    If r > 0 Then BellekOku = GuvenliMetin(ws.Cells(r, BLK_DEGER).Value)
End Function


' ------------------------------------------------------------
'  BILDIRIM SERIDI (UX-01)
' ------------------------------------------------------------
Public Sub Bildir(ByVal seviye As String, ByVal baslik As String, ByVal metin As String, _
                  Optional ByVal raporSatiri As Long = 0)

    Dim ws As Worksheet
    Dim r As Long, n As Long

    On Error GoTo Cikis
    Set ws = BellekSayfasi()
    r = ws.Cells(ws.Rows.Count, BLD_ZAMAN).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, BLD_ZAMAN).Value = Now
    ws.Cells(r, BLD_SEVIYE).Value = seviye
    ws.Cells(r, BLD_BASLIK).Value = HucreMetni(baslik)
    ws.Cells(r, BLD_METIN).Value = HucreMetni(Left$(metin, 30000))
    If raporSatiri > 0 Then ws.Cells(r, BLD_RAPOR).Value = raporSatiri

    ' Halka: iki kat dolunca en eskiler atilir
    n = r - 1
    If n > 2 * BILDIRIM_TUTULAN Then
        ws.Range(ws.Cells(2, BLD_ZAMAN), ws.Cells(n - BILDIRIM_TUTULAN + 1, BLD_RAPOR)).Delete Shift:=xlUp
    End If

    BildirimSeridiniYaz
    DurumCubugu ""
Cikis:

End Sub


' Seritteki tek satirlik bildirim: "hh:nn  BASLIK  ilk satir"
Public Sub BildirimSeridiniYaz()

    Dim pnl As Worksheet
    Dim ws As Worksheet
    Dim shp As Object
    Dim r As Long
    Dim seviye As String, metin As String, ilk As String

    On Error GoTo Cikis
    Set pnl = PanelSayfasi()
    If pnl Is Nothing Then Exit Sub
    Set shp = pnl.Shapes(BILDIRIM_ADI)
    If Not SayfaVarMi(SAYFA_BELLEK) Then
        shp.TextFrame2.TextRange.Text = TR("Bildirim yok")
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SAYFA_BELLEK)
    r = ws.Cells(ws.Rows.Count, BLD_ZAMAN).End(xlUp).Row
    If r < 2 Then
        shp.TextFrame2.TextRange.Text = TR("Bildirim yok")
        Exit Sub
    End If
    seviye = GuvenliMetin(ws.Cells(r, BLD_SEVIYE).Value)
    ilk = Replace(GuvenliMetin(ws.Cells(r, BLD_METIN).Value), vbCr, "")
    If InStr(ilk, vbLf) > 0 Then ilk = Left$(ilk, InStr(ilk, vbLf) - 1)
    metin = BildirimGlifi(seviye) & "  " & Format$(ws.Cells(r, BLD_ZAMAN).Value, "hh:nn") & "  " & _
            GuvenliMetin(ws.Cells(r, BLD_BASLIK).Value) & "  " & ChrW(183) & "  " & Left$(ilk, 90)
    shp.TextFrame2.TextRange.Text = metin
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = BildirimRengi(seviye)
Cikis:

End Sub


Public Function BildirimGlifi(ByVal seviye As String) As String
    Select Case seviye
        Case BS_OK:     BildirimGlifi = ChrW(10003)
        Case BS_DIKKAT: BildirimGlifi = "!"
        Case BS_ENGEL:  BildirimGlifi = ChrW(10005)
        Case Else:      BildirimGlifi = ChrW(8226)
    End Select
End Function


Public Function BildirimRengi(ByVal seviye As String) As Long
    Select Case seviye
        Case BS_OK:     BildirimRengi = Palet(prOk)
        Case BS_DIKKAT: BildirimRengi = Palet(prDikkat)
        Case BS_ENGEL:  BildirimRengi = Palet(prEngel)
        Case Else:      BildirimRengi = Palet(prIkincilMetin)
    End Select
End Function


' Bildirim seridine tiklaninca: son 5 bildirimi acar/kapatir. Listeye
' tiklamak en son bildirimin Rapor satirina gider.
Public Sub BildirimListesi()

    Dim pnl As Worksheet, ws As Worksheet
    Dim shp As Shape, serit As Shape
    Dim r As Long, i As Long
    Dim metin As String, ilk As String

    EkranKilitle
    On Error GoTo Cikis
    Set pnl = PanelSayfasi()
    Set serit = pnl.Shapes(BILDIRIM_ADI)
    On Error Resume Next
    Set shp = pnl.Shapes(BILDIRIM_LISTE_ADI)
    On Error GoTo Cikis
    If Not shp Is Nothing Then
        shp.Delete
        GoTo Cikis
    End If
    If SayfaVarMi(SAYFA_BELLEK) Then
        Set ws = ThisWorkbook.Worksheets(SAYFA_BELLEK)
        r = ws.Cells(ws.Rows.Count, BLD_ZAMAN).End(xlUp).Row
        For i = r To 2 Step -1
            If r - i >= BILDIRIM_GOSTERILEN Then Exit For
            ilk = Replace(GuvenliMetin(ws.Cells(i, BLD_METIN).Value), vbCr, "")
            If InStr(ilk, vbLf) > 0 Then ilk = Left$(ilk, InStr(ilk, vbLf) - 1)
            metin = metin & BildirimGlifi(GuvenliMetin(ws.Cells(i, BLD_SEVIYE).Value)) & "  " & _
                    Format$(ws.Cells(i, BLD_ZAMAN).Value, "dd.mm hh:nn") & "   " & _
                    GuvenliMetin(ws.Cells(i, BLD_BASLIK).Value) & "  " & ChrW(183) & "  " & Left$(ilk, 110) & vbLf
        Next i
    End If
    If metin = "" Then metin = TR("Bildirim yok") & vbLf
    metin = metin & TR("T#iklay#in: Rapor'da son kayda git")

    Set shp = pnl.Shapes.AddShape(1, serit.Left, serit.Top + serit.Height + 2, serit.Width + 160, _
                                  (BILDIRIM_GOSTERILEN + 1) * 15 + 10)
    shp.Name = BILDIRIM_LISTE_ADI
    shp.Fill.ForeColor.RGB = Palet(prZemin)
    shp.Line.ForeColor.RGB = Palet(prCizgi)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = 0
    shp.OnAction = "BildirimRaporaGit"
    With shp.TextFrame2
        .MarginLeft = 8
        .MarginRight = 8
        .MarginTop = 4
        .MarginBottom = 4
        .WordWrap = 0
        .TextRange.Text = metin
        .TextRange.Font.Name = PANEL_YAZI_TIPI
        .TextRange.Font.Size = YZ_DIPNOT
        .TextRange.Font.Fill.ForeColor.RGB = Palet(prMetin)
        .TextRange.ParagraphFormat.Alignment = 1
    End With
    shp.Placement = xlFreeFloating
    shp.ZOrder 0
Cikis:
    EkranAc
End Sub


Public Sub BildirimRaporaGit()
    Dim ws As Worksheet
    Dim r As Long, i As Long, satir As Long
    On Error Resume Next
    PanelSayfasi().Shapes(BILDIRIM_LISTE_ADI).Delete
    If SayfaVarMi(SAYFA_BELLEK) Then
        Set ws = ThisWorkbook.Worksheets(SAYFA_BELLEK)
        r = ws.Cells(ws.Rows.Count, BLD_ZAMAN).End(xlUp).Row
        For i = r To 2 Step -1
            satir = GuvenliSayi(ws.Cells(i, BLD_RAPOR).Value)
            If satir > 0 Then Exit For
        Next i
    End If
    ThisWorkbook.Worksheets(SAYFA_RAPOR).Activate
    If satir > 0 Then
        ActiveWindow.ScrollRow = satir
        ThisWorkbook.Worksheets(SAYFA_RAPOR).Cells(satir, R_ACIKLAMA).Select
    End If
    On Error GoTo 0
End Sub


' ------------------------------------------------------------
'  ANA DOSYA KILIDI (M12)
' ------------------------------------------------------------
Private Function KilitYolu(ByVal wb As Workbook) As String
    KilitYolu = ""
    If wb Is Nothing Then Exit Function
    If wb.Path = "" Then Exit Function
    If InStr(wb.FullName, "://") > 0 Then Exit Function
    KilitYolu = wb.FullName & KILIT_UZANTI
End Function


Private Function Makine() As String
    On Error Resume Next
    Makine = Environ$("COMPUTERNAME")
    If Makine = "" Then Makine = Environ$("HOSTNAME")
    On Error GoTo 0
End Function


' Kilit baskasinda ve tazeyse True; sahip/zaman doldurulur.
Public Function KilitBaskasinda(ByVal wb As Workbook, ByRef sahip As String, ByRef zaman As String) As Boolean

    Dim yol As String, satir As String
    Dim f As Integer
    Dim yas As Double

    KilitBaskasinda = False
    sahip = ""
    zaman = ""
    yol = KilitYolu(wb)
    If yol = "" Then Exit Function
    On Error GoTo Cikis
    If Dir(yol) = "" Then Exit Function
    yas = (Now - FileDateTime(yol)) * 1440
    f = FreeFile
    Open yol For Input As #f
    If Not EOF(f) Then Line Input #f, satir
    sahip = satir
    If Not EOF(f) Then Line Input #f, satir
    If satir <> "" Then sahip = sahip & " @ " & satir
    If Not EOF(f) Then Line Input #f, satir
    zaman = satir
    Close #f
    If yas > Esik(ekKilitDakika) Then Exit Function
    If InStr(1, sahip, KullaniciAdi() & " @ " & Makine(), vbTextCompare) = 1 Then Exit Function
    KilitBaskasinda = True
    Exit Function
Cikis:
    On Error Resume Next
    If f > 0 Then Close #f

End Function


' Kilidi yazar. Donus: "" = alindi; aksi halde neden (aktarim durmaz).
Public Function KilitAl(ByVal wb As Workbook) As String

    Dim yol As String
    Dim f As Integer

    yol = KilitYolu(wb)
    If yol = "" Then
        KilitAl = "kilit kurulamadi (kaydedilmemis ya da bulut yolu)"
        Exit Function
    End If
    On Error GoTo Hata
    f = FreeFile
    Open yol For Output As #f
    Print #f, KullaniciAdi()
    Print #f, Makine()
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, ThisWorkbook.Name
    Close #f
    KilitAl = ""
    Exit Function
Hata:
    KilitAl = "kilit dosyasi yazilamadi (" & Err.Description & ")"
    On Error Resume Next
    If f > 0 Then Close #f

End Function


Public Sub KilitBirak(ByVal wb As Workbook)
    Dim yol As String, sahip As String, zaman As String
    yol = KilitYolu(wb)
    If yol = "" Then Exit Sub
    On Error Resume Next
    If KilitBaskasinda(wb, sahip, zaman) Then Exit Sub
    If Dir(yol) <> "" Then Kill yol
    On Error GoTo 0
End Sub
