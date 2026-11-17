Attribute VB_Name = "M13_Koruma"
Option Explicit

' Accidental-edit protection only; not encryption or access control.
' UserInterfaceOnly is not persisted by Excel and must be reapplied on open.
Public Sub KorumalariUygula()
    Dim ws As Worksheet, pnl As Worksheet, k As Long, c As Long, r As Variant
    Dim onceki As Object
    Dim hata As String
    EkranKilitle
    On Error GoTo Hata
    Set onceki = ActiveSheet
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SAYFA_PANEL, SAYFA_RAPOR, SAYFA_GECMIS, SAYFA_DENETIM, SAYFA_PROFIL, SAYFA_SET, SAYFA_AYARLAR, SAYFA_ANALIZ
                ws.Unprotect
                ws.Cells.Locked = True
                ws.Cells.Font.Name = PANEL_YAZI_TIPI
                If ws.Name = SAYFA_PANEL Then
                    Set pnl = ws
                    For k = 1 To BlokSayisi(pnl)
                        c = BlokSutun(k)
                        For Each r In Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                                            SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, SATIR_PROFIL, _
                                            SATIR_TUTAR, SATIR_HASAR, SATIR_IHBAR, SATIR_ANAHTAR, SATIR_KIRILIM, SATIR_DONEM)
                            pnl.Cells(CLng(r), c + OFS_AYAR).Locked = False
                        Next r
                        pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_DEGER)).Locked = False
                        pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(SON_SATIR, c + OFS_ISLEM)).Locked = False
                    Next k
                ElseIf ws.Name = SAYFA_AYARLAR Then
                    ws.Range(ws.Cells(AYR_ILK, AYR_DEGER), ws.Cells(AYR_ILK + ESIK_SAYI - 1, AYR_DEGER)).Locked = False
                    ws.Range(ws.Cells(AYR_Z_ILK, AYR_Z_SAYFA), ws.Cells(AYR_Z_SON, AYR_Z_SUTUN)).Locked = False
                ElseIf ws.Name = SAYFA_ANALIZ Then
                    ws.Range(ws.Cells(AN_ILK_ALAN, AN_DEGER_SUTUN), ws.Cells(AN_SON_ALAN, AN_DEGER_SUTUN)).Locked = False
                ElseIf ws.Name = SAYFA_RAPOR Or ws.Name = SAYFA_GECMIS Then
                    If ws.AutoFilterMode Then ws.AutoFilterMode = False
                    ws.UsedRange.AutoFilter
                    ws.Activate
                    ActiveWindow.FreezePanes = False
                    ActiveWindow.SplitRow = 1
                    ActiveWindow.SplitColumn = 0
                    ActiveWindow.FreezePanes = True
                End If
                ws.Protect UserInterfaceOnly:=True, AllowFiltering:=True
        End Select
    Next ws
    If SayfaVarMi(SAYFA_SET) Then
        ThisWorkbook.Worksheets(SAYFA_SET).Columns(ST_ESKI_SABIT_SUTUN).Hidden = True
        ThisWorkbook.Worksheets(SAYFA_SET).Columns(ST_ESKI_SABIT_DEGER).Hidden = True
    End If
    SekmeleriDuzenle
    BaskiAyarlari
Cikis:
    If Not onceki Is Nothing Then
        If onceki.Visible = xlSheetVisible Then onceki.Activate
    End If
    EkranAc
    Exit Sub
Hata:
    hata = Err.Number & " - " & Err.Description
    EkranAc
    Err.Raise vbObjectError + 1103, "KorumalariUygula", hata
End Sub


' SY-07: her gorunur sayfa icin baski alani ve olcek. Hedef: bos yayin
' kitabinin tamami < 15 sayfa (Panel 1, Analiz 1, Denetim 1, Rapor 1,
' Gecmis 1, Profiller 1, Setler 1, Ayarlar 1, Kilavuz <= 6).
Public Sub BaskiAyarlari()

    Dim ws As Worksheet
    Dim son As Long, sonSut As Long

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            With ws.PageSetup
                .PaperSize = xlPaperA4
                .Zoom = False
                .CenterFooter = "&P / &N"
                .LeftHeader = "&F  " & ChrW(183) & "  &A"
                .RightHeader = "&D"
                Select Case ws.Name
                    Case SAYFA_PANEL
                        son = SATIR_DURUM
                        If BlokSonSatir(ws, 1) > son Then son = BlokSonSatir(ws, 1)
                        sonSut = BlokSayisi(ws) * BLOK_GENISLIK
                        If sonSut < BLOK_GENISLIK Then sonSut = BLOK_GENISLIK
                        .Orientation = xlLandscape
                        .PrintArea = ws.Range(ws.Cells(SATIR_BLOK_BASLIK, 1), ws.Cells(son, sonSut)).Address
                        .FitToPagesWide = 1
                        .FitToPagesTall = 1
                    Case SAYFA_RAPOR, SAYFA_GECMIS, SAYFA_PROFIL, SAYFA_SET
                        .Orientation = xlLandscape
                        .PrintTitleRows = "$1:$1"
                        .PrintArea = ""
                        .FitToPagesWide = 1
                        .FitToPagesTall = False
                    Case SAYFA_DENETIM, SAYFA_AYARLAR, SAYFA_ANALIZ
                        .Orientation = xlLandscape
                        .FitToPagesWide = 1
                        .FitToPagesTall = 1
                    Case Else
                        .Orientation = xlPortrait
                        .FitToPagesWide = 1
                        .FitToPagesTall = False
                End Select
            End With
        End If
    Next ws
    On Error GoTo 0

End Sub
