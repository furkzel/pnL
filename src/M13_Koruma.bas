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
            Case SAYFA_PANEL, SAYFA_RAPOR, SAYFA_GECMIS, SAYFA_DENETIM, SAYFA_PROFIL, SAYFA_SET
                ws.Unprotect
                ws.Cells.Locked = True
                ws.Cells.Font.Name = PANEL_YAZI_TIPI
                If ws.Name = SAYFA_PANEL Then
                    Set pnl = ws
                    For k = 1 To BlokSayisi(pnl)
                        c = BlokSutun(k)
                        For Each r In Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                                            SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, SATIR_PROFIL, _
                                            SATIR_TUTAR, SATIR_HASAR, SATIR_IHBAR)
                            pnl.Cells(CLng(r), c + OFS_AYAR).Locked = False
                        Next r
                        pnl.Range(pnl.Cells(SABIT_ILK, c + OFS_SABIT_SUTUN), pnl.Cells(SABIT_SON, c + OFS_SABIT_DEGER)).Locked = False
                        pnl.Range(pnl.Cells(ILK_SATIR, c + OFS_HEDEF), pnl.Cells(SON_SATIR, c + OFS_ISLEM)).Locked = False
                    Next k
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
