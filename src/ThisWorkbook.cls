' ============================================================
' AKTARIM PANELI  v1.0  -  "ThisWorkbook" KOD SAYFASI
'
' Bu dosya bir sinif modulu OLARAK ICE AKTARILMAZ. tools\build.py
' icerigi ThisWorkbook'un kod sayfasina yazar. Elle kurulumda: VBA
' editorunde "ThisWorkbook"a cift tikla, "Option Explicit" ve
' sonrasini yapistir.
'
'   Workbook_Open -> UygulamaDurumunuOnar (M08_Kurulum): bir makro yarida
'                    kesildiyse kapali kalmis olaylari/ekrani geri acar.
' ============================================================
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo Hata
    UygulamaDurumunuOnar
    KorumalariUygula
    KisayollariKur True
    Exit Sub
Hata:
    MsgBox "Acilis / koruma hatasi: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' UX-03: Ctrl+Shift+1..4 = adimlar; yalnizca bu kitap etkinken (Excel'in
' ayni kisayollari diger kitaplarda bozulmasin).
Private Sub Workbook_Activate()
    On Error Resume Next
    KisayollariKur True
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    KisayollariKur False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    KisayollariKur False
End Sub

' Analiz istek karti (sayfanin kendi kod modulu yok; olay buradan yonlenir)
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Sh.Name = SAYFA_ANALIZ Then AnalizDegisti Target
End Sub
