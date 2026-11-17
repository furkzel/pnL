Attribute VB_Name = "M11_Analiz"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M11_Analiz  (iskelet; istek karti E asamasinda)
' ============================================================

Public Sub AnalizAc()
    On Error Resume Next
    If SayfaVarMi(SAYFA_ANALIZ) Then ThisWorkbook.Worksheets(SAYFA_ANALIZ).Activate
    On Error GoTo 0
End Sub


Public Sub AnalizDegisti(ByVal Target As Range)
End Sub
