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
    On Error Resume Next
    UygulamaDurumunuOnar
    On Error GoTo 0
End Sub
