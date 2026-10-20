' ============================================================
' AKTARIM PANELI  v1.0  -  "Panel" SAYFASININ KOD SAYFASI
'
' Bu dosya bir sinif modulu OLARAK ICE AKTARILMAZ. tools\build.py
' icerigi Panel sayfasinin kendi kod sayfasina yazar. Elle kurulumda:
' VBA editorunde parantez icinde "(Panel)" yazan satira cift tikla,
' icindekini sil, bu dosyadaki "Option Explicit" ve sonrasini yapistir.
'
'   Worksheet_Change          -> HedefDegisti (M04_Yenile): oneri onayi,
'                                SABIT DEGERLER, "Bu paneli aktar", dosya secimi
'   Worksheet_SelectionChange -> ButonlariEkranaTasi (M03_Panel): serit
'                                gorunen alanin soluna gelir, AKTIF panel
'                                gostergesi guncellenir
'   Worksheet_Activate        -> ayni (sayfaya donuldugunde)
'   Worksheet_BeforeDoubleClick -> PanelCiftTik (M04_Yenile): turuncu oneriye
'                                cift tik = onay; ISLEM hucresine cift tik =
'                                SEC ac/kapa
' ============================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    HedefDegisti Target
    On Error GoTo 0
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    ButonlariEkranaTasi
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    g_SonAktifBlok = 0
    ButonlariEkranaTasi
    On Error GoTo 0
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    PanelCiftTik Target, Cancel
    On Error GoTo 0
End Sub
