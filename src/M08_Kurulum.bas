Attribute VB_Name = "M08_Kurulum"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.0  -  M08_Kurulum
'
'   SifirdanKur      : TEK TIKLA KURULUM / YUKSELTME. Kayit sayfalarini
'                      hazirlar, v33'ten kalan analiz sayfalarini (Ozet,
'                      Dosya Durumu, Pivot Kontrol, OzetVeri) ve adlarini
'                      siler, Paneli v34 duzenine tasir (ayarlar ve
'                      eslestirmeler KORUNUR), Kilavuz ve Gelistirici'yi
'                      kurar, sekmeleri duzenler.
'   PaneliYenidenKur : yalnizca Panel ekranini yeniden cizer; icerik korunur
'                      (v33 7-sutunlu bloklar 6-sutunlu v34 bloklara tasinir)
'   PanelKur         : Paneli BOS olarak sifirdan kurar (eslestirmeler silinir)
'   SayfaDuzeni      : v33 adiyla cagrilabilsin diye SifirdanKur'a yonlendirir
'   Tani             : surum, dosyalar, paneller; Excel durumunu onarir
'   Yardim           : Kilavuz'u acar
' ============================================================

Private Const AYAR_SAYISI As Long = 11

' Panelden okunan bir blogun tum icerigi (yeniden cizim / gecis icin)
Private Type BlokIcerik
    Var As Boolean
    Gizli As Boolean
    Ayarlar(1 To AYAR_SAYISI) As Variant
    Sabitler As String
    Adet As Long
    Kaynak() As String
    Hedef() As String
    Filtre() As String
    Islem() As String
    Oneri() As String
End Type


' ============================================================
'  SIFIRDAN KUR / YUKSELT
' ============================================================
Public Sub SifirdanKur()

    Dim eskiler As Collection
    Dim liste As String, sonuc As String, panelSonucu As String, hata As String
    Dim i As Long
    Dim panelVar As Boolean, kilavuzTamam As Boolean, gelistiriciTamam As Boolean
    Dim artik As String

    Set eskiler = EskiSayfalar(liste)
    panelVar = SayfaVarMi(SAYFA_PANEL)
    artik = ArtikModuller()

    If FazlaPanelVar() <> "" Then
        MsgBox "KURULUM DURDU - hicbir sey degistirilmedi" & vbNewLine & vbNewLine & _
               FazlaPanelVar() & vbNewLine & vbNewLine & _
               "v1.0 en fazla " & MAX_BLOK & " panel destekler. Fazla panelleri once" & vbNewLine & _
               "'Set Kaydet' ile saklayip v33'te 'Panel Kaldir' ile kaldir.", vbCritical, "Sifirdan Kur"
        Exit Sub
    End If

    If artik <> "" Then
        MsgBox "KURULUM DURDU" & vbNewLine & vbNewLine & _
               "VBA projesinde eski surumun modulleri duruyor: " & artik & vbNewLine & vbNewLine & _
               "Bunlar v1.0 modulleriyle ayni adli prosedurler icerir ve" & vbNewLine & _
               "'Ambiguous name' hatasina yol acar. VBA editorunde (Alt+F11)" & vbNewLine & _
               "bu modullere sag tikla > Remove (Export etmeden) ve tekrar dene." & vbNewLine & _
               "(tools/build.py bunu otomatik yapar.)", vbCritical, "Sifirdan Kur"
        Exit Sub
    End If

    If MsgBox("AKTARIM PANELI " & SURUM & " - KURULUM / YUKSELTME" & vbNewLine & vbNewLine & _
              "Yapilacaklar:" & vbNewLine & _
              "  - Rapor, Gecmis, Profiller, Setler hazirlanir (KAYITLAR KORUNUR)" & vbNewLine & _
              IIf(panelVar, "  - Panel yeni duzene tasinir (ayarlar ve eslestirmeler KORUNUR)", _
                            "  - Panel sayfasi kurulur") & vbNewLine & _
              "  - Kilavuz ve Gelistirici sayfalari yeniden yazilir" & vbNewLine & _
              IIf(eskiler.Count > 0, "  - Artik kullanilmayan su sayfalar SILINIR:" & vbNewLine & liste, "") & vbNewLine & _
              "Oneri: once dosyanin bir yedegini al." & vbNewLine & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbQuestion + vbDefaultButton2, "Sifirdan Kur") <> vbYes Then
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Hata

    RaporSayfasiHazirla
    GecmisSayfasiHazirla
    ProfilSayfasiHazirla
    SetSayfasiHazirla

    ' --- v33 analiz sayfalari ve adlari ---
    If eskiler.Count > 0 Then
        EskiAdlariSil
        For i = 1 To eskiler.Count
            On Error Resume Next
            ThisWorkbook.Worksheets(CStr(eskiler(i))).Delete
            On Error GoTo Hata
        Next i
        sonuc = sonuc & "Silinen eski sayfa: " & eskiler.Count & vbNewLine
    End If
    EskiAdlariSil

    ' --- Panel ---
    panelSonucu = PaneliYenidenKurIc()
    sonuc = sonuc & "Panel: " & panelSonucu & vbNewLine

    ' --- Kilavuz ve Gelistirici ---
    On Error Resume Next
    KilavuzKur
    kilavuzTamam = SayfaVarMi(KilavuzAdi())
    GelistiriciGuncelle
    gelistiriciTamam = SayfaVarMi(SAYFA_GELISTIRICI)
    On Error GoTo Hata

    DenetimSessiz
    SekmeleriDuzenle
    EkranAc

    On Error Resume Next
    ThisWorkbook.Worksheets(SAYFA_PANEL).Activate
    PaneleGit 1
    On Error GoTo 0

    OlayYaz "KURULUM", 0, 0, 0, "SifirdanKur (" & SURUM & "): " & TekSatir(sonuc) & " Kilavuz: " & _
            IIf(kilavuzTamam, "kuruldu", "KURULAMADI") & ", Gelistirici: " & _
            IIf(gelistiriciTamam, "guncellendi", "GUNCELLENEMEDI"), ""

    MsgBox "KURULUM TAMAMLANDI  (" & SURUM & ")" & vbNewLine & vbNewLine & _
           sonuc & _
           "Rapor, Gecmis, Profiller, Setler : hazir" & vbNewLine & _
           "Kilavuz      : " & IIf(kilavuzTamam, "kuruldu", "KURULAMADI") & vbNewLine & _
           "Gelistirici  : " & IIf(gelistiriciTamam, "guncellendi", "GUNCELLENEMEDI") & vbNewLine & vbNewLine & _
           "SIRADAKI ADIM: kaynak ve ana dosyalari ac, YENILE'ye bas," & vbNewLine & _
           "sonra bu dosyayi kaydet.", vbInformation, "Sifirdan Kur"
    Exit Sub

Hata:
    hata = Err.Number & " - " & Err.Description
    EkranAc True
    MsgBox "KURULUM SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Hata : " & hata & vbNewLine & vbNewLine & _
           "Paneldeki eski bilgiler, panel yeniden cizilmeden once" & vbNewLine & _
           "okunur; hata o asamadan sonra olduysa Setler/Profiller" & vbNewLine & _
           "ile yeniden kurabilirsin. Rapor ve Gecmis etkilenmez.", vbCritical, "Sifirdan Kur"

End Sub


' v33 kurulum notlarinda adi gecen makro
Public Sub SayfaDuzeni()
    SifirdanKur
End Sub


' Panel ekranini yeniden cizer; tum panel icerigini korur.
Public Sub PaneliYenidenKur()

    Dim sonuc As String

    If MsgBox("Panel ekrani yeniden cizilecek." & vbNewLine & vbNewLine & _
              "Butun panellerin ayarlari, SABIT DEGERLER ve eslestirmeleri" & vbNewLine & _
              "korunur. Eski (v33) duzendeki panel v1.0 duzenine tasinir." & vbNewLine & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbQuestion, "Paneli Yeniden Kur") <> vbYes Then Exit Sub

    EkranKilitle
    On Error GoTo Hata
    sonuc = PaneliYenidenKurIc()
    EkranAc
    PaneleGit 1
    OlayYaz "KURULUM", 0, 0, 0, "PaneliYenidenKur: " & sonuc, ""
    MsgBox "Panel yeniden cizildi." & vbNewLine & vbNewLine & sonuc & vbNewLine & vbNewLine & _
           "Dosyalari acip YENILE'ye bas.", vbInformation, "Paneli Yeniden Kur"
    Exit Sub

Hata:
    EkranAc True
    MsgBox "PANEL YENIDEN KURULAMADI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Paneli Yeniden Kur"

End Sub


' Paneli BOS olarak kurar (tum paneller ve eslestirmeler silinir)
Public Sub PanelKur()

    Dim pnl As Worksheet

    If SayfaVarMi(SAYFA_PANEL) Then
        If MsgBox("PANEL KUR (BOS)" & vbNewLine & vbNewLine & _
                  "Panel sayfasi SIFIRDAN kurulacak: tum panel bloklari" & vbNewLine & _
                  "ve paneldeki eslestirmeler SILINIR." & vbNewLine & vbNewLine & _
                  "Profiller, Setler, Rapor ve Gecmis KORUNUR." & vbNewLine & _
                  "(Icerigi koruyarak yeniden cizmek icin: PaneliYenidenKur)" & vbNewLine & vbNewLine & _
                  "Devam edilsin mi?", vbYesNo + vbExclamation + vbDefaultButton2, "Panel Kur") <> vbYes Then
            Exit Sub
        End If
    End If

    EkranKilitle
    On Error GoTo Hata
    RaporSayfasiHazirla
    GecmisSayfasiHazirla
    ProfilSayfasiHazirla
    SetSayfasiHazirla

    Set pnl = PanelHazirla()
    PanelTemizle pnl
    DuzenIsaretle
    BlokCiz pnl, 1
    PanelGorunumu pnl
    EkranAc

    YenileCalistir True
    PaneleGit 1
    OlayYaz "KURULUM", 0, 0, 0, "Panel sayfasi bos olarak kuruldu (PanelKur, " & SURUM & ")", ""

    MsgBox "Panel kuruldu." & vbNewLine & vbNewLine & _
           "AKIS:" & vbNewLine & _
           "1. Kaynak ve ana dosyalari ac (Enable Editing)." & vbNewLine & _
           "2. Panelde dosya/sayfa sec, YENILE'ye bas." & vbNewLine & _
           "3. Otomatik Eslestir, onerileri onayla, On Kontrol." & vbNewLine & _
           "4. Aktaracagin panellerde 'Bu paneli aktar' = Evet." & vbNewLine & _
           "5. AKTAR." & vbNewLine & vbNewLine & _
           "Ayrintili anlatim: KILAVUZ.", vbInformation, "Panel Kur"
    Exit Sub

Hata:
    EkranAc True
    MsgBox "PANEL KURULAMADI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Panel Kur"

End Sub


' ============================================================
'  PANELI YENIDEN KURMA (ic)
' ============================================================

' Paneli okur (v33 veya v34 duzeni), temizler ve v34 duzeninde yeniden
' cizer. Donus: kisa sonuc metni.
Private Function PaneliYenidenKurIc() As String

    Dim pnl As Worksheet
    Dim bloklar() As BlokIcerik
    Dim adet As Long, k As Long, gorunur As Long, tasinan As Long
    Dim v33 As Boolean, v34 As Boolean

    If Not SayfaVarMi(SAYFA_PANEL) Then
        Set pnl = PanelHazirla()
        PanelTemizle pnl
        DuzenIsaretle
        BlokCiz pnl, 1
        PanelGorunumu pnl
        PaneliYenidenKurIc = "yeni kuruldu (1 panel)"
        Exit Function
    End If

    Set pnl = PanelHazirla()
    v34 = (KayitliDuzen() = "v34")
    v33 = Not PanelDuzeniGuncel() And Not v34 And V33DuzeniMi(pnl)

    ' --- 1. OKU (hicbir sey silinmeden once) ---
    If v33 Then
        adet = V33BlokSayisi(pnl)
    ElseIf PanelDuzeniGuncel() Or v34 Then
        adet = BlokSayisi(pnl)
    Else
        adet = 0
    End If

    If adet > 0 Then
        ReDim bloklar(1 To adet)
        For k = 1 To adet
            bloklar(k) = BlokOku(pnl, k, v33, v34)
            If bloklar(k).Var And Not bloklar(k).Gizli Then
                gorunur = gorunur + 1
                If k > MAX_BLOK Then
                    Err.Raise vbObjectError + 600, , "PANEL " & k & " dolu; v1.0 en fazla " & MAX_BLOK & _
                              " panel destekler. Once fazla panelleri Set olarak kaydedip kaldir."
                End If
            End If
        Next k
    End If

    ' --- 2. TEMIZLE ---
    PanelTemizle pnl
    DuzenIsaretle

    ' --- 3. YENIDEN CIZ ---
    If adet > MAX_BLOK Then adet = MAX_BLOK
    If adet = 0 Then
        BlokCiz pnl, 1
    Else
        For k = 1 To adet
            BlokCiz pnl, k
            If bloklar(k).Gizli Or Not bloklar(k).Var Then
                AlanTemizle pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)), _
                                      pnl.Cells(SON_SATIR, BlokSutun(k) + BLOK_GENISLIK - 1))
                pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(k)).Value = "PANEL " & k
                pnl.Range(pnl.Columns(BlokSutun(k)), pnl.Columns(BlokSutun(k) + BLOK_GENISLIK - 1)).Hidden = True
            Else
                BlokYaz pnl, k, bloklar(k)
                tasinan = tasinan + 1
            End If
        Next k
    End If

    FazlaListeleriSil
    PanelGorunumu pnl

    If v33 Then
        PaneliYenidenKurIc = "v33 duzeninden tasindi (" & tasinan & " panel; ayarlar, sabit degerler ve eslestirmeler korundu)"
    ElseIf v34 Then
        PaneliYenidenKurIc = "v34 duzeninden tasindi (" & tasinan & " panel; FILTRE/ISLEM sola, HEDEF HARITASI saga alindi; icerik korundu)"
    ElseIf adet > 0 Then
        PaneliYenidenKurIc = "yeniden cizildi (" & tasinan & " panel, icerik korundu)"
    Else
        PaneliYenidenKurIc = "kuruldu (1 panel)"
    End If

End Function


Private Function PanelHazirla() As Worksheet
    Dim pnl As Worksheet
    Set pnl = SayfaOlustur(SAYFA_PANEL)
    pnl.Visible = xlSheetVisible
    Set PanelHazirla = pnl
End Function


' Calisma kitabindaki panel duzeni adi: "v1.0" (eski adi "v35"), "v34" ya da "" (v33 ve oncesi)
Private Function KayitliDuzen() As String
    Dim s As String
    On Error Resume Next
    s = ThisWorkbook.Names(DUZEN_ADI).RefersTo
    On Error GoTo 0
    s = Replace(Replace(s, "=", ""), Chr$(34), "")
    KayitliDuzen = Trim$(s)
End Function


Private Sub DuzenIsaretle()
    On Error Resume Next
    ThisWorkbook.Names(DUZEN_ADI).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=DUZEN_ADI, RefersTo:="=""" & DUZEN_DEGERI & """", Visible:=False
End Sub


' Panel sayfasini belirli bir aralikta temizler (ws.Cells.Clear YOK)
Private Sub PanelTemizle(ByVal pnl As Worksheet)

    Dim alan As Range

    On Error Resume Next
    pnl.Buttons.Delete
    On Error GoTo 0
    SekilleriSil pnl

    Set alan = pnl.Range(pnl.Cells(1, 1), pnl.Cells(SON_SATIR + 20, TEMIZLIK_SON_SUTUN))
    alan.EntireColumn.Hidden = False
    AlanTemizle alan
    pnl.Range(pnl.Columns(1), pnl.Columns(TEMIZLIK_SON_SUTUN)).ColumnWidth = 9
    pnl.Range(pnl.Rows(1), pnl.Rows(SON_SATIR + 20)).RowHeight = 15

    On Error Resume Next
    pnl.Cells.Font.Name = PANEL_YAZI_TIPI
    pnl.Cells.Font.Size = PANEL_YAZI_BOYUT
    On Error GoTo 0

End Sub


' Serit, dondurma, satir yukseklikleri, sekme rengi
Private Sub PanelGorunumu(ByVal pnl As Worksheet)
    PanelSatirYukseklikleri pnl
    ButonlariKur pnl
    SeritDondur pnl
    On Error Resume Next
    pnl.Tab.Color = RGB(22, 163, 74)
    On Error GoTo 0
End Sub


' v33 panelinde MAX_BLOK'tan sonra DOLU (gizli olmayan) blok var mi?
Private Function FazlaPanelVar() As String

    Dim pnl As Worksheet
    Dim k As Long, c As Long
    Dim metin As String

    FazlaPanelVar = ""
    If Not SayfaVarMi(SAYFA_PANEL) Then Exit Function
    Set pnl = ThisWorkbook.Worksheets(SAYFA_PANEL)
    If PanelDuzeniGuncel() Or Not V33DuzeniMi(pnl) Then Exit Function

    For k = MAX_BLOK + 1 To V33BlokSayisi(pnl)
        c = 1 + (k - 1) * V33_BLOK_GENISLIK
        If Not pnl.Columns(c).Hidden Then metin = metin & k & " "
    Next k
    If metin <> "" Then FazlaPanelVar = "Panelde " & MAX_BLOK & "'dan sonra dolu panel var: PANEL " & Trim$(metin)

End Function


' v33 paneli mi? (7 sutunlu bloklar: B28 "HEDEF SUTUN", A28 "KAYNAK SUTUN")
Private Function V33DuzeniMi(ByVal pnl As Worksheet) As Boolean
    V33DuzeniMi = (BH(Trim$(GuvenliMetin(pnl.Cells(BASLIK_SATIR, 1).Value))) = "KAYNAK SUTUN" And _
                   BH(Trim$(GuvenliMetin(pnl.Cells(BASLIK_SATIR, 2).Value))) = "HEDEF SUTUN") Or _
                  (Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SATIR_BLOK_BASLIK, 1).Value))), 5) = "PANEL" And _
                   Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SATIR_BLOK_BASLIK, 1 + V33_BLOK_GENISLIK).Value))), 7) = "PANEL 2")
End Function


Private Function V33BlokSayisi(ByVal pnl As Worksheet) As Long
    Dim k As Long
    For k = 1 To V33_MAX_BLOK
        If Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SATIR_BLOK_BASLIK, 1 + (k - 1) * V33_BLOK_GENISLIK).Value))), 5) <> "PANEL" Then
            V33BlokSayisi = k - 1
            Exit Function
        End If
    Next k
    V33BlokSayisi = V33_MAX_BLOK
End Function


' Bir blogun butun icerigini okur (v33 ya da v34 yerlesimi)
Private Function BlokOku(ByVal pnl As Worksheet, ByVal k As Long, ByVal v33 As Boolean, _
                         ByVal v34 As Boolean) As BlokIcerik

    Dim b As BlokIcerik
    Dim c As Long, i As Long, son As Long, n As Long
    Dim satirlar As Variant
    Dim oKaynak As Long, oHedef As Long, oFiltre As Long, oDeger As Long, oIslem As Long
    Dim oSabitS As Long, oSabitD As Long
    Dim sec As Boolean, formul As Boolean
    Dim r As Long, sabitler As String, eskiSutun As String, eskiDeger As String

    If v33 Then
        c = 1 + (k - 1) * V33_BLOK_GENISLIK
        oKaynak = 0
        oHedef = V33_OFS_HEDEF
        oFiltre = V33_OFS_FILTRE
        oDeger = 1
        oSabitS = V33_SABIT_SUTUN_OFS
        oSabitD = V33_SABIT_DEGER_OFS
        satirlar = Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                         SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, SATIR_PROFIL, _
                         V33_SATIR_TUTAR, V33_SATIR_HASAR, V33_SATIR_IHBAR)
    Else
        c = BlokSutun(k)
        oKaynak = OFS_KAYNAK
        oHedef = OFS_HEDEF
        oFiltre = IIf(v34, V34_OFS_FILTRE, OFS_FILTRE)
        oIslem = IIf(v34, V34_OFS_ISLEM, OFS_ISLEM)
        oDeger = OFS_AYAR
        oSabitS = OFS_SABIT_SUTUN
        oSabitD = OFS_SABIT_DEGER
        satirlar = AyarSatirlari()
    End If

    b.Var = (Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SATIR_BLOK_BASLIK, c).Value))), 5) = "PANEL")
    b.Gizli = pnl.Columns(c).Hidden
    If Not b.Var Or b.Gizli Then
        BlokOku = b
        Exit Function
    End If

    For i = 1 To AYAR_SAYISI
        b.Ayarlar(i) = pnl.Cells(satirlar(i - 1), c + oDeger).Value
    Next i

    ' --- Sabit degerler ---
    sabitler = ""
    If v33 And Left$(BH(Trim$(GuvenliMetin(pnl.Cells(SABIT_BASLIK, c + oSabitS).Value))), 14) <> "SABIT DEGERLER" Then
        ' v31: tek sabit deger soldaki 19-20. satirlarda
        If BH(Trim$(GuvenliMetin(pnl.Cells(V31_SABIT_SUTUN, c).Value))) = "SABIT DEGER SUTUNU" Then
            eskiSutun = Trim$(GuvenliMetin(pnl.Cells(V31_SABIT_SUTUN, c + 1).Value))
            eskiDeger = GuvenliMetin(pnl.Cells(V31_SABIT_DEGER, c + 1).Value)
            If eskiSutun <> "" And Trim$(eskiDeger) <> "" And eskiDeger <> "(bos birak)" Then
                sabitler = eskiSutun & vbTab & eskiDeger
            End If
        End If
    Else
        For r = SABIT_ILK To SABIT_SON
            If Trim$(GuvenliMetin(pnl.Cells(r, c + oSabitS).Value)) <> "" Then
                If sabitler <> "" Then sabitler = sabitler & vbLf
                sabitler = sabitler & Trim$(GuvenliMetin(pnl.Cells(r, c + oSabitS).Value)) & vbTab & _
                           GuvenliMetin(pnl.Cells(r, c + oSabitD).Value)
            End If
        Next r
    End If
    b.Sabitler = sabitler

    ' --- Eslesme tablosu ---
    son = pnl.Cells(SON_SATIR + 1, c + oKaynak).End(xlUp).Row
    If son > SON_SATIR Then son = SON_SATIR
    n = son - ILK_SATIR + 1
    If n < 1 Then
        b.Adet = 0
        BlokOku = b
        Exit Function
    End If

    b.Adet = n
    ReDim b.Kaynak(1 To n)
    ReDim b.Hedef(1 To n)
    ReDim b.Filtre(1 To n)
    ReDim b.Islem(1 To n)
    ReDim b.Oneri(1 To n)
    For i = 1 To n
        r = ILK_SATIR + i - 1
        b.Kaynak(i) = GuvenliMetin(pnl.Cells(r, c + oKaynak).Value)
        b.Hedef(i) = GuvenliMetin(pnl.Cells(r, c + oHedef).Value)
        b.Filtre(i) = GuvenliMetin(pnl.Cells(r, c + oFiltre).Value)
        If v33 Then
            sec = (BH(Trim$(GuvenliMetin(pnl.Cells(r, c + V33_OFS_GERIAL).Value))) = "EVET")
            formul = (BH(Trim$(GuvenliMetin(pnl.Cells(r, c + V33_OFS_FORMUL).Value))) = "EVET")
            b.Islem(i) = IslemMetni(sec, formul)
        Else
            b.Islem(i) = GuvenliMetin(pnl.Cells(r, c + oIslem).Value)
        End If
        If OneriMi(pnl.Cells(r, c + oHedef)) Then b.Oneri(i) = OneriMetni(pnl.Cells(r, c + oHedef))
    Next i

    BlokOku = b

End Function


Private Function AyarSatirlari() As Variant
    AyarSatirlari = Array(SATIR_DAHIL, SATIR_KAYNAK_DOSYA, SATIR_KAYNAK_SAYFA, SATIR_SON_SATIR, _
                          SATIR_HEDEF_DOSYA, SATIR_HEDEF_SAYFA, SATIR_YAZIM, SATIR_PROFIL, _
                          SATIR_TUTAR, SATIR_HASAR, SATIR_IHBAR)
End Function


' Okunan icerigi v34 blogu olarak yazar
Private Sub BlokYaz(ByVal pnl As Worksheet, ByVal k As Long, ByRef b As BlokIcerik)

    Dim c As Long, i As Long, r As Long
    Dim satirlar As Variant
    Dim kaynak() As Variant, hedef() As Variant, filtre() As Variant, islem() As Variant

    c = BlokSutun(k)
    satirlar = AyarSatirlari()

    For i = 1 To AYAR_SAYISI
        r = satirlar(i - 1)
        Select Case r
            Case SATIR_DAHIL
                If BH(Trim$(GuvenliMetin(b.Ayarlar(i)))) = "EVET" Then
                    Ayar(pnl, k, r).Value = EvetMetni()
                Else
                    Ayar(pnl, k, r).Value = HayirMetni()
                End If
            Case SATIR_YAZIM
                Ayar(pnl, k, r).Value = YazimMetni(YazimModuOku(b.Ayarlar(i)))
            Case SATIR_SON_SATIR
                If Trim$(GuvenliMetin(b.Ayarlar(i))) <> "" Then Ayar(pnl, k, r).Value = b.Ayarlar(i)
            Case Else
                If VarType(b.Ayarlar(i)) = vbString Then
                    Ayar(pnl, k, r).Value = HucreMetni(CStr(b.Ayarlar(i)))
                Else
                    Ayar(pnl, k, r).Value = b.Ayarlar(i)
                End If
        End Select
    Next i
    DahilBoya pnl, k

    SabitleriYaz pnl, k, b.Sabitler
    SabitAlaniniDuzenle pnl, k, ListeSayfasi(k), Nothing

    If b.Adet > 0 Then
        ReDim kaynak(1 To b.Adet, 1 To 1)
        ReDim hedef(1 To b.Adet, 1 To 1)
        ReDim filtre(1 To b.Adet, 1 To 1)
        ReDim islem(1 To b.Adet, 1 To 1)
        For i = 1 To b.Adet
            kaynak(i, 1) = HucreMetni(b.Kaynak(i))
            hedef(i, 1) = HucreMetni(b.Hedef(i))
            filtre(i, 1) = b.Filtre(i)
            islem(i, 1) = IslemMetni(IslemSecili(b.Islem(i)), IslemFormul(b.Islem(i)))
        Next i
        pnl.Cells(ILK_SATIR, c + OFS_KAYNAK).Resize(b.Adet, 1).Value = kaynak
        pnl.Cells(ILK_SATIR, c + OFS_KAYNAK).Resize(b.Adet, 1).Interior.Color = Palet(prKaynakZemin)
        pnl.Cells(ILK_SATIR, c + OFS_HEDEF).Resize(b.Adet, 1).Value = hedef
        pnl.Cells(ILK_SATIR, c + OFS_HEDEF).Resize(b.Adet, 1).Interior.Color = Palet(prHedefZemin)
        With pnl.Cells(ILK_SATIR, c + OFS_FILTRE).Resize(b.Adet, 1)
            .NumberFormat = "@"
            .Value = filtre
            .Interior.Color = Palet(prFiltreZemin)
        End With
        With pnl.Cells(ILK_SATIR, c + OFS_ISLEM).Resize(b.Adet, 1)
            .Value = islem
            .Interior.Color = Palet(prIslemZemin)
            .HorizontalAlignment = xlCenter
        End With
        For i = 1 To b.Adet
            If b.Oneri(i) <> "" Then OneriIsaretle pnl.Cells(ILK_SATIR + i - 1, c + OFS_HEDEF), b.Oneri(i)
        Next i
    End If

    pnl.Cells(SATIR_DURUM, c).Value = TR("DURUM: ") & SURUM & TR(" d#uzenine ta#s#ind#i  #-  dosyalar#i a#c#ip YEN#ILE'ye bas#in")
    pnl.Cells(SATIR_DURUM, c).Interior.Color = Palet(prDurumUyari)
    If AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA) <> "" Then
        pnl.Cells(SATIR_BLOK_BASLIK, c).Value = HucreMetni("PANEL " & k & "   " & ChrW(8212) & "   " & _
                                                          AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA) & "  /  " & _
                                                          AyarMetni(pnl, k, SATIR_KAYNAK_SAYFA))
    End If

End Sub


' MAX_BLOK'tan sonraki Listeler sayfalari (v33'te 20'ye kadar) silinir
Private Sub FazlaListeleriSil()
    Dim k As Long
    On Error Resume Next
    For k = MAX_BLOK + 1 To V33_MAX_BLOK
        If SayfaVarMi(SAYFA_LISTE_ONEK & k) Then ThisWorkbook.Worksheets(SAYFA_LISTE_ONEK & k).Delete
    Next k
    For k = 1 To MAX_BLOK
        If SayfaVarMi(SAYFA_LISTE_ONEK & k) Then ThisWorkbook.Worksheets(SAYFA_LISTE_ONEK & k).UsedRange.Clear
    Next k
    On Error GoTo 0
End Sub


' ============================================================
'  ESKI SAYFALAR VE ADLAR (v33 Module3 ve daha eskileri)
' ============================================================
Private Function EskiSayfalar(ByRef liste As String) As Collection

    Dim ws As Worksheet
    Dim c As New Collection
    Dim ad As String, b As String

    liste = ""
    For Each ws In ThisWorkbook.Worksheets
        ad = ws.Name
        b = BH(ad)
        If b = "OZET" Or b = "DOSYA DURUMU" Or b = "PIVOT KONTROL" Or b = "OZETVERI" Or _
           b = "ONIZLEME" Or b = "HEDEFSUTUNLAR" Or b = "KONTROL" Or b = "RAPOR_GECIS" Or _
           Left$(b, 11) = "GECMIS_ESKI" Then
            c.Add ad
            liste = liste & "       " & ad & vbNewLine
        End If
    Next ws
    Set EskiSayfalar = c

End Function


' Module3'un (Ozet) olusturdugu adlar: OZ_*, ve OzetVeri / Ozet /
' Dosya Durumu / Pivot Kontrol sayfalarina bakan ya da #REF! olmus
' V_*, T_*, L_* adlari.
Private Sub EskiAdlariSil()

    Dim i As Long
    Dim nm As Name
    Dim ad As String, ref As String, b As String

    On Error Resume Next
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        ad = nm.Name
        ref = BH(nm.RefersTo)
        b = BH(ad)
        If Left$(b, 3) = "OZ_" Or InStr(ref, "OZETVERI") > 0 Or InStr(ref, "DOSYA DURUMU") > 0 Or _
           InStr(ref, "PIVOT KONTROL") > 0 Or InStr(ref, "OZET!") > 0 Or InStr(ref, "'OZET'!") > 0 Or _
           ((Left$(b, 2) = "V_" Or Left$(b, 2) = "T_" Or Left$(b, 2) = "L_") And InStr(ref, "#REF!") > 0) Then
            nm.Delete
        End If
    Next i
    On Error GoTo 0

End Sub


' VBA projesinde v33'ten kalan Module1/2/3 var mi? (erisim yoksa "")
Private Function ArtikModuller() As String

    Dim bilesen As Object
    Dim metin As String

    On Error Resume Next
    For Each bilesen In ThisWorkbook.VBProject.VBComponents
        If bilesen.Type = 1 Then
            Select Case bilesen.Name
                Case "Module1", "Module2", "Module3"
                    If bilesen.CodeModule.Find("AKTARIM PANELI", 1, 1, 30, 1) Then
                        metin = metin & bilesen.Name & " "
                    End If
            End Select
        End If
    Next bilesen
    On Error GoTo 0
    ArtikModuller = Trim$(metin)

End Function


Public Sub SekmeleriDuzenle()

    Dim sira As Variant, renk As Variant
    Dim i As Long, konum As Long
    Dim eskiSayfa As Object

    On Error Resume Next
    Set eskiSayfa = ActiveSheet

    sira = Array(KilavuzAdi(), SAYFA_PANEL, SAYFA_DENETIM, SAYFA_RAPOR, SAYFA_PROFIL, SAYFA_SET, SAYFA_GECMIS, SAYFA_GELISTIRICI)
    renk = Array(RGB(29, 78, 216), RGB(22, 163, 74), RGB(14, 116, 144), RGB(234, 88, 12), RGB(124, 58, 237), _
                 RGB(124, 58, 237), RGB(100, 116, 139), RGB(100, 116, 139))

    konum = 0
    For i = 0 To UBound(sira)
        If SayfaVarMi(CStr(sira(i))) Then
            konum = konum + 1
            If ThisWorkbook.Sheets(CStr(sira(i))).Index <> konum Then
                ThisWorkbook.Sheets(CStr(sira(i))).Move Before:=ThisWorkbook.Sheets(konum)
            End If
            ThisWorkbook.Sheets(CStr(sira(i))).Tab.Color = renk(i)
        End If
    Next i

    For i = 1 To MAX_BLOK
        If SayfaVarMi(SAYFA_LISTE_ONEK & i) Then ThisWorkbook.Worksheets(SAYFA_LISTE_ONEK & i).Visible = xlSheetHidden
    Next i
    If SayfaVarMi(SAYFA_YEDEK) Then ThisWorkbook.Worksheets(SAYFA_YEDEK).Visible = xlSheetHidden

    If Not eskiSayfa Is Nothing Then eskiSayfa.Activate
    On Error GoTo 0

End Sub


' ============================================================
'  YARDIM / TANI / ONARIM
' ============================================================
Public Sub Yardim()
    KilavuzAc
End Sub


' Workbook_Open ve Tani'dan: Excel'i her durumda normale dondurur
' (bir makro yarida kesilip olaylar kapali kaldiysa).
Public Sub UygulamaDurumunuOnar()
    EkranAc True
    g_SonAktifBlok = 0
End Sub


Public Sub Tani()

    Dim pnl As Worksheet
    Dim wb As Workbook
    Dim metin As String, durum As String
    Dim k As Long, adet As Long

    durum = "olaylar " & IIf(Application.EnableEvents, "acik", "KAPALIYDI") & _
            ", hesaplama " & IIf(Application.Calculation = xlCalculationAutomatic, "otomatik", "ELLE")
    UygulamaDurumunuOnar

    metin = "SURUM: " & SURUM & "     KULLANICI: " & KullaniciAdi() & vbNewLine
    metin = metin & "Excel durumu: " & durum & "  -> olaylar ve ekran acildi" & vbNewLine
    metin = metin & "Kilavuz sayfasi: " & IIf(SayfaVarMi(KilavuzAdi()), "var", "YOK (SifirdanKur calistir)") & vbNewLine
    metin = metin & "Panel duzeni: " & IIf(PanelDuzeniGuncel(), DUZEN_DEGERI, "ESKI - SifirdanKur calistir") & vbNewLine & vbNewLine

    metin = metin & "EXCEL'IN BU MAKROYA GOSTERDIGI DOSYALAR:" & vbNewLine
    For Each wb In Workbooks
        metin = metin & "  - " & Left$(wb.Name, 50) & IIf(wb.ReadOnly, "  (SALT OKUNUR)", "") & vbNewLine
    Next wb

    Set pnl = PanelSayfasi()
    If Not pnl Is Nothing And PanelDuzeniGuncel() Then
        adet = BlokSayisi(pnl)
        metin = metin & vbNewLine & "PANEL BLOKLARI: " & adet & "   (bulundugun: " & AktifBlok() & ")" & vbNewLine
        For k = 1 To adet
            If BlokGizliMi(pnl, k) Then
                metin = metin & "  P" & k & ": (kaldirilmis, gizli)" & vbNewLine
            Else
                metin = metin & "  P" & k & ": dahil=" & AyarMetni(pnl, k, SATIR_DAHIL) & _
                        "  kaynak=" & IIf(BlokKaynakSayfasi(pnl, k) Is Nothing, _
                                          "BULUNAMADI (" & YolDurumu(Ayar(pnl, k, SATIR_KAYNAK_DOSYA).Value) & ")", "tamam") & _
                        "  hedef=" & IIf(BlokHedefSayfasi(pnl, k) Is Nothing, _
                                         "BULUNAMADI (" & YolDurumu(Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value) & ")", "tamam") & vbNewLine
            End If
        Next k
    End If

    If g_SonAcmaHatasi <> "" Then
        metin = metin & vbNewLine & "SON DOSYA ACMA HATASI:" & vbNewLine & g_SonAcmaHatasi
    End If

    MsgBox Sigdir(metin, MSG_SINIR), vbInformation, "Tani"

End Sub
