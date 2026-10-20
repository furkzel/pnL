Attribute VB_Name = "M05_Eslestirme"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.0  -  M05_Eslestirme
'
'   OtomatikEslestir : 1. tur birebir ad (+ ayni adli sutunlarda SIRA),
'                      2. tur Levenshtein benzerligi (>= %70) -> ONERI
'   ONERI            : turuncu HEDEF hucresi + "ONERI:" ile baslayan not
'   OneriOnayla      : AKILLI onay - secili satirlarda oneri varsa
'                      onlari, yoksa bulundugun paneldeki tumunu onaylar
'   Profiller        : ProfilKaydet / ProfilUygula / ProfilSil
'                      (diger ad ogrenme, bilerek bos birakilan sutunlar,
'                      dosyada olmayan eski eslesmeleri koruma)
'   Setler           : SetKaydet / SetYukle / SetSil
' ============================================================

Private Const ONERI_ONEK As String = "ONERI:"


' ============================================================
'  OTOMATIK ESLESTIRME  (aktarim yapmaz, yalnizca onerir)
' ============================================================
Public Sub OtomatikEslestir()

    Dim pnl As Worksheet, lst As Worksheet
    Dim kaynakWs As Worksheet, hedefWs As Worksheet
    Dim k As Long, c As Long, son As Long, hedefSon As Long
    Dim i As Long, h As Long, t As Long
    Dim kaynakAd As String, kNorm As String, adim As String
    Dim hedefDeger() As String, hedefNorm() As String, hedefSira() As Long
    Dim kSira() As Long
    Dim kullanilan As Sozluk, tipOnb As Sozluk, normSay As Sozluk
    Dim birebir As Long, oneriSay As Long, bosKalan As Long, elenen As Long
    Dim skor As Double
    Dim adaylar(1 To 3) As Long, adaySkor(1 To 3) As Double
    Dim secilen As Long, secilenSkor As Double, kNo As Long
    Dim kTip As String, hTip As String, adayMetni As String
    Dim hv As Variant

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    k = AktifBlok()
    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)

    son = BlokSonSatir(pnl, k)
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    If son < ILK_SATIR Or Trim$(GuvenliMetin(lst.Cells(1, LST_HEDEF_SUTUNLAR).Value)) = "" Then
        MsgBox "Once 'YENILE' calistir; PANEL " & k & " icin kaynak ve hedef" & vbNewLine & _
               "sutunlari yuklensin.", vbExclamation
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Hata
    adim = "dosyalar bulunuyor"

    Set kaynakWs = BlokKaynakSayfasi(pnl, k)
    Set hedefWs = BlokHedefSayfasi(pnl, k)

    ' --- Hedef listesi: deger, sade ad, ayni adlilar arasinda sira ---
    hv = IkiBoyutlu(lst.Range(lst.Cells(1, LST_HEDEF_SUTUNLAR), lst.Cells(hedefSon, LST_HEDEF_SUTUNLAR)))
    ReDim hedefDeger(1 To hedefSon)
    ReDim hedefNorm(1 To hedefSon)
    ReDim hedefSira(1 To hedefSon)
    Set normSay = New Sozluk
    For h = 1 To hedefSon
        hedefDeger(h) = GuvenliMetin(hv(h, 1))
        hedefNorm(h) = Normalize(SutunAdi(hedefDeger(h)))
        If hedefNorm(h) <> "" Then hedefSira(h) = normSay.Artir("H" & hedefNorm(h))
    Next h

    ' --- Kaynak sirasi (ayni adli kaynak sutunlar) ---
    ReDim kSira(ILK_SATIR To son)
    For i = ILK_SATIR To son
        kNorm = Normalize(SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value))
        If kNorm <> "" Then kSira(i) = normSay.Artir("K" & kNorm)
    Next i

    Set kullanilan = New Sozluk
    Set tipOnb = New Sozluk
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) <> "" Then
            kullanilan.Item(BH(Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)))) = True
        End If
    Next i

    ' --- 1. tur: adi ayni olanlar ---
    ' Buyuk/kucuk harf, Turkce karakter, bosluk ve noktalama farki yok
    ' sayilir. Ayni adli birden fazla sutun varsa once AYNI SIRADAKI
    ' (2. "Tutar" -> 2. "Tutar") aranir. Veri tipi de uyuyorsa dogrudan
    ' eslesir; tip farkliysa ONERI olarak isaretlenir.
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" And _
           Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) = "" Then

            kaynakAd = SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value)
            adim = "1. tur, aylik sutun " & kaynakAd
            kNorm = Normalize(kaynakAd)
            If kNorm <> "" Then
                secilen = 0
                For h = 1 To hedefSon
                    If hedefNorm(h) = kNorm And hedefSira(h) = kSira(i) Then
                        If Not kullanilan.Exists(BH(Trim$(hedefDeger(h)))) Then secilen = h
                        Exit For
                    End If
                Next h
                If secilen = 0 Then
                    For h = 1 To hedefSon
                        If hedefNorm(h) = kNorm Then
                            If Not kullanilan.Exists(BH(Trim$(hedefDeger(h)))) Then
                                secilen = h
                                Exit For
                            End If
                        End If
                    Next h
                End If

                If secilen > 0 Then
                    kNo = SutunNo(pnl.Cells(i, c + OFS_KAYNAK).Value)
                    kTip = OnbellekTip(tipOnb, kaynakWs, kNo, True)
                    hTip = OnbellekTip(tipOnb, hedefWs, SutunNo(hedefDeger(secilen)), False)
                    pnl.Cells(i, c + OFS_HEDEF).Value = HucreMetni(hedefDeger(secilen))
                    kullanilan.Item(BH(Trim$(hedefDeger(secilen)))) = True
                    If TipCatisir(kTip, hTip) Then
                        OneriIsaretle pnl.Cells(i, c + OFS_HEDEF), _
                            "Ad ayni ama veri tipi farkli." & vbLf & _
                            "Aylik: " & kTip & ", ana dosya: " & hTip & "."
                        oneriSay = oneriSay + 1
                    Else
                        birebir = birebir + 1
                    End If
                End If
            End If
        End If
    Next i

    ' --- 2. tur: adi benzeyenler (her zaman ONERI) ---
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" And _
           Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) = "" Then

            kaynakAd = SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value)
            adim = "2. tur, aylik sutun " & kaynakAd
            For t = 1 To 3
                adaylar(t) = 0
                adaySkor(t) = 0
            Next t

            kNo = SutunNo(pnl.Cells(i, c + OFS_KAYNAK).Value)
            kTip = OnbellekTip(tipOnb, kaynakWs, kNo, True)

            ' Veri tipi kesin celisen adaylar hic listeye alinmaz
            For h = 1 To hedefSon
                If hedefNorm(h) <> "" Then
                    If Not kullanilan.Exists(BH(Trim$(hedefDeger(h)))) Then
                        skor = Benzerlik(kaynakAd, SutunAdi(hedefDeger(h)))
                        If skor >= BENZERLIK_ESIGI Then
                            hTip = OnbellekTip(tipOnb, hedefWs, SutunNo(hedefDeger(h)), False)
                            If TipCatisir(kTip, hTip) Then
                                elenen = elenen + 1
                            Else
                                AdayEkle adaylar, adaySkor, h, skor
                            End If
                        End If
                    End If
                End If
            Next h

            secilen = adaylar(1)
            secilenSkor = adaySkor(1)

            If secilen > 0 Then
                adayMetni = ""
                For t = 1 To 3
                    If adaylar(t) > 0 Then
                        adayMetni = adayMetni & "  " & t & ") " & SutunAdi(hedefDeger(adaylar(t))) & _
                                    "  %" & Format$(adaySkor(t) * 100, "0") & "  (" & _
                                    BosIse(OnbellekTip(tipOnb, hedefWs, SutunNo(hedefDeger(adaylar(t))), False), "?") & _
                                    ")" & vbLf
                    End If
                Next t

                pnl.Cells(i, c + OFS_HEDEF).Value = HucreMetni(hedefDeger(secilen))
                kullanilan.Item(BH(Trim$(hedefDeger(secilen)))) = True
                OneriIsaretle pnl.Cells(i, c + OFS_HEDEF), _
                    "Otomatik eslestirme, ad benzerligi %" & Format$(secilenSkor * 100, "0") & vbLf & _
                    "Aylik: " & kaynakAd & "  (" & BosIse(kTip, "?") & ")" & vbLf & _
                    "En iyi adaylar:" & vbLf & adayMetni
                oneriSay = oneriSay + 1
            End If
        End If
    Next i

    adim = "menuler ve renkler tazeleniyor"
    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" And _
           Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) = "" Then bosKalan = bosKalan + 1
    Next i

    BlokMenuleriTazele pnl, k
    EkranAc

    OlayYaz "ESLESTIRME", k, 0, 0, "Otomatik eslestirme: " & birebir & " ayni ad, " & _
            oneriSay & " oneri, " & bosKalan & " bos kaldi, " & elenen & " aday tip nedeniyle elendi", ""

    MsgBox "PANEL " & k & " - OTOMATIK ESLESTIRME" & vbNewLine & vbNewLine & _
           "Adi ayni, eslestirildi          : " & birebir & vbNewLine & _
           "ONERI (turuncu, onay bekliyor)  : " & oneriSay & vbNewLine & _
           "Veri tipi uymadigi icin elenen  : " & elenen & vbNewLine & _
           "Bos kalan                       : " & bosKalan & vbNewLine & vbNewLine & _
           "Turuncu hucrenin notunda en iyi 3 aday ve benzerlik" & vbNewLine & _
           "oranlari yazar. Onay: satiri sec > 'Oneri Onayla'" & vbNewLine & _
           "(hicbir satir secili degilse paneldeki tumu onaylanir)." & vbNewLine & vbNewLine & _
           "Hicbir sey aktarilmadi.", vbInformation, "Otomatik Eslestirme"
    Exit Sub

Hata:
    EkranAc
    MsgBox "OTOMATIK ESLESTIRME SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Adim : " & adim & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Hata"

End Sub


Private Function BosIse(ByVal s As String, ByVal yerine As String) As String
    If s = "" Then
        BosIse = yerine
    Else
        BosIse = s
    End If
End Function


' Ilk 3 adayi skora gore sirali tutar
Private Sub AdayEkle(adaylar() As Long, adaySkor() As Double, ByVal h As Long, ByVal skor As Double)

    Dim t As Long, u As Long

    For t = 1 To 3
        If skor > adaySkor(t) Then
            For u = 3 To t + 1 Step -1
                adaylar(u) = adaylar(u - 1)
                adaySkor(u) = adaySkor(u - 1)
            Next u
            adaylar(t) = h
            adaySkor(t) = skor
            Exit Sub
        End If
    Next t

End Sub


' ============================================================
'  ONERI ONAYI
' ============================================================

' AKILLI ONAY:
'   - Panel tablosunda onerili (turuncu) satir(lar) seciliyse yalnizca
'     onlar onaylanir (tek oneride soru sorulmaz).
'   - Secimde oneri yoksa bulundugun paneldeki TUM oneriler listelenir
'     ve onay istenir.
Public Sub OneriOnayla()

    Dim pnl As Worksheet
    Dim k As Long, c As Long, i As Long, r As Long
    Dim alan As Range, parca As Range
    Dim satirlar As Sozluk
    Dim anahtar As Variant
    Dim adet As Long
    Dim metin As String, tamMetin As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    k = AktifBlok()
    c = BlokSutun(k)

    If ActiveSheet Is pnl And TypeName(Selection) = "Range" Then
        Set alan = Intersect(Selection, pnl.Range(pnl.Cells(ILK_SATIR, c), _
                                                  pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1)))
    End If

    If Not alan Is Nothing Then
        Set satirlar = New Sozluk
        For Each parca In alan.Areas
            For i = parca.Row To parca.Row + parca.Rows.Count - 1
                If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then satirlar.Item(i) = True
            Next i
        Next parca

        If satirlar.Count > 0 Then
            For Each anahtar In satirlar.Keys
                r = CLng(anahtar)
                adet = adet + 1
                If adet <= 8 Then
                    metin = metin & "   " & SutunAdi(pnl.Cells(r, c + OFS_KAYNAK).Value) & "  ->  " & _
                            SutunAdi(pnl.Cells(r, c + OFS_HEDEF).Value) & vbNewLine
                End If
                tamMetin = tamMetin & SutunAdi(pnl.Cells(r, c + OFS_KAYNAK).Value) & " -> " & _
                           SutunAdi(pnl.Cells(r, c + OFS_HEDEF).Value) & "; "
            Next anahtar

            If Not SatirOnerileriniKaldir(pnl, k, satirlar) Then Exit Sub

            OlayYaz "ONERI ONAY", k, 0, 0, adet & " oneri secerek onaylandi: " & tamMetin, "Onaylandi"

            If adet = 1 Then
                DurumCubugu "PANEL " & k & ": oneri onaylandi  (" & Left$(tamMetin, 80) & ")"
                On Error Resume Next
                Application.OnTime Now + TimeValue("00:00:06"), "DurumCubuguTemizle"
                On Error GoTo 0
            Else
                MsgBox "PANEL " & k & ": " & adet & " oneri onaylandi." & vbNewLine & vbNewLine & _
                       metin & IIf(adet > 8, "   ... ve " & (adet - 8) & " tane daha" & vbNewLine, ""), _
                       vbInformation, "Oneri onayi"
            End If
            Exit Sub
        End If
    End If

    TumOnerileriOnayla

End Sub


' Verilen satirlarin onerilerini kaldirir ve paneli tazeler
Private Function SatirOnerileriniKaldir(ByVal pnl As Worksheet, ByVal k As Long, _
                                        ByVal satirlar As Sozluk) As Boolean

    Dim anahtar As Variant
    Dim c As Long

    c = BlokSutun(k)
    EkranKilitle
    On Error GoTo Hata
    For Each anahtar In satirlar.Keys
        OneriKaldir pnl.Cells(CLng(anahtar), c + OFS_HEDEF)
    Next anahtar
    BlokMenuleriTazele pnl, k
    EkranAc
    SatirOnerileriniKaldir = True
    Exit Function

Hata:
    EkranAc
    MsgBox "Oneri onaylanirken hata: " & Err.Number & " - " & Err.Description, vbExclamation

End Function


' Bulundugun paneldeki tum onerileri tek seferde onaylar
Public Sub TumOnerileriOnayla()

    Dim pnl As Worksheet
    Dim k As Long, c As Long, i As Long, son As Long, adet As Long
    Dim metin As String, tamMetin As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    k = AktifBlok()
    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)

    For i = ILK_SATIR To son
        If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then
            adet = adet + 1
            If adet <= 12 Then
                metin = metin & "   " & SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value) & "  ->  " & _
                        SutunAdi(pnl.Cells(i, c + OFS_HEDEF).Value) & vbNewLine
            End If
            tamMetin = tamMetin & SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value) & " -> " & _
                       SutunAdi(pnl.Cells(i, c + OFS_HEDEF).Value) & "; "
        End If
    Next i

    If adet = 0 Then
        MsgBox "PANEL " & k & ": onay bekleyen oneri yok.", vbInformation
        Exit Sub
    End If
    If adet > 12 Then metin = metin & "   ... ve " & (adet - 12) & " tane daha" & vbNewLine

    If MsgBox("PANEL " & k & " - TUM ONERILERI ONAYLA" & vbNewLine & vbNewLine & _
              adet & " oneri onaylanacak:" & vbNewLine & vbNewLine & _
              Sigdir(metin, 600) & vbNewLine & _
              "Her birinin nedeni turuncu hucrenin notunda yazar." & vbNewLine & _
              "(Yalnizca birkacini onaylamak icin once o satirlari sec.)" & vbNewLine & vbNewLine & _
              "Hepsi onaylansin mi?", vbYesNo + vbQuestion + vbDefaultButton2, "Oneri onayi") <> vbYes Then
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Cikis
    For i = ILK_SATIR To son
        OneriKaldir pnl.Cells(i, c + OFS_HEDEF)
    Next i
    BlokMenuleriTazele pnl, k
Cikis:
    If Err.Number <> 0 Then metin = "HATA"
    EkranAc
    If metin = "HATA" Then
        MsgBox "Oneriler onaylanirken hata olustu; paneli kontrol et.", vbExclamation
        Exit Sub
    End If

    OlayYaz "ONERI ONAY", k, 0, 0, "Toplu onay, " & adet & " oneri: " & tamMetin, "Onaylandi (toplu)"
    MsgBox "PANEL " & k & ": " & adet & " oneri onaylandi.", vbInformation, "Oneri onayi"

End Sub


Public Sub DurumCubuguTemizle()
    Application.StatusBar = False
End Sub


' Hucrede onay bekleyen bir oneri var mi?
Public Function OneriMi(ByVal h As Range) As Boolean

    Dim t As String

    OneriMi = False
    On Error Resume Next
    If Not h.Comment Is Nothing Then
        t = h.Comment.Text
        OneriMi = (Left$(t, Len(ONERI_ONEK)) = ONERI_ONEK)
    End If
    On Error GoTo 0

End Function


' Onerinin aciklama kismi (Yenile sirasinda geri yazmak icin)
Public Function OneriMetni(ByVal h As Range) As String

    Dim t As String
    Dim p As Long

    OneriMetni = ""
    If Not OneriMi(h) Then Exit Function

    On Error Resume Next
    t = h.Comment.Text
    On Error GoTo 0

    t = Mid$(t, Len(ONERI_ONEK) + 2)
    p = InStr(t, vbLf & "---")
    If p > 0 Then t = Left$(t, p - 1)
    OneriMetni = t

End Function


Public Sub OneriIsaretle(ByVal h As Range, ByVal aciklama As String)

    On Error Resume Next
    h.ClearComments
    h.AddComment ONERI_ONEK & " " & aciklama & vbLf & "---" & vbLf & _
                 "Onay: satiri sec > 'Oneri Onayla'" & vbLf & _
                 "veya HEDEF hucresini menuden yeniden sec."
    h.Comment.Visible = False
    h.Comment.Shape.TextFrame.AutoSize = True
    On Error GoTo 0

    h.Interior.Color = Palet(prOneri)

End Sub


Public Sub OneriKaldir(ByVal h As Range)
    If OneriMi(h) Then
        On Error Resume Next
        h.ClearComments
        On Error GoTo 0
    End If
End Sub


' ============================================================
'  PROFILLER
'
'  Profil = bir aylik dosyanin sutun semasi. Kaydedilirken:
'   - aylik dosyanin TUM sutunlari yazilir (eslestirilmeyenler de;
'     boylece "bilerek bos birakilan" sutun her ay sorulmaz ve
'     "dosyada yeni sutun" dogru tespit edilir)
'   - her sutunun sirasi (ayni adli sutunlar), yeri (harf), veri tipi
'   - ayni hedefe giden sutunun adi degismisse eski ad DIGER ADLAR'a
'     eklenir (ogrenme)
'   - bu ay dosyada olmayan eski eslesmeler KORUNUR
' ============================================================
Public Sub ProfilKaydet()

    Dim pnl As Worksheet, prf As Worksheet, lst As Worksheet
    Dim kaynakWs As Worksheet
    Dim k As Long, c As Long, son As Long, i As Long, e As Long, t As Long, h As Long
    Dim ad As String
    Dim oneriSay As Long, prfSon As Long, hedefSon As Long, eskiSay As Long, yaz As Long
    Dim v As Variant
    Dim eSatir() As Long, eAdlar() As String, eAnahtar() As String, eKullanildi() As Boolean
    Dim w() As Variant
    Dim kaynakDeger As String, kaynakAdi As String, kNorm As String, harf As String
    Dim hedefDeger As String, hedefAdi As String, hedefSira As Long, anahtarMetni As String
    Dim normSay As Sozluk, hedefSiralari As Sozluk, hedefNormSay As Sozluk
    Dim digerler As Sozluk, kullanilanAnahtar As Sozluk
    Dim anahtar As Variant
    Dim digerMetni As String, nm As String
    Dim eslesen As Long, bilerekBos As Long, korunan As Long, ogrenilen As Long
    Dim ornekOgrenme As String, kayitMetni As String, veriTipiMetni As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    k = AktifBlok()
    c = BlokSutun(k)
    Set lst = ListeSayfasi(k)
    son = BlokSonSatir(pnl, k)

    If son < ILK_SATIR Then
        MsgBox "PANEL " & k & " bos. Once kaynak dosyayi secip 'YENILE' calistir.", vbExclamation
        Exit Sub
    End If

    ad = Trim$(InputBox("PANEL " & k & " icin eslesme hangi isimle kaydedilsin?" & vbNewLine & _
                        "Ornek: A Sirketi" & vbNewLine & vbNewLine & _
                        "Ayni isimde profil varsa GUNCELLENIR (eski bilgiler korunarak).", _
                        "Profil Kaydet", AyarMetni(pnl, k, SATIR_PROFIL)))
    If ad = "" Then Exit Sub

    For i = ILK_SATIR To son
        If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then oneriSay = oneriSay + 1
    Next i
    If oneriSay > 0 Then
        If MsgBox("PANEL " & k & ": " & oneriSay & " oneri henuz onaylanmadi (turuncu)." & vbNewLine & vbNewLine & _
                  "Profili kaydedersen bu oneriler ONAYLANMIS sayilir ve" & vbNewLine & _
                  "profile dogru eslesme olarak yazilir." & vbNewLine & vbNewLine & _
                  "Devam edilsin mi?", vbYesNo + vbQuestion + vbDefaultButton2, "Profil Kaydet") <> vbYes Then
            Exit Sub
        End If
    End If

    EkranKilitle
    On Error GoTo Hata

    If oneriSay > 0 Then
        For i = ILK_SATIR To son
            OneriKaldir pnl.Cells(i, c + OFS_HEDEF)
        Next i
    End If

    ' --- Eski profil kayitlari ---
    ProfilSayfasiHazirla
    Set prf = ThisWorkbook.Worksheets(SAYFA_PROFIL)
    prfSon = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row

    eskiSay = 0
    If prfSon >= 2 Then
        v = IkiBoyutlu(prf.Range(prf.Cells(2, 1), prf.Cells(prfSon, P_SON)))
        ReDim eSatir(1 To UBound(v, 1))
        ReDim eAdlar(1 To UBound(v, 1))
        ReDim eAnahtar(1 To UBound(v, 1))
        ReDim eKullanildi(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            If BH(Trim$(GuvenliMetin(v(i, P_AD)))) = BH(ad) Then
                eskiSay = eskiSay + 1
                eSatir(eskiSay) = i
                eAdlar(eskiSay) = ProfilAdlari(GuvenliMetin(v(i, P_KAYNAK)), GuvenliMetin(v(i, P_DIGER)))
                If Trim$(GuvenliMetin(v(i, P_HEDEF))) <> "" Then
                    eAnahtar(eskiSay) = Normalize(GuvenliMetin(v(i, P_HEDEF))) & "#" & SiraDegeri(v(i, P_HEDEF_SIRA))
                End If
            End If
        Next i
    End If

    ' --- Ana dosya sutunlarinin sirasi (ayni adli hedefler icin) ---
    Set hedefSiralari = New Sozluk
    Set hedefNormSay = New Sozluk
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    For h = 1 To hedefSon
        hedefDeger = GuvenliMetin(lst.Cells(h, LST_HEDEF_SUTUNLAR).Value)
        If hedefDeger <> "" Then
            nm = Normalize(SutunAdi(hedefDeger))
            hedefSiralari.Item(BH(Trim$(hedefDeger))) = hedefNormSay.Artir(nm)
        End If
    Next h

    Set kaynakWs = BlokKaynakSayfasi(pnl, k)

    Set normSay = New Sozluk
    Set kullanilanAnahtar = New Sozluk
    ReDim w(1 To (son - ILK_SATIR + 1) + eskiSay + 1, 1 To P_SON)
    yaz = 0
    kayitMetni = Format$(Now, "dd.mm.yyyy hh:nn") & " " & KullaniciAdi()

    For i = ILK_SATIR To son
        kaynakDeger = GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)
        If Trim$(kaynakDeger) <> "" Then

            kaynakAdi = SutunAdi(kaynakDeger)
            kNorm = Normalize(kaynakAdi)
            normSay.Artir kNorm
            harf = SutunHarfi(SutunNo(kaynakDeger))

            hedefDeger = GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)
            hedefAdi = SutunAdi(hedefDeger)
            hedefSira = 1
            If hedefSiralari.Exists(BH(Trim$(hedefDeger))) Then hedefSira = hedefSiralari.Item(BH(Trim$(hedefDeger)))
            anahtarMetni = ""
            If hedefAdi <> "" Then anahtarMetni = Normalize(hedefAdi) & "#" & hedefSira
            If anahtarMetni <> "" Then kullanilanAnahtar.Item(anahtarMetni) = True

            ' --- Diger adlar: ayni hedefe giden eski kayit + ayni adli eski kayit ---
            Set digerler = New Sozluk
            For e = 1 To eskiSay
                If Not eKullanildi(e) Then
                    If anahtarMetni <> "" And eAnahtar(e) = anahtarMetni Then
                        eKullanildi(e) = True
                        AdlariEkle digerler, GuvenliMetin(v(eSatir(e), P_KAYNAK)), GuvenliMetin(v(eSatir(e), P_DIGER))
                        If Normalize(GuvenliMetin(v(eSatir(e), P_KAYNAK))) <> kNorm Then
                            ogrenilen = ogrenilen + 1
                            If ornekOgrenme = "" Then ornekOgrenme = kaynakAdi & " = " & GuvenliMetin(v(eSatir(e), P_KAYNAK))
                        End If
                    ElseIf InStr("|" & eAdlar(e) & "|", "|" & kNorm & "|") > 0 Then
                        eKullanildi(e) = True
                        AdlariEkle digerler, GuvenliMetin(v(eSatir(e), P_KAYNAK)), GuvenliMetin(v(eSatir(e), P_DIGER))
                    End If
                End If
            Next e

            digerMetni = ""
            For Each anahtar In digerler.Keys
                If anahtar <> kNorm Then digerMetni = digerMetni & digerler.Item(anahtar) & " | "
            Next anahtar
            If Len(digerMetni) > 3 Then digerMetni = Left$(digerMetni, Len(digerMetni) - 3)

            veriTipiMetni = ""
            If Not kaynakWs Is Nothing Then veriTipiMetni = VeriTipiKaynak(kaynakWs, SutunNo(kaynakDeger))

            yaz = yaz + 1
            w(yaz, P_AD) = HucreMetni(ad)
            w(yaz, P_KAYNAK) = HucreMetni(kaynakAdi)
            w(yaz, P_HEDEF) = HucreMetni(hedefAdi)
            w(yaz, P_FILTRE) = HucreMetni(GuvenliMetin(pnl.Cells(i, c + OFS_FILTRE).Value))
            w(yaz, P_FORMUL) = IIf(IslemFormul(pnl.Cells(i, c + OFS_ISLEM).Value), "EVET", "")
            w(yaz, P_DIGER) = HucreMetni(digerMetni)
            w(yaz, P_SIRA) = normSay.Item(kNorm)
            w(yaz, P_YER) = harf
            w(yaz, P_TIP) = veriTipiMetni
            If hedefAdi <> "" Then
                w(yaz, P_HEDEF_SIRA) = hedefSira
                eslesen = eslesen + 1
            Else
                bilerekBos = bilerekBos + 1
            End If
            w(yaz, P_KAYIT) = kayitMetni
        End If
    Next i

    ' --- Bu ay dosyada olmayan eski eslesmeler korunur ---
    For e = 1 To eskiSay
        If Not eKullanildi(e) And eAnahtar(e) <> "" Then
            If Not kullanilanAnahtar.Exists(eAnahtar(e)) Then
                yaz = yaz + 1
                For t = 1 To P_SON
                    w(yaz, t) = v(eSatir(e), t)
                    If VarType(w(yaz, t)) = vbString Then w(yaz, t) = HucreMetni(CStr(w(yaz, t)))
                Next t
                w(yaz, P_AD) = HucreMetni(ad)
                korunan = korunan + 1
            End If
        End If
    Next e

    ProfilSatirlariSil prf, ad
    prfSon = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row
    If yaz > 0 Then prf.Cells(prfSon + 1, 1).Resize(yaz, P_SON).Value = w

    Ayar(pnl, k, SATIR_PROFIL).Value = HucreMetni(ad)
    ProfilListesiYaz lst
    MenuKur Ayar(pnl, k, SATIR_PROFIL), lst, LST_PROFILLER
    BlokMenuleriTazele pnl, k

    EkranAc

    OlayYaz "PROFIL KAYDET", k, 0, 0, "'" & ad & "' kaydedildi: " & eslesen & " eslesme, " & _
            bilerekBos & " eslestirilmemis sutun, " & korunan & " eski eslesme korundu, " & _
            ogrenilen & " yeni diger ad ogrenildi" & IIf(ornekOgrenme <> "", " (or. " & ornekOgrenme & ")", ""), _
            IIf(oneriSay > 0, oneriSay & " oneri kayitla birlikte onaylandi", "")

    MsgBox "'" & ad & "' profili kaydedildi (PANEL " & k & ")." & vbNewLine & vbNewLine & _
           "Eslesen sutun                     : " & eslesen & vbNewLine & _
           "Bilerek eslestirilmemis sutun     : " & bilerekBos & vbNewLine & _
           "Bu ay dosyada olmayan, korunan    : " & korunan & vbNewLine & _
           "Ogrenilen yeni ad (diger ad)      : " & ogrenilen & _
           IIf(ornekOgrenme <> "", vbNewLine & "   ornek: " & ornekOgrenme, "") & vbNewLine & vbNewLine & _
           "Gelecek ay sutun adi bu adlardan biriyle gelirse" & vbNewLine & _
           "dogrudan eslesir.", vbInformation, "Profil Kaydet"
    Exit Sub

Hata:
    EkranAc
    MsgBox "PROFIL KAYDEDILEMEDI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Profil Kaydet"

End Sub


' Birincil ad + diger adlari sozluge ekler (anahtar: sade ad)
Private Sub AdlariEkle(ByVal sozluk As Sozluk, ByVal birincil As String, ByVal digerleri As String)

    Dim parca As Variant

    If Trim$(birincil) <> "" Then
        If Not sozluk.Exists(Normalize(birincil)) Then sozluk.Item(Normalize(birincil)) = Trim$(birincil)
    End If
    For Each parca In Split(digerleri, "|")
        If Trim$(parca) <> "" Then
            If Not sozluk.Exists(Normalize(CStr(parca))) Then sozluk.Item(Normalize(CStr(parca))) = Trim$(parca)
        End If
    Next parca

End Sub


' Birincil ad + diger adlar -> "sade1|sade2|..." (arama icin)
Private Function ProfilAdlari(ByVal birincil As String, ByVal digerleri As String) As String

    Dim parca As Variant
    Dim metin As String

    metin = Normalize(birincil)
    For Each parca In Split(digerleri, "|")
        If Trim$(parca) <> "" Then metin = metin & "|" & Normalize(CStr(parca))
    Next parca
    ProfilAdlari = metin

End Function


' Bos / hatali sira degeri 1 sayilir
Private Function SiraDegeri(ByVal x As Variant) As Long
    SiraDegeri = 1
    If IsError(x) Then Exit Function
    If IsNumeric(x) And Trim$(GuvenliMetin(x)) <> "" Then
        If CLng(x) >= 1 Then SiraDegeri = CLng(x)
    End If
End Function


Public Sub ProfilSil()

    Dim pnl As Worksheet, prf As Worksheet
    Dim k As Long
    Dim ad As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    ProfilSayfasiHazirla
    Set prf = ThisWorkbook.Worksheets(SAYFA_PROFIL)
    k = AktifBlok()

    ad = AyarMetni(pnl, k, SATIR_PROFIL)
    If ad = "" Then
        MsgBox "Once bu panelde silinecek profili sec (Profil hucresi).", vbExclamation
        Exit Sub
    End If

    If MsgBox("'" & ad & "' profili silinsin mi?" & vbNewLine & vbNewLine & _
              "Bu profili kullanan tum paneller ve setler etkilenir.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Profil Sil") <> vbYes Then Exit Sub

    EkranKilitle
    On Error GoTo Cikis
    ProfilSatirlariSil prf, ad
    Ayar(pnl, k, SATIR_PROFIL).ClearContents
    ProfilListesiYaz ListeSayfasi(k)
    MenuKur Ayar(pnl, k, SATIR_PROFIL), ListeSayfasi(k), LST_PROFILLER
Cikis:
    EkranAc

    OlayYaz "PROFIL SIL", k, 0, 0, "'" & ad & "' profili silindi", ""
    MsgBox "'" & ad & "' profili silindi.", vbInformation

End Sub


' Profili bir panele uygular. Donus: profil bulundu mu.
'   1. tur: ad (veya ogrenilmis diger ad) + sira birebir -> dogrudan eslesir
'   2. tur: bulunamayanlar icin aday: ad benzerligi, ayni yer (harf),
'           ayni veri tipi. Bulunan ONERI olur. Tip kesin farkliysa elenir.
Public Function ProfilUygula(ByVal pnl As Worksheet, ByVal k As Long, ByVal kaynakWs As Worksheet, _
                             ByVal hedefWs As Worksheet, ByVal ad As String, _
                             ByRef tamam As Long, ByRef oneriSay As Long, _
                             ByRef eksik As String, ByRef yeni As String, _
                             ByRef hedefYok As String) As Boolean

    Dim prf As Worksheet, lst As Worksheet
    Dim c As Long, son As Long, prfSon As Long, hedefSon As Long
    Dim i As Long, e As Long, h As Long, t As Long, r As Long, p As Long
    Dim v As Variant
    Dim gSay As Long, satirSay As Long
    Dim gSatir() As Long, gAdlar() As String, gHedefDeger() As String
    Dim gEslesti() As Boolean, gPanel() As Long
    Dim pAd() As String, pNorm() As String, pHarf() As String, pSira() As Long
    Dim pAlindi() As Boolean
    Dim normSay As Sozluk, profilAdSozluk As Sozluk, kullanilan As Sozluk, birinciTur As Sozluk
    Dim hDeger() As String, hNorm() As String, hSira() As Long, hNormSay As Sozluk
    Dim adlar() As String
    Dim hucre As Range
    Dim tipOnb As Sozluk
    Dim adSkor As Double, skor As Double, enIyiSkor As Double
    Dim yerAyni As Boolean, tipAyni As Boolean
    Dim kTip As String, gTip As String, gYer As String
    Dim ciftE() As Long, ciftR() As Long, ciftSkor() As Double, ciftAd() As Double
    Dim ciftYer() As Boolean, ciftTip() As String, ciftSay As Long
    Dim enIyi As Long
    Dim eAtandi() As Boolean

    ProfilUygula = False
    tamam = 0
    oneriSay = 0
    eksik = ""
    yeni = ""
    hedefYok = ""

    ProfilSayfasiHazirla
    Set prf = ThisWorkbook.Worksheets(SAYFA_PROFIL)
    Set lst = ListeSayfasi(k)
    c = BlokSutun(k)

    prfSon = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row
    son = BlokSonSatir(pnl, k)
    If prfSon < 2 Or son < ILK_SATIR Then Exit Function

    ' --- Profilin kayitlari ---
    v = IkiBoyutlu(prf.Range(prf.Cells(2, 1), prf.Cells(prfSon, P_SON)))
    ReDim gSatir(1 To UBound(v, 1))
    ReDim gAdlar(1 To UBound(v, 1))
    ReDim gHedefDeger(1 To UBound(v, 1))
    ReDim gEslesti(1 To UBound(v, 1))
    ReDim gPanel(1 To UBound(v, 1))
    Set profilAdSozluk = New Sozluk

    gSay = 0
    For i = 1 To UBound(v, 1)
        If BH(Trim$(GuvenliMetin(v(i, P_AD)))) = BH(ad) Then
            gSay = gSay + 1
            gSatir(gSay) = i
            gAdlar(gSay) = ProfilAdlari(GuvenliMetin(v(i, P_KAYNAK)), GuvenliMetin(v(i, P_DIGER)))
            adlar = Split(gAdlar(gSay), "|")
            For t = 0 To UBound(adlar)
                If adlar(t) <> "" Then profilAdSozluk.Item(adlar(t)) = True
            Next t
        End If
    Next i

    If gSay = 0 Then Exit Function
    ProfilUygula = True

    ' --- Paneldeki kaynak sutunlari ---
    satirSay = son - ILK_SATIR + 1
    ReDim pAd(1 To satirSay)
    ReDim pNorm(1 To satirSay)
    ReDim pHarf(1 To satirSay)
    ReDim pSira(1 To satirSay)
    ReDim pAlindi(1 To satirSay)
    Set normSay = New Sozluk
    For r = 1 To satirSay
        pAd(r) = SutunAdi(pnl.Cells(ILK_SATIR + r - 1, c + OFS_KAYNAK).Value)
        pNorm(r) = Normalize(pAd(r))
        pHarf(r) = SutunHarfi(SutunNo(pnl.Cells(ILK_SATIR + r - 1, c + OFS_KAYNAK).Value))
        If pNorm(r) <> "" Then pSira(r) = normSay.Artir(pNorm(r))
    Next r

    ' --- Ana dosya sutunlari (liste) ---
    hedefSon = lst.Cells(lst.Rows.Count, LST_HEDEF_SUTUNLAR).End(xlUp).Row
    ReDim hDeger(1 To hedefSon)
    ReDim hNorm(1 To hedefSon)
    ReDim hSira(1 To hedefSon)
    Set hNormSay = New Sozluk
    For h = 1 To hedefSon
        hDeger(h) = GuvenliMetin(lst.Cells(h, LST_HEDEF_SUTUNLAR).Value)
        hNorm(h) = Normalize(SutunAdi(hDeger(h)))
        hSira(h) = hNormSay.Artir(hNorm(h))
    Next h

    For e = 1 To gSay
        gHedefDeger(e) = ""
        If Trim$(GuvenliMetin(v(gSatir(e), P_HEDEF))) <> "" Then
            gHedefDeger(e) = HedefListedeBul(hDeger, hNorm, hSira, _
                                             Normalize(GuvenliMetin(v(gSatir(e), P_HEDEF))), _
                                             SiraDegeri(v(gSatir(e), P_HEDEF_SIRA)))
        End If
    Next e

    ' ============ 1. TUR: birebir ad + sira ============
    For e = 1 To gSay
        adlar = Split(gAdlar(e), "|")
        For t = 0 To UBound(adlar)
            If adlar(t) <> "" Then
                For r = 1 To satirSay
                    If Not pAlindi(r) Then
                        If pNorm(r) = adlar(t) And pSira(r) = SiraDegeri(v(gSatir(e), P_SIRA)) Then
                            gEslesti(e) = True
                            gPanel(e) = r
                            pAlindi(r) = True
                            Exit For
                        End If
                    End If
                Next r
                If gEslesti(e) Then Exit For
            End If
        Next t
    Next e

    ' Profilin birinci turda yazacagi hedefler; baska satirda (eski
    ' oturumdan kalma) ayni hedef varsa profil kazanir.
    Set birinciTur = New Sozluk
    For e = 1 To gSay
        If gEslesti(e) And gHedefDeger(e) <> "" Then birinciTur.Item(BH(Trim$(gHedefDeger(e)))) = e
    Next e

    For r = 1 To satirSay
        If Not pAlindi(r) Then
            Set hucre = pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF)
            If birinciTur.Exists(BH(Trim$(GuvenliMetin(hucre.Value)))) Then
                hucre.ClearContents
                OneriKaldir hucre
            End If
        End If
    Next r

    For e = 1 To gSay
        If gEslesti(e) Then
            If Trim$(GuvenliMetin(v(gSatir(e), P_HEDEF))) <> "" Then
                If gHedefDeger(e) = "" Then
                    hedefYok = hedefYok & GuvenliMetin(v(gSatir(e), P_HEDEF)) & ", "
                ElseIf birinciTur.Item(BH(Trim$(gHedefDeger(e)))) = e Then
                    Set hucre = pnl.Cells(ILK_SATIR + gPanel(e) - 1, c + OFS_HEDEF)
                    hucre.Value = HucreMetni(gHedefDeger(e))
                    OneriKaldir hucre
                    ProfilEkleriYaz pnl, c, ILK_SATIR + gPanel(e) - 1, v, gSatir(e)
                    tamam = tamam + 1
                End If
            End If
        End If
    Next e

    ' Su an kullanilan tum hedefler
    Set kullanilan = New Sozluk
    For r = 1 To satirSay
        If Trim$(GuvenliMetin(pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF).Value)) <> "" Then
            kullanilan.Item(BH(Trim$(GuvenliMetin(pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF).Value)))) = True
        End If
    Next r

    ' ============ 2. TUR: ad degismis olabilir -> ONERI ============
    Set tipOnb = New Sozluk
    ReDim ciftE(1 To gSay * satirSay + 1)
    ReDim ciftR(1 To gSay * satirSay + 1)
    ReDim ciftSkor(1 To gSay * satirSay + 1)
    ReDim ciftAd(1 To gSay * satirSay + 1)
    ReDim ciftYer(1 To gSay * satirSay + 1)
    ReDim ciftTip(1 To gSay * satirSay + 1)
    ciftSay = 0

    For e = 1 To gSay
        If Not gEslesti(e) And gHedefDeger(e) <> "" Then
            If Not kullanilan.Exists(BH(Trim$(gHedefDeger(e)))) Then

                gTip = GuvenliMetin(v(gSatir(e), P_TIP))
                gYer = BH(Trim$(GuvenliMetin(v(gSatir(e), P_YER))))
                adlar = Split(gAdlar(e), "|")

                For r = 1 To satirSay
                    If Not pAlindi(r) And pNorm(r) <> "" Then
                        If Trim$(GuvenliMetin(pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF).Value)) = "" Then

                            adSkor = 0
                            For t = 0 To UBound(adlar)
                                If adlar(t) <> "" Then
                                    skor = Benzerlik(adlar(t), pAd(r))
                                    If skor > adSkor Then adSkor = skor
                                End If
                            Next t

                            yerAyni = (gYer <> "" And gYer = BH(pHarf(r)))

                            kTip = ""
                            If gTip = "TARIH" Or gTip = "SAYI" Or gTip = "METIN" Then
                                kTip = OnbellekTip(tipOnb, kaynakWs, SutunNo(pHarf(r)), True)
                            End If
                            tipAyni = (kTip <> "" And kTip = gTip)

                            If Not TipCatisir(gTip, kTip) Then
                                If adSkor >= 0.6 Or (yerAyni And adSkor >= 0.3) Or (yerAyni And tipAyni) Then
                                    ciftSay = ciftSay + 1
                                    ciftE(ciftSay) = e
                                    ciftR(ciftSay) = r
                                    ciftAd(ciftSay) = adSkor
                                    ciftYer(ciftSay) = yerAyni
                                    ciftTip(ciftSay) = kTip
                                    ciftSkor(ciftSay) = adSkor + IIf(yerAyni, 0.3, 0) + IIf(tipAyni, 0.1, 0)
                                End If
                            End If

                        End If
                    End If
                Next r

            End If
        End If
    Next e

    ' En yuksek skorlu ciftten baslayarak ata (her kayit ve satir bir kez)
    ReDim eAtandi(1 To gSay)
    Do
        enIyi = 0
        enIyiSkor = 0
        For p = 1 To ciftSay
            If Not eAtandi(ciftE(p)) And Not pAlindi(ciftR(p)) Then
                If Not kullanilan.Exists(BH(Trim$(gHedefDeger(ciftE(p))))) Then
                    If ciftSkor(p) > enIyiSkor Then
                        enIyiSkor = ciftSkor(p)
                        enIyi = p
                    End If
                End If
            End If
        Next p
        If enIyi = 0 Then Exit Do

        e = ciftE(enIyi)
        r = ciftR(enIyi)
        eAtandi(e) = True
        pAlindi(r) = True
        gEslesti(e) = True
        kullanilan.Item(BH(Trim$(gHedefDeger(e)))) = True

        Set hucre = pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF)
        hucre.Value = HucreMetni(gHedefDeger(e))
        ProfilEkleriYaz pnl, c, ILK_SATIR + r - 1, v, gSatir(e)

        If ciftYer(enIyi) Then
            gYer = "ayni (" & pHarf(r) & ")"
        Else
            gYer = "farkli (" & GuvenliMetin(v(gSatir(e), P_YER)) & " -> " & pHarf(r) & ")"
        End If

        OneriIsaretle hucre, _
            "Profilde '" & GuvenliMetin(v(gSatir(e), P_KAYNAK)) & "' idi," & vbLf & _
            "bu dosyada '" & pAd(r) & "' bulundu." & vbLf & _
            "Ad benzerligi %" & Format$(ciftAd(enIyi) * 100, "0") & vbLf & _
            "Yer: " & gYer & vbLf & _
            "Veri tipi: " & TipMetni(GuvenliMetin(v(gSatir(e), P_TIP)), ciftTip(enIyi))
        oneriSay = oneriSay + 1
    Loop

    ' --- Profilde olup bulunamayanlar ---
    For e = 1 To gSay
        If Not gEslesti(e) And gHedefDeger(e) <> "" Then
            eksik = eksik & GuvenliMetin(v(gSatir(e), P_KAYNAK)) & " (-> " & GuvenliMetin(v(gSatir(e), P_HEDEF)) & "), "
        ElseIf Not gEslesti(e) And Trim$(GuvenliMetin(v(gSatir(e), P_HEDEF))) <> "" Then
            If InStr(hedefYok, GuvenliMetin(v(gSatir(e), P_HEDEF)) & ",") = 0 Then
                hedefYok = hedefYok & GuvenliMetin(v(gSatir(e), P_HEDEF)) & ", "
            End If
        End If
    Next e

    ' --- Dosyada olup profilde hic olmayanlar ---
    For r = 1 To satirSay
        If Not pAlindi(r) And pNorm(r) <> "" Then
            If Not profilAdSozluk.Exists(pNorm(r)) Then
                If Trim$(GuvenliMetin(pnl.Cells(ILK_SATIR + r - 1, c + OFS_HEDEF).Value)) = "" Then
                    yeni = yeni & pAd(r) & ", "
                End If
            End If
        End If
    Next r

    If Len(eksik) > 2 Then eksik = Left$(eksik, Len(eksik) - 2)
    If Len(yeni) > 2 Then yeni = Left$(yeni, Len(yeni) - 2)
    If Len(hedefYok) > 2 Then hedefYok = Left$(hedefYok, Len(hedefYok) - 2)

End Function


' Profil kaydinin FILTRE ve FORMUL bilgisini panel satirina yazar
Private Sub ProfilEkleriYaz(ByVal pnl As Worksheet, ByVal c As Long, ByVal satir As Long, _
                            ByVal v As Variant, ByVal i As Long)

    Dim mevcut As Variant

    If Trim$(GuvenliMetin(v(i, P_FILTRE))) <> "" Then
        pnl.Cells(satir, c + OFS_FILTRE).Value = v(i, P_FILTRE)
    End If
    If BH(Trim$(GuvenliMetin(v(i, P_FORMUL)))) = "EVET" Then
        mevcut = pnl.Cells(satir, c + OFS_ISLEM).Value
        pnl.Cells(satir, c + OFS_ISLEM).Value = IslemMetni(IslemSecili(mevcut), True)
    End If

End Sub


' Ana dosya listesinde, verilen sade ada sahip "sira"inci sutunu bulur
Private Function HedefListedeBul(hDeger() As String, hNorm() As String, hSira() As Long, _
                                 ByVal adNorm As String, ByVal sira As Long) As String

    Dim h As Long

    HedefListedeBul = ""
    If adNorm = "" Then Exit Function
    If sira < 1 Then sira = 1
    For h = LBound(hDeger) To UBound(hDeger)
        If hNorm(h) = adNorm And hSira(h) = sira Then
            HedefListedeBul = hDeger(h)
            Exit Function
        End If
    Next h

End Function


' Aktarim onay ekrani icin: profille paneldeki su anki durum farki
Public Sub ProfilFarklari(ByVal pnl As Worksheet, ByVal k As Long, ByVal ad As String, _
                          ByRef eksik As String, ByRef yeni As String)

    Dim prf As Worksheet
    Dim prfSon As Long, i As Long, t As Long, c As Long, son As Long
    Dim v As Variant
    Dim profilAdlariSoz As Sozluk, kullanilanHedef As Sozluk
    Dim adlar() As String
    Dim kaynakAdi As String, hedefAdi As String

    eksik = ""
    yeni = ""

    ProfilSayfasiHazirla
    Set prf = ThisWorkbook.Worksheets(SAYFA_PROFIL)
    prfSon = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row
    If prfSon < 2 Then Exit Sub

    c = BlokSutun(k)
    son = BlokSonSatir(pnl, k)
    v = IkiBoyutlu(prf.Range(prf.Cells(2, 1), prf.Cells(prfSon, P_SON)))

    Set profilAdlariSoz = New Sozluk
    Set kullanilanHedef = New Sozluk

    For i = ILK_SATIR To son
        hedefAdi = SutunAdi(pnl.Cells(i, c + OFS_HEDEF).Value)
        If hedefAdi <> "" Then kullanilanHedef.Item(Normalize(hedefAdi)) = True
    Next i

    For i = 1 To UBound(v, 1)
        If BH(Trim$(GuvenliMetin(v(i, P_AD)))) = BH(ad) Then
            adlar = Split(ProfilAdlari(GuvenliMetin(v(i, P_KAYNAK)), GuvenliMetin(v(i, P_DIGER))), "|")
            For t = 0 To UBound(adlar)
                If adlar(t) <> "" Then profilAdlariSoz.Item(adlar(t)) = True
            Next t
            hedefAdi = Trim$(GuvenliMetin(v(i, P_HEDEF)))
            If hedefAdi <> "" Then
                If Not kullanilanHedef.Exists(Normalize(hedefAdi)) Then
                    eksik = eksik & GuvenliMetin(v(i, P_KAYNAK)) & " (-> " & hedefAdi & "), "
                End If
            End If
        End If
    Next i

    If profilAdlariSoz.Count = 0 Then
        eksik = ""
        Exit Sub
    End If

    For i = ILK_SATIR To son
        kaynakAdi = SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value)
        If kaynakAdi <> "" And Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) = "" Then
            If Not profilAdlariSoz.Exists(Normalize(kaynakAdi)) Then yeni = yeni & kaynakAdi & ", "
        End If
    Next i

    If Len(eksik) > 2 Then eksik = Left$(eksik, Len(eksik) - 2)
    If Len(yeni) > 2 Then yeni = Left$(yeni, Len(yeni) - 2)

End Sub


' Profiller sayfasinin basliklari (eski sayfalara eksik sutunlari ekler)
Public Sub ProfilSayfasiHazirla()

    Dim prf As Worksheet
    Dim basliklar As Variant
    Dim i As Long

    Set prf = SayfaOlustur(SAYFA_PROFIL)
    basliklar = Array("PROFIL", "KAYNAK SUTUN", "HEDEF SUTUN", "FILTRE", "FORMUL", _
                      "DIGER ADLAR", "SIRA", "SON YER", "VERI TIPI", "HEDEF SIRA", "SON KAYIT")

    If Trim$(GuvenliMetin(prf.Cells(1, P_DIGER).Value)) = "" Or _
       Trim$(GuvenliMetin(prf.Cells(1, P_KAYIT).Value)) = "" Then
        For i = 0 To UBound(basliklar)
            prf.Cells(1, i + 1).Value = basliklar(i)
        Next i
        With prf.Range(prf.Cells(1, 1), prf.Cells(1, P_SON))
            .Font.Bold = True
            .Interior.Color = Palet(prTabloBaslik)
        End With
        prf.Range(prf.Columns(1), prf.Columns(P_SON)).ColumnWidth = 18
        prf.Columns(P_KAYNAK).ColumnWidth = 26
        prf.Columns(P_HEDEF).ColumnWidth = 26
        prf.Columns(P_DIGER).ColumnWidth = 34
        prf.Columns(P_KAYIT).ColumnWidth = 24
    End If

End Sub


Public Sub ProfilListesiYaz(ByVal lst As Worksheet)

    Dim prf As Worksheet
    Dim son As Long, i As Long, n As Long
    Dim gorulen As Sozluk
    Dim ad As String
    Dim adlar() As Variant

    ProfilSayfasiHazirla
    Set prf = ThisWorkbook.Worksheets(SAYFA_PROFIL)
    Set gorulen = New Sozluk

    lst.Columns(LST_PROFILLER).ClearContents
    son = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row
    If son < 2 Then Exit Sub

    ReDim adlar(0 To son)
    For i = 2 To son
        ad = Trim$(GuvenliMetin(prf.Cells(i, P_AD).Value))
        If ad <> "" Then
            If Not gorulen.Exists(BH(ad)) Then
                gorulen.Item(BH(ad)) = True
                adlar(n) = ad
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then Exit Sub
    ReDim Preserve adlar(0 To n - 1)
    ListeYaz lst, LST_PROFILLER, adlar, True

End Sub


Public Sub ProfilSatirlariSil(ByVal prf As Worksheet, ByVal ad As String)

    Dim son As Long, i As Long
    Dim silinecek As Range

    son = prf.Cells(prf.Rows.Count, P_AD).End(xlUp).Row
    For i = 2 To son
        If BH(Trim$(GuvenliMetin(prf.Cells(i, P_AD).Value))) = BH(Trim$(ad)) Then
            If silinecek Is Nothing Then
                Set silinecek = prf.Rows(i)
            Else
                Set silinecek = Union(silinecek, prf.Rows(i))
            End If
        End If
    Next i
    If Not silinecek Is Nothing Then silinecek.Delete

End Sub


' ============================================================
'  SETLER  (birden fazla paneli tek isimle saklama)
'
'  Her panel icin hedef dosya/sayfa, profil, sabit degerler, tutar,
'  hasar/ihbar sutunlari ve yazim modu saklanir. Kaynak dosya ve
'  "kaynak son satir" her ay degistigi icin SAKLANMAZ.
' ============================================================

Public Sub SetKaydet()

    Dim pnl As Worksheet, st As Worksheet
    Dim ad As String
    Dim adet As Long, k As Long, yaz As Long, kaydedilen As Long

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    SetSayfasiHazirla
    Set st = ThisWorkbook.Worksheets(SAYFA_SET)

    adet = BlokSayisi(pnl)
    If GorunurBlokSayisi(pnl) = 0 Then
        MsgBox "Kaydedilecek panel yok.", vbExclamation
        Exit Sub
    End If

    ad = Trim$(InputBox("Bu panel duzeni hangi isimle kaydedilsin?" & vbNewLine & _
                        "Ornek: HAYAT" & vbNewLine & vbNewLine & _
                        SetListesi(st, st.Cells(st.Rows.Count, ST_AD).End(xlUp).Row), "Set Kaydet"))
    If ad = "" Then Exit Sub

    EkranKilitle
    On Error GoTo Cikis
    SetSatirlariSil st, ad
    yaz = st.Cells(st.Rows.Count, ST_AD).End(xlUp).Row

    For k = 1 To adet
        If Not BlokGizliMi(pnl, k) Then
            yaz = yaz + 1
            kaydedilen = kaydedilen + 1
            st.Cells(yaz, ST_AD).Value = HucreMetni(ad)
            st.Cells(yaz, ST_PANEL).Value = k
            st.Cells(yaz, ST_HEDEF_DOSYA).Value = HucreMetni(AyarMetni(pnl, k, SATIR_HEDEF_DOSYA))
            st.Cells(yaz, ST_HEDEF_SAYFA).NumberFormat = "@"
            st.Cells(yaz, ST_HEDEF_SAYFA).Value = AyarMetni(pnl, k, SATIR_HEDEF_SAYFA)
            st.Cells(yaz, ST_PROFIL).Value = HucreMetni(AyarMetni(pnl, k, SATIR_PROFIL))
            st.Cells(yaz, ST_TUTAR).Value = HucreMetni(AyarMetni(pnl, k, SATIR_TUTAR))
            st.Cells(yaz, ST_HASAR).Value = HucreMetni(AyarMetni(pnl, k, SATIR_HASAR))
            st.Cells(yaz, ST_IHBAR).Value = HucreMetni(AyarMetni(pnl, k, SATIR_IHBAR))
            st.Cells(yaz, ST_YAZIM).Value = AyarMetni(pnl, k, SATIR_YAZIM)
            st.Cells(yaz, ST_SABITLER).Value = HucreMetni(SabitlerMetni(pnl, k))
        End If
    Next k
Cikis:
    If Err.Number <> 0 Then ad = ""
    EkranAc
    If ad = "" Then
        MsgBox "Set kaydedilemedi (bir hata olustu). Setler sayfasini kontrol et.", vbCritical
        Exit Sub
    End If

    OlayYaz "SET KAYDET", 0, 0, 0, "'" & ad & "' seti kaydedildi (" & kaydedilen & " panel)", ""

    MsgBox "'" & ad & "' seti kaydedildi." & vbNewLine & _
           kaydedilen & " panelin ayarlari (sabit degerler dahil) saklandi." & vbNewLine & vbNewLine & _
           "Kaynak dosya ve kaynak son satir sete dahil edilmedi;" & vbNewLine & _
           "bunlari her ay yeniden secersin.", vbInformation

End Sub


Public Sub SetYukle()

    Dim pnl As Worksheet, st As Worksheet
    Dim ad As String, liste As String, atlanan As String
    Dim son As Long, i As Long, k As Long
    Dim setPanel As Long, mevcut As Long, c As Long
    Dim sabitler As String, eskiSutun As String, eskiDeger As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    SetSayfasiHazirla
    Set st = ThisWorkbook.Worksheets(SAYFA_SET)

    son = st.Cells(st.Rows.Count, ST_AD).End(xlUp).Row
    If son < 2 Then
        MsgBox "Kayitli set yok. Once panelleri ayarlayip 'Set Kaydet' calistir.", vbExclamation
        Exit Sub
    End If

    liste = SetListesi(st, son)
    ad = Trim$(InputBox("Hangi set yuklensin?" & vbNewLine & vbNewLine & liste & vbNewLine & _
                        "Set adini yaz:", "Set Yukle"))
    If ad = "" Then Exit Sub

    setPanel = 0
    For i = 2 To son
        If BH(Trim$(GuvenliMetin(st.Cells(i, ST_AD).Value))) = BH(ad) Then
            If GuvenliSayi(st.Cells(i, ST_PANEL).Value) > setPanel Then setPanel = GuvenliSayi(st.Cells(i, ST_PANEL).Value)
        End If
    Next i

    If setPanel = 0 Then
        MsgBox "'" & ad & "' adinda bir set bulunamadi.", vbExclamation
        Exit Sub
    End If

    If MsgBox("'" & ad & "' seti yuklenecek." & vbNewLine & vbNewLine & _
              "Setteki panellerin mevcut ayarlari bu setin ayarlariyla degisecek." & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbQuestion, "Set Yukle") <> vbYes Then
        Exit Sub
    End If

    EkranKilitle
    On Error GoTo Hata

    ' Gerekli sayida blok olustur (MAX_BLOK'a kadar)
    mevcut = BlokSayisi(pnl)
    Do While mevcut < setPanel And mevcut < MAX_BLOK
        mevcut = mevcut + 1
        AlanTemizle pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, BlokSutun(mevcut)), _
                              pnl.Cells(SON_SATIR, BlokSutun(mevcut) + BLOK_GENISLIK - 1))
        BlokCiz pnl, mevcut
    Loop

    For i = 2 To son
        If BH(Trim$(GuvenliMetin(st.Cells(i, ST_AD).Value))) = BH(ad) Then
            k = GuvenliSayi(st.Cells(i, ST_PANEL).Value)
            If k > MAX_BLOK Then
                atlanan = atlanan & k & " "
            ElseIf k >= 1 And k <= mevcut Then
                c = BlokSutun(k)

                ' Kaldirilmis (gizli) panel ise yeniden ac
                If BlokGizliMi(pnl, k) Then
                    pnl.Range(pnl.Columns(c), pnl.Columns(c + BLOK_GENISLIK - 1)).Hidden = False
                    AlanTemizle pnl.Range(pnl.Cells(SATIR_BLOK_BASLIK, c), pnl.Cells(SON_SATIR, c + BLOK_GENISLIK - 1))
                    BlokCiz pnl, k
                End If

                Ayar(pnl, k, SATIR_HEDEF_DOSYA).Value = GuvenliMetin(st.Cells(i, ST_HEDEF_DOSYA).Value)
                Ayar(pnl, k, SATIR_HEDEF_SAYFA).Value = GuvenliMetin(st.Cells(i, ST_HEDEF_SAYFA).Value)
                Ayar(pnl, k, SATIR_PROFIL).Value = HucreMetni(GuvenliMetin(st.Cells(i, ST_PROFIL).Value))
                Ayar(pnl, k, SATIR_TUTAR).Value = HucreMetni(GuvenliMetin(st.Cells(i, ST_TUTAR).Value))
                Ayar(pnl, k, SATIR_HASAR).Value = HucreMetni(GuvenliMetin(st.Cells(i, ST_HASAR).Value))
                Ayar(pnl, k, SATIR_IHBAR).Value = HucreMetni(GuvenliMetin(st.Cells(i, ST_IHBAR).Value))
                Ayar(pnl, k, SATIR_YAZIM).Value = YazimMetni(YazimModuOku(st.Cells(i, ST_YAZIM).Value))

                ' --- Sabit degerler ---
                sabitler = GuvenliMetin(st.Cells(i, ST_SABITLER).Value)
                If sabitler = "" Then
                    ' v31 ve oncesi: tek sabit deger (F, G)
                    eskiSutun = Trim$(GuvenliMetin(st.Cells(i, ST_ESKI_SABIT_SUTUN).Value))
                    eskiDeger = GuvenliMetin(st.Cells(i, ST_ESKI_SABIT_DEGER).Value)
                    If eskiSutun <> "" And eskiDeger <> "" And eskiDeger <> "(bos birak)" Then
                        sabitler = eskiSutun & vbTab & eskiDeger
                    End If
                End If
                SabitleriYaz pnl, k, sabitler

                ' Her ay degisenler temizlenir
                Ayar(pnl, k, SATIR_KAYNAK_DOSYA).ClearContents
                Ayar(pnl, k, SATIR_KAYNAK_SAYFA).ClearContents
                Ayar(pnl, k, SATIR_SON_SATIR).ClearContents
                Ayar(pnl, k, SATIR_DAHIL).Value = HayirMetni()
                DahilBoya pnl, k
            End If
        End If
    Next i

    ButonlariKur pnl
    EkranAc

    YenileCalistir True

    OlayYaz "SET YUKLE", 0, 0, 0, "'" & ad & "' seti yuklendi (" & setPanel & " panele kadar)" & _
            IIf(atlanan <> "", "; " & MAX_BLOK & "'dan buyuk panel numaralari atlandi: " & atlanan, ""), ""

    MsgBox "'" & ad & "' seti yuklendi." & vbNewLine & vbNewLine & _
           IIf(atlanan <> "", "DIKKAT: panel " & atlanan & "en fazla " & MAX_BLOK & " panel siniri nedeniyle atlandi." & _
               vbNewLine & vbNewLine, "") & _
           "SIRADAKI ADIM:" & vbNewLine & _
           "Her panele o ayin kaynak dosyasini ve sayfasini sec," & vbNewLine & _
           "'YENILE' calistir. Eslestirmeler profilden otomatik gelir." & vbNewLine & _
           "Sonra aktaracagin panellerde 'Bu paneli aktar' = Evet yap.", vbInformation
    Exit Sub

Hata:
    EkranAc
    MsgBox "SET YUKLENEMEDI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description, vbCritical, "Set Yukle"

End Sub


Public Sub SetSil()

    Dim st As Worksheet
    Dim ad As String, liste As String
    Dim son As Long

    SetSayfasiHazirla
    Set st = ThisWorkbook.Worksheets(SAYFA_SET)

    son = st.Cells(st.Rows.Count, ST_AD).End(xlUp).Row
    If son < 2 Then
        MsgBox "Kayitli set yok.", vbInformation
        Exit Sub
    End If

    liste = SetListesi(st, son)
    ad = Trim$(InputBox("Hangi set silinsin?" & vbNewLine & vbNewLine & liste, "Set Sil"))
    If ad = "" Then Exit Sub

    If MsgBox("'" & ad & "' seti silinsin mi?", vbYesNo + vbQuestion + vbDefaultButton2, "Set Sil") <> vbYes Then Exit Sub

    EkranKilitle
    On Error GoTo Cikis
    SetSatirlariSil st, ad
Cikis:
    EkranAc
    OlayYaz "SET SIL", 0, 0, 0, "'" & ad & "' seti silindi", ""
    MsgBox "'" & ad & "' seti silindi.", vbInformation

End Sub


Private Function SetListesi(ByVal st As Worksheet, ByVal son As Long) As String

    Dim i As Long
    Dim gorulen As Sozluk
    Dim ad As String, metin As String

    Set gorulen = New Sozluk
    For i = 2 To son
        ad = Trim$(GuvenliMetin(st.Cells(i, ST_AD).Value))
        If ad <> "" Then
            If Not gorulen.Exists(BH(ad)) Then
                gorulen.Item(BH(ad)) = True
                metin = metin & "   " & ad & vbNewLine
            End If
        End If
    Next i
    If metin = "" Then metin = "   (yok)" & vbNewLine
    SetListesi = "KAYITLI SETLER:" & vbNewLine & metin

End Function


Private Sub SetSatirlariSil(ByVal st As Worksheet, ByVal ad As String)

    Dim son As Long, i As Long
    Dim silinecek As Range

    son = st.Cells(st.Rows.Count, ST_AD).End(xlUp).Row
    For i = 2 To son
        If BH(Trim$(GuvenliMetin(st.Cells(i, ST_AD).Value))) = BH(ad) Then
            If silinecek Is Nothing Then
                Set silinecek = st.Rows(i)
            Else
                Set silinecek = Union(silinecek, st.Rows(i))
            End If
        End If
    Next i
    If Not silinecek Is Nothing Then silinecek.Delete

End Sub


Public Sub SetSayfasiHazirla()

    Dim st As Worksheet
    Dim basliklar As Variant
    Dim i As Long

    Set st = SayfaOlustur(SAYFA_SET)

    If Trim$(GuvenliMetin(st.Cells(1, ST_AD).Value)) = "" Then
        basliklar = Array("SET ADI", "PANEL NO", "HEDEF DOSYA", "HEDEF SAYFA", "PROFIL", _
                          "SABIT SUTUN (v31)", "SABIT DEGER (v31)", "TUTAR SUTUN", "HASAR SUTUN", _
                          "IHBAR SUTUN", "YAZIM BASLANGICI")
        For i = 0 To UBound(basliklar)
            st.Cells(1, i + 1).Value = basliklar(i)
        Next i
        st.Range(st.Columns(1), st.Columns(ST_YAZIM)).ColumnWidth = 20
    End If

    If Trim$(GuvenliMetin(st.Cells(1, ST_SABITLER).Value)) = "" Then
        st.Cells(1, ST_SABITLER).Value = "SABIT DEGERLER"
        st.Columns(ST_SABITLER).ColumnWidth = 40
    End If

    With st.Range(st.Cells(1, 1), st.Cells(1, ST_SABITLER))
        .Font.Bold = True
        .Interior.Color = Palet(prTabloBaslik)
    End With

End Sub
