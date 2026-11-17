Attribute VB_Name = "M06_Aktarim"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.0  -  M06_Aktarim  (AKTARIM MOTORU)
'
' GIRIS MAKROLARI
'   Aktar         : "Bu paneli aktar = Evet" olan panelleri SIRAYLA aktarir
'   AktarSecili   : bulundugun panelde yalnizca ISLEM = SEC sutunlar
'                   (SECILI SUTUN YAMASI)
'   OnKontrol     : KURU CALISMA - hicbir sey yazmadan aktarimin ne
'                   yapacagini ve kaynak verideki sorunlari raporlar
'   KontrolTemizle: v33'un ana dosyada biraktigi filtre/olcut izlerini siler
'
' AKTARIM ASAMALARI (AktarIc)
'   1 PlanHazirla        panel okunur, basliklar dogrulanir   -> ENGEL
'   2 Mukerrer           ayni kaynak daha once aktarilmis mi  -> soru
'   3 SonSatirBelirle    supheli alt satir / toplam satiri    -> soru
'   4 VeriOku+SatirSec   sutun sutun okuma, filtre
'   5 HedefAralikBelirle yazim modu, satir siniri             -> ENGEL
'   6 EkDikkat           tip, yesil, profil, tarih, bicim     -> DIKKAT
'   7 Onay               tek onay ekrani (DIKKAT varsa varsayilan Hayir)
'   8 Yaz                (satir ekleme) yedek, bicim, yazma, imza, Gecmis
'   9 Mutabakat          kaynak / yazilan / hedef onceki / hedef yeni
'
' FILTRE SOZDIZIMI (FILTRE sutunu)
'   AXA            tam deger (buyuk/kucuk harf, Turkce harf farki yok)
'   AXA, ALLIANZ   birden fazla deger (virgul; deger virgul iceriyorsa
'                  ';' kullan). Metin, sutunda AYNEN varsa tek deger sayilir.
'   AX*  A?A       joker
'   >100 <=500 =0  sayi; >01.01.2026 tarih karsilastirma
'   <>IPTAL        esit olmayan (birden fazla <> hepsi saglanmali)
'   BOS / DOLU     bos ya da dolu hucre
' ============================================================

Private Type AktarimPlani
    k As Long
    sadeceIsaretli As Boolean
    partiNo As Long
    sira As Long
    toplam As Long
    kuru As Boolean                 ' OnKontrol: soru sorma, yazma

    pnl As Worksheet
    kaynakWb As Workbook
    hedefWb As Workbook
    kaynakWs As Worksheet
    hedefWs As Worksheet
    kaynakTanim As String
    hedefTanim As String
    mod As YazimModu

    ' Eslesmeler: 1..gercekAdet kaynaktan, gercekAdet+1..adet sabit deger
    adet As Long
    gercekAdet As Long
    kaynakNo() As Long
    hedefNo() As Long
    kaynakAdlari() As String
    hedefAdlari() As String
    formul() As Boolean
    sabitDeger() As Variant
    korumaTip() As String
    formulSayisi As Long
    sabitMetni As String
    bozukMetni As String
    oneriSay As Long
    oneriListesi As String

    ' Filtre
    filtreAdet As Long
    filtreler() As FiltreSutun
    filtreMetni As String

    ' Tutar mutabakati
    tutarJ As Long
    tutarHedefNo As Long
    tutarAd As String

    ' Satirlar
    otomatikSon As Long
    sonSatir As Long
    supheliNot As String
    n As Long
    uygun() As Long
    kv() As Variant
    kf() As Variant

    ' Hedef aralik
    baslangic() As Long
    enKucukBas As Long
    enBuyukBas As Long
    eklenecekSatir As Long
    boslukIlk As Long
    boslukSon As Long
    yama As Boolean
    yamaAktarim As Long

    ' Mesajlar
    engel As String
    engelSay As Long
    dikkat As String
    dikkatSay As Long

    ' Aktueryal kontrol noktalari K1..K14 (M10_Denetim): KD_* durumlari
    kDurum() As String

    ' --- v1.1 ---
    basZaman As Double              ' Timer, adim suresi (ms) icin
    kTip() As String                ' EkDikkat: kaynak / hedef veri tipi
    hTip() As String
    anahtarHedefNo As Long          ' K9: ana dosya anahtar sutunu
    anahtarJ As Long                '     ve onu dolduran eslesme
    anahtarAd As String
    kirilimHedefNo As Long          ' K12/K13: kirilim (brans/sirket)
    kirilimJ As Long
    kirilimAd As String
    hasarJ As Long
    ihbarJ As Long
    donemSonu As Date               ' K10: rapor donemi sonu
    donemNotu As String
    mutAdet As Long                 ' M1: mutabakat yapilan sutunlar (1 = ana tutar)
    mutJ() As Long
    grupAnahtar As Sozluk           ' K13: kirilim degeri -> grup no
    grupAdet As Long
    grupN() As Long                 ' grup basina satir
    grupT() As Currency             ' grup x mutabakat sutunu toplami (1-boyutlu: (g-1)*mutAdet + m)
    kaynakIzi As String             ' K14: kaynak icerik izi
    planIzi As String               ' K14: izin karsilastirilabilir oldugu plan
    icerikIzi As String             ' K2 (M7): icerik izi
    soyYol As String                ' M15: kaynak soyu
    soyBoyut As Double
    soyZaman As String
    kilitNotu As String             ' M12
    ' Metrikler ve kalite skoru (-1 = uygulanamaz)
    mTamlik As Double
    mGecerlilik As Double
    mTekillik As Double
    mMakulluk As Double
    mMutabakat As Double
    zN As Variant                   ' K11: satir ve tutar z-skoru (Empty = gecmis yetersiz)
    zT As Variant
    benfordMad As Variant
    buyukSay As Long
    metrik As String
    kalite As Double
End Type


' ============================================================
'  GIRIS MAKROLARI
' ============================================================

' "Bu paneli aktar = Evet" olan bloklari SIRAYLA aktarir.
Public Sub Aktar()

    Dim pnl As Worksheet
    Dim adet As Long, k As Long, i As Long
    Dim secili() As Long, secilenSayi As Long
    Dim liste As String
    Dim partiNo As Long, basarili As Long
    Dim sonuc As AktarimSonucu

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    adet = BlokSayisi(pnl)

    ReDim secili(1 To adet)
    For k = 1 To adet
        If Not BlokGizliMi(pnl, k) Then
            If DahilMi(pnl, k) Then
                secilenSayi = secilenSayi + 1
                secili(secilenSayi) = k
                liste = liste & "  " & secilenSayi & ") PANEL " & k & "  " & HazirlikEtiketi(k) & vbNewLine & "       " & _
                        AyarMetni(pnl, k, SATIR_KAYNAK_DOSYA) & " / " & AyarMetni(pnl, k, SATIR_KAYNAK_SAYFA) & vbNewLine & _
                        "       -> " & AyarMetni(pnl, k, SATIR_HEDEF_DOSYA) & " / " & _
                        AyarMetni(pnl, k, SATIR_HEDEF_SAYFA) & vbNewLine
            End If
        End If
    Next k

    If secilenSayi = 0 Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "Aktarilacak panel secilmemis." & vbNewLine & vbNewLine & _
               "Aktarmak istedigin panellerde 'Bu paneli aktar' hucresini" & vbNewLine & _
               "'Evet' yap, sonra tekrar dene."
        Exit Sub
    End If

    If MsgBox("SIRALI AKTARIM" & vbNewLine & vbNewLine & _
              Sigdir(liste, 600) & vbNewLine & _
              "Toplam " & secilenSayi & " panel sirayla aktarilacak." & vbNewLine & _
              "Her biri icin ayrica onay sorulacak." & vbNewLine & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbQuestion, "Aktarim") <> vbYes Then
        Exit Sub
    End If

    partiNo = YeniPartiNo()
    g_MutabakatOzet = ""

    For i = 1 To secilenSayi
        k = secili(i)
        sonuc = AktarCalistir(k, False, partiNo, i, secilenSayi)

        ' Basarili panelden sonra soru yok: siradaki panelin kendi onay
        ' ekrani zaten var. Yalnizca iptal / engel sonrasi karar sorulur.
        If sonuc = asTamam Then
            basarili = basarili + 1
        ElseIf i < secilenSayi Then
            If MsgBox("PANEL " & k & IIf(sonuc = asIptal, " aktarilmadi (iptal).", " aktarilamadi (ENGEL / hata).") & _
                      vbNewLine & vbNewLine & "Kalan panellere (" & (secilenSayi - i) & ") devam edilsin mi?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Sirali aktarim") <> vbYes Then Exit For
        End If
    Next i

    If basarili > 0 Then
        YedekSaklamaUygula
        DenetimSessiz
    End If

    Bildir BS_OK, "Aktarim", "SIRALI AKTARIM BITTI" & vbNewLine & vbNewLine & _
           "Parti no           : " & partiNo & vbNewLine & _
           "Secilen panel      : " & secilenSayi & vbNewLine & _
           "Aktarilan panel    : " & basarili & vbNewLine & vbNewLine & _
           "TUTAR MUTABAKATI" & vbNewLine & _
           IIf(g_MutabakatOzet = "", "   (aktarim yok)" & vbNewLine, Sigdir(g_MutabakatOzet, 500)) & vbNewLine & _
           "HATIRLATMA: Ana dosyayi SIMDI kaydet, sonra bu paneli de kaydet." & vbNewLine & vbNewLine & _
           "Hepsini birden geri almak icin 'Geri Al' calistirip" & vbNewLine & _
           "'parti' secenegini sec." & vbNewLine & vbNewLine & _
           "Yonetici ozeti: Denetim sayfasi (guncellendi)."

End Sub


' SECILI SUTUN YAMASI: bulundugun panelde yalnizca ISLEM = SEC sutunlar.
' Yazim modu "Sayfanin en alti" ise bu kaynagin SON AKTIF aktariminin
' satirlarina UZERINE yazar (eski degerler yedeklenir; Geri Al geri getirir).
Public Sub AktarSecili()

    Dim pnl As Worksheet

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    g_MutabakatOzet = ""
    If AktarCalistir(AktifBlok(), True, YeniPartiNo(), 1, 1) = asTamam Then
        YedekSaklamaUygula
        DenetimSessiz
    End If

End Sub


' KURU CALISMA: bulundugun panel icin aktarimin butun kontrollerini
' calistirir, kaynak tarih kontrollerini yapar; HICBIR SEY YAZMAZ.
Public Sub OnKontrol()

    Dim pnl As Worksheet
    Dim p As AktarimPlani
    Dim metin As String, ozet As String, mukerrer As String
    Dim hata As String
    Dim kisaltildi As Boolean
    Dim raporIlk As Long
    Dim ekranDikkat As String

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub

    p.k = AktifBlok()
    p.kuru = True
    KBaslat p
    p.toplam = 1
    p.sira = 1
    g_OkumaUyarisi = ""

    EkranKilitle
    On Error GoTo Hata
    DurumCubugu "ON KONTROL: PANEL " & p.k & " okunuyor ..."

    If PlanHazirla(p) Then
        If p.engel = "" Then
            ' Mukerrer aktarim: AktarIc'taki soru, kuru calismada DIKKAT olur
            mukerrer = OncekiAktarimlar(p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name))
            KTamam p, 2
            If mukerrer <> "" Then
                KEkle p, 2, False, "MUKERRER AKTARIM UYARISI: bu kaynak ayni ana dosyaya daha once " & _
                          "aktarilmis ve geri alinmamis; tekrar aktarilirsa kayitlar IKI KEZ bulunur: " & TekSatir(mukerrer)
            End If
            SonSatirBelirle p
            KTamam p, 3
            VeriOku p
            SatirSec p
            If p.n > 0 Then
                HedefAralikBelirle p
                EkDikkat p
            Else
                KEkle p, 3, True, "Filtreye uyan (dolu) satir yok."
            End If
        End If
    End If
    GoTo Sonuc

Hata:
    hata = Err.Number & " - " & Err.Description
    Resume Sonuc

Sonuc:
    On Error Resume Next
    DurumCubugu ""
    EkranAc

    ozet = "ON KONTROL  -  PANEL " & p.k & "   (hicbir sey yazilmadi)" & vbNewLine & vbNewLine
    If p.kaynakTanim <> "" Then ozet = ozet & "Aylik : " & Left$(p.kaynakTanim, 70) & vbNewLine
    If p.hedefTanim <> "" Then ozet = ozet & "Ana   : " & Left$(p.hedefTanim, 70) & vbNewLine
    If p.sonSatir >= 2 Then
        ozet = ozet & "Okunacak satirlar : 2 - " & SayiMetni(p.sonSatir)
        If p.sonSatir < p.otomatikSon Then ozet = ozet & "   (dosyada " & SayiMetni(p.otomatikSon) & ". satira kadar veri var)"
        ozet = ozet & vbNewLine
    End If
    If p.filtreMetni <> "" Then ozet = ozet & "Filtre : " & Left$(p.filtreMetni, 120) & vbNewLine
    If p.n > 0 Then
        ozet = ozet & "Aktarilacak : " & SayiMetni(p.n) & " satir  x  " & p.adet & " sutun"
        If p.enKucukBas > 0 Then
            ozet = ozet & "  ->  ana dosya " & SayiMetni(p.enKucukBas) & " - " & SayiMetni(p.enBuyukBas + p.n - 1) & _
                   IIf(p.yama, "  (YAMA: uzerine yazilir)", "")
        End If
        ozet = ozet & vbNewLine
    End If
    If p.supheliNot <> "" Then ozet = ozet & vbNewLine & "SUPHELI ALT SATIRLAR: " & Left$(p.supheliNot, 220) & vbNewLine
    ozet = ozet & vbNewLine & "KONTROL NOKTALARI:" & vbNewLine & KarneSatiri(p.kDurum) & vbNewLine

    If hata <> "" Then
        ozet = ozet & vbNewLine & "KONTROL SIRASINDA HATA: " & hata & vbNewLine
    End If
    If p.engel <> "" Then
        metin = MaddeKisalt(p.engel, 330)
        If metin <> p.engel Then kisaltildi = True
        ozet = ozet & vbNewLine & "ENGEL (" & p.engelSay & " madde, aktarim YAPILAMAZ):" & vbNewLine & metin
    End If
    If p.dikkat <> "" Then
        ' Supheli kuyruk yukarida ayrica gosterildi; ekranda tekrar edilmez
        ' (Rapor'a yine tam yazilir).
        ekranDikkat = MaddeCikar(p.dikkat, "- [K3] KAYNAK SON SATIR SUPHELI")
        If ekranDikkat <> "" Then
            metin = MaddeKisalt(ekranDikkat, MSGBOX_SINIR - Len(ozet) - 110)
            If metin <> ekranDikkat Then kisaltildi = True
            ozet = ozet & vbNewLine & "DIKKAT (" & p.dikkatSay & " madde):" & vbNewLine & metin
        End If
    End If
    If p.engel = "" And p.dikkat = "" And hata = "" Then
        ozet = ozet & vbNewLine & "Sorun bulunamadi. Aktarima hazir."
    End If
    If p.kalite >= 0 Then ozet = ozet & vbNewLine & "Veri kalitesi : " & Format$(p.kalite, "0") & " / 100"
    If kisaltildi Then ozet = ozet & vbNewLine & "Maddelerin tamami Rapor sayfasinda."

    ' Rapor: once ozet satiri, sonra HER madde kendi satirinda, kisaltilmadan
    raporIlk = OlayYaz("KONTROL", p.k, 0, 0, "On kontrol: " & p.n & " satir, " & p.engelSay & " ENGEL, " & _
                       p.dikkatSay & " DIKKAT" & IIf(p.supheliNot <> "", ". Supheli alt satirlar: " & p.supheliNot, ""), _
                       TekSatir(KarneSatiri(p.kDurum)), _
                       DosyaAdiGuvenli(p.kaynakWb), SayfaAdiGuvenli(p.kaynakWs), DosyaAdiGuvenli(p.hedefWb), SayfaAdiGuvenli(p.hedefWs))
    If p.metrik = "" Then KaliteHesapla p
    RaporEkAlanlar raporIlk, p.soyYol, p.soyBoyut, p.soyZaman, p.kaynakIzi, p.icerikIzi, SureMs(p.basZaman), "", _
                   p.kalite, p.metrik
    MaddeleriRaporaYaz p, p.engel, "KONTROL", "ENGEL", KD_ENGEL
    MaddeleriRaporaYaz p, p.dikkat, "KONTROL", "DIKKAT", KD_DIKKAT
    If Not p.pnl Is Nothing Then
        KarneYaz p.pnl, p.k, p.kDurum, TR("3 #On Kontrol: ") & Format$(Now, "dd.mm hh:nn") & _
                 IIf(p.n > 0, "  " & ChrW(183) & "  " & SayiMetni(p.n) & TR(" sat#ir"), ""), p.kalite
    End If

    ' UX-01: sonuc modal degil. Karne + rozet + HAZIRLIK ekranda; ozet
    ' bildirim seridinde, her madde Rapor'da (seride tiklayinca acilir).
    Bildir IIf(p.engel <> "" Or hata <> "", BS_ENGEL, IIf(p.dikkat <> "", BS_DIKKAT, BS_OK)), _
           "On Kontrol  PANEL " & p.k, _
           SayiMetni(p.n) & " satir, " & p.engelSay & " ENGEL, " & p.dikkatSay & " DIKKAT" & _
           IIf(p.kalite >= 0, ", kalite " & Format$(p.kalite, "0"), "") & vbLf & ozet, raporIlk

End Sub


' Listeden, verilen onekle baslayan maddeleri cikarir
Private Function MaddeCikar(ByVal liste As String, ByVal onek As String) As String
    Dim parcalar() As String
    Dim i As Long
    Dim sonuc As String
    parcalar = Split(liste, vbNewLine)
    For i = 0 To UBound(parcalar)
        If Trim$(parcalar(i)) <> "" And Left$(parcalar(i), Len(onek)) <> onek Then
            sonuc = sonuc & parcalar(i) & vbNewLine
        End If
    Next i
    MaddeCikar = sonuc
End Function


' Madde listesinin her maddesini Rapor'a ayri bir KONTROL satiri olarak,
' KISALTMADAN yazar (MsgBox'a sigmayan ayrinti kaybolmasin).
Private Sub MaddeleriRaporaYaz(ByRef p As AktarimPlani, ByVal liste As String, ByVal olay As String, _
                               ByVal karar As String, ByVal seviye As String)

    Dim parcalar() As String
    Dim i As Long, kapa As Long
    Dim m As String, kod As String

    If liste = "" Then Exit Sub
    parcalar = Split(liste, vbNewLine)
    For i = 0 To UBound(parcalar)
        m = Trim$(parcalar(i))
        If Left$(m, 2) = "- " Then m = Mid$(m, 3)
        If m <> "" Then
            kod = ""
            If Left$(m, 2) = "[K" Then
                kapa = InStr(m, "]")
                If kapa > 2 Then kod = Mid$(m, 2, kapa - 2) & " " & seviye
            End If
            OlayYaz olay, p.k, p.partiNo, 0, m, karar, DosyaAdiGuvenli(p.kaynakWb), SayfaAdiGuvenli(p.kaynakWs), _
                    DosyaAdiGuvenli(p.hedefWb), SayfaAdiGuvenli(p.hedefWs), kod
        End If
    Next i

End Sub


' ------------------------------------------------------------
'  KONTROL NOKTASI DURUMLARI  (K1..K8, bkz. M10_Denetim)
' ------------------------------------------------------------
Private Sub KBaslat(ByRef p As AktarimPlani)
    Dim n As Long
    ReDim p.kDurum(1 To K_SAYI)
    For n = 1 To K_SAYI
        p.kDurum(n) = KD_BEKLIYOR
    Next n
    p.basZaman = Timer
    p.mTamlik = -1
    p.mGecerlilik = -1
    p.mTekillik = -1
    p.mMakulluk = -1
    p.mMutabakat = -1
    p.kalite = -1
    p.zN = Empty
    p.zT = Empty
    p.benfordMad = Empty
End Sub


' "[Kn] metin" maddesini ENGEL ya da DIKKAT listesine ekler, durumu yukseltir
Private Sub KEkle(ByRef p As AktarimPlani, ByVal n As Long, ByVal engelMi As Boolean, ByVal metin As String)
    If engelMi Then
        MaddeEkle p.engel, p.engelSay, "[K" & n & "] " & metin
        KSeviye p, n, KD_ENGEL
    Else
        MaddeEkle p.dikkat, p.dikkatSay, "[K" & n & "] " & metin
        KSeviye p, n, KD_DIKKAT
    End If
End Sub


Private Sub KSeviye(ByRef p As AktarimPlani, ByVal n As Long, ByVal d As String)
    If KSira(d) > KSira(p.kDurum(n)) Then p.kDurum(n) = d
End Sub


' Kontrol calisti ve sorun bulmadi (daha once DIKKAT/ENGEL/YOK olduysa degismez)
Private Sub KTamam(ByRef p As AktarimPlani, ByVal n As Long)
    If p.kDurum(n) = KD_BEKLIYOR Then p.kDurum(n) = KD_OK
End Sub


' Kontrol bu aktarimda uygulanmaz (or. tutar sutunu secilmemis)
Private Sub KYok(ByRef p As AktarimPlani, ByVal n As Long)
    If p.kDurum(n) = KD_BEKLIYOR Then p.kDurum(n) = KD_YOK
End Sub


' Rapor sayfasini acip verilen satira gider
Private Sub RaporuGoster(ByVal satir As Long)
    On Error Resume Next
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(SAYFA_RAPOR).Activate
    If satir > 0 Then
        ActiveWindow.ScrollRow = satir
        ThisWorkbook.Worksheets(SAYFA_RAPOR).Cells(satir, R_ACIKLAMA).Select
    End If
    On Error GoTo 0
End Sub


' ============================================================
'  TEK PANEL AKTARIMI
'  Donus: asTamam = aktarildi, asIptal = kullanici vazgecti, asHata = yapilamadi
' ============================================================
Private Function AktarCalistir(ByVal k As Long, ByVal sadeceIsaretli As Boolean, _
                               ByVal partiNo As Long, ByVal sira As Long, _
                               ByVal toplam As Long) As AktarimSonucu

    Dim hata As String
    AktarCalistir = asHata
    EkranKilitle
    On Error GoTo Hata
    AktarCalistir = AktarIc(k, sadeceIsaretli, partiNo, sira, toplam)
Cikis:
    HesapAc
    EkranAc
    Exit Function
Hata:
    hata = Err.Number & " - " & Err.Description
    On Error Resume Next
    OlayYaz "HATA", k, partiNo, 0, "AktarCalistir: " & hata, "Aktarim tamamlanamadi"
    On Error GoTo 0
    GoTo Cikis

End Function


Private Function AktarIc(ByVal k As Long, ByVal sadeceIsaretli As Boolean, _
                         ByVal partiNo As Long, ByVal sira As Long, _
                         ByVal toplam As Long) As AktarimSonucu

    Dim p As AktarimPlani
    Dim mukerrer As String, karar As String, hataMetni As String
    Dim kilitSahip As String, kilitZaman As String
    Dim kilitAlindi As Boolean
    Dim cevap As VbMsgBoxResult

    AktarIc = asHata
    p.k = k
    p.sadeceIsaretli = sadeceIsaretli
    p.partiNo = partiNo
    p.sira = sira
    p.toplam = toplam
    g_OkumaUyarisi = ""
    KBaslat p

    On Error GoTo Hata

    ' ============ 1. PLAN ============
    If Not PlanHazirla(p) Or p.engel <> "" Then
        EngelBildir p
        Exit Function
    End If

    ' ============ 2. MUKERRER AKTARIM ============
    If Not sadeceIsaretli Then
        mukerrer = OncekiAktarimlar(p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name))
        If mukerrer <> "" Then
            cevap = MsgBox("DIKKAT - BU KAYNAK DAHA ONCE AKTARILMIS  (PANEL " & k & ")" & vbNewLine & vbNewLine & _
                           Left$(p.kaynakTanim, 80) & vbNewLine & _
                           "ayni ana dosyaya daha once aktarilmis ve geri alinmamis:" & vbNewLine & vbNewLine & _
                           mukerrer & vbNewLine & _
                           "Tekrar aktarirsan ayni kayitlar ana dosyada IKI KEZ bulunur." & vbNewLine & _
                           "Tutar mutabakati bunu yakalamaz (her iki yazim da kendi" & vbNewLine & _
                           "icinde tutarlidir)." & vbNewLine & vbNewLine & _
                           "Yine de devam edilsin mi?", _
                           vbYesNo + vbExclamation + vbDefaultButton2, "Mukerrer aktarim")
            karar = IIf(cevap = vbYes, "Yine de devam edildi", "Aktarim iptal edildi")
            OlayYaz "UYARI", k, partiNo, 0, "Mukerrer aktarim uyarisi, onceki aktarimlar: " & TekSatir(mukerrer), karar, _
                    p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
            If cevap <> vbYes Then
                KSeviye p, 2, KD_DIKKAT
                KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR iptal: m#ukerrer aktar#im  (") & Format$(Now, "dd.mm hh:nn") & ")"
                AktarIc = asIptal
                Exit Function
            End If
            KEkle p, 2, False, "MUKERRER AKTARIM: bu kaynak daha once aktarilmis ve geri alinmamis; kullanici " & _
                  "yine de devam etti. Onceki: " & TekSatir(mukerrer)
        End If
        KTamam p, 2
    Else
        KYok p, 2
    End If

    ' ============ 3. KAYNAK SON SATIR ============
    If Not SonSatirBelirle(p) Then
        KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR iptal: son sat#ir sorusu  (") & Format$(Now, "dd.mm hh:nn") & ")"
        AktarIc = asIptal
        Exit Function
    End If
    KTamam p, 3
    If p.sonSatir < 2 Then
        KEkle p, 3, True, "Kaynak sayfada aktarilacak veri yok."
        EngelBildir p
        Exit Function
    End If

    ' ============ 4. VERIYI OKU, SATIRLARI SEC ============
    DurumCubugu "PANEL " & k & ": aylik dosya okunuyor ..."
    VeriOku p
    SatirSec p
    DurumCubugu ""
    If p.n = 0 Then
        KEkle p, 3, True, "Filtreye uyan (dolu) satir yok. Filtre: " & p.filtreMetni
        EngelBildir p
        Exit Function
    End If

    ' ============ 5. HEDEF ARALIK ============
    HedefAralikBelirle p
    If p.engel <> "" Then
        EngelBildir p
        Exit Function
    End If

    ' ============ 6. DIKKAT MADDELERI ============
    EkDikkat p
    If p.engel <> "" Then
        EngelBildir p
        Exit Function
    End If

    ' ============ 7. ONAY ============
    If Not OnayAl(p) Then
        KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR: onay ekran#inda vazge#cildi  (") & Format$(Now, "dd.mm hh:nn") & ")"
        AktarIc = asIptal
        Exit Function
    End If

    ' ============ 8-9. YAZ VE MUTABAKAT (ana dosya kilidi altinda, M12) ============
    If KilitBaskasinda(p.hedefWb, kilitSahip, kilitZaman) Then
        KEkle p, 1, True, "ANA DOSYA KILITLI: " & kilitSahip & " (" & kilitZaman & ")."
        EngelBildir p
        Exit Function
    End If
    p.kilitNotu = KilitAl(p.hedefWb)
    kilitAlindi = (p.kilitNotu = "")
    Yaz p
    If kilitAlindi Then KilitBirak p.hedefWb
    AktarIc = asTamam
    Exit Function

Hata:
    hataMetni = Err.Number & " - " & Err.Description
    On Error Resume Next
    If kilitAlindi Then KilitBirak p.hedefWb
    HesapAc
    DurumCubugu ""
    OlayYaz "HATA", k, partiNo, 0, "Aktarim sirasinda hata: " & hataMetni, _
            "Ana dosya yarim yazilmis olabilir; Gecmis'e islenen sutunlar Geri Al ile alinabilir"
    On Error GoTo 0
    AktarIc = asHata
    Application.ScreenUpdating = True
    MsgBox "AKTARIM SIRASINDA HATA  -  PANEL " & k & vbNewLine & vbNewLine & _
           "Hata : " & hataMetni & vbNewLine & vbNewLine & _
           "Ana dosyayi kontrol et; yarim yazilmis olabilir." & vbNewLine & _
           "Yazilan sutunlar Gecmis'e islendi: gerekirse 'Geri Al'.", vbCritical, "Hata"

End Function


' ============================================================
'  1. PLAN: panel okunur, basliklar dogrulanir
'  Donus False: dosya/sayfa yok gibi temel eksik (p.engel doludur)
' ============================================================
Private Function PlanHazirla(ByRef p As AktarimPlani) As Boolean

    Dim pnl As Worksheet
    Dim c As Long, i As Long, j As Long, son As Long, kNo As Long, hNo As Long, sabitNo As Long
    Dim eskiNo As Long
    Dim durum As BaslikDurumu
    Dim gorulen As Sozluk
    Dim islem As Variant, filtreHam As String

    PlanHazirla = False
    Set pnl = PanelSayfasi()
    Set p.pnl = pnl
    c = BlokSutun(p.k)

    Set p.kaynakWb = DosyaBul(Ayar(pnl, p.k, SATIR_KAYNAK_DOSYA).Value)
    Set p.hedefWb = DosyaBul(Ayar(pnl, p.k, SATIR_HEDEF_DOSYA).Value)
    Set p.kaynakWs = SayfaBul(p.kaynakWb, Ayar(pnl, p.k, SATIR_KAYNAK_SAYFA).Value)
    Set p.hedefWs = SayfaBul(p.hedefWb, Ayar(pnl, p.k, SATIR_HEDEF_SAYFA).Value)

    If p.kaynakWs Is Nothing Or p.hedefWs Is Nothing Then
        KEkle p, 1, True, "Kaynak veya hedef dosya/sayfa bulunamadi. Secimleri kontrol edip 'YENILE' calistir."
        Exit Function
    End If

    p.kaynakTanim = p.kaynakWb.Name & " / " & CStr(p.kaynakWs.Name)
    p.hedefTanim = p.hedefWb.Name & " / " & CStr(p.hedefWs.Name)
    p.mod = YazimModuOku(Ayar(pnl, p.k, SATIR_YAZIM).Value)

    If p.hedefWb.ReadOnly And Not p.kuru Then
        KEkle p, 1, True, "ANA DOSYA SALT OKUNUR ACIK (" & p.hedefWb.Name & "). Muhtemelen baska biri " & _
                  "acik tutuyor; yazilanlar kaydedilemez."
    ElseIf p.hedefWb.ReadOnly Then
        KEkle p, 1, False, "Ana dosya SALT OKUNUR acik; bu haliyle aktarim yapilamaz."
    End If

    ' --- Panelden eslesmeleri oku ---
    son = BlokSonSatir(pnl, p.k)
    ReDim p.kaynakNo(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.hedefNo(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.kaynakAdlari(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.hedefAdlari(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.formul(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.sabitDeger(1 To EN_FAZLA_SUTUN + SABIT_EN_FAZLA)
    ReDim p.filtreler(1 To EN_FAZLA_SUTUN)

    For i = ILK_SATIR To son
        If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_KAYNAK).Value)) <> "" Then
            islem = pnl.Cells(i, c + OFS_ISLEM).Value

            If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_HEDEF).Value)) <> "" And _
               (Not p.sadeceIsaretli Or IslemSecili(islem)) Then

                kNo = SutunNo(pnl.Cells(i, c + OFS_KAYNAK).Value)
                hNo = SutunNo(pnl.Cells(i, c + OFS_HEDEF).Value)
                If kNo = 0 Or hNo = 0 Then
                    p.bozukMetni = p.bozukMetni & i & " "
                ElseIf p.adet < EN_FAZLA_SUTUN Then
                    p.adet = p.adet + 1
                    p.kaynakNo(p.adet) = kNo
                    p.hedefNo(p.adet) = hNo
                    p.kaynakAdlari(p.adet) = SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value)
                    p.hedefAdlari(p.adet) = SutunAdi(pnl.Cells(i, c + OFS_HEDEF).Value)
                    p.formul(p.adet) = IslemFormul(islem)
                    If p.formul(p.adet) Then p.formulSayisi = p.formulSayisi + 1
                    If OneriMi(pnl.Cells(i, c + OFS_HEDEF)) Then
                        p.oneriSay = p.oneriSay + 1
                        If p.oneriSay <= 3 Then
                            p.oneriListesi = p.oneriListesi & p.kaynakAdlari(p.adet) & " -> " & p.hedefAdlari(p.adet) & "; "
                        End If
                    End If
                End If
            End If

            ' Filtre: secili sutun yamasinda da AYNI filtre uygulanir
            ' (yoksa satirlar ilk aktarimla hizalanmazdi)
            filtreHam = Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_FILTRE).Value))
            If filtreHam <> "" Then
                kNo = SutunNo(pnl.Cells(i, c + OFS_KAYNAK).Value)
                If kNo > 0 And p.filtreAdet < EN_FAZLA_SUTUN Then
                    p.filtreAdet = p.filtreAdet + 1
                    p.filtreler(p.filtreAdet) = FiltreAyristir(filtreHam)
                    p.filtreler(p.filtreAdet).Sutun = kNo
                    p.filtreler(p.filtreAdet).Ad = SutunAdi(pnl.Cells(i, c + OFS_KAYNAK).Value)
                End If
            End If
        End If
    Next i

    p.gercekAdet = p.adet

    ' --- Sabit degerler (yamada YAZILMAZ: sabit sutunlarin altina fazladan
    '     satir eklenmesin / yamada zaten yazilmis olsunlar) ---
    If Not p.sadeceIsaretli Then
        For i = SABIT_ILK To SABIT_SON
            If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_SABIT_SUTUN).Value)) = "" Then
                If Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_SABIT_DEGER).Value)) <> "" Then
                    KEkle p, 5, True, "Sabit deger satiri " & (i - SABIT_ILK + 1) & _
                              ": deger yazilmis ama ana dosya SUTUNU secilmemis."
                End If
            Else
                sabitNo = SutunNo(pnl.Cells(i, c + OFS_SABIT_SUTUN).Value)
                If sabitNo = 0 Then
                    KEkle p, 5, True, "Sabit deger satiri " & (i - SABIT_ILK + 1) & _
                              ": ana dosya sutunu okunamadi. Menuden yeniden sec."
                ElseIf Trim$(GuvenliMetin(pnl.Cells(i, c + OFS_SABIT_DEGER).Value)) = "" Then
                    KEkle p, 5, True, "Sabit deger satiri " & (i - SABIT_ILK + 1) & " (" & _
                              SutunAdi(pnl.Cells(i, c + OFS_SABIT_SUTUN).Value) & "): DEGER bos."
                Else
                    p.adet = p.adet + 1
                    p.kaynakNo(p.adet) = 0
                    p.hedefNo(p.adet) = sabitNo
                    p.sabitDeger(p.adet) = pnl.Cells(i, c + OFS_SABIT_DEGER).Value
                    p.kaynakAdlari(p.adet) = "(sabit: " & GuvenliMetin(p.sabitDeger(p.adet)) & ")"
                    p.hedefAdlari(p.adet) = SutunAdi(pnl.Cells(i, c + OFS_SABIT_SUTUN).Value)
                    p.sabitMetni = p.sabitMetni & p.hedefAdlari(p.adet) & " = " & GuvenliMetin(p.sabitDeger(p.adet)) & "; "
                End If
            End If
        Next i
    End If

    If p.gercekAdet = 0 Then
        If p.sadeceIsaretli Then
            KEkle p, 1, True, "ISLEM sutununda SEC isaretli eslesme yok."
        Else
            KEkle p, 1, True, "Hic eslesme yok."
        End If
        PlanHazirla = True
        Exit Function
    End If

    ' --- Basliklar hala panelde gorunen yerde mi? ---
    For j = 1 To p.adet
        If p.kaynakNo(j) > 0 Then
            eskiNo = p.kaynakNo(j)
            durum = BaslikDogrula(p.kaynakWs, p.kaynakNo(j), p.kaynakAdlari(j))
            If durum = bdYerDegisti Then
                KEkle p, 1, False, "Aylik dosyada '" & p.kaynakAdlari(j) & "' " & SutunHarfi(eskiNo) & _
                          " yerine " & SutunHarfi(p.kaynakNo(j)) & " sutununda; oradan okunacak."
            ElseIf durum = bdYok Then
                KEkle p, 1, True, "Aylik dosyada '" & p.kaynakAdlari(j) & "' sutunu yok (" & _
                          SutunHarfi(eskiNo) & " sutununun basligi degismis). 'YENILE' calistir."
            ElseIf durum = bdBelirsiz Then
                KEkle p, 1, True, "Aylik dosyada '" & p.kaynakAdlari(j) & "' adiyla birden fazla" & _
                          " sutun var ve hicbiri beklenen yerde degil. 'YENILE' calistir."
            End If
        End If

        eskiNo = p.hedefNo(j)
        durum = BaslikDogrula(p.hedefWs, p.hedefNo(j), p.hedefAdlari(j))
        If durum = bdYerDegisti Then
            KEkle p, 1, False, "Ana dosyada '" & p.hedefAdlari(j) & "' " & SutunHarfi(eskiNo) & _
                      " yerine " & SutunHarfi(p.hedefNo(j)) & " sutununda; oraya yazilacak."
        ElseIf durum = bdYok Then
            KEkle p, 1, True, "Ana dosyada '" & p.hedefAdlari(j) & "' sutunu yok (" & _
                      SutunHarfi(eskiNo) & " sutununun basligi degismis). 'YENILE' calistir."
        ElseIf durum = bdBelirsiz Then
            KEkle p, 1, True, "Ana dosyada '" & p.hedefAdlari(j) & "' adiyla birden fazla" & _
                      " sutun var ve hicbiri beklenen yerde degil. 'YENILE' calistir."
        End If
    Next j

    For j = 1 To p.filtreAdet
        eskiNo = p.filtreler(j).Sutun
        durum = BaslikDogrula(p.kaynakWs, p.filtreler(j).Sutun, p.filtreler(j).Ad)
        If durum = bdYerDegisti Then
            KEkle p, 1, False, "Filtre sutunu '" & p.filtreler(j).Ad & "' " & SutunHarfi(eskiNo) & _
                      " yerine " & SutunHarfi(p.filtreler(j).Sutun) & " sutununda."
        ElseIf durum >= bdYok Then
            KEkle p, 1, True, "Filtre sutunu '" & p.filtreler(j).Ad & "' aylik dosyada bulunamadi. 'YENILE' calistir."
        End If
        p.filtreMetni = p.filtreMetni & p.filtreler(j).Ad & " = " & p.filtreler(j).Ham & ";  "
    Next j

    ' --- Ayni hedef sutuna iki kaynak (veya sabit + eslesme) yazilamaz ---
    Set gorulen = New Sozluk
    For j = 1 To p.adet
        If gorulen.Exists(p.hedefNo(j)) Then
            KEkle p, 1, True, "Ana dosyadaki " & SutunHarfi(p.hedefNo(j)) & " (" & p.hedefAdlari(j) & _
                      ") sutununa iki kaynak yaziliyor: " & gorulen.Item(p.hedefNo(j)) & " ve " & p.kaynakAdlari(j) & _
                      ". Biri digerinin uzerine yazardi."
        Else
            gorulen.Item(p.hedefNo(j)) = p.kaynakAdlari(j)
        End If
    Next j

    ' --- Tutar sutunu (mutabakat): ana dosyanin bir sutunu ---
    p.tutarHedefNo = SutunNo(Ayar(pnl, p.k, SATIR_TUTAR).Value)
    If p.tutarHedefNo > 0 Then
        p.tutarAd = SutunAdi(Ayar(pnl, p.k, SATIR_TUTAR).Value)
        If BaslikDogrula(p.hedefWs, p.tutarHedefNo, p.tutarAd) >= bdYok Then
            KEkle p, 7, False, "MUTABAKAT YOK: tutar sutunu '" & p.tutarAd & "' ana dosyada bulunamadi."
            p.tutarHedefNo = 0
        End If
        If p.tutarHedefNo > 0 Then
            For j = 1 To p.adet
                If p.hedefNo(j) = p.tutarHedefNo And p.kaynakNo(j) > 0 Then p.tutarJ = j
            Next j
            If p.tutarJ = 0 And Not p.sadeceIsaretli Then
                KEkle p, 7, False, "MUTABAKAT YOK: tutar sutunu (" & SutunHarfi(p.tutarHedefNo) & " " & p.tutarAd & _
                          ") bu panelde hicbir aylik sutunla eslesmemis."
            End If
        End If
    End If

    PlanEki p
    KTamam p, 1
    PlanHazirla = True

End Function


Private Sub EngelBildir(ByRef p As AktarimPlani)

    Bildir BS_ENGEL, "Engel", "AKTARIM YAPILAMAZ  -  PANEL " & p.k & vbNewLine & vbNewLine & _
           "KONTROL NOKTALARI:" & vbNewLine & KarneSatiri(p.kDurum) & vbNewLine & vbNewLine & _
           MaddeKisalt(p.engel, 650) & vbNewLine & _
           "Hicbir sey yazilmadi. Sorunu giderip tekrar dene." & vbNewLine & _
           "(Ipucu: 'On Kontrol' butun sorunlari tek seferde listeler.)"

    MaddeleriRaporaYaz p, p.engel, "ENGEL", "Aktarim yapilmadi", KD_ENGEL
    KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR: ENGEL, hi#cbir #sey yaz#ilmad#i  (") & Format$(Now, "dd.mm hh:nn") & ")"

End Sub


' ============================================================
'  3. KAYNAK SON SATIR  (cogunluk mantigi + alt bolum ayraci)
'  Donus False = kullanici iptal etti
' ============================================================
Private Function SonSatirBelirle(ByRef p As AktarimPlani) As Boolean

    Dim sonlar() As Long
    Dim j As Long, sinirSon As Long, cogunluk As Long, onerilen As Long, ayrac As Long
    Dim metin As String, tasanlar As String, karar As String
    Dim cevap As VbMsgBoxResult

    SonSatirBelirle = True
    ReDim sonlar(1 To p.adet)
    p.otomatikSon = 0
    For j = 1 To p.adet
        If p.kaynakNo(j) > 0 Then
            sonlar(j) = p.kaynakWs.Cells(p.kaynakWs.Rows.Count, p.kaynakNo(j)).End(xlUp).Row
            If sonlar(j) > p.otomatikSon Then p.otomatikSon = sonlar(j)
        End If
    Next j

    metin = AyarMetni(p.pnl, p.k, SATIR_SON_SATIR)
    If metin <> "" And IsNumeric(metin) Then sinirSon = CLng(metin)

    p.sonSatir = p.otomatikSon

    If sinirSon >= 2 Then
        ' Kullanici kendisi sinir yazmis: ona uyulur, soru sorulmaz
        If sinirSon < p.otomatikSon Then p.sonSatir = sinirSon
        Exit Function
    End If

    cogunluk = CogunlukSatiri(sonlar, p.kaynakNo, p.adet)
    onerilen = 0
    If cogunluk >= 2 And p.otomatikSon > cogunluk Then
        For j = 1 To p.adet
            If p.kaynakNo(j) > 0 Then
                If sonlar(j) > cogunluk Then
                    tasanlar = tasanlar & "   " & p.kaynakAdlari(j) & " (" & SutunHarfi(p.kaynakNo(j)) & _
                               "): " & sonlar(j) & ". satira kadar" & vbNewLine
                End If
            End If
        Next j
        onerilen = cogunluk
    End If

    ayrac = AltBlokAyraci(p.kaynakWs, p.kaynakNo, p.adet, p.otomatikSon)
    If ayrac >= 2 Then
        If onerilen = 0 Or ayrac < onerilen Then onerilen = ayrac
    End If

    If onerilen < 2 Or onerilen >= p.otomatikSon Then Exit Function

    p.supheliNot = "veri " & onerilen & ". satirda bitiyor gibi, dosyada " & p.otomatikSon & ". satira kadar deger var"

    If p.kuru Then
        ' Kuru calismada soru sorulmaz; disarida kalacak satirlarin icerigi
        ' gosterilir ki kullanici neyin elendigini gorsun.
        p.sonSatir = onerilen
        p.supheliNot = p.supheliNot & ". Icerik: " & _
                       TekSatir(KuyrukOnizleme(p.kaynakWs, p.kaynakNo, p.adet, onerilen + 1, p.otomatikSon))
        KEkle p, 3, False, "KAYNAK SON SATIR SUPHELI: " & p.supheliNot & _
                  " (toplam/not satiri olabilir). Kontrol onerilen satira gore yapildi."
        Exit Function
    End If

    metin = "KAYNAK SON SATIR SUPHELI  (PANEL " & p.k & ")" & vbNewLine & vbNewLine
    If tasanlar <> "" Then
        metin = metin & "Eslesen sutunlarin cogu " & cogunluk & ". satirda bitiyor, ama:" & vbNewLine & _
                Sigdir(tasanlar, 250) & vbNewLine
    End If
    If ayrac >= 2 Then
        metin = metin & "Veri " & ayrac & ". satirda bitiyor gibi; altinda bos satirla" & vbNewLine & _
                "ayrilmis bir bolum var (" & p.otomatikSon & ". satira kadar)." & vbNewLine & vbNewLine
    End If
    metin = metin & "Supheli satirlarin icerigi (genel toplam / not olabilir):" & vbNewLine & _
            KuyrukOnizleme(p.kaynakWs, p.kaynakNo, p.adet, onerilen + 1, p.otomatikSon) & vbNewLine & _
            "EVET   = " & onerilen & ". satira kadar oku (onerilen, temiz veri)" & vbNewLine & _
            "HAYIR  = hepsini oku (" & p.otomatikSon & ". satira kadar)" & vbNewLine & _
            "IPTAL  = aktarimi durdur" & vbNewLine & vbNewLine & _
            "(Baska bir sinir icin panelde 'Son satir' hucresine yaz.)"

    KSeviye p, 3, KD_DIKKAT
    cevap = MsgBox(Sigdir(metin, 900), vbYesNoCancel + vbExclamation + vbDefaultButton1, "Kaynak son satir")
    Select Case cevap
        Case vbYes
            p.sonSatir = onerilen
            karar = "Onerilen sinir kullanildi: " & onerilen & ". satir"
        Case vbNo
            p.sonSatir = p.otomatikSon
            karar = "Hepsi alindi (" & p.otomatikSon & ". satira kadar)"
        Case Else
            karar = "Aktarim iptal edildi"
            SonSatirBelirle = False
    End Select

    OlayYaz "UYARI", p.k, p.partiNo, 0, "Kaynak son satir supheli: cogunluk " & cogunluk & _
            ". satir, en gec " & p.otomatikSon & ". satir. " & TekSatir(tasanlar) & _
            IIf(ayrac >= 2, " Bos satirla ayrilmis alt bolum: " & ayrac & ". satirdan sonra.", ""), _
            karar, p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)

End Function


' Eslesen kaynak sutunlarinin EN COK hangi satirda bittigi.
' Esitlikte kucuk olan secilir (daha temkinli).
Private Function CogunlukSatiri(sonlar() As Long, kaynakNo() As Long, ByVal adet As Long) As Long

    Dim say As Sozluk
    Dim j As Long, enIyi As Long, enIyiSay As Long, n As Long
    Dim anahtar As Variant

    Set say = New Sozluk
    For j = 1 To adet
        If kaynakNo(j) > 0 And sonlar(j) >= 2 Then say.Artir sonlar(j)
    Next j

    For Each anahtar In say.Keys
        n = say.Item(anahtar)
        If n > enIyiSay Then
            enIyi = CLng(anahtar)
            enIyiSay = n
        ElseIf n = enIyiSay And CLng(anahtar) < enIyi Then
            enIyi = CLng(anahtar)
        End If
    Next anahtar

    CogunlukSatiri = enIyi

End Function


' Supheli alt satirlarin kisa icerigi (en fazla 4 satir, satir basina
' ilk 5 dolu deger)
Private Function KuyrukOnizleme(ByVal ws As Worksheet, kaynakNo() As Long, ByVal adet As Long, _
                                ByVal ilk As Long, ByVal son As Long) As String

    Dim r As Long, j As Long, say As Long
    Dim satir As String, deger As String, metin As String

    If son - ilk > 3 Then ilk = son - 3
    If ilk < 2 Then ilk = 2

    For r = ilk To son
        satir = ""
        say = 0
        For j = 1 To adet
            If kaynakNo(j) > 0 And say < 5 Then
                deger = Trim$(GuvenliMetin(ws.Cells(r, kaynakNo(j)).Text))
                If deger <> "" Then
                    satir = satir & Left$(deger, 14) & " | "
                    say = say + 1
                End If
            End If
        Next j
        If satir = "" Then
            satir = "(bos)"
        Else
            satir = Left$(satir, Len(satir) - 3)
        End If
        metin = metin & "   " & r & ". satir: " & satir & vbNewLine
    Next r

    KuyrukOnizleme = metin

End Function


' Kaynagin en altinda, eslesen sutunlarin HEPSINDE bos bir satirla
' ayrilmis kucuk bir bolum (toplam/not) var mi? Varsa bosluktan
' onceki son veri satirini, yoksa 0 dondurur.
Private Function AltBlokAyraci(ByVal ws As Worksheet, kaynakNo() As Long, ByVal adet As Long, _
                               ByVal sonSatirDeger As Long) As Long

    Dim ilk As Long, r As Long, i As Long, j As Long, kuyruk As Long
    Dim dolu() As Boolean
    Dim v As Variant

    AltBlokAyraci = 0
    ilk = sonSatirDeger - KUYRUK_PENCERE + 1
    If ilk < 2 Then ilk = 2
    If sonSatirDeger - ilk < 2 Then Exit Function

    ReDim dolu(ilk To sonSatirDeger)
    For j = 1 To adet
        If kaynakNo(j) > 0 Then
            v = SutunOku(ws, kaynakNo(j), ilk, sonSatirDeger)
            For i = 1 To UBound(v, 1)
                If Trim$(GuvenliMetin(v(i, 1))) <> "" Then dolu(ilk + i - 1) = True
            Next i
        End If
    Next j

    ' Asagidan yukari: once dolu kuyruk, sonra bos satir(lar), sonra veri
    r = sonSatirDeger
    Do While r >= ilk
        If Not dolu(r) Then Exit Do
        r = r - 1
        kuyruk = kuyruk + 1
    Loop
    If r < ilk Then Exit Function
    If kuyruk > KUYRUK_EN_FAZLA Then Exit Function

    Do While r >= ilk
        If dolu(r) Then Exit Do
        r = r - 1
    Loop
    If r < ilk Then Exit Function

    AltBlokAyraci = r

End Function


' ============================================================
'  4. VERI OKUMA VE SATIR SECIMI
' ============================================================
Private Sub VeriOku(ByRef p As AktarimPlani)

    Dim j As Long

    ReDim p.kv(1 To p.adet)
    ReDim p.kf(1 To p.adet)
    For j = 1 To p.adet
        If p.kaynakNo(j) > 0 Then
            p.kv(j) = SutunOku(p.kaynakWs, p.kaynakNo(j), 2, p.sonSatir, False)
            If p.formul(j) Then p.kf(j) = SutunOku(p.kaynakWs, p.kaynakNo(j), 2, p.sonSatir, True)
        End If
    Next j

End Sub


' Tamamen bos satirlari atlar, filtreyi uygular; p.uygun / p.n doldurur
Private Sub SatirSec(ByRef p As AktarimPlani)

    Dim i As Long, j As Long, satirSayisi As Long
    Dim fv() As Variant
    Dim tut() As Boolean
    Dim kol As Variant

    satirSayisi = p.sonSatir - 1
    p.n = 0
    If satirSayisi < 1 Then Exit Sub

    If p.filtreAdet > 0 Then
        ReDim fv(1 To p.filtreAdet)
        For j = 1 To p.filtreAdet
            fv(j) = SutunOku(p.kaynakWs, p.filtreler(j).Sutun, 2, p.sonSatir, False)
            FiltreBelirsizligiGider p.filtreler(j), fv(j)
        Next j
    End If

    ' SUTUN SUTUN ilerlenir: her sutun dizisi bir kez yerel degiskene alinir.
    ' (p.kv(j)(i, 1) gibi UDT icinden hucre hucre erisim VBA'da cok yavas.)
    ReDim tut(1 To satirSayisi)
    For j = 1 To p.gercekAdet
        kol = p.kv(j)
        For i = 1 To satirSayisi
            If Not tut(i) Then
                If Not IsEmpty(kol(i, 1)) Then
                    If Trim$(GuvenliMetin(kol(i, 1))) <> "" Then tut(i) = True
                End If
            End If
        Next i
    Next j

    For j = 1 To p.filtreAdet
        kol = fv(j)
        For i = 1 To satirSayisi
            If tut(i) Then
                If Not FiltreUyar(p.filtreler(j), kol(i, 1)) Then tut(i) = False
            End If
        Next i
    Next j

    ReDim p.uygun(1 To satirSayisi)
    For i = 1 To satirSayisi
        If tut(i) Then
            p.n = p.n + 1
            p.uygun(p.n) = i
        End If
    Next i

End Sub


' ============================================================
'  5. HEDEF ARALIK  (yazim modu, yama, satir siniri)
' ============================================================
Private Sub HedefAralikBelirle(ByRef p As AktarimPlani)

    Dim j As Long, genelSon As Long, boy As Long, sonYazilacak As Long, kalanSatir As Long
    Dim ilkY As Long, sonY As Long
    Dim kaynakEslesti As Boolean

    ReDim p.baslangic(1 To p.adet)
    p.eklenecekSatir = 0
    p.yama = False

    If p.sadeceIsaretli And p.mod = ymSayfaSonu Then

        ' --- SECILI SUTUN YAMASI: son aktif aktarimin satirlarina ---
        If Not SonAktifAralik(p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name), _
                              p.yamaAktarim, ilkY, sonY, kaynakEslesti) Then
            KEkle p, 1, True, "Secili sutun yamasi icin bu ana dosya sayfasinda AKTIF (geri alinmamis) " & _
                      "bir aktarim bulunamadi. Yamanin yazilacagi satirlar bilinmiyor. Yazim baslangicini " & _
                      "'Geri alinan araliga' veya 'Sutunun kendi sonu' yap ya da once tam aktarim yap."
            Exit Sub
        End If
        boy = sonY - ilkY + 1
        If p.n <> boy Then
            KEkle p, 1, True, "Satir sayisi tutmuyor: son aktarim (no " & p.yamaAktarim & ") ana dosyada " & _
                      ilkY & " - " & sonY & " (" & boy & " satir), simdi okunan " & p.n & " satir. Ayni filtre ve " & _
                      "ayni son satirla tekrar dene; yoksa satirlar kayardi."
            Exit Sub
        End If
        For j = 1 To p.adet
            p.baslangic(j) = ilkY
        Next j
        p.yama = True
        KEkle p, 1, False, "SECILI SUTUN YAMASI: secili sutunlar ana dosyada " & ilkY & " - " & sonY & _
                  " satirlarinin (aktarim no " & p.yamaAktarim & ") UZERINE yazilacak. Eski degerler yedeklenir; " & _
                  "Geri Al onlari geri getirir."
        If Not kaynakEslesti Then
            KEkle p, 1, False, "Son aktif aktarim BASKA bir aylik dosyadan yapilmis; yamanin dogru " & _
                      "satirlara gittiginden emin ol."
        End If

    ElseIf p.mod = ymGeriAlinan Then

        If Not GeriAlinanAralik(p.hedefWb.Name, CStr(p.hedefWs.Name), p.boslukIlk, p.boslukSon) Then
            KEkle p, 1, True, "'Geri alinan araliga' secili ama bu ana dosya sayfasinda geri alinmis " & _
                      "bir aralik yok. Bu mod, once 'Geri Al' ile bosaltilmis yere duzeltilmis veriyi yazmak icindir."
            Exit Sub
        End If
        boy = p.boslukSon - p.boslukIlk + 1
        If p.n > boy Then
            p.eklenecekSatir = p.n - boy
            KEkle p, 1, False, "SATIR EKLENECEK: gelen veri " & p.n & " satir, bosluk " & boy & " satir (" & _
                      p.boslukIlk & "-" & p.boslukSon & "); ana dosyaya " & p.eklenecekSatir & _
                      " satir eklenecek, altindaki veriler asagi kayacak (Gecmis ve Rapor satir numaralari da kaydirilir)."
        ElseIf p.n < boy Then
            KEkle p, 1, False, "BOSLUK KALACAK: gelen veri " & p.n & " satir, bosluk " & boy & _
                      " satir; " & (p.boslukIlk + p.n) & " - " & p.boslukSon & " arasi bos kalacak."
        End If
        For j = 1 To p.adet
            p.baslangic(j) = p.boslukIlk
        Next j

    ElseIf p.mod = ymSutunSonu Then

        For j = 1 To p.adet
            p.baslangic(j) = SutunSonDolu(p.hedefWs, p.hedefNo(j)) + 1
            If p.baslangic(j) < 2 Then p.baslangic(j) = 2
        Next j
        If p.adet > 1 And Not p.sadeceIsaretli Then
            KEkle p, 1, False, "SATIR HIZASI RISKI: 'Sutunun kendi sonu' ile " & p.adet & _
                      " sutun birden aktariliyor; her sutun kendi son dolu hucresinden baslar." & _
                      " Normal aylik aktarim icin dogru secim 'Sayfanin en alti'."
        End If

    Else

        genelSon = SonDoluSatir(p.hedefWs)
        If genelSon < 1 Then genelSon = 1
        For j = 1 To p.adet
            p.baslangic(j) = genelSon + 1
        Next j

    End If

    p.enKucukBas = p.baslangic(1)
    p.enBuyukBas = p.baslangic(1)
    For j = 2 To p.adet
        If p.baslangic(j) < p.enKucukBas Then p.enKucukBas = p.baslangic(j)
        If p.baslangic(j) > p.enBuyukBas Then p.enBuyukBas = p.baslangic(j)
    Next j

    If p.mod = ymSutunSonu And p.enKucukBas <> p.enBuyukBas Then
        KEkle p, 1, False, "Sutunlarin baslangic satirlari FARKLI (" & p.enKucukBas & " ile " & _
                  p.enBuyukBas & " arasi): ayni kaydin parcalari farkli satirlara yazilacak."
    End If

    ' --- Excel satir siniri ---
    If p.eklenecekSatir > 0 Then
        genelSon = SonDoluSatir(p.hedefWs)
        If genelSon + p.eklenecekSatir > p.hedefWs.Rows.Count Then
            KEkle p, 1, True, "Satir eklenemez: ana dosya Excel'in satir sinirini (" & _
                      SayiMetni(p.hedefWs.Rows.Count) & ") asar."
        End If
        kalanSatir = p.hedefWs.Rows.Count - (genelSon + p.eklenecekSatir)
    Else
        sonYazilacak = p.enBuyukBas + p.n - 1
        If sonYazilacak > p.hedefWs.Rows.Count Then
            KEkle p, 1, True, "Excel'in satir siniri asiliyor: bir sayfa en fazla " & _
                      SayiMetni(p.hedefWs.Rows.Count) & " satir alir, bu aktarim " & _
                      SayiMetni(sonYazilacak) & ". satira kadar yazacakti."
        End If
        kalanSatir = p.hedefWs.Rows.Count - sonYazilacak
    End If

    If kalanSatir < SATIR_UYARI_ESIGI And p.engel = "" Then
        KEkle p, 1, False, "SATIR SINIRINA YAKIN: aktarimdan sonra ana dosyada yalnizca " & SayiMetni(kalanSatir) & _
                  " satirlik yer kalacak (Excel siniri " & SayiMetni(p.hedefWs.Rows.Count) & ")."
    End If

End Sub


' ============================================================
'  6. EK DIKKAT MADDELERI
' ============================================================
Private Sub EkDikkat(ByRef p As AktarimPlani)

    Dim j As Long, i As Long, korunan As Long, hedefHata As Long
    Dim kod As Long, tip As String
    Dim tipOnb As Sozluk, dolacak As Sozluk
    Dim kTip As String, hTip As String, tipUyusmaz As String, yesilBos As String
    Dim profilAdi As String, pEksik As String, pYeni As String
    Dim korumaMetni As String, tarihMetni As String, sabitUyusmaz As String
    Dim x As Variant, kol As Variant
    Dim u() As Long
    Dim zorunluBos As Long, zorunluHucre As Long, onceki As String

    If p.tutarHedefNo > 0 Then
        hedefHata = TutarHataSayisi(p.hedefWs, p.tutarHedefNo)
        If hedefHata > 0 Then KEkle p, 7, True, "Ana dosya tutar sutununda " & hedefHata & " hata hucresi var; toplam dogrulanamaz."
    End If

    If g_OkumaUyarisi <> "" Then
        KEkle p, 6, False, "TARIH TASMASI (#####): su sutunlarin tarihleri SAYI olarak aktarilacak: " & Trim$(g_OkumaUyarisi)
    End If
    If p.bozukMetni <> "" Then
        KEkle p, 1, False, "Panelde okunamayan eslesme hucreleri atlanacak (panel satiri: " & _
                  Trim$(p.bozukMetni) & ")."
    End If
    If p.oneriSay > 0 Then
        KEkle p, 1, False, "ONAYLANMAMIS ONERI: " & p.oneriSay & " eslesme henuz onaylanmadi" & _
                  " (turuncu): " & p.oneriListesi
    End If
    If p.formulSayisi > 0 Then
        KEkle p, 6, False, p.formulSayisi & " sutun FORMUL olarak tasinacak; goreli satir referanslari" & _
                  " yeni satirlara kaydirilir, sutun harfleri ayni kalir."
    End If

    ' --- Veri tipi ve bicim korumasi ---
    u = p.uygun
    Set tipOnb = New Sozluk
    ReDim p.korumaTip(1 To p.adet)
    ReDim p.kTip(1 To p.adet)
    ReDim p.hTip(1 To p.adet)
    For j = 1 To p.gercekAdet
        kTip = OnbellekTip(tipOnb, p.kaynakWs, p.kaynakNo(j), True)
        hTip = OnbellekTip(tipOnb, p.hedefWs, p.hedefNo(j), False)
        p.kTip(j) = kTip
        p.hTip(j) = hTip
        If TipCatisir(kTip, hTip) And Not p.sadeceIsaretli Then
            tipUyusmaz = tipUyusmaz & p.kaynakAdlari(j) & " (" & kTip & ") -> " & p.hedefAdlari(j) & " (" & hTip & "); "
        End If
        If hTip = "TARIH" Or hTip = "SAYI" Or hTip = "METIN" Then
            p.korumaTip(j) = hTip
        Else
            p.korumaTip(j) = kTip
        End If

        korunan = 0
        kol = p.kv(j)
        tip = p.korumaTip(j)
        For i = 1 To p.n
            x = kol(u(i), 1)
            If VarType(x) = vbString Then
                ' Harfle baslayan metin icin cagri yapilmaz (KorumaGerekir hizli yolu)
                If Len(x) > 0 Then
                    kod = AscW(x)
                    If Not ((kod >= 65 And kod <= 90) Or (kod >= 97 And kod <= 122) Or kod > 127) Then
                        If KorumaGerekir(CStr(x), tip) Then korunan = korunan + 1
                    End If
                End If
            End If
        Next i
        If korunan > 0 And Not p.formul(j) Then
            korumaMetni = korumaMetni & p.hedefAdlari(j) & ": " & SayiMetni(korunan) & "; "
        End If
    Next j
    For j = p.gercekAdet + 1 To p.adet
        p.korumaTip(j) = OnbellekTip(tipOnb, p.hedefWs, p.hedefNo(j), False)
        ' SABIT DEGER ana dosya sutununun tipiyle celisiyor mu?
        ' (or. tarih sutununa "EYLUL" yazilmis)
        kTip = DegerTipi(p.sabitDeger(j))
        If TipCatisir(kTip, p.korumaTip(j)) Then
            sabitUyusmaz = sabitUyusmaz & p.hedefAdlari(j) & " = '" & GuvenliMetin(p.sabitDeger(j)) & "' (" & kTip & _
                           ", ana dosyada " & p.korumaTip(j) & "); "
        End If
    Next j

    If sabitUyusmaz <> "" Then
        KEkle p, 5, False, "SABIT DEGER TIPI UYUSMAZ: " & sabitUyusmaz
    End If
    If tipUyusmaz <> "" Then
        KEkle p, 5, False, "VERI TIPI UYUSMAZ, yanlis eslesme olabilir: " & tipUyusmaz
    End If
    If korumaMetni <> "" Then
        KEkle p, 6, False, "BICIM KORUMASI: bastaki sifir / uzun numara / formul gibi gorunen " & _
                  "metinler Excel bozmasin diye METIN olarak yazilacak: " & korumaMetni
    End If

    ' --- Kaynak tarih kontrolleri (hasar / ihbar) ---
    tarihMetni = KaynakTarihKontrolu(p)
    If tarihMetni <> "" Then KEkle p, 8, False, tarihMetni

    ' --- M1: mutabakat sutunlari; tutar butunlugu (mutabakat bozuk metni yakalamaz) ---
    MutabakatSutunlariSec p
    tarihMetni = TutarButunlugu(p)
    If tarihMetni <> "" Then KEkle p, 7, False, tarihMetni
    tarihMetni = EkTutarButunlugu(p)
    If tarihMetni <> "" Then KEkle p, 7, False, tarihMetni
    If p.mutAdet = 0 Then KYok p, 7

    ' --- Zorunlu sutunlarda satir bazinda bos hucre ---
    tarihMetni = YesilBosHucreler(p, zorunluBos, zorunluHucre)
    If tarihMetni <> "" Then KEkle p, 4, False, tarihMetni

    If Not p.sadeceIsaretli Then
        Set dolacak = New Sozluk
        For j = 1 To p.adet
            dolacak.Item(p.hedefNo(j)) = True
        Next j
        yesilBos = YesilBosSutunlar(p.hedefWs, dolacak)
        If yesilBos <> "" Then
            KEkle p, 4, False, "YESIL SUTUN BOS KALACAK (ana dosyada zorunlu): " & yesilBos
        End If
    End If

    ' --- K2 (M7): icerik tabanli mukerrer (dosya adi degisse de yakalanir) ---
    If Not p.sadeceIsaretli Then
        p.icerikIzi = IcerikIziHesapla(p)
        onceki = IcerikIziBul(p.hedefWb.Name, CStr(p.hedefWs.Name), p.icerikIzi)
        If onceki <> "" Then
            KEkle p, 2, False, "ICERIK AYNI: bu verinin icerik izi (" & p.icerikIzi & ") ana dosyaya daha once aktarilmis ve " & _
                      "geri alinmamis bir aktarimla birebir ayni (dosya adi farkli olabilir): " & TekSatir(onceki)
        End If
    End If

    ' --- v1.1 kontrol noktalari ---
    K9Kontrol p
    K10Kontrol p
    K11Kontrol p
    K13Hazirla p
    K12Kontrol p
    K14Kontrol p
    TamlikHesapla p, zorunluBos, zorunluHucre
    KaliteHesapla p

    KTamam p, 1
    KTamam p, 2
    KTamam p, 4
    KTamam p, 5
    KTamam p, 6
    KTamam p, 7
    KTamam p, 8

End Sub


' Bir degerin tipi: TARIH / SAYI / METIN (bos ise "")
Private Function DegerTipi(ByVal x As Variant) As String

    Dim s As String

    DegerTipi = ""
    If IsError(x) Or IsEmpty(x) Then Exit Function
    Select Case VarType(x)
        Case vbDate
            DegerTipi = "TARIH"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            DegerTipi = "SAYI"
        Case vbString
            s = Trim$(CStr(x))
            If s = "" Then
                DegerTipi = ""
            ElseIf MetinTarihMi(s) Then
                DegerTipi = "TARIH"
            ElseIf KodGibiMi(s) Then
                DegerTipi = "METIN"
            ElseIf IsNumeric(s) Then
                DegerTipi = "SAYI"
            Else
                DegerTipi = "METIN"
            End If
        Case Else
            DegerTipi = "METIN"
    End Select

End Function


' Tutar (mutabakat) sutununa giden aylik degerler: sayi olmayan dolu
' hucre ("1.200 TL", "YOK", #N/A), bos hucre ve eksi tutar sayilir.
' SayiDegeri ve AGGREGATE bunlari sessizce atladigi icin mutabakat
' "TUTUYOR" dese de veri eksik olabilir.
Private Function TutarButunlugu(ByRef p As AktarimPlani) As String

    Dim u() As Long
    Dim i As Long, r As Long
    Dim bozuk As Long, bos As Long, eksi As Long
    Dim ornBozuk As String, ornBos As String, ornEksi As String
    Dim ondalik As Long
    Dim x As Variant, kol As Variant
    Dim d As Double
    Dim metin As String

    TutarButunlugu = ""
    If p.tutarJ = 0 Then Exit Function

    kol = p.kv(p.tutarJ)
    u = p.uygun
    For i = 1 To p.n
        r = u(i)
        x = kol(r, 1)
        If IsError(x) Then
            bozuk = bozuk + 1
            If bozuk <= KONTROL_ORNEK Then ornBozuk = ornBozuk & (r + 1) & " "
        ElseIf Trim$(GuvenliMetin(x)) = "" Then
            bos = bos + 1
            If bos <= KONTROL_ORNEK Then ornBos = ornBos & (r + 1) & " "
        ElseIf SayiDegeri(x, d) Then
            If d < 0 Then
                eksi = eksi + 1
                If eksi <= KONTROL_ORNEK Then ornEksi = ornEksi & (r + 1) & " "
            End If
            If Abs(d * 100 - Round(d * 100, 0)) > 0.000001 Then ondalik = ondalik + 1
        Else
            bozuk = bozuk + 1
            If bozuk <= KONTROL_ORNEK Then ornBozuk = ornBozuk & (r + 1) & " "
        End If
    Next i

    If bozuk > 0 Then metin = metin & "sayi OLMAYAN deger " & bozuk & " hucre, toplama girmez (aylik satir " & _
                                Trim$(ornBozuk) & IIf(bozuk > KONTROL_ORNEK, " ...", "") & "); "
    If bos > 0 Then metin = metin & "BOS " & bos & " hucre (satir " & Trim$(ornBos) & IIf(bos > KONTROL_ORNEK, " ...", "") & "); "
    If eksi > 0 Then metin = metin & "EKSI tutar " & eksi & " hucre (satir " & Trim$(ornEksi) & IIf(eksi > KONTROL_ORNEK, " ...", "") & "); "
    If metin <> "" Then TutarButunlugu = "TUTAR SUTUNU (" & p.kaynakAdlari(p.tutarJ) & "): " & metin
    If p.n - bos > 0 Then p.mGecerlilik = OrtalamaEkle(p.mGecerlilik, (p.n - bos - bozuk) / (p.n - bos))
    If ondalik > 0 Then
        KEkle p, 6, False, "ONDALIK HANE: tutar sutununda (" & p.kaynakAdlari(p.tutarJ) & ") " & SayiMetni(ondalik) & _
                  " deger 2'den fazla ondalik hane tasiyor; mutabakat Currency ile 4 haneye kadar kesindir, " & _
                  "raporlar kurusa yuvarlar."
    End If

End Function


' M1: ek mutabakat sutunlarinda sayi OLMAYAN dolu hucre (toplama girmez,
' ana dosyada metin kalirsa yazimdan sonra FARK cikar)
Private Function EkTutarButunlugu(ByRef p As AktarimPlani) As String

    Dim m As Long, i As Long, bozuk As Long
    Dim kol As Variant, x As Variant
    Dim u() As Long
    Dim d As Double
    Dim metin As String

    u = p.uygun
    For m = 1 To p.mutAdet
        If p.mutJ(m) <> p.tutarJ Then
            bozuk = 0
            kol = p.kv(p.mutJ(m))
            For i = 1 To p.n
                x = kol(u(i), 1)
                If IsError(x) Then
                    bozuk = bozuk + 1
                ElseIf Not IsEmpty(x) Then
                    If Not SayiDegeri(x, d) Then
                        If Trim$(GuvenliMetin(x)) <> "" Then bozuk = bozuk + 1
                    End If
                End If
            Next i
            If bozuk > 0 Then metin = metin & p.kaynakAdlari(p.mutJ(m)) & ": " & SayiMetni(bozuk) & " hucre; "
        End If
    Next m
    If metin <> "" Then EkTutarButunlugu = "EK TUTAR SUTUNLARINDA sayi OLMAYAN deger (mutabakata girmez): " & metin

End Function


' Ana dosyada YESIL (zorunlu) olup eslestirilmis sutunlarda, aktarilacak
' satirlardan bos olanlar (YesilBosSutunlar yalnizca hic eslesmeyeni bulur).
Private Function YesilBosHucreler(ByRef p As AktarimPlani, ByRef bosToplam As Long, ByRef hucreToplam As Long) As String

    Dim u() As Long
    Dim renkler() As Long, renkVar() As Boolean
    Dim sonSut As Long, j As Long, i As Long, r As Long, bos As Long
    Dim orn As String, metin As String
    Dim kol As Variant

    YesilBosHucreler = ""
    If p.gercekAdet = 0 Then Exit Function
    sonSut = BaslikRenkleri(p.hedefWs, renkler, renkVar)

    For j = 1 To p.gercekAdet
        If p.hedefNo(j) >= 1 And p.hedefNo(j) <= sonSut Then
            If renkVar(p.hedefNo(j)) Then
                If YesilMi(renkler(p.hedefNo(j))) Then
                    bos = 0
                    orn = ""
                    kol = p.kv(j)
                    u = p.uygun
                    For i = 1 To p.n
                        r = u(i)
                        If Trim$(GuvenliMetin(kol(r, 1))) = "" Then
                            bos = bos + 1
                            If bos <= KONTROL_ORNEK Then orn = orn & (r + 1) & " "
                        End If
                    Next i
                    bosToplam = bosToplam + bos
                    hucreToplam = hucreToplam + p.n
                    If bos > 0 Then
                        metin = metin & p.hedefAdlari(j) & " sutununda " & bos & " satirda deger bos (%" & _
                                Format$((p.n - bos) / p.n * 100, "0.0") & " dolu; ornek aylik satir: " & _
                                Trim$(orn) & IIf(bos > KONTROL_ORNEK, " ...", "") & "); "
                    End If
                End If
            End If
        End If
    Next j

    If metin <> "" Then YesilBosHucreler = "ZORUNLU (YESIL) SUTUNDA BOS HUCRE: " & metin

End Function


' Hasar / ihbar tarihi sutunlari (panel ayari, ana dosya sutunu) hangi
' aylik sutundan dolduruluyorsa, aktarilacak satirlarda kontrol eder:
'   ihbar tarihi hasardan once, bos tarih, tarihe cevrilemeyen deger.
' Ana dosyaya HICBIR SEY yazmaz (v33'teki Gelismis Filtre kontrollerinin
' yerine, aktarimdan ONCE ve kaynak uzerinde).
Private Function KaynakTarihKontrolu(ByRef p As AktarimPlani) As String

    Dim u() As Long
    Dim hasarNo As Long, ihbarNo As Long, jH As Long, jI As Long, j As Long, i As Long, r As Long
    Dim xH As Variant, xI As Variant, kolH As Variant, kolI As Variant
    Dim dH As Double, dI As Double
    Dim okH As Boolean, okI As Boolean
    Dim once As Long, bos As Long, bozuk As Long, metinTarih As Long
    Dim ornekOnce As String, ornekBos As String, ornekBozuk As String
    Dim metin As String

    KaynakTarihKontrolu = ""
    hasarNo = SutunNo(Ayar(p.pnl, p.k, SATIR_HASAR).Value)
    ihbarNo = SutunNo(Ayar(p.pnl, p.k, SATIR_IHBAR).Value)
    If hasarNo = 0 And ihbarNo = 0 Then
        KYok p, 8
        Exit Function
    End If

    For j = 1 To p.gercekAdet
        If p.hedefNo(j) = hasarNo And hasarNo > 0 Then jH = j
        If p.hedefNo(j) = ihbarNo And ihbarNo > 0 Then jI = j
    Next j
    If jH = 0 And jI = 0 Then
        KYok p, 8
        Exit Function
    End If
    If jH > 0 Then kolH = p.kv(jH)
    If jI > 0 Then kolI = p.kv(jI)

    u = p.uygun
    For i = 1 To p.n
        r = u(i)
        okH = False
        okI = False
        If jH > 0 Then
            xH = kolH(r, 1)
            If Trim$(GuvenliMetin(xH)) = "" Then
                bos = bos + 1
                If bos <= KONTROL_ORNEK Then ornekBos = ornekBos & (r + 1) & " "
            ElseIf TarihSayisi(xH, dH) Then
                okH = True
            ElseIf VarType(xH) = vbString And MetinTarihMi(CStr(xH)) Then
                metinTarih = metinTarih + 1
            Else
                bozuk = bozuk + 1
                If bozuk <= KONTROL_ORNEK Then ornekBozuk = ornekBozuk & (r + 1) & " "
            End If
        End If
        If jI > 0 Then
            xI = kolI(r, 1)
            If Trim$(GuvenliMetin(xI)) = "" Then
                bos = bos + 1
                If bos <= KONTROL_ORNEK Then ornekBos = ornekBos & (r + 1) & " "
            ElseIf TarihSayisi(xI, dI) Then
                okI = True
            ElseIf VarType(xI) = vbString And MetinTarihMi(CStr(xI)) Then
                metinTarih = metinTarih + 1
            Else
                bozuk = bozuk + 1
                If bozuk <= KONTROL_ORNEK Then ornekBozuk = ornekBozuk & (r + 1) & " "
            End If
        End If
        If okH And okI Then
            If dI < dH Then
                once = once + 1
                If once <= KONTROL_ORNEK Then ornekOnce = ornekOnce & (r + 1) & " "
            End If
        End If
    Next i

    If once > 0 Then metin = metin & "ihbar tarihi hasardan ONCE: " & once & " satir (aylik satir " & Trim$(ornekOnce) & IIf(once > KONTROL_ORNEK, " ...", "") & "); "
    If bos > 0 Then metin = metin & "BOS tarih: " & bos & " hucre (satir " & Trim$(ornekBos) & IIf(bos > KONTROL_ORNEK, " ...", "") & "); "
    If bozuk > 0 Then metin = metin & "tarihe cevrilemeyen deger: " & bozuk & " hucre (satir " & Trim$(ornekBozuk) & IIf(bozuk > KONTROL_ORNEK, " ...", "") & "); "
    If metinTarih > 0 Then metin = metin & "METIN olarak yazilmis tarih: " & metinTarih & " hucre; "

    If metin <> "" Then KaynakTarihKontrolu = "KAYNAK TARIH KONTROLU: " & metin

End Function


' ============================================================
'  BICIM KORUMASI
'
'  Range.Value'ya verilen METIN, Excel tarafindan elle yazilmis gibi
'  yorumlanir: "0004821" -> 4821, "1234567890123456789" -> 1,23E+18,
'  "=A1" -> formul. Bu fonksiyon hangi metnin oldugu gibi (tek
'  tirnakla) yazilmasi gerektigine karar verir. Sayi, tarih ve
'  normal metinler v33'teki gibi Excel'e birakilir.
' ============================================================
Private Function KorumaGerekir(ByVal s As String, ByVal tip As String) As Boolean

    Dim t As String, ilk As String
    Dim kod As Long

    KorumaGerekir = False
    If Len(s) = 0 Then Exit Function

    ' HIZLI YOL: harfle baslayan metin (isim, sirket, aciklama) hicbir
    ' kurala takilmaz. 200.000 satirlik metin sutunlarinda belirleyici.
    kod = AscW(s)
    If (kod >= 65 And kod <= 90) Or (kod >= 97 And kod <= 122) Or kod > 127 Then Exit Function

    t = Trim$(s)
    If t = "" Then Exit Function
    ilk = Left$(t, 1)

    ' Formul gibi gorunen metin (formul enjeksiyonu / #NAME?)
    If ilk = "=" Then
        KorumaGerekir = True
        Exit Function
    End If
    If (ilk = "+" Or ilk = "-" Or ilk = "@") And Not IsNumeric(t) Then
        KorumaGerekir = True
        Exit Function
    End If

    ' Yalnizca rakamdan olusan kodlar
    If Not (t Like "*[!0-9]*") Then
        If Len(t) >= 2 And ilk = "0" Then
            KorumaGerekir = True            ' bastaki sifir (police / hasar no)
        ElseIf Len(t) >= 16 Then
            KorumaGerekir = True            ' 15 haneden uzun: hassasiyet kaybi
        ElseIf tip = "METIN" Then
            KorumaGerekir = True            ' metin sutununa sayi gibi kod
        End If
        Exit Function
    End If

    ' "12E3" gibi bilimsel gosterime benzeyen kodlar
    If t Like "*[0-9][Ee][0-9]*" And tip <> "SAYI" Then
        If IsNumeric(t) Then KorumaGerekir = True
        Exit Function
    End If

    ' Metin sutununa sayi / tarih gibi gorunen metin
    If tip = "METIN" Then
        If IsNumeric(t) Or MetinTarihMi(t) Then KorumaGerekir = True
    End If

End Function


Private Function HucreDegeri(ByVal x As Variant, ByVal tip As String) As Variant
    If VarType(x) = vbString Then
        If KorumaGerekir(CStr(x), tip) Then
            HucreDegeri = "'" & CStr(x)
            Exit Function
        End If
    End If
    HucreDegeri = x
End Function


' Kaynak sutunun R1C1 formulunu, SUTUN HARFLERI ayni kalacak ve goreli
' SATIR referanslari hedef satira kayacak bicime cevirir:
'   kaynak D5: =B5*C5   -> hedef K870005: =B870005*C870005
' Ayni formul deseni bir kez cevrilir (onbellek). 255 karakteri asan
' formulde ConvertFormula calismaz; formul oldugu gibi birakilir.
Private Function FormulSatirKaydir(ByVal r1c1 As String, ByVal kaynakSutun As Long, ByVal hedefSutun As Long, _
                                   ByVal onb As Sozluk, ByRef kaydirilamayan As Long) As String

    Dim a1 As String, sonuc As String
    Dim anahtar As String

    anahtar = r1c1
    If onb.Exists(anahtar) Then
        FormulSatirKaydir = onb.Item(anahtar)
        Exit Function
    End If

    sonuc = r1c1
    On Error Resume Next
    a1 = Application.ConvertFormula(r1c1, xlR1C1, xlA1, xlRelative, ThisWorkbook.Worksheets(SAYFA_PANEL).Cells(1000, kaynakSutun))
    If Err.Number = 0 And Not IsError(a1) And a1 <> "" Then
        sonuc = Application.ConvertFormula(a1, xlA1, xlR1C1, xlRelative, ThisWorkbook.Worksheets(SAYFA_PANEL).Cells(1000, hedefSutun))
        If Err.Number <> 0 Or IsError(sonuc) Then
            sonuc = r1c1
            kaydirilamayan = kaydirilamayan + 1
        End If
    Else
        kaydirilamayan = kaydirilamayan + 1
    End If
    Err.Clear
    On Error GoTo 0

    onb.Item(anahtar) = sonuc
    FormulSatirKaydir = sonuc

End Function


' ============================================================
'  7. ONAY
' ============================================================
Private Function OnayAl(ByRef p As AktarimPlani) As Boolean

    Dim ozet As String, karar As String
    Dim kalanMetin As Long
    Dim cevap As VbMsgBoxResult

    ozet = "AKTARIM ONAYI  -  PANEL " & p.k
    If p.toplam > 1 Then ozet = ozet & "   (" & p.sira & "/" & p.toplam & ")"
    ozet = ozet & vbNewLine & vbNewLine
    ozet = ozet & "Aylik dosya : " & Left$(p.kaynakTanim, 70) & vbNewLine
    ozet = ozet & "Ana dosya   : " & Left$(p.hedefTanim, 70) & vbNewLine
    If p.yama Then
        ozet = ozet & "Kapsam      : SECILI SUTUN YAMASI (uzerine yazar)" & vbNewLine
    ElseIf p.sadeceIsaretli Then
        ozet = ozet & "Kapsam      : SADECE SEC ISARETLI SUTUNLAR" & vbNewLine
    Else
        ozet = ozet & "Kapsam      : TUM ESLESMELER" & vbNewLine
    End If
    ozet = ozet & "Okunacak satirlar (aylik) : 2 - " & SayiMetni(p.sonSatir)
    If p.sonSatir < p.otomatikSon Then ozet = ozet & "   (dosyada " & SayiMetni(p.otomatikSon) & ". satira kadar veri var)"
    ozet = ozet & vbNewLine
    If p.filtreMetni <> "" Then ozet = ozet & "Filtre : " & Left$(p.filtreMetni, 120) & vbNewLine
    ozet = ozet & "Aktarilacak : " & SayiMetni(p.n) & " satir  x  " & p.adet & " sutun" & vbNewLine
    If p.enKucukBas = p.enBuyukBas Then
        ozet = ozet & "Ana dosyada yazilacak satirlar : " & SayiMetni(p.enKucukBas) & " - " & _
               SayiMetni(p.enKucukBas + p.n - 1) & vbNewLine
    Else
        ozet = ozet & "Ana dosyada baslangic satirlari : " & SayiMetni(p.enKucukBas) & " ile " & _
               SayiMetni(p.enBuyukBas) & " arasi (sutuna gore)" & vbNewLine
    End If
    If p.sabitMetni <> "" Then ozet = ozet & "Sabit degerler : " & Left$(p.sabitMetni, 120) & vbNewLine
    If p.mutAdet > 0 Then ozet = ozet & "Mutabakat   : " & Left$(MutabakatAdlari(p), 90) & _
                                    IIf(p.grupAdet > 0, "  x " & p.grupAdet & " " & p.kirilimAd, "") & vbNewLine
    If p.kalite >= 0 Then ozet = ozet & "Veri kalitesi : " & Format$(p.kalite, "0") & " / 100" & vbNewLine
    ozet = ozet & vbNewLine & "KONTROL NOKTALARI:" & vbNewLine & KarneSatiri(p.kDurum) & vbNewLine

    If p.dikkat <> "" Then
        kalanMetin = MSGBOX_SINIR - Len(ozet) - 90
        If kalanMetin < 120 Then kalanMetin = 120
        ozet = ozet & vbNewLine & "DIKKAT (" & p.dikkatSay & " madde, tamami Rapor'a yazilir):" & vbNewLine & _
               MaddeKisalt(p.dikkat, kalanMetin) & vbNewLine & _
               "Maddeleri okudun mu? Yine de aktarilsin mi?"
        cevap = MsgBox(ozet, vbYesNo + vbExclamation + vbDefaultButton2, "Onay - DIKKAT")
        karar = IIf(cevap = vbYes, "Okundu, devam edildi", "Aktarim iptal edildi")
        ' Denetim izi: her DIKKAT maddesi KISALTILMADAN, kendi satirinda
        MaddeleriRaporaYaz p, p.dikkat, "DIKKAT", karar, KD_DIKKAT
    Else
        ozet = ozet & vbNewLine & "Uyari yok. Aktarilsin mi?"
        cevap = MsgBox(ozet, vbYesNo + vbQuestion, "Onay")
        If cevap <> vbYes Then
            OlayYaz "IPTAL", p.k, p.partiNo, 0, "Onay ekraninda vazgecildi", "Aktarim yapilmadi", _
                    p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
        End If
    End If

    OnayAl = (cevap = vbYes)

End Function


' ============================================================
'  8. YAZMA  +  9. MUTABAKAT
' ============================================================
Private Sub Yaz(ByRef p As AktarimPlani)

    Dim j As Long, i As Long, n As Long
    Dim aktarimNo As Long, kaydirilamayan As Long
    Dim rng As Range
    Dim cikti() As Variant
    Dim formulOnb As Sozluk
    Dim yedekRef As String, imza As String, x As Variant, f As String
    Dim kaynakToplam As Currency, yazilanToplam As Double, deger As Double
    Dim hedefOnceki As Double, hedefYeni As Double, eskiAralik As Double
    Dim tutarVar As Boolean, tutarTuttu As Boolean
    Dim mutabakatMetni As String, tutarHarf As String
    Dim raporIlk As Long, raporSon As Long
    Dim u() As Long
    Dim kol As Variant, kolF As Variant, sabitHucre As Variant
    Dim korumaTip As String, kod As Long
    ' v1.1: coklu tutar (M1), adet (M2), katmanli (K13)
    Dim mKaynak() As Currency, mKaynakAdet() As Long, mYazilan() As Double, mYazAdet() As Long
    Dim m As Long, jm As Long, sonOnce As Long, sonSonra As Long, beklenenSon As Long, raporSatiri As Long
    Dim tol As Double
    Dim adetTuttu As Boolean, ekTuttu As Boolean, k13Tuttu As Boolean, hepsiTuttu As Boolean
    Dim adetMetni As String, ekMetni As String, k13Metni As String

    n = p.n
    u = p.uygun
    tol = Esik(ekMutabakatTolerans)

    ' --- Mutabakat: kaynak toplamlari (ana tutar = mutJ(1), varsa) ---
    tutarVar = (p.tutarJ > 0)
    If p.mutAdet > 0 Then
        ReDim mKaynak(1 To p.mutAdet)
        ReDim mKaynakAdet(1 To p.mutAdet)
        ReDim mYazilan(1 To p.mutAdet)
        ReDim mYazAdet(1 To p.mutAdet)
        For m = 1 To p.mutAdet
            kol = p.kv(p.mutJ(m))
            For i = 1 To n
                If SayiDegeri(kol(u(i), 1), deger) Then
                    mKaynak(m) = mKaynak(m) + CCur(deger)
                    mKaynakAdet(m) = mKaynakAdet(m) + 1
                End If
            Next i
        Next m
    End If
    If tutarVar Then
        kaynakToplam = mKaynak(1)
        hedefOnceki = KolonToplami(p.hedefWs, p.tutarHedefNo)
    End If

    ' --- Satir ekleme ONAYDAN SONRA (iptalde bos satir kalmasin) ---
    If p.eklenecekSatir > 0 Then
        p.hedefWs.Rows(p.boslukSon + 1).Resize(p.eklenecekSatir).Insert Shift:=xlDown
        GecmisSatirlariKaydir p.hedefWb.Name, CStr(p.hedefWs.Name), p.boslukSon + 1, p.eklenecekSatir
        OlayYaz "SATIR EKLEME", p.k, p.partiNo, 0, "Ana dosyaya " & p.eklenecekSatir & " satir eklendi (" & _
                (p.boslukSon + 1) & ". satirdan itibaren); Gecmis ve Rapor satir numaralari kaydirildi.", _
                "", p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If

    ' --- Mutabakat: once ---
    If tutarVar Then
        eskiAralik = AralikToplami(p.hedefWs, p.tutarHedefNo, p.baslangic(p.tutarJ), p.baslangic(p.tutarJ) + n - 1)
    End If
    sonOnce = SonDoluSatir(p.hedefWs)

    HesapDurdur
    aktarimNo = YeniAktarimNo()
    Set formulOnb = New Sozluk

    For j = 1 To p.adet
        DurumCubugu "PANEL " & p.k & ": yaziliyor " & j & "/" & p.adet & "  (" & p.hedefAdlari(j) & ")"
        Set rng = p.hedefWs.Cells(p.baslangic(j), p.hedefNo(j)).Resize(n, 1)

        ' Uzerine yazilacak dolu hucre varsa eski degerler yedeklenir
        yedekRef = ""
        If Application.WorksheetFunction.CountA(rng) > 0 Then
            yedekRef = AralikYedekle(rng, aktarimNo, p.hedefNo(j))
        End If

        ' Yeni satirlar ustteki satirin sayi bicimini alir
        BicimDevral rng

        ReDim cikti(1 To n, 1 To 1)
        korumaTip = p.korumaTip(j)
        If p.kaynakNo(j) > 0 Then kol = p.kv(j)
        If p.formul(j) Then kolF = p.kf(j)
        If p.kaynakNo(j) = 0 Then
            sabitHucre = HucreDegeri(p.sabitDeger(j), korumaTip)
            For i = 1 To n
                cikti(i, 1) = sabitHucre
            Next i
            rng.Value = cikti
        ElseIf p.formul(j) Then
            For i = 1 To n
                f = GuvenliMetin(kolF(u(i), 1))
                If Left$(f, 1) = "=" Then
                    cikti(i, 1) = FormulSatirKaydir(f, p.kaynakNo(j), p.hedefNo(j), formulOnb, kaydirilamayan)
                Else
                    x = kol(u(i), 1)
                    If VarType(x) = vbDate Then x = CDbl(x)
                    cikti(i, 1) = HucreDegeri(x, korumaTip)
                End If
            Next i
            rng.FormulaR1C1 = cikti
        Else
            For i = 1 To n
                x = kol(u(i), 1)
                If VarType(x) = vbString Then
                    If Len(x) > 0 Then
                        kod = AscW(x)
                        If Not ((kod >= 65 And kod <= 90) Or (kod >= 97 And kod <= 122) Or kod > 127) Then
                            If KorumaGerekir(CStr(x), korumaTip) Then x = "'" & x
                        End If
                    End If
                End If
                cikti(i, 1) = x
            Next i
            rng.Value = cikti
        End If

        imza = AralikImzasi(p.hedefWs, p.hedefNo(j), p.baslangic(j), p.baslangic(j) + n - 1, p.formul(j))
        LogYaz p.partiNo, aktarimNo, p.hedefWb.Name, CStr(p.hedefWs.Name), p.hedefNo(j), _
               p.kaynakAdlari(j), p.baslangic(j), p.baslangic(j) + n - 1, n, _
               GuvenliMetin(p.hedefWs.Cells(1, p.hedefNo(j)).Value), imza, yedekRef
        If imza = "?" Then Err.Raise vbObjectError + 1101, "Yaz", "Parmak izi hesaplanamadi; aktarim durduruldu."
    Next j

    HesapAc
    DurumCubugu ""

    If p.sadeceIsaretli Then SecIsaretleriniKaldir p.pnl, p.k

    ' --- Mutabakat: sonra ---
    ' (M2) adet: son dolu satir beklenen yerde mi; (M1) her mutabakat
    ' sutununda kaynak = yazilan (toplam ve sayisal hucre adedi); ana
    ' tutarda ayrica hedef sutun degisimi = yazilan - uzerine yazilan.
    sonSonra = SonDoluSatir(p.hedefWs)
    beklenenSon = sonOnce
    If p.enBuyukBas + n - 1 > beklenenSon Then beklenenSon = p.enBuyukBas + n - 1
    adetTuttu = (sonSonra = beklenenSon)
    For m = 1 To p.mutAdet
        jm = p.mutJ(m)
        mYazilan(m) = AralikSayisal(p.hedefWs, p.hedefNo(jm), p.baslangic(jm), p.baslangic(jm) + n - 1, mYazAdet(m))
        If mYazAdet(m) <> mKaynakAdet(m) Then adetTuttu = False
    Next m
    adetMetni = IIf(adetTuttu, "TUTUYOR", "FARK VAR") & ": " & SayiMetni(n) & " satir; son dolu satir " & SayiMetni(sonOnce) & _
                " -> " & SayiMetni(sonSonra) & " (beklenen " & SayiMetni(beklenenSon) & ")"
    For m = 1 To p.mutAdet
        adetMetni = adetMetni & "; " & p.hedefAdlari(p.mutJ(m)) & " sayisal hucre " & SayiMetni(mKaynakAdet(m)) & " = " & SayiMetni(mYazAdet(m))
    Next m

    ekTuttu = True
    For m = 1 To p.mutAdet
        If p.mutJ(m) <> p.tutarJ Then
            If Abs(mKaynak(m) - mYazilan(m)) >= tol Then
                ekTuttu = False
                ekMetni = ekMetni & p.hedefAdlari(p.mutJ(m)) & " " & TutarMetni(mKaynak(m)) & " / " & TutarMetni(mYazilan(m)) & " FARK; "
            Else
                ekMetni = ekMetni & p.hedefAdlari(p.mutJ(m)) & " " & TutarMetni(mKaynak(m)) & " TUTUYOR; "
            End If
        End If
    Next m

    tutarHarf = ""
    mutabakatMetni = "(tutar sutunu secilmemis)"
    tutarTuttu = True
    If tutarVar Then
        tutarHarf = SutunHarfi(p.tutarHedefNo)
        yazilanToplam = mYazilan(1)
        hedefYeni = KolonToplami(p.hedefWs, p.tutarHedefNo)
        tutarTuttu = (Abs(kaynakToplam - yazilanToplam) < tol) And _
                     (Abs((hedefYeni - hedefOnceki) - (yazilanToplam - eskiAralik)) < tol)
    End If
    hepsiTuttu = tutarTuttu And ekTuttu And adetTuttu
    If tutarVar Then
        mutabakatMetni = IIf(hepsiTuttu, "TUTUYOR", "FARK VAR") & ": kaynaktan okunan " & TutarMetni(kaynakToplam) & _
                         ", hedefe yazilan " & TutarMetni(yazilanToplam) & _
                         " (fark " & TutarMetni(kaynakToplam - yazilanToplam) & "); hedef sutun toplami " & _
                         TutarMetni(hedefOnceki) & " -> " & TutarMetni(hedefYeni) & "; tolerans " & Format$(tol, "0.000")
    ElseIf p.mutAdet > 0 Then
        mutabakatMetni = IIf(hepsiTuttu, "TUTUYOR", "FARK VAR") & ": ana tutar secilmemis; ek sutunlar"
    End If
    If ekMetni <> "" Then mutabakatMetni = mutabakatMetni & "; EK: " & ekMetni
    If Not adetTuttu Then mutabakatMetni = mutabakatMetni & "; ADET: " & adetMetni

    ' --- K13: katmanli mutabakat (geri okuma) ---
    k13Tuttu = K13Dogrula(p, k13Metni)
    p.mMutabakat = IIf(hepsiTuttu And k13Tuttu, 1, 0)
    KaliteHesapla p

    If tutarVar Then
        g_MutabakatOzet = g_MutabakatOzet & "   PANEL " & p.k & ": " & IIf(hepsiTuttu, "TUTUYOR (", "FARK VAR (") & _
                          TutarMetni(kaynakToplam) & " / " & TutarMetni(yazilanToplam) & ")" & vbNewLine
    Else
        g_MutabakatOzet = g_MutabakatOzet & "   PANEL " & p.k & ": tutar sutunu secilmemis" & vbNewLine
    End If

    If p.tutarJ > 0 Then
        raporIlk = p.baslangic(p.tutarJ)
    Else
        raporIlk = p.enKucukBas
    End If
    raporSon = raporIlk + n - 1

    raporSatiri = RaporYaz(p.partiNo, aktarimNo, p.k, p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name), _
                           p.sonSatir, n, p.adet, raporIlk, raporSon, p.filtreMetni, p.sabitMetni, mutabakatMetni, _
                           tutarVar, kaynakToplam, tutarHarf, yazilanToplam, hedefOnceki, hedefYeni)
    RaporEkAlanlar raporSatiri, p.soyYol, p.soyBoyut, p.soyZaman, p.kaynakIzi, p.icerikIzi, SureMs(p.basZaman), _
                   adetMetni, p.kalite, p.metrik & IIf(p.kilitNotu <> "", ";kilit=" & p.kilitNotu, "")

    ' --- K7 / K13 son hukmu ve karne ---
    If p.mutAdet > 0 And Not hepsiTuttu Then KSeviye p, 7, KD_ENGEL
    If Not adetTuttu Then KSeviye p, 7, KD_ENGEL
    If Not k13Tuttu Then
        KSeviye p, 13, KD_ENGEL
        OlayYaz "UYARI", p.k, p.partiNo, aktarimNo, "Katmanli mutabakat (" & p.kirilimAd & ") tutmadi: " & k13Metni, _
                "Aktarim yapildi, kontrol gerekiyor", p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, _
                CStr(p.hedefWs.Name), "K13 " & KD_ENGEL
    ElseIf k13Metni <> "" Then
        OlayYaz "MUTABAKAT", p.k, p.partiNo, aktarimNo, "Katmanli mutabakat (" & p.kirilimAd & "): " & k13Metni, _
                IIf(Left$(k13Metni, 5) = "sutun", "Yapilmadi", "TUTUYOR"), p.kaynakWb.Name, CStr(p.kaynakWs.Name), _
                p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If
    KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR: aktar#im no ") & aktarimNo & "  " & ChrW(183) & "  " & _
             IIf(tutarVar, IIf(hepsiTuttu, "FARK 0,00", "FARK " & TutarMetni(kaynakToplam - yazilanToplam)), _
                 IIf(p.mutAdet > 0, IIf(hepsiTuttu, "FARK 0,00", "FARK VAR"), TR("tutar se#cilmemi#s"))) & _
             "  (" & Format$(Now, "dd.mm hh:nn") & ")", p.kalite

    ' --- Durum bellegi: K12 sema, K14 izi tuketildi ---
    If Not p.sadeceIsaretli Then BellekYaz "SEMA|" & p.k, SemaMetni(p.kaynakWs)
    BellekYaz "OKIZ|" & p.k, ""

    If kaydirilamayan > 0 Then
        OlayYaz "UYARI", p.k, p.partiNo, aktarimNo, kaydirilamayan & " formul deseni satir kaydirmasi yapilamadan " & _
                "(255 karakter siniri) oldugu gibi yazildi", "Kontrol gerekiyor", _
                p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If

    If p.mutAdet > 0 And Not hepsiTuttu And Not tutarVar Then
        OlayYaz "UYARI", p.k, p.partiNo, aktarimNo, "Mutabakat tutmadi: " & mutabakatMetni, _
                "Aktarim yapildi, kontrol gerekiyor", p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If
    If tutarVar And Not hepsiTuttu Then
        OlayYaz "UYARI", p.k, p.partiNo, aktarimNo, "Tutar mutabakati tutmadi: " & mutabakatMetni, _
                "Aktarim yapildi, kontrol gerekiyor", p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
        MsgBox "TUTAR MUTABAKATI TUTMADI  -  PANEL " & p.k & vbNewLine & vbNewLine & _
               "Kaynaktan okunan  (" & p.kaynakAdlari(p.tutarJ) & ") : " & TutarMetni(kaynakToplam) & vbNewLine & _
               "Hedefe yazilan    (" & p.tutarAd & ", " & raporIlk & "-" & raporSon & ") : " & TutarMetni(yazilanToplam) & vbNewLine & _
               "FARK              : " & TutarMetni(kaynakToplam - yazilanToplam) & vbNewLine & vbNewLine & _
               "Hedef sutun toplami once : " & TutarMetni(hedefOnceki) & vbNewLine & _
               "Hedef sutun toplami sonra: " & TutarMetni(hedefYeni) & vbNewLine & _
               "Beklenen artis           : " & TutarMetni(yazilanToplam - eskiAralik) & vbNewLine & vbNewLine & _
               "Olasi nedenler: metin olarak yazilmis sayilar, formul," & vbNewLine & _
               "ana dosyada o satirlarda onceden duran veri." & vbNewLine & vbNewLine & _
               "Aktarim YAPILDI. Satirlari kontrol et; gerekirse 'Geri Al'.", vbExclamation, "Mutabakat"
    End If

    If p.toplam = 1 Then
        Bildir BS_OK, "Aktarim", "PANEL " & p.k & IIf(p.yama, " YAMASI", "") & " aktarildi." & vbNewLine & vbNewLine & _
               "Parti no        : " & p.partiNo & vbNewLine & _
               "Aktarim no      : " & aktarimNo & vbNewLine & _
               "Aktarilan satir : " & SayiMetni(n) & "   (ana dosya " & SayiMetni(p.enKucukBas) & " - " & _
               SayiMetni(p.enBuyukBas + n - 1) & ")" & vbNewLine & _
               "Aktarilan sutun : " & p.adet & vbNewLine & _
               "Ana dosya       : " & Left$(p.hedefTanim, 70) & vbNewLine & vbNewLine & _
               "TUTAR MUTABAKATI" & vbNewLine & _
               IIf(tutarVar, _
                   "   Kaynaktan okunan   : " & TutarMetni(kaynakToplam) & vbNewLine & _
                   "   Hedefe yazilan     : " & TutarMetni(yazilanToplam) & vbNewLine & _
                   "   Hedef onceki toplam: " & TutarMetni(hedefOnceki) & vbNewLine & _
                   "   Hedef yeni toplam  : " & TutarMetni(hedefYeni) & vbNewLine & _
                   "   FARK               : " & TutarMetni(kaynakToplam - yazilanToplam) & _
                   IIf(tutarTuttu, "   (TUTUYOR)", "   (FARK VAR)") & vbNewLine, _
                   "   (tutar sutunu secilmemis)" & vbNewLine) & vbNewLine & _
               "HATIRLATMA: Ana dosyayi SIMDI kaydet, sonra bu paneli de kaydet." & vbNewLine & _
               "Ikisi birlikte kaydedilmezse Gecmis kaydi ile dosyanin" & vbNewLine & _
               "gercek durumu birbirinden ayrisir."
    End If

End Sub


' Yeni yazilacak aralik "Genel" bicimdeyse ustundeki satirin sayi
' bicimini alir (tutar 1.234,56; tarih gg.aa.yyyy ...). Metin ("@")
' bicimi devralinmaz: sayilar metne donmesin.
Private Sub BicimDevral(ByVal rng As Range)

    Dim ust As String, mevcut As String

    If rng.Row <= 2 Then Exit Sub
    On Error Resume Next
    ust = rng.Cells(1, 1).Offset(-1, 0).NumberFormat
    mevcut = rng.Cells(1, 1).NumberFormat
    If mevcut = "General" And ust <> "General" And ust <> "@" And ust <> "" Then
        rng.NumberFormat = ust
    End If
    On Error GoTo 0

End Sub


Public Function AralikToplami(ByVal ws As Worksheet, ByVal sutun As Long, ByVal ilk As Long, ByVal son As Long) As Double
    Dim adet As Long
    AralikToplami = AralikSayisal(ws, sutun, ilk, son, adet)
End Function


' Araligin Currency toplami ve sayisal hucre adedi. Excel SUM anlami:
' metin olarak duran sayi toplanmaz. Hata hucresi acik hatadir (M4).
Public Function AralikSayisal(ByVal ws As Worksheet, ByVal sutun As Long, ByVal ilk As Long, ByVal son As Long, _
                             ByRef adet As Long) As Double
    Dim v As Variant, i As Long, toplam As Currency, deger As Double
    adet = 0
    If ilk < 1 Or son < ilk Or sutun < 1 Then Exit Function
    v = SutunOku(ws, sutun, ilk, son)
    For i = 1 To UBound(v, 1)
        If IsError(v(i, 1)) Then
            Err.Raise vbObjectError + 1102, "AralikSayisal", "Hedef tutar araliginda Excel hata hucresi var."
        End If
        If VarType(v(i, 1)) <> vbString Then
            If SayiDegeri(v(i, 1), deger) Then
                toplam = toplam + CCur(deger)
                adet = adet + 1
            End If
        End If
    Next i
    AralikSayisal = CDbl(toplam)
End Function

Public Function TutarHataSayisi(ByVal ws As Worksheet, ByVal sutun As Long) As Long
    Dim v As Variant, i As Long, son As Long
    son = ws.Cells(ws.Rows.Count, sutun).End(xlUp).Row
    If son < 2 Then Exit Function
    v = SutunOku(ws, sutun, 2, son)
    For i = 1 To UBound(v, 1)
        If IsError(v(i, 1)) Then TutarHataSayisi = TutarHataSayisi + 1
    Next i
End Function


Public Function KolonToplami(ByVal ws As Worksheet, ByVal sutun As Long) As Double
    Dim son As Long
    son = ws.Cells(ws.Rows.Count, sutun).End(xlUp).Row
    If son < 2 Then
        KolonToplami = 0
    Else
        KolonToplami = AralikToplami(ws, sutun, 2, son)
    End If
End Function


' ============================================================
'  v1.1  KONTROL NOKTALARI K9..K14, COKLU MUTABAKAT, SOY, METRIK
'
'  Aktueryal varsayimlar (Kilavuz ve README'de de yazilidir):
'   - K9 bilesik anahtar = anahtar + tutar (2 hane) + hasar tarihi.
'     Ayni hasarin ayni gun ayni tutarli iki odemesi tekrar sayilir
'     (DIKKAT); TUM eslenen sutunlari ayni satir "birebir kopya"dir.
'   - K9 ana dosya ornegi anahtar + tutar + sabit degerlerle aranir.
'     Muallak gibi anlik goruntulerde donem sabiti yoksa eslesme
'     beklenir; bu yuzden yalnizca DIKKAT'tir, ENGEL degildir.
'   - K10 donem sonu bossa gecen ay sonu varsayilir (aylik kapanis).
'   - K11 gecmisi: ayni ana dosya sayfasina AYNI panelden yapilan ETKIN
'     aktarimlarin son 12 takvim ayi (icinde bulunulan ay haric).
'   - K13 kaynakta metin olarak duran sayilar toplanir, ana dosyada
'     metin kalan hucre toplanmaz (K7 ile ayni anlam): fark ENGEL'dir.
' ============================================================

Private Function SureMs(ByVal bas As Double) As Double
    Dim t As Double
    t = Timer - bas
    If t < 0 Then t = t + 86400
    SureMs = Round(t * 1000, 0)
End Function


' Ana dosya sutununu dolduran (kaynakli) eslesme; yoksa 0
Private Function EslesmeBul(ByRef p As AktarimPlani, ByVal hedefNo As Long) As Long
    Dim j As Long
    If hedefNo = 0 Then Exit Function
    For j = 1 To p.gercekAdet
        If p.hedefNo(j) = hedefNo And p.kaynakNo(j) > 0 Then
            EslesmeBul = j
            Exit Function
        End If
    Next j
End Function


' PlanHazirla'nin v1.1 eki: anahtar, kirilim, tarih sutunlari, donem
' sonu, ana dosya kilidi.
Private Sub PlanEki(ByRef p As AktarimPlani)

    Dim sahip As String, zaman As String

    p.anahtarHedefNo = SutunNo(Ayar(p.pnl, p.k, SATIR_ANAHTAR).Value)
    If p.anahtarHedefNo > 0 Then
        p.anahtarAd = SutunAdi(Ayar(p.pnl, p.k, SATIR_ANAHTAR).Value)
        If BaslikDogrula(p.hedefWs, p.anahtarHedefNo, p.anahtarAd) >= bdYok Then
            KEkle p, 9, False, "ANAHTAR SUTUNU '" & p.anahtarAd & "' ana dosyada bulunamadi; tekillik kontrolu yapilamadi."
            p.anahtarHedefNo = 0
        End If
    End If
    p.anahtarJ = EslesmeBul(p, p.anahtarHedefNo)

    p.kirilimHedefNo = SutunNo(Ayar(p.pnl, p.k, SATIR_KIRILIM).Value)
    If p.kirilimHedefNo > 0 Then
        p.kirilimAd = SutunAdi(Ayar(p.pnl, p.k, SATIR_KIRILIM).Value)
        If BaslikDogrula(p.hedefWs, p.kirilimHedefNo, p.kirilimAd) >= bdYok Then
            KEkle p, 13, False, "KIRILIM SUTUNU '" & p.kirilimAd & "' ana dosyada bulunamadi; katmanli mutabakat yapilamadi."
            p.kirilimHedefNo = 0
        End If
    End If
    p.kirilimJ = EslesmeBul(p, p.kirilimHedefNo)

    p.hasarJ = EslesmeBul(p, SutunNo(Ayar(p.pnl, p.k, SATIR_HASAR).Value))
    p.ihbarJ = EslesmeBul(p, SutunNo(Ayar(p.pnl, p.k, SATIR_IHBAR).Value))
    DonemBelirle p

    If KilitBaskasinda(p.hedefWb, sahip, zaman) Then
        KEkle p, 1, Not p.kuru, "ANA DOSYA KILITLI: " & sahip & " (" & zaman & ") bu ana dosyaya yaziyor olabilir. " & _
                  "Ayni anda iki aktarim satirlari karistirir. Bitmesini bekle; kilit " & CLng(Esik(ekKilitDakika)) & _
                  " dakikadan eskiyse bayat sayilir."
    End If

End Sub


Private Sub DonemBelirle(ByRef p As AktarimPlani)

    Dim x As Variant
    Dim varsayilan As Date

    varsayilan = DateSerial(Year(Date), Month(Date), 0)
    x = Ayar(p.pnl, p.k, SATIR_DONEM).Value
    p.donemNotu = ""
    If VarType(x) = vbDate Then
        p.donemSonu = CDate(x)
    ElseIf VarType(x) = vbDouble Then
        If x > 0 And x < 2958466 Then
            p.donemSonu = CDate(x)
        Else
            p.donemSonu = varsayilan
            p.donemNotu = "donem sonu okunamadi"
        End If
    ElseIf VarType(x) = vbString Then
        If Trim$(CStr(x)) = "" Then
            p.donemSonu = varsayilan
            p.donemNotu = "donem sonu bos, gecen ay sonu varsayildi"
        ElseIf IsDate(x) Then
            p.donemSonu = CDate(x)
        Else
            p.donemSonu = varsayilan
            p.donemNotu = "donem sonu '" & CStr(x) & "' okunamadi, gecen ay sonu varsayildi"
        End If
    Else
        p.donemSonu = varsayilan
        p.donemNotu = "donem sonu bos, gecen ay sonu varsayildi"
    End If

End Sub


' M1: ana tutar + sayisal (kaynak ve ana dosyada SAYI) eslenen sutunlar
Private Sub MutabakatSutunlariSec(ByRef p As AktarimPlani)

    Dim j As Long, enFazla As Long

    enFazla = CLng(Esik(ekMutabakatSutun))
    If enFazla < 1 Then enFazla = 1
    ReDim p.mutJ(1 To enFazla)
    p.mutAdet = 0
    If p.tutarJ > 0 Then
        p.mutAdet = 1
        p.mutJ(1) = p.tutarJ
    End If
    For j = 1 To p.gercekAdet
        If p.mutAdet >= enFazla Then Exit For
        If j <> p.tutarJ And j <> p.anahtarJ And Not p.formul(j) And p.kaynakNo(j) > 0 Then
            If p.kTip(j) = "SAYI" And p.hTip(j) = "SAYI" Then
                p.mutAdet = p.mutAdet + 1
                p.mutJ(p.mutAdet) = j
            End If
        End If
    Next j

End Sub


Private Function MutabakatAdlari(ByRef p As AktarimPlani) As String
    Dim m As Long
    For m = 1 To p.mutAdet
        MutabakatAdlari = MutabakatAdlari & p.hedefAdlari(p.mutJ(m)) & IIf(m < p.mutAdet, ", ", "")
    Next m
End Function


' ------------------------------------------------------------
'  K9  ANAHTAR TEKILLIGI
' ------------------------------------------------------------
Private Function BilesikAnahtar(ByVal a As Variant, ByVal t As Variant, ByVal h As Variant) As String

    Dim s As String
    Dim d As Double
    Dim c As Currency

    s = Trim$(GuvenliMetin(a))
    If Not IsEmpty(t) Then
        If ParaDegeri(t, c) Then
            s = s & "|" & Format$(Round(c, 2), "0.00")
        Else
            s = s & "|" & Trim$(GuvenliMetin(t))
        End If
    End If
    If Not IsEmpty(h) Then
        If TarihSayisi(h, d) Then
            s = s & "|" & CStr(Int(d))
        Else
            s = s & "|" & Trim$(GuvenliMetin(h))
        End If
    End If
    BilesikAnahtar = s

End Function


' Iki kaynak satiri TUM eslenen sutunlarda ayni mi?
Private Function SatirlarAyni(ByRef p As AktarimPlani, ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim j As Long
    SatirlarAyni = False
    For j = 1 To p.gercekAdet
        If GuvenliMetin(p.kv(j)(r1, 1)) <> GuvenliMetin(p.kv(j)(r2, 1)) Then Exit Function
    Next j
    SatirlarAyni = True
End Function


Private Sub K9Kontrol(ByRef p As AktarimPlani)

    Dim u() As Long
    Dim i As Long, r As Long, ilk As Long
    Dim kolA As Variant, kolT As Variant, kolH As Variant, t As Variant, h As Variant
    Dim anahtarlar As Collection
    Dim dup As Long, kopya As Long, bosA As Long, dolu As Long
    Dim ornDup As String, ornKopya As String, a As String

    If p.anahtarHedefNo = 0 Then
        KYok p, 9
        Exit Sub
    End If
    If p.anahtarJ = 0 Then
        KEkle p, 9, False, "ANAHTAR SUTUNU (" & p.anahtarAd & ") bu panelde hicbir aylik sutunla eslesmemis; tekillik kontrolu yapilamadi."
        Exit Sub
    End If

    u = p.uygun
    kolA = p.kv(p.anahtarJ)
    If p.tutarJ > 0 Then kolT = p.kv(p.tutarJ)
    If p.hasarJ > 0 Then kolH = p.kv(p.hasarJ)
    Set anahtarlar = New Collection

    For i = 1 To p.n
        r = u(i)
        If Trim$(GuvenliMetin(kolA(r, 1))) = "" Then
            bosA = bosA + 1
        Else
            t = Empty
            h = Empty
            If p.tutarJ > 0 Then t = kolT(r, 1)
            If p.hasarJ > 0 Then h = kolH(r, 1)
            a = "k" & BilesikAnahtar(kolA(r, 1), t, h)
            On Error Resume Next
            anahtarlar.Add i, a
            If Err.Number <> 0 Then
                Err.Clear
                ilk = 0
                ilk = anahtarlar.Item(a)
                Err.Clear
                On Error GoTo 0
                dup = dup + 1
                If dup <= KONTROL_ORNEK Then ornDup = ornDup & (r + 1) & " "
                If ilk > 0 Then
                    If SatirlarAyni(p, u(ilk), r) Then
                        kopya = kopya + 1
                        If kopya <= KONTROL_ORNEK Then ornKopya = ornKopya & (u(ilk) + 1) & "=" & (r + 1) & " "
                    End If
                End If
            End If
            On Error GoTo 0
        End If
    Next i

    dolu = p.n - bosA
    If dolu > 0 Then p.mTekillik = (dolu - dup) / dolu
    If bosA > 0 Then
        KEkle p, 9, False, "ANAHTAR BOS: " & p.kaynakAdlari(p.anahtarJ) & " sutununda " & SayiMetni(bosA) & " satir bos."
    End If
    If kopya > 0 Then
        KEkle p, 9, (Esik(ekK9KopyaEngel) >= 1), "BIREBIR KOPYA SATIR: " & SayiMetni(kopya) & _
              " satir, eslenen TUM sutunlarda daha onceki bir satirla ayni (aylik satir " & Trim$(ornKopya) & _
              IIf(kopya > KONTROL_ORNEK, " ...", "") & "). Kaynagi duzelt; mesru tekrarlar icin Ayarlar: K9_KOPYA_ENGEL = 0."
    End If
    If dup - kopya > 0 Then
        KEkle p, 9, False, "ANAHTAR TEKRARI: " & SayiMetni(dup - kopya) & " satirda " & p.anahtarAd & _
              IIf(p.tutarJ > 0, " + tutar", "") & IIf(p.hasarJ > 0, " + hasar tarihi", "") & _
              " daha once gecmis (aylik satir " & Trim$(ornDup) & IIf(dup > KONTROL_ORNEK, " ...", "") & _
              "). Kismi odeme mi, cift kayit mi kontrol et."
    End If

    If p.tutarJ > 0 Then K9AnaDosyaOrnegi p

End Sub


' Tabakali ornegin anahtar + tutar + sabitleri ana dosyanin kuyrugunda var mi?
' Tutara gore ikili arama on suzgeci: ana dosya satirlarinin cogu hic
' metin birlestirmeye girmez.
Private Sub K9AnaDosyaOrnegi(ByRef p As AktarimPlani)

    Dim u() As Long
    Dim ornekSay As Long, i As Long, j As Long, t As Long, r As Long, son As Long, ilk As Long, pencere As Long
    Dim tutarlar() As Currency, gecici As Currency
    Dim ornek As Collection, bulunan As Collection
    Dim sabitMetin As String, a As String
    Dim kolA As Variant, kolT As Variant, vA As Variant, vT As Variant
    Dim vS() As Variant
    Dim sabitSay As Long, eslesen As Long
    Dim d As Double, c As Currency

    u = p.uygun
    ornekSay = CLng(Esik(ekK9Ornek))
    If ornekSay < 1 Then Exit Sub
    If ornekSay > p.n Then ornekSay = p.n

    For j = p.gercekAdet + 1 To p.adet
        sabitMetin = sabitMetin & "|" & Trim$(GuvenliMetin(p.sabitDeger(j)))
    Next j
    sabitSay = p.adet - p.gercekAdet

    kolA = p.kv(p.anahtarJ)
    kolT = p.kv(p.tutarJ)
    Set ornek = New Collection
    ReDim tutarlar(1 To ornekSay)
    For t = 1 To ornekSay
        If ornekSay = 1 Then
            i = 1
        Else
            i = 1 + Int((t - 1) * (p.n - 1) / (ornekSay - 1))
        End If
        r = u(i)
        If ParaDegeri(kolT(r, 1), c) Then
            c = Round(c, 2)
        Else
            c = 0
        End If
        tutarlar(t) = c
        a = "k" & Trim$(GuvenliMetin(kolA(r, 1))) & "|" & Format$(c, "0.00") & sabitMetin
        On Error Resume Next
        ornek.Add t, a
        Err.Clear
        On Error GoTo 0
    Next t

    ' tutarlari sirala (ornek kucuk: araya sokma)
    For i = 2 To ornekSay
        gecici = tutarlar(i)
        j = i - 1
        Do While j >= 1
            If tutarlar(j) <= gecici Then Exit Do
            tutarlar(j + 1) = tutarlar(j)
            j = j - 1
        Loop
        tutarlar(j + 1) = gecici
    Next i

    son = p.hedefWs.Cells(p.hedefWs.Rows.Count, p.tutarHedefNo).End(xlUp).Row
    If son < 2 Then Exit Sub
    pencere = CLng(Esik(ekK9Pencere))
    ilk = son - pencere + 1
    If ilk < 2 Then ilk = 2
    vA = SutunOku(p.hedefWs, p.anahtarHedefNo, ilk, son)
    vT = SutunOku(p.hedefWs, p.tutarHedefNo, ilk, son)
    If sabitSay > 0 Then
        ReDim vS(1 To sabitSay)
        For j = 1 To sabitSay
            vS(j) = SutunOku(p.hedefWs, p.hedefNo(p.gercekAdet + j), ilk, son)
        Next j
    End If

    Set bulunan = New Collection
    For i = 1 To UBound(vT, 1)
        If ParaDegeri(vT(i, 1), c) Then
            c = Round(c, 2)
            If TutarVarMi(tutarlar, ornekSay, c) Then
                a = "k" & Trim$(GuvenliMetin(vA(i, 1))) & "|" & Format$(c, "0.00")
                For j = 1 To sabitSay
                    a = a & "|" & Trim$(GuvenliMetin(vS(j)(i, 1)))
                Next j
                On Error Resume Next
                t = 0
                t = ornek.Item(a)
                If Err.Number = 0 And t > 0 Then
                    bulunan.Add t, "b" & t
                    If Err.Number = 0 Then eslesen = eslesen + 1
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i

    If eslesen > 0 Then
        KEkle p, 9, False, "ANA DOSYADA AYNI KAYIT: tabakali ornekteki " & ornekSay & " satirin " & eslesen & "'i (%" & _
              Format$(eslesen / ornekSay * 100, "0") & ") ana dosyanin son " & SayiMetni(son - ilk + 1) & _
              " satirinda anahtar + tutar" & IIf(sabitSay > 0, " + sabit degerler", "") & " ile birebir var. " & _
              "Odenen dosyasinda mukerrer aktarim isaretidir; donem sabiti olmayan anlik goruntulerde (muallak) beklenebilir."
    End If

End Sub


Private Function TutarVarMi(tutarlar() As Currency, ByVal n As Long, ByVal c As Currency) As Boolean
    Dim alt As Long, ust As Long, orta As Long
    alt = 1
    ust = n
    Do While alt <= ust
        orta = (alt + ust) \ 2
        If tutarlar(orta) = c Then
            TutarVarMi = True
            Exit Function
        ElseIf tutarlar(orta) < c Then
            alt = orta + 1
        Else
            ust = orta - 1
        End If
    Loop
End Function


' ------------------------------------------------------------
'  K10  DONEM TUTARLILIGI  (M8 makul tarih araligi)
' ------------------------------------------------------------
Private Sub K10Kontrol(ByRef p As AktarimPlani)

    Dim u() As Long
    Dim i As Long, r As Long, s As Long, j As Long
    Dim kol As Variant
    Dim d As Double, alt As Double, ust As Double, donem As Double
    Dim gelecek As Long, disari As Long, gecerli As Long, dolu As Long
    Dim ornG As String, ornD As String, adlar As String, metin As String

    If p.hasarJ = 0 And p.ihbarJ = 0 Then
        KYok p, 10
        Exit Sub
    End If
    alt = CDbl(DateSerial(CLng(Esik(ekTarihAltYil)), 1, 1))
    ust = CDbl(DateAdd("m", CLng(Esik(ekTarihUstAy)), p.donemSonu))
    donem = CDbl(p.donemSonu) + 1

    u = p.uygun
    For s = 1 To 2
        If s = 1 Then j = p.hasarJ Else j = p.ihbarJ
        If j > 0 Then
            adlar = adlar & p.kaynakAdlari(j) & " "
            kol = p.kv(j)
            For i = 1 To p.n
                r = u(i)
                If Not IsEmpty(kol(r, 1)) Then
                    If Trim$(GuvenliMetin(kol(r, 1))) <> "" Then dolu = dolu + 1
                    If TarihSayisi(kol(r, 1), d) Then
                        If d < alt Or d > ust Then
                            disari = disari + 1
                            If disari <= KONTROL_ORNEK Then ornD = ornD & (r + 1) & " "
                        Else
                            gecerli = gecerli + 1
                            If d >= donem Then
                                gelecek = gelecek + 1
                                If gelecek <= KONTROL_ORNEK Then ornG = ornG & (r + 1) & " "
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next s

    If dolu > 0 Then p.mGecerlilik = OrtalamaEkle(p.mGecerlilik, gecerli / dolu)
    If gelecek > 0 Then metin = metin & "donem sonundan (" & Format$(p.donemSonu, "dd.mm.yyyy") & ") SONRA: " & _
                                SayiMetni(gelecek) & " hucre (aylik satir " & Trim$(ornG) & IIf(gelecek > KONTROL_ORNEK, " ...", "") & "); "
    If disari > 0 Then metin = metin & "makul aralik (" & Format$(alt, "dd.mm.yyyy") & " - " & Format$(ust, "dd.mm.yyyy") & _
                               ") DISINDA: " & SayiMetni(disari) & " hucre (aylik satir " & Trim$(ornD) & _
                               IIf(disari > KONTROL_ORNEK, " ...", "") & "; tutar gibi bir sayi tarih sutununa girmis olabilir); "
    If metin <> "" Then
        If p.donemNotu <> "" Then metin = metin & "(" & p.donemNotu & ")"
        KEkle p, 10, False, "DONEM TUTARLILIGI (" & Trim$(adlar) & "): " & metin
    End If
    KTamam p, 10

End Sub


Private Function OrtalamaEkle(ByVal mevcut As Double, ByVal yeni As Double) As Double
    If mevcut < 0 Then
        OrtalamaEkle = yeni
    Else
        OrtalamaEkle = (mevcut + yeni) / 2
    End If
End Function


' ------------------------------------------------------------
'  K11  MAKULLUK VE AYKIRI DEGER
' ------------------------------------------------------------
Private Sub K11Kontrol(ByRef p As AktarimPlani)

    Dim aylikN() As Double, aylikT() As Double
    Dim ayVar() As Boolean, tutarVar() As Boolean
    Dim ay As Long, nAy As Long, nAyT As Long, i As Long, r As Long, rakam As Long
    Dim ort As Double, sapma As Double, ortT As Double, sapmaT As Double
    Dim z As Double, zEsik As Double, buyukEsik As Double, mad As Double, beklenen As Double
    Dim toplam As Double, d As Double, f As Double, buyukToplam As Double
    Dim u() As Long
    Dim kol As Variant
    Dim ilkHane(1 To 9) As Long, benfordN As Long, enKotu As Long
    Dim metin As String, ornB As String, uygulandi As Boolean

    zEsik = Esik(ekZSkor)
    AylikGecmis p.hedefWb.Name, CStr(p.hedefWs.Name), p.k, Date, aylikN, aylikT, ayVar, tutarVar

    u = p.uygun
    If p.tutarJ > 0 Then kol = p.kv(p.tutarJ)
    buyukEsik = Esik(ekBuyukHasar)
    For i = 1 To p.n
        If p.tutarJ = 0 Then Exit For
        r = u(i)
        If SayiDegeri(kol(r, 1), d) Then
            toplam = toplam + d
            If buyukEsik > 0 And Abs(d) >= buyukEsik Then
                p.buyukSay = p.buyukSay + 1
                buyukToplam = buyukToplam + d
                If p.buyukSay <= 3 Then ornB = ornB & (r + 1) & " (" & TutarMetni(d) & ") "
            End If
            If d >= 10 Then
                f = d / 10 ^ Int(Log(d) / Log(10#))
                If f >= 10 Then f = f / 10
                If f < 1 Then f = f * 10
                rakam = Int(f)
                If rakam >= 1 And rakam <= 9 Then
                    ilkHane(rakam) = ilkHane(rakam) + 1
                    benfordN = benfordN + 1
                End If
            End If
        End If
    Next i

    ' --- z-skorlari (gecmis yeterliyse) ---
    For ay = 1 To 12
        If ayVar(ay) Then
            nAy = nAy + 1
            ort = ort + aylikN(ay)
        End If
        If tutarVar(ay) Then
            nAyT = nAyT + 1
            ortT = ortT + aylikT(ay)
        End If
    Next ay
    p.zN = Empty
    p.zT = Empty
    If nAy >= Esik(ekZMinAy) And nAy >= 2 Then
        uygulandi = True
        ort = ort / nAy
        For ay = 1 To 12
            If ayVar(ay) Then sapma = sapma + (aylikN(ay) - ort) ^ 2
        Next ay
        sapma = Sqr(sapma / (nAy - 1))
        If sapma > 0 Then
            z = (p.n - ort) / sapma
            p.zN = Round(z, 2)
            If Abs(z) > zEsik Then metin = metin & "satir sayisi " & SayiMetni(p.n) & ", son " & nAy & " ay ortalamasi " & _
                                         SayiMetni(ort) & " (z = " & Format$(z, "0.0") & "); "
        End If
    End If
    If p.tutarJ > 0 And nAyT >= Esik(ekZMinAy) And nAyT >= 2 Then
        uygulandi = True
        ortT = ortT / nAyT
        For ay = 1 To 12
            If tutarVar(ay) Then sapmaT = sapmaT + (aylikT(ay) - ortT) ^ 2
        Next ay
        sapmaT = Sqr(sapmaT / (nAyT - 1))
        If sapmaT > 0 Then
            z = (toplam - ortT) / sapmaT
            p.zT = Round(z, 2)
            If Abs(z) > zEsik Then metin = metin & "toplam tutar " & TutarMetni(toplam) & ", son " & nAyT & _
                                         " ay ortalamasi " & TutarMetni(ortT) & " (z = " & Format$(z, "0.0") & "); "
        End If
    End If
    p.mMakulluk = -1
    If Not IsEmpty(p.zN) Or Not IsEmpty(p.zT) Then
        z = 0
        If Not IsEmpty(p.zN) Then z = Abs(p.zN)
        If Not IsEmpty(p.zT) Then
            If Abs(p.zT) > z Then z = Abs(p.zT)
        End If
        If z <= zEsik Then
            p.mMakulluk = 1
        Else
            p.mMakulluk = 1 - (z - zEsik) / zEsik
            If p.mMakulluk < 0 Then p.mMakulluk = 0
        End If
    End If
    If metin <> "" Then KEkle p, 11, False, "MAKULLUK (z-skor, esik " & Format$(zEsik, "0.0") & "): " & metin

    ' --- buyuk hasar ---
    If p.tutarJ > 0 And buyukEsik > 0 Then
        uygulandi = True
        If p.buyukSay > 0 Then
            KEkle p, 11, False, "BUYUK HASAR (>= " & TutarMetni(buyukEsik) & "): " & p.buyukSay & " satir, toplam " & _
                  TutarMetni(buyukToplam) & IIf(toplam <> 0, " (tutarin %" & Format$(buyukToplam / toplam * 100, "0.0") & "'i)", "") & _
                  "; ornek: " & Trim$(ornB)
        End If
    End If

    ' --- Benford ilk hane (Nigrini MAD) ---
    If p.tutarJ > 0 And Esik(ekBenfordMad) > 0 And benfordN >= Esik(ekBenfordMinN) And benfordN > 0 Then
        uygulandi = True
        mad = 0
        For rakam = 1 To 9
            beklenen = Log(1 + 1 / rakam) / Log(10#)
            mad = mad + Abs(ilkHane(rakam) / benfordN - beklenen)
            If enKotu = 0 Then
                enKotu = rakam
            ElseIf Abs(ilkHane(rakam) / benfordN - beklenen) > _
                   Abs(ilkHane(enKotu) / benfordN - Log(1 + 1 / enKotu) / Log(10#)) Then
                enKotu = rakam
            End If
        Next rakam
        mad = mad / 9
        p.benfordMad = Round(mad, 4)
        If mad > Esik(ekBenfordMad) Then
            KEkle p, 11, False, "BENFORD SAPMASI: ilk hane MAD " & Format$(mad, "0.0000") & " > " & _
                  Format$(Esik(ekBenfordMad), "0.000") & " (" & SayiMetni(benfordN) & " tutar); en sapan hane " & enKotu & _
                  " (%" & Format$(ilkHane(enKotu) / benfordN * 100, "0.0") & ", beklenen %" & _
                  Format$(Log(1 + 1 / enKotu) / Log(10#) * 100, "0.0") & "). Yuvarlak/limit tutar veya eksik dilim olabilir."
        End If
    End If

    If uygulandi Then
        KTamam p, 11
    Else
        KYok p, 11
    End If

End Sub


' ------------------------------------------------------------
'  K12  SEMA VE KATEGORI KAYMASI
' ------------------------------------------------------------
Private Function SemaMetni(ByVal ws As Worksheet) As String
    Dim sonSut As Long, i As Long
    Dim v As Variant
    Dim s As String
    sonSut = SonSutun(ws, EN_FAZLA_SUTUN)
    If sonSut < 1 Then Exit Function
    v = IkiBoyutlu(ws.Range(ws.Cells(1, 1), ws.Cells(1, sonSut)))
    For i = 1 To sonSut
        If Trim$(GuvenliMetin(v(1, i))) <> "" Then s = s & "|" & Trim$(GuvenliMetin(v(1, i)))
    Next i
    SemaMetni = s & "|"
End Function


Private Function SemaFarki(ByVal eski As String, ByVal yeni As String) As String
    Dim parcalar() As String
    Dim i As Long
    Dim metin As String
    parcalar = Split(eski, "|")
    For i = 0 To UBound(parcalar)
        If parcalar(i) <> "" Then
            If InStr(1, yeni, "|" & parcalar(i) & "|", vbTextCompare) = 0 Then metin = metin & parcalar(i) & ", "
        End If
    Next i
    If Len(metin) > 2 Then metin = Left$(metin, Len(metin) - 2)
    SemaFarki = metin
End Function


Private Sub K12Kontrol(ByRef p As AktarimPlani)

    Dim profilAdi As String, pEksik As String, pYeni As String
    Dim eskiSema As String, yeniSema As String, kayip As String, yeni As String
    Dim ana As Sozluk, onceki As Sozluk
    Dim v As Variant, anahtar As Variant
    Dim son As Long, ilk As Long, i As Long, sinir As Long, aktarimNo As Long, oIlk As Long, oSon As Long
    Dim kaynakEslesti As Boolean, tasti As Boolean
    Dim yeniKod As String, kayipKod As String, k As String
    Dim yeniSay As Long, kayipSay As Long

    ' --- a) profil farki (v1.0'da K1'de idi) ---
    If Not p.sadeceIsaretli Then
        profilAdi = AyarMetni(p.pnl, p.k, SATIR_PROFIL)
        If profilAdi <> "" Then
            ProfilFarklari p.pnl, p.k, profilAdi, pEksik, pYeni
            If pEksik <> "" Then KEkle p, 12, False, "PROFIL FARKI - profilde olup bu dosyada eslesmeyen: " & pEksik
            If pYeni <> "" Then KEkle p, 12, False, "PROFIL FARKI - dosyada yeni (profilde yok, eslestirilmemis): " & pYeni
        End If
    End If

    ' --- b) son basarili aktarima gore kaynak semasi ---
    yeniSema = SemaMetni(p.kaynakWs)
    eskiSema = BellekOku("SEMA|" & p.k)
    If eskiSema <> "" And yeniSema <> "" And BH(eskiSema) <> BH(yeniSema) Then
        kayip = SemaFarki(eskiSema, yeniSema)
        yeni = SemaFarki(yeniSema, eskiSema)
        If kayip <> "" Or yeni <> "" Then
            KEkle p, 12, False, "SEMA KAYMASI (son aktarima gore): " & IIf(kayip <> "", "kaybolan sutun: " & kayip & "; ", "") & _
                  IIf(yeni <> "", "yeni sutun: " & yeni, "")
        ElseIf Not p.sadeceIsaretli Then
            KEkle p, 12, False, "SEMA KAYMASI: kaynak sutunlarinin SIRASI son aktarimdan farkli."
        End If
    End If

    ' --- c) kirilim kodlari: ana dosyada gorulmemis / bu ay kaybolan ---
    If p.kirilimJ > 0 And Not p.grupAnahtar Is Nothing Then
        sinir = CLng(Esik(ekAnalizGrup))
        son = p.hedefWs.Cells(p.hedefWs.Rows.Count, p.kirilimHedefNo).End(xlUp).Row
        If son >= 2 Then
            ilk = son - CLng(Esik(ekK12Pencere)) + 1
            If ilk < 2 Then ilk = 2
            v = SutunOku(p.hedefWs, p.kirilimHedefNo, ilk, son)
            Set ana = New Sozluk
            For i = 1 To UBound(v, 1)
                k = Trim$(GuvenliMetin(v(i, 1)))
                If k <> "" Then
                    If Not ana.Exists(k) Then
                        If ana.Count >= sinir Then
                            tasti = True
                            Exit For
                        End If
                        ana.Item(k) = True
                    End If
                End If
            Next i
            If Not tasti Then
                For Each anahtar In p.grupAnahtar.Keys
                    If CStr(anahtar) <> "" And Not ana.Exists(anahtar) Then
                        yeniSay = yeniSay + 1
                        If yeniSay <= 5 Then yeniKod = yeniKod & CStr(anahtar) & " (" & _
                                                  SayiMetni(p.grupN(p.grupAnahtar.Item(anahtar))) & " satir), "
                    End If
                Next anahtar
            End If
        End If
        ' onceki aktarimin araligi
        If SonAktifAralik("", "", p.hedefWb.Name, CStr(p.hedefWs.Name), aktarimNo, oIlk, oSon, kaynakEslesti) Then
            v = SutunOku(p.hedefWs, p.kirilimHedefNo, oIlk, oSon)
            Set onceki = New Sozluk
            For i = 1 To UBound(v, 1)
                k = Trim$(GuvenliMetin(v(i, 1)))
                If k <> "" Then
                    If Not onceki.Exists(k) Then
                        If onceki.Count >= sinir Then Exit For
                        onceki.Item(k) = True
                        If Not p.grupAnahtar.Exists(k) Then
                            kayipSay = kayipSay + 1
                            If kayipSay <= 5 Then kayipKod = kayipKod & k & ", "
                        End If
                    End If
                End If
            Next i
        End If
        If yeniSay > 0 Then KEkle p, 12, False, "YENI KOD (" & p.kirilimAd & "): ana dosyanin son satirlarinda hic gorulmemis " & _
                                  yeniSay & " kod: " & yeniKod & IIf(yeniSay > 5, "...", "")
        If kayipSay > 0 Then KEkle p, 12, False, "KAYBOLAN KOD (" & p.kirilimAd & "): son aktarimda (no " & aktarimNo & _
                                   ") olup bu dosyada olmayan " & kayipSay & " kod: " & kayipKod & IIf(kayipSay > 5, "...", "")
    End If
    KTamam p, 12

End Sub


' ------------------------------------------------------------
'  K13  KATMANLI MUTABAKAT  (once kaynak gruplari, yazimdan sonra dogrulama)
' ------------------------------------------------------------
Private Sub K13Hazirla(ByRef p As AktarimPlani)

    Dim u() As Long
    Dim i As Long, r As Long, g As Long, m As Long, sinir As Long
    Dim kolK As Variant, kol As Variant
    Dim k As String
    Dim c As Currency

    Set p.grupAnahtar = Nothing
    p.grupAdet = 0
    If p.kirilimHedefNo = 0 Then
        KYok p, 13
        Exit Sub
    End If
    If p.kirilimJ = 0 Then
        KEkle p, 13, False, "KIRILIM SUTUNU (" & p.kirilimAd & ") bu panelde hicbir aylik sutunla eslesmemis; katmanli mutabakat yapilamadi."
        Exit Sub
    End If

    sinir = CLng(Esik(ekAnalizGrup))
    If sinir < 1 Then sinir = 1
    Set p.grupAnahtar = New Sozluk
    ReDim p.grupN(1 To sinir)
    ReDim p.grupT(1 To sinir * IIf(p.mutAdet > 0, p.mutAdet, 1))
    u = p.uygun
    kolK = p.kv(p.kirilimJ)
    For i = 1 To p.n
        r = u(i)
        k = Trim$(GuvenliMetin(kolK(r, 1)))
        g = GuvenliSayi(p.grupAnahtar.Item(k))
        If g = 0 Then
            If p.grupAdet >= sinir Then
                KEkle p, 13, False, "KIRILIM (" & p.kirilimAd & ") " & sinir & "'den fazla farkli deger iceriyor; " & _
                      "katmanli mutabakat yapilmadi (kirilim sutunu dogru mu?)."
                Set p.grupAnahtar = Nothing
                Exit Sub
            End If
            p.grupAdet = p.grupAdet + 1
            g = p.grupAdet
            p.grupAnahtar.Item(k) = g
        End If
        p.grupN(g) = p.grupN(g) + 1
    Next i
    For m = 1 To p.mutAdet
        kol = p.kv(p.mutJ(m))
        For i = 1 To p.n
            r = u(i)
            If ParaDegeri(kol(r, 1), c) Then
                g = p.grupAnahtar.Item(Trim$(GuvenliMetin(kolK(r, 1))))
                p.grupT((g - 1) * p.mutAdet + m) = p.grupT((g - 1) * p.mutAdet + m) + c
            End If
        Next i
    Next m
    KTamam p, 13

End Sub


' Yazimdan sonra: ana dosyaya yazilan satirlari geri okuyup kirilim
' bazinda adet ve her mutabakat sutununu kaynakla karsilastirir.
Private Function K13Dogrula(ByRef p As AktarimPlani, ByRef metin As String) As Boolean

    Dim vK As Variant, v As Variant
    Dim yN() As Long, yT() As Currency
    Dim i As Long, g As Long, m As Long, j As Long, bas As Long, fark As Long
    Dim k As String
    Dim tol As Double
    Dim anahtar As Variant

    K13Dogrula = True
    metin = ""
    If p.grupAnahtar Is Nothing Or p.grupAdet = 0 Then Exit Function
    bas = p.baslangic(p.kirilimJ)
    For m = 1 To p.mutAdet
        If p.baslangic(p.mutJ(m)) <> bas Then
            metin = "sutunlar farkli satirlardan basliyor; katmanli mutabakat yapilmadi"
            Exit Function
        End If
    Next m
    tol = Esik(ekMutabakatTolerans)
    ReDim yN(1 To p.grupAdet)
    ReDim yT(1 To p.grupAdet * IIf(p.mutAdet > 0, p.mutAdet, 1))
    vK = SutunOku(p.hedefWs, p.hedefNo(p.kirilimJ), bas, bas + p.n - 1)
    For i = 1 To p.n
        k = Trim$(GuvenliMetin(vK(i, 1)))
        g = GuvenliSayi(p.grupAnahtar.Item(k))
        If g = 0 Then
            fark = fark + 1
            If fark <= 5 Then metin = metin & "ana dosyada kaynakta olmayan grup '" & k & "'; "
        Else
            yN(g) = yN(g) + 1
        End If
    Next i
    For m = 1 To p.mutAdet
        j = p.mutJ(m)
        v = SutunOku(p.hedefWs, p.hedefNo(j), bas, bas + p.n - 1)
        For i = 1 To p.n
            If VarType(v(i, 1)) <> vbString And Not IsError(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
                g = GuvenliSayi(p.grupAnahtar.Item(Trim$(GuvenliMetin(vK(i, 1)))))
                If g > 0 And IsNumeric(v(i, 1)) Then
                    If Abs(v(i, 1)) < PARA_SINIR Then yT((g - 1) * p.mutAdet + m) = yT((g - 1) * p.mutAdet + m) + CCur(v(i, 1))
                End If
            End If
        Next i
    Next m
    For Each anahtar In p.grupAnahtar.Keys
        g = p.grupAnahtar.Item(anahtar)
        If yN(g) <> p.grupN(g) Then
            fark = fark + 1
            If fark <= 5 Then metin = metin & "'" & CStr(anahtar) & "' adet " & p.grupN(g) & " -> " & yN(g) & "; "
        End If
        For m = 1 To p.mutAdet
            If Abs(yT((g - 1) * p.mutAdet + m) - p.grupT((g - 1) * p.mutAdet + m)) >= tol Then
                fark = fark + 1
                If fark <= 5 Then metin = metin & "'" & CStr(anahtar) & "' " & p.hedefAdlari(p.mutJ(m)) & " " & _
                                         TutarMetni(p.grupT((g - 1) * p.mutAdet + m)) & " -> " & _
                                         TutarMetni(yT((g - 1) * p.mutAdet + m)) & "; "
            End If
        Next m
    Next anahtar
    If fark > 0 Then
        K13Dogrula = False
        metin = fark & " fark: " & metin
    Else
        metin = p.grupAdet & " grup x (adet + " & p.mutAdet & " tutar) TUTUYOR"
    End If

End Function


' ------------------------------------------------------------
'  K14  KAYNAK BUTUNLUGU VE SOY  (M11 TOCTOU, M15)
' ------------------------------------------------------------
Private Sub SoyBilgisi(ByRef p As AktarimPlani)
    p.soyYol = p.kaynakWb.FullName
    p.soyBoyut = 0
    p.soyZaman = ""
    If p.kaynakWb.Path = "" Or InStr(p.soyYol, "://") > 0 Then Exit Sub
    On Error Resume Next
    p.soyBoyut = FileLen(p.soyYol)
    p.soyZaman = Format$(FileDateTime(p.soyYol), "dd.mm.yyyy hh:nn:ss")
    On Error GoTo 0
End Sub


Private Function PlanIziHesapla(ByRef p As AktarimPlani) As String
    Dim j As Long
    Dim s As String
    s = p.kaynakWb.Name & "|" & CStr(p.kaynakWs.Name) & "|" & p.sonSatir & "|" & IIf(p.sadeceIsaretli, "S", "T") & "|"
    For j = 1 To p.gercekAdet
        s = s & p.kaynakNo(j) & ">" & p.hedefNo(j) & ","
    Next j
    PlanIziHesapla = s & "|" & p.filtreMetni & "|" & p.sabitMetni
End Function


' Kaynak izi: satir sayisi + anahtar/tutar/tarih/kirilim (eslenmisse;
' yoksa ilk eslesme) sutunlarinin secili satirlarinin CRC32'si.
Private Function KaynakIziHesapla(ByRef p As AktarimPlani) As String

    Dim kolonlar(1 To 5) As Long
    Dim adet As Long, j As Long, i As Long, t As Long
    Dim crc As Long
    Dim kol As Variant
    Dim u() As Long
    Dim zaten As Boolean

    kolonlar(1) = p.anahtarJ
    kolonlar(2) = p.tutarJ
    kolonlar(3) = p.hasarJ
    kolonlar(4) = p.ihbarJ
    kolonlar(5) = p.kirilimJ
    crc = KarmaBaslat()
    KarmaDeger crc, p.n
    u = p.uygun
    For j = 1 To 5
        zaten = False
        For t = 1 To j - 1
            If kolonlar(t) = kolonlar(j) Then zaten = True
        Next t
        If kolonlar(j) > 0 And Not zaten Then
            adet = adet + 1
            kol = p.kv(kolonlar(j))
            For i = 1 To p.n
                KarmaDeger crc, kol(u(i), 1)
            Next i
        End If
    Next j
    If adet = 0 And p.gercekAdet > 0 Then
        kol = p.kv(1)
        For i = 1 To p.n
            KarmaDeger crc, kol(u(i), 1)
        Next i
    End If
    KaynakIziHesapla = "n=" & p.n & ";c=" & KarmaBitir(crc)

End Function


' K2 (M7): icerik izi = satir + tutar toplami + anahtar (yoksa ilk sutun) CRC
Private Function IcerikIziHesapla(ByRef p As AktarimPlani) As String

    Dim crc As Long, i As Long, j As Long
    Dim kol As Variant
    Dim u() As Long
    Dim toplam As Currency, c As Currency

    u = p.uygun
    crc = KarmaBaslat()
    j = p.anahtarJ
    If j = 0 Then j = 1
    If j <= p.gercekAdet Then
        kol = p.kv(j)
        For i = 1 To p.n
            KarmaDeger crc, kol(u(i), 1)
        Next i
    End If
    If p.tutarJ > 0 Then
        kol = p.kv(p.tutarJ)
        For i = 1 To p.n
            If ParaDegeri(kol(u(i), 1), c) Then toplam = toplam + c
        Next i
    End If
    IcerikIziHesapla = "n=" & p.n & ";t=" & Format$(toplam, "0.00") & ";c=" & KarmaBitir(crc)

End Function


Private Sub K14Kontrol(ByRef p As AktarimPlani)

    Dim anahtar As String, onceki As String
    Dim parcalar() As String

    SoyBilgisi p
    If Not p.kaynakWb.Saved Then
        KEkle p, 14, False, "KAYNAK DOSYADA KAYDEDILMEMIS DEGISIKLIK (" & p.kaynakWb.Name & "): aktarilan veri diskteki " & _
                  "dosyadan farkli olabilir; denetim izi icin once kaynagi kaydet."
    End If
    p.planIzi = PlanIziHesapla(p)
    p.kaynakIzi = KaynakIziHesapla(p)
    anahtar = "OKIZ|" & p.k
    If p.kuru Then
        BellekYaz anahtar, p.planIzi & vbTab & p.kaynakIzi & vbTab & Format$(Now, "dd.mm.yyyy hh:nn")
    Else
        onceki = BellekOku(anahtar)
        If onceki <> "" Then
            parcalar = Split(onceki, vbTab)
            If UBound(parcalar) >= 2 Then
                If parcalar(0) = p.planIzi And parcalar(1) <> p.kaynakIzi Then
                    KEkle p, 14, True, "KAYNAK ON KONTROL'DEN SONRA DEGISTI (On Kontrol: " & parcalar(2) & "; iz " & _
                              parcalar(1) & " -> " & p.kaynakIzi & "). Kontrol edilen veri ile yazilacak veri ayni degil: " & _
                              "On Kontrol'u yeniden calistir."
                End If
            End If
        End If
    End If
    KTamam p, 14

End Sub


' ------------------------------------------------------------
'  METRIKLER VE VERI KALITESI SKORU (DN-06)
' ------------------------------------------------------------
Private Sub TamlikHesapla(ByRef p As AktarimPlani, ByVal zorunluBos As Long, ByVal zorunluHucre As Long)

    Dim kolonlar(1 To 5) As Long
    Dim j As Long, t As Long, i As Long, bos As Long, hucre As Long
    Dim kol As Variant
    Dim u() As Long
    Dim zaten As Boolean

    kolonlar(1) = p.anahtarJ
    kolonlar(2) = p.tutarJ
    kolonlar(3) = p.hasarJ
    kolonlar(4) = p.ihbarJ
    kolonlar(5) = p.kirilimJ
    u = p.uygun
    bos = zorunluBos
    hucre = zorunluHucre
    For j = 1 To 5
        zaten = False
        For t = 1 To j - 1
            If kolonlar(t) = kolonlar(j) Then zaten = True
        Next t
        If kolonlar(j) > 0 And Not zaten Then
            kol = p.kv(kolonlar(j))
            For i = 1 To p.n
                If Trim$(GuvenliMetin(kol(u(i), 1))) = "" Then bos = bos + 1
            Next i
            hucre = hucre + p.n
        End If
    Next j
    If hucre > 0 Then p.mTamlik = (hucre - bos) / hucre

End Sub


Private Sub KaliteHesapla(ByRef p As AktarimPlani)

    Dim pay As Double, payda As Double

    If p.mTamlik >= 0 Then pay = pay + KS_TAMLIK * p.mTamlik: payda = payda + KS_TAMLIK
    If p.mGecerlilik >= 0 Then pay = pay + KS_GECERLILIK * p.mGecerlilik: payda = payda + KS_GECERLILIK
    If p.mTekillik >= 0 Then pay = pay + KS_TEKILLIK * p.mTekillik: payda = payda + KS_TEKILLIK
    If p.mMakulluk >= 0 Then pay = pay + KS_MAKULLUK * p.mMakulluk: payda = payda + KS_MAKULLUK
    If p.mMutabakat >= 0 Then pay = pay + KS_MUTABAKAT * p.mMutabakat: payda = payda + KS_MUTABAKAT
    If payda > 0 Then
        p.kalite = Round(pay / payda * 100, 1)
    Else
        p.kalite = -1
    End If

    p.metrik = "n=" & p.n & ";sutun=" & p.adet & ";tamlik=" & MetrikYuzde(p.mTamlik) & _
               ";gecerlilik=" & MetrikYuzde(p.mGecerlilik) & ";tekillik=" & MetrikYuzde(p.mTekillik) & _
               ";makulluk=" & MetrikYuzde(p.mMakulluk) & ";mutabakat=" & MetrikYuzde(p.mMutabakat) & _
               ";z_satir=" & MetrikDeger(p.zN) & ";z_tutar=" & MetrikDeger(p.zT) & ";benford_mad=" & MetrikDeger(p.benfordMad) & _
               ";buyuk_hasar=" & p.buyukSay & ";mutabakat_sutun=" & p.mutAdet & ";grup=" & p.grupAdet & _
               ";engel=" & p.engelSay & ";dikkat=" & p.dikkatSay & ";sure_ms=" & SureMs(p.basZaman)

End Sub


Private Function MetrikYuzde(ByVal x As Double) As String
    If x < 0 Then
        MetrikYuzde = ""
    Else
        MetrikYuzde = Trim$(Str$(Round(x * 100, 2)))
    End If
End Function


Private Function MetrikDeger(ByVal x As Variant) As String
    If IsEmpty(x) Then
        MetrikDeger = ""
    Else
        MetrikDeger = Trim$(Str$(x))
    End If
End Function


Private Function KaliteMetni(ByVal kalite As Double) As String
    If kalite < 0 Then
        KaliteMetni = "-"
    Else
        KaliteMetni = Format$(kalite, "0") & " / 100"
    End If
End Function


' Para tutari olarak okunabilir mi (Currency tasmasi olmadan)?
Private Function ParaDegeri(ByVal x As Variant, ByRef c As Currency) As Boolean
    Dim d As Double
    ParaDegeri = False
    c = 0
    If Not SayiDegeri(x, d) Then Exit Function
    If Abs(d) >= PARA_SINIR Then Exit Function
    c = CCur(d)
    ParaDegeri = True
End Function


' ============================================================
'  v33 KONTROL IZLERI (ana dosyada kalmis Gelismis Filtre)
' ============================================================

' v33'un "Kontrol" butonlari ana dosyaya gecici olcut hucresi ve
' Gelismis Filtre birakiyordu. v34 ana dosyaya kontrol icin hic
' yazmaz; bu makro eski izleri temizlemek icin durur.
Public Sub KontrolTemizle()

    Dim pnl As Worksheet, hedefWs As Worksheet

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    Set hedefWs = BlokHedefSayfasi(pnl, AktifBlok())
    If hedefWs Is Nothing Then
        Bildir BS_DIKKAT, "Aktarim Paneli", "Hedef dosya/sayfa bulunamadi."
        Exit Sub
    End If

    EkranKilitle
    If EskiKontrolIzleriniSil(hedefWs) Then
        EkranAc
        Bildir BS_OK, "Aktarim Paneli", "Eski kontrol izleri temizlendi (filtre ve olcut hucreleri)." & vbNewLine & _
               hedefWs.Parent.Name & " / " & hedefWs.Name
    Else
        EkranAc
        Bildir BS_DIKKAT, "Aktarim Paneli", "TEMIZLEME TAM YAPILAMADI" & vbNewLine & vbNewLine & _
               "Filtre veya gecici olcut hucreleri hala duruyor." & vbNewLine & _
               "Sayfa korumali olabilir (Review > Unprotect Sheet)."
    End If

End Sub


Private Function EskiKontrolIzleriniSil(ByVal ws As Worksheet) As Boolean

    Dim sonSut As Long, i As Long
    Dim kalanIz As Boolean

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Rows.Hidden = False

    sonSut = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 5
    If sonSut > ws.Columns.Count Then sonSut = ws.Columns.Count
    For i = sonSut To 1 Step -1
        If Trim$(GuvenliMetin(ws.Cells(1, i).Value)) = KONTROL_BASLIK Then
            If Trim$(GuvenliMetin(ws.Cells(2, i).Value)) = "" And ws.Cells(3, i).HasFormula Then
                ws.Range(ws.Cells(1, i), ws.Cells(3, i)).ClearContents
            Else
                ws.Columns(i).Delete
            End If
        End If
    Next i

    kalanIz = ws.FilterMode
    sonSut = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 5
    If sonSut > ws.Columns.Count Then sonSut = ws.Columns.Count
    For i = 1 To sonSut
        If Trim$(GuvenliMetin(ws.Cells(1, i).Value)) = KONTROL_BASLIK Then kalanIz = True
    Next i

    EskiKontrolIzleriniSil = Not kalanIz
    Err.Clear
    On Error GoTo 0

End Function
