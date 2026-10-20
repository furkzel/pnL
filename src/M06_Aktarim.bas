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

Private Const F_ESIT As Long = 1
Private Const F_DESEN As Long = 2
Private Const F_KARSILASTIR As Long = 3
Private Const F_BOS As Long = 4
Private Const F_DOLU As Long = 5

Private Type FiltreOge
    Tur As Long
    Negatif As Boolean
    Metin As String
    Esik As Double
    Op As String
End Type

Private Type FiltreSutun
    Sutun As Long
    Ad As String
    Ham As String
    Adet As Long
    Ogeler() As FiltreOge
End Type

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

    ' Aktueryal kontrol noktalari K1..K8 (M10_Denetim): KD_* durumlari
    kDurum() As String
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
        MsgBox "Aktarilacak panel secilmemis." & vbNewLine & vbNewLine & _
               "Aktarmak istedigin panellerde 'Bu paneli aktar' hucresini" & vbNewLine & _
               "'Evet' yap, sonra tekrar dene.", vbExclamation
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

        Select Case sonuc
            Case asTamam
                basarili = basarili + 1
                If i < secilenSayi Then
                    If MsgBox(i & ". panel aktarildi." & vbNewLine & vbNewLine & _
                              "Siradaki: PANEL " & secili(i + 1) & vbNewLine & _
                              "Devam edilsin mi?", vbYesNo + vbQuestion, "Sirali aktarim") <> vbYes Then Exit For
                End If
            Case asIptal
                If i < secilenSayi Then
                    If MsgBox("PANEL " & k & " aktarilmadi (iptal)." & vbNewLine & vbNewLine & _
                              "Kalan panellere devam edilsin mi?", vbYesNo + vbQuestion, "Sirali aktarim") <> vbYes Then Exit For
                End If
            Case Else
                If i < secilenSayi Then
                    If MsgBox("PANEL " & k & " aktarilamadi." & vbNewLine & vbNewLine & _
                              "Kalan panellere devam edilsin mi?", vbYesNo + vbExclamation, "Sirali aktarim") <> vbYes Then Exit For
                End If
        End Select
    Next i

    If basarili > 0 Then DenetimSessiz

    MsgBox "SIRALI AKTARIM BITTI" & vbNewLine & vbNewLine & _
           "Parti no           : " & partiNo & vbNewLine & _
           "Secilen panel      : " & secilenSayi & vbNewLine & _
           "Aktarilan panel    : " & basarili & vbNewLine & vbNewLine & _
           "TUTAR MUTABAKATI" & vbNewLine & _
           IIf(g_MutabakatOzet = "", "   (aktarim yok)" & vbNewLine, Sigdir(g_MutabakatOzet, 500)) & vbNewLine & _
           "HATIRLATMA: Ana dosyayi SIMDI kaydet, sonra bu paneli de kaydet." & vbNewLine & vbNewLine & _
           "Hepsini birden geri almak icin 'Geri Al' calistirip" & vbNewLine & _
           "'parti' secenegini sec." & vbNewLine & vbNewLine & _
           "Yonetici ozeti: Denetim sayfasi (guncellendi).", vbInformation, "Aktarim"

End Sub


' SECILI SUTUN YAMASI: bulundugun panelde yalnizca ISLEM = SEC sutunlar.
' Yazim modu "Sayfanin en alti" ise bu kaynagin SON AKTIF aktariminin
' satirlarina UZERINE yazar (eski degerler yedeklenir; Geri Al geri getirir).
Public Sub AktarSecili()

    Dim pnl As Worksheet

    Set pnl = PanelSayfasi()
    If Not PanelHazirMi(pnl) Then Exit Sub
    g_MutabakatOzet = ""
    If AktarCalistir(AktifBlok(), True, YeniPartiNo(), 1, 1) = asTamam Then DenetimSessiz

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

    ' Rapor: once ozet satiri, sonra HER madde kendi satirinda, kisaltilmadan
    raporIlk = OlayYaz("KONTROL", p.k, 0, 0, "On kontrol: " & p.n & " satir, " & p.engelSay & " ENGEL, " & _
                       p.dikkatSay & " DIKKAT" & IIf(p.supheliNot <> "", ". Supheli alt satirlar: " & p.supheliNot, ""), _
                       TekSatir(KarneSatiri(p.kDurum)), _
                       DosyaAdiGuvenli(p.kaynakWb), SayfaAdiGuvenli(p.kaynakWs), DosyaAdiGuvenli(p.hedefWb), SayfaAdiGuvenli(p.hedefWs))
    MaddeleriRaporaYaz p, p.engel, "KONTROL", "ENGEL", KD_ENGEL
    MaddeleriRaporaYaz p, p.dikkat, "KONTROL", "DIKKAT", KD_DIKKAT
    If Not p.pnl Is Nothing Then
        KarneYaz p.pnl, p.k, p.kDurum, TR("3 #On Kontrol: ") & Format$(Now, "dd.mm hh:nn") & _
                 IIf(p.n > 0, "  " & ChrW(183) & "  " & SayiMetni(p.n) & TR(" sat#ir"), "")
    End If

    If kisaltildi Then
        ozet = ozet & vbNewLine & "Maddelerin tamami Rapor sayfasinda. Rapor simdi acilsin mi?"
        If MsgBox(ozet, vbYesNo + IIf(p.engel <> "" Or hata <> "", vbExclamation, vbInformation) + vbDefaultButton2, _
                  "On Kontrol") = vbYes Then RaporuGoster raporIlk
    Else
        MsgBox ozet, IIf(p.engel <> "" Or hata <> "", vbExclamation, vbInformation), "On Kontrol"
    End If

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

    AktarCalistir = asHata
    EkranKilitle
    On Error Resume Next
    AktarCalistir = AktarIc(k, sadeceIsaretli, partiNo, sira, toplam)
    On Error GoTo 0
    HesapAc
    EkranAc

End Function


Private Function AktarIc(ByVal k As Long, ByVal sadeceIsaretli As Boolean, _
                         ByVal partiNo As Long, ByVal sira As Long, _
                         ByVal toplam As Long) As AktarimSonucu

    Dim p As AktarimPlani
    Dim mukerrer As String, karar As String, hataMetni As String
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

    ' ============ 7. ONAY ============
    If Not OnayAl(p) Then
        KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR: onay ekran#inda vazge#cildi  (") & Format$(Now, "dd.mm hh:nn") & ")"
        AktarIc = asIptal
        Exit Function
    End If

    ' ============ 8-9. YAZ VE MUTABAKAT ============
    Yaz p
    AktarIc = asTamam
    Exit Function

Hata:
    hataMetni = Err.Number & " - " & Err.Description
    On Error Resume Next
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

    If p.tutarHedefNo = 0 Then KYok p, 7
    KTamam p, 1
    PlanHazirla = True

End Function


Private Sub EngelBildir(ByRef p As AktarimPlani)

    MsgBox "AKTARIM YAPILAMAZ  -  PANEL " & p.k & vbNewLine & vbNewLine & _
           "KONTROL NOKTALARI:" & vbNewLine & KarneSatiri(p.kDurum) & vbNewLine & vbNewLine & _
           MaddeKisalt(p.engel, 650) & vbNewLine & _
           "Hicbir sey yazilmadi. Sorunu giderip tekrar dene." & vbNewLine & _
           "(Ipucu: 'On Kontrol' butun sorunlari tek seferde listeler.)", vbCritical, "Engel"

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
'  FILTRE MOTORU
' ============================================================
Private Function FiltreAyristir(ByVal ham As String) As FiltreSutun

    Dim f As FiltreSutun
    Dim parcalar() As String
    Dim t As String
    Dim i As Long

    f.Ham = Trim$(ham)
    t = f.Ham

    If TekKarsilastirmaMi(t) Then
        ReDim parcalar(0 To 0)
        parcalar(0) = t
    ElseIf InStr(t, ";") > 0 Then
        parcalar = Split(t, ";")
    Else
        parcalar = Split(t, ",")
    End If

    ReDim f.Ogeler(1 To UBound(parcalar) - LBound(parcalar) + 1)
    For i = LBound(parcalar) To UBound(parcalar)
        If Trim$(parcalar(i)) <> "" Then
            f.Adet = f.Adet + 1
            f.Ogeler(f.Adet) = FiltreOgesi(Trim$(parcalar(i)))
        End If
    Next i

    FiltreAyristir = f

End Function


' ">1,5" gibi tek bir karsilastirma mi? (Turkce ondalik virgulu bolunmesin)
Private Function TekKarsilastirmaMi(ByVal t As String) As Boolean

    Dim op As String, geri As String

    TekKarsilastirmaMi = False
    op = KarsilastirmaOp(t, geri)
    If op = "" Then Exit Function
    If IsNumeric(geri) Or (IsDate(geri) And InStr(geri, ",") = 0) Then TekKarsilastirmaMi = True

End Function


Private Function KarsilastirmaOp(ByVal t As String, ByRef geri As String) As String

    KarsilastirmaOp = ""
    geri = ""
    If Left$(t, 2) = ">=" Or Left$(t, 2) = "<=" Then
        KarsilastirmaOp = Left$(t, 2)
        geri = Trim$(Mid$(t, 3))
    ElseIf Left$(t, 2) = "<>" Then
        Exit Function
    ElseIf Left$(t, 1) = ">" Or Left$(t, 1) = "<" Or Left$(t, 1) = "=" Then
        KarsilastirmaOp = Left$(t, 1)
        geri = Trim$(Mid$(t, 2))
    End If

End Function


Private Function FiltreOgesi(ByVal t As String) As FiltreOge

    Dim o As FiltreOge
    Dim tb As String, op As String, geri As String

    If Left$(t, 2) = "<>" Then
        o.Negatif = True
        t = Trim$(Mid$(t, 3))
    End If
    tb = BH(t)

    If t = "" Or tb = "BOS" Then
        o.Tur = F_BOS
    ElseIf tb = "DOLU" Then
        o.Tur = F_DOLU
    Else
        op = KarsilastirmaOp(t, geri)
        If op <> "" And IsNumeric(geri) Then
            o.Tur = F_KARSILASTIR
            o.Op = op
            o.Esik = CDbl(geri)
        ElseIf op <> "" And IsDate(geri) Then
            o.Tur = F_KARSILASTIR
            o.Op = op
            o.Esik = CDbl(CDate(geri))
        ElseIf op = "=" Then
            o.Tur = F_ESIT
            o.Metin = BH(geri)
        ElseIf InStr(t, "*") > 0 Or InStr(t, "?") > 0 Then
            o.Tur = F_DESEN
            o.Metin = Replace(Replace(tb, "[", "[[]"), "#", "[#]")
        Else
            o.Tur = F_ESIT
            o.Metin = tb
        End If
    End If

    FiltreOgesi = o

End Function


' Virgulle bolunmus filtre metni sutunda AYNEN bir deger olarak varsa
' ("ALLIANZ, INC" gibi) tek deger olarak yorumlanir.
Private Sub FiltreBelirsizligiGider(ByRef f As FiltreSutun, ByVal v As Variant)

    Dim i As Long
    Dim aranan As String

    If f.Adet <= 1 Then Exit Sub
    aranan = BH(f.Ham)
    For i = 1 To UBound(v, 1)
        If BH(Trim$(GuvenliMetin(v(i, 1)))) = aranan Then
            f.Adet = 1
            ReDim f.Ogeler(1 To 1)
            f.Ogeler(1).Tur = F_ESIT
            f.Ogeler(1).Negatif = False
            f.Ogeler(1).Metin = aranan
            Exit Sub
        End If
    Next i

End Sub


' Pozitif ogelerden EN AZ BIRI uymali (varsa); negatif (<>) ogelerin HICBIRI uymamali
Private Function FiltreUyar(ByRef f As FiltreSutun, ByVal x As Variant) As Boolean

    Dim i As Long
    Dim s As String, sb As String
    Dim pozVar As Boolean, pozUydu As Boolean, sonuc As Boolean

    s = Trim$(GuvenliMetin(x))
    sb = BH(s)

    For i = 1 To f.Adet
        sonuc = OgeUyar(f.Ogeler(i), x, s, sb)
        If f.Ogeler(i).Negatif Then
            If sonuc Then
                FiltreUyar = False
                Exit Function
            End If
        Else
            pozVar = True
            If sonuc Then pozUydu = True
        End If
    Next i

    FiltreUyar = (Not pozVar) Or pozUydu

End Function


Private Function OgeUyar(ByRef o As FiltreOge, ByVal x As Variant, ByVal s As String, ByVal sb As String) As Boolean

    Dim d As Double

    Select Case o.Tur
        Case F_BOS
            OgeUyar = (s = "")
        Case F_DOLU
            OgeUyar = (s <> "")
        Case F_ESIT
            OgeUyar = (sb = o.Metin)
        Case F_DESEN
            OgeUyar = (sb Like o.Metin)
        Case F_KARSILASTIR
            If SayiDegeri(x, d) Then
                Select Case o.Op
                    Case ">":  OgeUyar = (d > o.Esik)
                    Case ">=": OgeUyar = (d >= o.Esik)
                    Case "<":  OgeUyar = (d < o.Esik)
                    Case "<=": OgeUyar = (d <= o.Esik)
                    Case Else: OgeUyar = (Abs(d - o.Esik) < 0.0000001)
                End Select
            Else
                OgeUyar = False
            End If
    End Select

End Function


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

    Dim j As Long, i As Long, korunan As Long
    Dim kod As Long, tip As String
    Dim tipOnb As Sozluk, dolacak As Sozluk
    Dim kTip As String, hTip As String, tipUyusmaz As String, yesilBos As String
    Dim profilAdi As String, pEksik As String, pYeni As String
    Dim korumaMetni As String, tarihMetni As String, sabitUyusmaz As String
    Dim x As Variant, kol As Variant
    Dim u() As Long

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
    For j = 1 To p.gercekAdet
        kTip = OnbellekTip(tipOnb, p.kaynakWs, p.kaynakNo(j), True)
        hTip = OnbellekTip(tipOnb, p.hedefWs, p.hedefNo(j), False)
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

    ' --- Tutar sutunu butunlugu (mutabakat bozuk metni yakalamaz) ---
    tarihMetni = TutarButunlugu(p)
    If tarihMetni <> "" Then KEkle p, 7, False, tarihMetni

    ' --- Zorunlu (yesil) sutunlarda satir bazinda bos hucre ---
    tarihMetni = YesilBosHucreler(p)
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

        profilAdi = AyarMetni(p.pnl, p.k, SATIR_PROFIL)
        If profilAdi <> "" Then
            ProfilFarklari p.pnl, p.k, profilAdi, pEksik, pYeni
            If pEksik <> "" Then KEkle p, 1, False, "PROFIL FARKI - profilde olup bu dosyada eslesmeyen: " & pEksik
            If pYeni <> "" Then KEkle p, 1, False, "PROFIL FARKI - dosyada yeni (profilde yok, eslestirilmemis): " & pYeni
        End If
    End If

    KTamam p, 1
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
    Dim x As Variant, kol As Variant
    Dim d As Double
    Dim metin As String

    TutarButunlugu = ""
    If p.tutarJ = 0 Then
        KYok p, 7
        Exit Function
    End If

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

End Function


' Ana dosyada YESIL (zorunlu) olup eslestirilmis sutunlarda, aktarilacak
' satirlardan bos olanlar (YesilBosSutunlar yalnizca hic eslesmeyeni bulur).
Private Function YesilBosHucreler(ByRef p As AktarimPlani) As String

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
                    If bos > 0 Then
                        metin = metin & p.hedefAdlari(j) & " sutununda " & bos & " satirda deger bos (ornek aylik satir: " & _
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
    If p.tutarJ > 0 Then ozet = ozet & "Tutar mutabakati : " & p.tutarAd & " (" & SutunHarfi(p.tutarHedefNo) & ")" & vbNewLine
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
    Dim kaynakToplam As Double, yazilanToplam As Double, deger As Double
    Dim hedefOnceki As Double, hedefYeni As Double, eskiAralik As Double
    Dim tutarVar As Boolean, tutarTuttu As Boolean
    Dim mutabakatMetni As String, tutarHarf As String
    Dim raporIlk As Long, raporSon As Long
    Dim u() As Long
    Dim kol As Variant, kolF As Variant, sabitHucre As Variant
    Dim korumaTip As String, kod As Long

    n = p.n
    u = p.uygun

    ' --- Satir ekleme ONAYDAN SONRA (iptalde bos satir kalmasin) ---
    If p.eklenecekSatir > 0 Then
        p.hedefWs.Rows(p.boslukSon + 1).Resize(p.eklenecekSatir).Insert Shift:=xlDown
        GecmisSatirlariKaydir p.hedefWb.Name, CStr(p.hedefWs.Name), p.boslukSon + 1, p.eklenecekSatir
        OlayYaz "SATIR EKLEME", p.k, p.partiNo, 0, "Ana dosyaya " & p.eklenecekSatir & " satir eklendi (" & _
                (p.boslukSon + 1) & ". satirdan itibaren); Gecmis ve Rapor satir numaralari kaydirildi.", _
                "", p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If

    ' --- Mutabakat: once ---
    tutarVar = (p.tutarJ > 0)
    If tutarVar Then
        kol = p.kv(p.tutarJ)
        For i = 1 To n
            If SayiDegeri(kol(u(i), 1), deger) Then kaynakToplam = kaynakToplam + deger
        Next i
        hedefOnceki = KolonToplami(p.hedefWs, p.tutarHedefNo)
        eskiAralik = AralikToplami(p.hedefWs, p.tutarHedefNo, p.baslangic(p.tutarJ), p.baslangic(p.tutarJ) + n - 1)
    End If

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
    Next j

    HesapAc
    DurumCubugu ""

    If p.sadeceIsaretli Then SecIsaretleriniKaldir p.pnl, p.k

    ' --- Mutabakat: sonra ---
    tutarHarf = ""
    mutabakatMetni = "(tutar sutunu secilmemis)"
    tutarTuttu = True
    If tutarVar Then
        tutarHarf = SutunHarfi(p.tutarHedefNo)
        yazilanToplam = AralikToplami(p.hedefWs, p.tutarHedefNo, p.baslangic(p.tutarJ), p.baslangic(p.tutarJ) + n - 1)
        hedefYeni = KolonToplami(p.hedefWs, p.tutarHedefNo)
        tutarTuttu = (Abs(kaynakToplam - yazilanToplam) < 0.005) And _
                     (Abs((hedefYeni - hedefOnceki) - (yazilanToplam - eskiAralik)) < 0.005)
        mutabakatMetni = IIf(tutarTuttu, "TUTUYOR", "FARK VAR") & ": kaynaktan okunan " & TutarMetni(kaynakToplam) & _
                         ", hedefe yazilan " & TutarMetni(yazilanToplam) & _
                         " (fark " & TutarMetni(kaynakToplam - yazilanToplam) & "); hedef sutun toplami " & _
                         TutarMetni(hedefOnceki) & " -> " & TutarMetni(hedefYeni)
    End If

    If tutarVar Then
        g_MutabakatOzet = g_MutabakatOzet & "   PANEL " & p.k & ": " & IIf(tutarTuttu, "TUTUYOR (", "FARK VAR (") & _
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

    RaporYaz p.partiNo, aktarimNo, p.k, p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name), _
             p.sonSatir, n, p.adet, raporIlk, raporSon, p.filtreMetni, p.sabitMetni, mutabakatMetni, _
             tutarVar, kaynakToplam, tutarHarf, yazilanToplam, hedefOnceki, hedefYeni

    ' --- K7 son hukmu: 4'lu mutabakat ve karne ---
    If tutarVar And Not tutarTuttu Then KSeviye p, 7, KD_ENGEL
    KarneYaz p.pnl, p.k, p.kDurum, TR("4 AKTAR: aktar#im no ") & aktarimNo & "  " & ChrW(183) & "  " & _
             IIf(tutarVar, IIf(tutarTuttu, "FARK 0,00", "FARK " & TutarMetni(kaynakToplam - yazilanToplam)), _
                 TR("tutar se#cilmemi#s")) & "  (" & Format$(Now, "dd.mm hh:nn") & ")"

    If kaydirilamayan > 0 Then
        OlayYaz "UYARI", p.k, p.partiNo, aktarimNo, kaydirilamayan & " formul deseni satir kaydirmasi yapilamadan " & _
                "(255 karakter siniri) oldugu gibi yazildi", "Kontrol gerekiyor", _
                p.kaynakWb.Name, CStr(p.kaynakWs.Name), p.hedefWb.Name, CStr(p.hedefWs.Name)
    End If

    If tutarVar And Not tutarTuttu Then
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
        MsgBox "PANEL " & p.k & IIf(p.yama, " YAMASI", "") & " aktarildi." & vbNewLine & vbNewLine & _
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
               "gercek durumu birbirinden ayrisir.", vbInformation, "Aktarim"
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

    Dim rng As Range

    AralikToplami = 0
    If ilk < 1 Or son < ilk Or sutun < 1 Then Exit Function
    Set rng = ws.Range(ws.Cells(ilk, sutun), ws.Cells(son, sutun))
    On Error Resume Next
    AralikToplami = Application.WorksheetFunction.Aggregate(9, 6, rng)
    If Err.Number <> 0 Then
        Err.Clear
        AralikToplami = Application.WorksheetFunction.Sum(rng)
    End If
    Err.Clear
    On Error GoTo 0

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
        MsgBox "Hedef dosya/sayfa bulunamadi.", vbExclamation
        Exit Sub
    End If

    EkranKilitle
    If EskiKontrolIzleriniSil(hedefWs) Then
        EkranAc
        MsgBox "Eski kontrol izleri temizlendi (filtre ve olcut hucreleri)." & vbNewLine & _
               hedefWs.Parent.Name & " / " & hedefWs.Name, vbInformation
    Else
        EkranAc
        MsgBox "TEMIZLEME TAM YAPILAMADI" & vbNewLine & vbNewLine & _
               "Filtre veya gecici olcut hucreleri hala duruyor." & vbNewLine & _
               "Sayfa korumali olabilir (Review > Unprotect Sheet).", vbExclamation
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
