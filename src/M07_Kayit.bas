Attribute VB_Name = "M07_Kayit"
Option Explicit

' ============================================================
' AKTARIM PANELI  v34  -  M07_Kayit  (GERI AL, GECMIS, RAPOR)
'
'   GeriAl        : IKI ASAMALI, ATOMIK geri alma
'                   1) butun kayitlar dogrulanir: dosya acik mi, salt
'                      okunur mu, baslik ayni mi, PARMAK IZI tutuyor mu
'                      (aktarimdan sonra satir eklenip silindiyse yakalanir)
'                   2) ancak hepsi gecerse hucreler temizlenir; uzerine
'                      yazilmis hucreler (yama) YEDEKTEN geri yuklenir
'   Gecmis        : Geri Al'in defteri (sutun basina bir satir) - ELLE
'                   DEGISTIRILMEZ
'   Rapor         : insanlar icin olay gunlugu
'   Yedek (gizli) : uzerine yazilan eski degerler
'   TutarMutabakatiSon : son aktif aktarimin tutarini simdi yeniden kontrol eder
' ============================================================

Private Const YEDEK_ONEK As String = "A"
Private Const IMZA_YOK As String = ""


' ============================================================
'  GERI AL
' ============================================================
Public Sub GeriAl()

    Dim pnl As Worksheet, gec As Worksheet, hedefWs As Worksheet
    Dim hedefWb As Workbook
    Dim v As Variant
    Dim gecSon As Long, i As Long, k As Long, son As Long, t As Long, c As Long
    Dim sonParti As Long, sonAktarim As Long, partiAdet As Long
    Dim partiModu As Boolean
    Dim isaretli As Sozluk, isaretSayisi As Long
    Dim metin As String, sutunHarf As String, sorun As String, imzaFarki As String
    Dim hNo As Long, gNo As Long, ilkS As Long, sonS As Long
    Dim cevap As VbMsgBoxResult
    Dim secilen() As Long, secilenSay As Long
    Dim hedefler() As Worksheet, sutunlar() As Long
    Dim kayitliAd As String, mevcutAd As String, kayitliImza As String, yedekRef As String
    Dim yapan As String, yapanTarih As String, baskasi As Boolean
    Dim temizlenen As Long, geriYuklenen As Long, enKucukIlk As Long, enBuyukSon As Long
    Dim sonHedefWs As Worksheet, sonHedefWb As Workbook
    Dim yeniSonSatir As Long
    Dim hataMetni As String

    Set pnl = PanelSayfasi()
    GecmisSayfasiHazirla
    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)

    gecSon = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    If gecSon < 2 Then
        MsgBox "Gecmis kaydi yok, geri alinacak aktarim bulunamadi.", vbInformation
        Exit Sub
    End If
    v = IkiBoyutlu(gec.Range(gec.Cells(2, 1), gec.Cells(gecSon, G_SON_SUTUN)))

    ' --- Aktif (geri alinmamis) en son parti ve aktarim ---
    For i = 1 To UBound(v, 1)
        If AktifMi(v(i, G_DURUM)) Then
            If GuvenliSayi(v(i, G_PARTI)) > sonParti Then sonParti = GuvenliSayi(v(i, G_PARTI))
        End If
    Next i
    If sonParti = 0 Then
        MsgBox "Geri alinabilecek aktif aktarim yok.", vbInformation
        Exit Sub
    End If
    For i = 1 To UBound(v, 1)
        If AktifMi(v(i, G_DURUM)) And GuvenliSayi(v(i, G_PARTI)) = sonParti Then
            If GuvenliSayi(v(i, G_AKTARIM)) > sonAktarim Then sonAktarim = GuvenliSayi(v(i, G_AKTARIM))
        End If
    Next i

    partiAdet = PartiAktarimSayisi(v, sonParti)
    If partiAdet > 1 Then
        cevap = MsgBox("Son parti (" & sonParti & ") " & partiAdet & " ayri aktarim iceriyor." & vbNewLine & vbNewLine & _
                       "EVET  = partinin TAMAMINI geri al" & vbNewLine & _
                       "HAYIR = sadece SON aktarimi geri al" & vbNewLine & _
                       "IPTAL = vazgec", vbYesNoCancel + vbQuestion, "Geri Al")
        If cevap = vbCancel Then Exit Sub
        partiModu = (cevap = vbYes)
    End If

    ' --- Panelde SEC isaretli hedef sutunlari (varsa yalnizca onlar) ---
    Set isaretli = New Sozluk
    If Not pnl Is Nothing Then
        If PanelDuzeniGuncel() Then
            For k = 1 To BlokSayisi(pnl)
                c = BlokSutun(k)
                son = BlokSonSatir(pnl, k)
                For i = ILK_SATIR To son
                    If IslemSecili(pnl.Cells(i, c + OFS_ISLEM).Value) Then
                        hNo = SutunNo(pnl.Cells(i, c + OFS_HEDEF).Value)
                        If hNo > 0 Then
                            isaretli.Item(BH(SutunHarfi(hNo))) = True
                            isaretSayisi = isaretSayisi + 1
                        End If
                    End If
                Next i
            Next k
        End If
    End If

    ' --- 1. ASAMA: geri alinacak kayitlari topla ve HEPSINI DOGRULA ---
    ReDim secilen(1 To UBound(v, 1))
    ReDim hedefler(1 To UBound(v, 1))
    ReDim sutunlar(1 To UBound(v, 1))

    For i = 1 To UBound(v, 1)
        If GeriAlinacakMi(v, i, sonParti, sonAktarim, partiModu) Then
            sutunHarf = BH(Trim$(GuvenliMetin(v(i, G_HEDEF_SUTUN))))
            If isaretSayisi = 0 Or isaretli.Exists(sutunHarf) Then

                Set hedefWb = DosyaBul(v(i, G_HEDEF_DOSYA))
                Set hedefWs = SayfaBul(hedefWb, v(i, G_HEDEF_SAYFA))
                gNo = SutunNo(sutunHarf)
                ilkS = GuvenliSayi(v(i, G_ILK))
                sonS = GuvenliSayi(v(i, G_SON))

                If hedefWb Is Nothing Then
                    sorun = sorun & "- Ana dosya acik degil: " & GuvenliMetin(v(i, G_HEDEF_DOSYA)) & vbNewLine
                ElseIf hedefWs Is Nothing Then
                    sorun = sorun & "- Sayfa bulunamadi: " & GuvenliMetin(v(i, G_HEDEF_SAYFA)) & vbNewLine
                ElseIf hedefWb.ReadOnly Then
                    sorun = sorun & "- Ana dosya SALT OKUNUR acik: " & hedefWb.Name & vbNewLine
                ElseIf gNo = 0 Or ilkS < 2 Or sonS < ilkS Then
                    sorun = sorun & "- Gecmis kaydi okunamadi (satir " & (i + 1) & ")" & vbNewLine
                Else
                    ' Gecmis sutunu HARF olarak saklar; ana dosyaya sonradan sutun
                    ' eklenmis/silinmisse ayni harf baska sutundur.
                    kayitliAd = Trim$(GuvenliMetin(v(i, G_HEDEF_ADI)))
                    mevcutAd = Trim$(GuvenliMetin(hedefWs.Cells(1, gNo).Value))
                    yedekRef = Trim$(GuvenliMetin(v(i, G_YEDEK)))
                    If kayitliAd <> "" And BH(kayitliAd) <> BH(mevcutAd) Then
                        sorun = sorun & "- " & sutunHarf & " sutununun basligi degismis: aktarimda '" & _
                                kayitliAd & "', simdi '" & mevcutAd & "'" & vbNewLine
                    ElseIf yedekRef <> "" And Not YedekVarMi(yedekRef) Then
                        sorun = sorun & "- " & sutunHarf & " sutununun yedegi (Yedek!" & yedekRef & ") bulunamadi" & vbNewLine
                    Else
                        ' PARMAK IZI: aktarimdan sonra satir eklendi/silindi mi?
                        kayitliImza = Trim$(GuvenliMetin(v(i, G_IMZA)))
                        If kayitliImza <> IMZA_YOK Then
                            If AralikImzasi(hedefWs, gNo, ilkS, sonS, (Left$(kayitliImza, 2) = "F;")) <> kayitliImza Then
                                imzaFarki = imzaFarki & "   " & sutunHarf & " (" & GuvenliMetin(v(i, G_KAYNAK_SUTUN)) & _
                                            ") satir " & ilkS & " - " & sonS & vbNewLine
                            End If
                        End If

                        secilenSay = secilenSay + 1
                        secilen(secilenSay) = i
                        Set hedefler(secilenSay) = hedefWs
                        sutunlar(secilenSay) = gNo
                        If secilenSay <= 12 Then
                            metin = metin & "   " & sutunHarf & " (" & GuvenliMetin(v(i, G_KAYNAK_SUTUN)) & _
                                    ")  satir " & ilkS & " - " & sonS & IIf(yedekRef <> "", "  [yedekten geri yuklenir]", "") & vbNewLine
                        End If
                        If yapan = "" Then
                            yapan = Trim$(GuvenliMetin(v(i, G_KULLANICI)))
                            yapanTarih = TarihMetni(v(i, G_TARIH))
                        End If
                        Set sonHedefWb = hedefWb
                        Set sonHedefWs = hedefWs
                    End If
                End If
            End If
        End If
    Next i

    If sorun <> "" Then
        MsgBox "GERI ALMA YAPILMADI - hicbir hucre silinmedi" & vbNewLine & vbNewLine & _
               Sigdir(sorun, 600) & vbNewLine & _
               "Ana dosyaya sutun eklenmis/silinmis olabilir veya dosya" & vbNewLine & _
               "kapali / salt okunur. Devam edilseydi YANLIS hucreler temizlenebilirdi.", vbCritical, "Geri Al"
        OlayYaz "ENGEL", 0, sonParti, 0, "Geri alma durduruldu: " & TekSatir(sorun), "Hicbir sey silinmedi"
        Exit Sub
    End If

    If secilenSay = 0 Then
        MsgBox "Geri alinacak kayit bulunamadi." & vbNewLine & _
               "Panelde SEC isareti koyduysan, o sutunlar son aktarimda olmayabilir.", vbInformation
        Exit Sub
    End If

    If imzaFarki <> "" Then
        cevap = MsgBox("DIKKAT - ANA DOSYA AKTARIMDAN SONRA DEGISMIS" & vbNewLine & vbNewLine & _
                       "Su araliklarin icerigi aktarim anindaki parmak iziyle TUTMUYOR:" & vbNewLine & _
                       Sigdir(imzaFarki, 400) & vbNewLine & _
                       "Olasi neden: ana dosyada bu satirlarin USTUNE elle satir" & vbNewLine & _
                       "eklenmis/silinmis (veriler kaymis) ya da hucreler elle" & vbNewLine & _
                       "duzeltilmis. Satirlar kaydiysa geri alma BASKA kayitlari silebilir!" & vbNewLine & vbNewLine & _
                       "Ana dosyayi kontrol etmeden devam etme. Yine de geri alinsin mi?", _
                       vbYesNo + vbCritical + vbDefaultButton2, "Geri Al - parmak izi")
        OlayYaz "UYARI", 0, sonParti, 0, "Geri almada parmak izi tutmadi: " & TekSatir(imzaFarki), _
                IIf(cevap = vbYes, "Kullanici yine de devam etti", "Geri alma iptal edildi")
        If cevap <> vbYes Then Exit Sub
    End If

    If secilenSay > 12 Then metin = metin & "   ... ve " & (secilenSay - 12) & " sutun daha" & vbNewLine
    baskasi = (yapan <> "" And BH(yapan) <> BH(KullaniciAdi()))
    If yapan = "" Then yapan = "(bilinmiyor - eski kayit)"

    If MsgBox("GERI ALMA ONAYI" & vbNewLine & vbNewLine & _
              "Parti      : " & sonParti & vbNewLine & _
              "Kapsam     : " & IIf(partiModu, "PARTININ TAMAMI", "SON AKTARIM") & vbNewLine & _
              "Sutunlar   : " & IIf(isaretSayisi = 0, "HEPSI", "SADECE SEC ISARETLILER") & vbNewLine & _
              "Aktarimi yapan : " & yapan & "   (" & yapanTarih & ")" & vbNewLine & _
              IIf(baskasi, "DIKKAT: Bu aktarimi BASKA bir kullanici yapti." & vbNewLine, "") & vbNewLine & _
              "Temizlenecek alanlar:" & vbNewLine & metin & vbNewLine & _
              "Bu hucreler SILINECEK (yedegi olanlar eski haline donecek)." & vbNewLine & _
              "Devam edilsin mi?", vbYesNo + vbExclamation + vbDefaultButton2, "Geri Al") <> vbYes Then
        Exit Sub
    End If

    ' --- 2. ASAMA: temizle / geri yukle ---
    EkranKilitle
    HesapDurdur
    On Error GoTo Hata

    ' Ters sirada: ayni hucreye birden fazla kayit dokunduysa en son
    ' yazilan once geri alinir.
    For t = secilenSay To 1 Step -1
        i = secilen(t)
        ilkS = GuvenliSayi(v(i, G_ILK))
        sonS = GuvenliSayi(v(i, G_SON))
        yedekRef = Trim$(GuvenliMetin(v(i, G_YEDEK)))

        With hedefler(t)
            If yedekRef <> "" Then
                YedekGeriYukle yedekRef, .Range(.Cells(ilkS, sutunlar(t)), .Cells(sonS, sutunlar(t)))
                geriYuklenen = geriYuklenen + 1
            Else
                .Range(.Cells(ilkS, sutunlar(t)), .Cells(sonS, sutunlar(t))).ClearContents
            End If
        End With

        gec.Cells(i + 1, G_DURUM).Value = DURUM_GERI_ALINDI
        gec.Cells(i + 1, G_GERI_ALMA_TARIHI).Value = Now
        temizlenen = temizlenen + 1
        If enKucukIlk = 0 Or ilkS < enKucukIlk Then enKucukIlk = ilkS
        If sonS > enBuyukSon Then enBuyukSon = sonS
    Next t

    If Not sonHedefWs Is Nothing Then yeniSonSatir = SonDoluSatir(sonHedefWs)
    HesapAc
    EkranAc

    OlayYaz "GERI ALMA", 0, sonParti, IIf(partiModu, 0, sonAktarim), _
            temizlenen & " sutun geri alindi (" & geriYuklenen & " tanesi yedekten), ana dosya satir " & _
            enKucukIlk & " - " & enBuyukSon & " (" & (enBuyukSon - enKucukIlk + 1) & " satir). Islem sonrasi son dolu satir: " & _
            yeniSonSatir & ". Aktarimi yapan: " & yapan, _
            IIf(partiModu, "Partinin tamami", "Son aktarim") & IIf(isaretSayisi = 0, ", tum sutunlar", ", SEC isaretli sutunlar"), _
            "", "", DosyaAdiGuvenli(sonHedefWb), SayfaAdiGuvenli(sonHedefWs)

    MsgBox "Geri alma tamamlandi." & vbNewLine & vbNewLine & _
           "Geri alinan sutun        : " & temizlenen & IIf(geriYuklenen > 0, "  (" & geriYuklenen & " tanesi yedekten geri yuklendi)", "") & vbNewLine & _
           "Satir araligi            : " & enKucukIlk & " - " & enBuyukSon & vbNewLine & _
           "Ana dosya : " & DosyaAdiGuvenli(sonHedefWb) & " / " & SayfaAdiGuvenli(sonHedefWs) & vbNewLine & _
           "Islem sonrasi son dolu satir : " & yeniSonSatir & vbNewLine & vbNewLine & _
           "Duzeltip tekrar aktarmak icin:" & vbNewLine & _
           "1. Eslesmeyi duzelt." & vbNewLine & _
           "2. Yazim baslangici 'Geri alinan araliga' (bosalan yere yazar)." & vbNewLine & _
           "3. Tum sutunlar icin AKTAR; yalnizca bazi sutunlar icin" & vbNewLine & _
           "   ISLEM = SEC yapip 'Secili Sutun Yamasi'." & vbNewLine & vbNewLine & _
           "HATIRLATMA: Ana dosyayi ve paneli kaydet.", vbInformation, "Geri Al"
    Exit Sub

Hata:
    hataMetni = Err.Number & " - " & Err.Description
    On Error Resume Next
    HesapAc
    EkranAc
    OlayYaz "HATA", 0, sonParti, 0, "Geri alma sirasinda hata: " & hataMetni, "Gecmis'teki DURUM sutununa bak"
    On Error GoTo 0
    MsgBox "GERI ALMA SIRASINDA HATA" & vbNewLine & vbNewLine & _
           "Hata : " & hataMetni & vbNewLine & vbNewLine & _
           "Gecmis sayfasindaki DURUM sutununa bakip hangi kayitlarin" & vbNewLine & _
           "geri alindigini gorebilirsin.", vbCritical, "Hata"

End Sub


Private Function AktifMi(ByVal x As Variant) As Boolean
    AktifMi = (BH(Trim$(GuvenliMetin(x))) = DURUM_AKTIF)
End Function


Private Function GeriAlinacakMi(ByVal v As Variant, ByVal i As Long, ByVal sonParti As Long, _
                                ByVal sonAktarim As Long, ByVal partiModu As Boolean) As Boolean
    GeriAlinacakMi = False
    If Not AktifMi(v(i, G_DURUM)) Then Exit Function
    If GuvenliSayi(v(i, G_PARTI)) <> sonParti Then Exit Function
    If Not partiModu Then
        If GuvenliSayi(v(i, G_AKTARIM)) <> sonAktarim Then Exit Function
    End If
    GeriAlinacakMi = True
End Function


Private Function PartiAktarimSayisi(ByVal v As Variant, ByVal parti As Long) As Long
    Dim i As Long
    Dim gorulen As Sozluk
    Set gorulen = New Sozluk
    For i = 1 To UBound(v, 1)
        If GuvenliSayi(v(i, G_PARTI)) = parti And AktifMi(v(i, G_DURUM)) Then
            gorulen.Item(CStr(GuvenliSayi(v(i, G_AKTARIM)))) = True
        End If
    Next i
    PartiAktarimSayisi = gorulen.Count
End Function


' ============================================================
'  PARMAK IZI
'
'  Yazilan araligin kisa ozeti: dolu hucre sayisi, sayisal toplam,
'  ilk ve son dolu deger. Formul sutunlarinda (F;) hesap sonucu
'  degisebilecegi icin formul metni kullanilir. Sayilar Str$ ile
'  yazilir: bolgesel ayardan (virgul/nokta) etkilenmez.
' ============================================================
Public Function AralikImzasi(ByVal ws As Worksheet, ByVal sutun As Long, ByVal ilk As Long, _
                             ByVal son As Long, ByVal formulMu As Boolean) As String

    Dim rng As Range
    Dim v As Variant, x As Variant
    Dim i As Long, n As Long
    Dim toplam As Double
    Dim s As String, ilkM As String, sonM As String

    On Error GoTo Hata
    Set rng = ws.Range(ws.Cells(ilk, sutun), ws.Cells(son, sutun))
    If formulMu Then
        v = rng.FormulaR1C1
    Else
        v = rng.Value2
    End If
    If Not IsArray(v) Then
        x = v
        ReDim v(1 To 1, 1 To 1)
        v(1, 1) = x
    End If

    For i = 1 To UBound(v, 1)
        x = v(i, 1)
        s = ""
        If IsError(x) Then
            s = "#E"
        ElseIf IsEmpty(x) Then
            s = ""
        ElseIf VarType(x) = vbString Then
            s = CStr(x)
        ElseIf VarType(x) = vbBoolean Then
            s = IIf(x, "D", "Y")
        ElseIf IsNumeric(x) Then
            toplam = toplam + CDbl(x)
            s = Trim$(Str$(x))
        End If
        If s <> "" Then
            n = n + 1
            If ilkM = "" Then ilkM = Left$(s, 30)
            sonM = Left$(s, 30)
        End If
    Next i

    AralikImzasi = IIf(formulMu, "F;", "V;") & "n=" & n & ";t=" & Trim$(Str$(Round(toplam, 4))) & _
                   ";i=" & ilkM & ";s=" & sonM
    Exit Function

Hata:
    AralikImzasi = "?"

End Function


' ============================================================
'  YEDEK (gizli "Yedek" sayfasi)
'
'  Uzerine yazilacak dolu hucreler (secili sutun yamasi, geri alinan
'  araliga yazim) once buraya kopyalanir. Her yedek bir sutundur:
'  1. satir anahtar "A<aktarim>|<harf>|<satir sayisi>", alti degerler.
'  Formuller "F|" + R1C1, "F|"/"V|" ile baslayan metinler "V|" +
'  metin olarak saklanir. Geri Al yedegi geri yukleyip sutunu bosaltir.
' ============================================================
Private Function YedekSayfasi() As Worksheet
    Dim ws As Worksheet
    Set ws = SayfaOlustur(SAYFA_YEDEK)
    If ws.Visible <> xlSheetHidden Then ws.Visible = xlSheetHidden
    Set YedekSayfasi = ws
End Function


Public Function AralikYedekle(ByVal rng As Range, ByVal aktarimNo As Long, ByVal hedefSutun As Long) As String

    Dim ws As Worksheet
    Dim col As Long, n As Long, i As Long
    Dim v As Variant, f As Variant, x As Variant
    Dim cikti() As Variant
    Dim fs As String

    Set ws = YedekSayfasi()
    n = rng.Rows.Count
    If n + 1 > ws.Rows.Count Then Err.Raise vbObjectError + 513, , "Yedek icin aralik cok buyuk"

    col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If Trim$(GuvenliMetin(ws.Cells(1, col).Value)) <> "" Then col = col + 1
    If col > ws.Columns.Count Then Err.Raise vbObjectError + 514, , "Yedek sayfasi dolu (16.384 sutun)"

    v = IkiBoyutlu(rng)
    If n = 1 Then
        ReDim f(1 To 1, 1 To 1)
        f(1, 1) = rng.FormulaR1C1
    Else
        f = rng.FormulaR1C1
    End If

    ReDim cikti(1 To n, 1 To 1)
    For i = 1 To n
        fs = GuvenliMetin(f(i, 1))
        x = v(i, 1)
        If Left$(fs, 1) = "=" Then
            cikti(i, 1) = "'F|" & fs
        ElseIf VarType(x) = vbString Then
            cikti(i, 1) = "'V|" & CStr(x)
        Else
            cikti(i, 1) = x
        End If
    Next i

    ws.Cells(1, col).Value = "'" & YEDEK_ONEK & aktarimNo & "|" & SutunHarfi(hedefSutun) & "|" & n
    ws.Cells(2, col).Resize(n, 1).Value = cikti
    AralikYedekle = SutunHarfi(col)

End Function


Private Function YedekVarMi(ByVal ref As String) As Boolean
    Dim col As Long
    YedekVarMi = False
    If Not SayfaVarMi(SAYFA_YEDEK) Then Exit Function
    col = SutunNo(ref)
    If col < 1 Then Exit Function
    YedekVarMi = (Left$(GuvenliMetin(ThisWorkbook.Worksheets(SAYFA_YEDEK).Cells(1, col).Value), 1) = YEDEK_ONEK)
End Function


Private Sub YedekGeriYukle(ByVal ref As String, ByVal hedef As Range)

    Dim ws As Worksheet
    Dim col As Long, n As Long, i As Long
    Dim v As Variant, x As Variant
    Dim cikti() As Variant
    Dim s As String

    Set ws = YedekSayfasi()
    col = SutunNo(ref)
    n = hedef.Rows.Count
    v = SutunOku(ws, col, 2, n + 1)

    ReDim cikti(1 To n, 1 To 1)
    For i = 1 To n
        x = v(i, 1)
        If VarType(x) = vbString Then
            s = CStr(x)
            If Left$(s, 2) = "F|" Then
                cikti(i, 1) = Mid$(s, 3)
            ElseIf Left$(s, 2) = "V|" Then
                cikti(i, 1) = "'" & Mid$(s, 3)
            Else
                cikti(i, 1) = "'" & s
            End If
        ElseIf VarType(x) = vbDate Then
            cikti(i, 1) = CDbl(x)
        Else
            cikti(i, 1) = x
        End If
    Next i
    hedef.FormulaR1C1 = cikti

    ' Kullanilan yedek bosaltilir
    ws.Range(ws.Cells(1, col), ws.Cells(n + 1, col)).ClearContents

End Sub


' ============================================================
'  GECMIS
' ============================================================
Public Sub GecmisSayfasiHazirla()

    Dim gec As Worksheet
    Dim eskiAd As String
    Dim i As Long
    Dim basliklar As Variant

    Set gec = SayfaOlustur(SAYFA_GECMIS)

    ' Cok eski bir surumun gecmis sayfasi farkli sutun duzenindeydi
    ' (A sutunu "AKTARIM NO"). Yanlis okunmasin diye kenara alinir.
    If BH(Trim$(GuvenliMetin(gec.Range("A1").Value))) = "AKTARIM NO" Then
        eskiAd = "Gecmis_eski"
        i = 1
        Do While SayfaVarMi(eskiAd)
            i = i + 1
            eskiAd = "Gecmis_eski" & i
        Loop
        On Error Resume Next
        gec.Name = eskiAd
        On Error GoTo 0
        Set gec = SayfaOlustur(SAYFA_GECMIS)
    End If

    basliklar = Array("PARTI", "AKTARIM NO", "TARIH", "HEDEF DOSYA", "HEDEF SAYFA", "HEDEF SUTUN", _
                      "KAYNAK SUTUN", "ILK SATIR", "SON SATIR", "SATIR SAYISI", "DURUM", _
                      "GERI ALMA TARIHI", "HEDEF SUTUN ADI", "KULLANICI", "IMZA", "YEDEK")
    For i = 0 To UBound(basliklar)
        If Trim$(GuvenliMetin(gec.Cells(1, i + 1).Value)) = "" Then
            gec.Cells(1, i + 1).Value = basliklar(i)
            gec.Columns(i + 1).ColumnWidth = 16
        End If
    Next i
    With gec.Range(gec.Cells(1, 1), gec.Cells(1, G_SON_SUTUN))
        .Font.Bold = True
        .Interior.Color = Palet(prTabloBaslik)
    End With

End Sub


Private Function GecmisEnBuyuk(ByVal sutun As Long) As Long

    Dim gec As Worksheet
    Dim son As Long, i As Long, enBuyuk As Long
    Dim v As Variant

    GecmisSayfasiHazirla
    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    son = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    If son >= 2 Then
        v = SutunOku(gec, sutun, 2, son)
        For i = 1 To UBound(v, 1)
            If GuvenliSayi(v(i, 1)) > enBuyuk Then enBuyuk = GuvenliSayi(v(i, 1))
        Next i
    End If
    GecmisEnBuyuk = enBuyuk

End Function


Public Function YeniPartiNo() As Long
    YeniPartiNo = GecmisEnBuyuk(G_PARTI) + 1
End Function


Public Function YeniAktarimNo() As Long
    YeniAktarimNo = GecmisEnBuyuk(G_AKTARIM) + 1
End Function


Public Sub LogYaz(ByVal partiNo As Long, ByVal aktarimNo As Long, ByVal dosyaAdi As String, _
                  ByVal sayfaAdi As String, ByVal hedefSutunNo As Long, ByVal kaynakAdi As String, _
                  ByVal ilkSatir As Long, ByVal sonSatirDeger As Long, ByVal satirSayisi As Long, _
                  ByVal hedefSutunAdi As String, ByVal imza As String, ByVal yedekRef As String)

    Dim gec As Worksheet
    Dim yaz As Long
    Dim w(1 To 1, 1 To G_SON_SUTUN) As Variant

    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    yaz = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row + 1

    w(1, G_PARTI) = partiNo
    w(1, G_AKTARIM) = aktarimNo
    w(1, G_TARIH) = Now
    w(1, G_HEDEF_DOSYA) = HucreMetni(dosyaAdi)
    w(1, G_HEDEF_SAYFA) = "'" & sayfaAdi
    w(1, G_HEDEF_SUTUN) = SutunHarfi(hedefSutunNo)
    w(1, G_KAYNAK_SUTUN) = HucreMetni(kaynakAdi)
    w(1, G_ILK) = ilkSatir
    w(1, G_SON) = sonSatirDeger
    w(1, G_SATIR_SAYISI) = satirSayisi
    w(1, G_DURUM) = DURUM_AKTIF
    w(1, G_HEDEF_ADI) = HucreMetni(hedefSutunAdi)
    w(1, G_KULLANICI) = HucreMetni(KullaniciAdi())
    w(1, G_IMZA) = "'" & imza
    w(1, G_YEDEK) = yedekRef
    gec.Cells(yaz, 1).Resize(1, G_SON_SUTUN).Value = w
    gec.Cells(yaz, G_TARIH).NumberFormat = "dd.mm.yyyy hh:mm"

End Sub


' Gecmiste, verilen hedef dosya/sayfa icin en son GERI ALINDI olan
' aktarimin satir araligi
Public Function GeriAlinanAralik(ByVal dosyaAdi As String, ByVal sayfaAdi As String, _
                                 ByRef ilkSatir As Long, ByRef sonSatirDeger As Long) As Boolean

    Dim gec As Worksheet
    Dim son As Long, i As Long, enSon As Long
    Dim v As Variant

    GeriAlinanAralik = False
    ilkSatir = 0
    sonSatirDeger = 0

    GecmisSayfasiHazirla
    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    son = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    If son < 2 Then Exit Function
    v = IkiBoyutlu(gec.Range(gec.Cells(2, 1), gec.Cells(son, G_SON_SUTUN)))

    For i = 1 To UBound(v, 1)
        If BH(Trim$(GuvenliMetin(v(i, G_DURUM)))) = DURUM_GERI_ALINDI Then
            If BH(Trim$(GuvenliMetin(v(i, G_HEDEF_DOSYA)))) = BH(dosyaAdi) And _
               BH(Trim$(GuvenliMetin(v(i, G_HEDEF_SAYFA)))) = BH(sayfaAdi) Then
                If GuvenliSayi(v(i, G_AKTARIM)) > enSon Then enSon = GuvenliSayi(v(i, G_AKTARIM))
            End If
        End If
    Next i
    If enSon = 0 Then Exit Function

    For i = 1 To UBound(v, 1)
        If GuvenliSayi(v(i, G_AKTARIM)) = enSon And BH(Trim$(GuvenliMetin(v(i, G_DURUM)))) = DURUM_GERI_ALINDI Then
            If ilkSatir = 0 Or GuvenliSayi(v(i, G_ILK)) < ilkSatir Then ilkSatir = GuvenliSayi(v(i, G_ILK))
            If GuvenliSayi(v(i, G_SON)) > sonSatirDeger Then sonSatirDeger = GuvenliSayi(v(i, G_SON))
        End If
    Next i

    GeriAlinanAralik = (ilkSatir > 0 And sonSatirDeger >= ilkSatir)

End Function


' Bu kaynaktan bu hedefe yapilmis SON AKTIF aktarimin satir araligi
' (secili sutun yamasi icin). Kaynak eslesmezse hedefe yapilan son
' aktif aktarim alinir ve kaynakEslesti = False olur.
Public Function SonAktifAralik(ByVal kaynakDosya As String, ByVal kaynakSayfa As String, _
                               ByVal hedefDosya As String, ByVal hedefSayfa As String, _
                               ByRef aktarimNo As Long, ByRef ilkSatir As Long, ByRef sonSatirDeger As Long, _
                               ByRef kaynakEslesti As Boolean) As Boolean

    Dim rpr As Worksheet, gec As Worksheet
    Dim son As Long, i As Long, adayHedef As Long, adayKaynak As Long
    Dim v As Variant
    Dim aktifler As Sozluk
    Dim no As String

    SonAktifAralik = False
    aktarimNo = 0
    ilkSatir = 0
    sonSatirDeger = 0
    kaynakEslesti = False

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    If son < 2 Then Exit Function

    Set aktifler = AktifAktarimlar()
    v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))
    For i = UBound(v, 1) To 1 Step -1
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And _
           BH(Trim$(GuvenliMetin(v(i, R_HEDEF_DOSYA)))) = BH(hedefDosya) And _
           BH(Trim$(GuvenliMetin(v(i, R_HEDEF_SAYFA)))) = BH(hedefSayfa) Then
            no = Trim$(GuvenliMetin(v(i, R_AKTARIM)))
            If aktifler.Exists(no) Then
                If adayHedef = 0 Then adayHedef = GuvenliSayi(no)
                If BH(Trim$(GuvenliMetin(v(i, R_KAYNAK_DOSYA)))) = BH(kaynakDosya) And _
                   BH(Trim$(GuvenliMetin(v(i, R_KAYNAK_SAYFA)))) = BH(kaynakSayfa) Then
                    adayKaynak = GuvenliSayi(no)
                    Exit For
                End If
            End If
        End If
    Next i

    If adayKaynak > 0 Then
        aktarimNo = adayKaynak
        kaynakEslesti = True
    Else
        aktarimNo = adayHedef
    End If
    If aktarimNo = 0 Then Exit Function

    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    son = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    v = IkiBoyutlu(gec.Range(gec.Cells(2, 1), gec.Cells(son, G_SON_SUTUN)))
    For i = 1 To UBound(v, 1)
        If GuvenliSayi(v(i, G_AKTARIM)) = aktarimNo And AktifMi(v(i, G_DURUM)) Then
            If ilkSatir = 0 Or GuvenliSayi(v(i, G_ILK)) < ilkSatir Then ilkSatir = GuvenliSayi(v(i, G_ILK))
            If GuvenliSayi(v(i, G_SON)) > sonSatirDeger Then sonSatirDeger = GuvenliSayi(v(i, G_SON))
        End If
    Next i

    SonAktifAralik = (ilkSatir >= 2 And sonSatirDeger >= ilkSatir)

End Function


' Hedef dosyaya satir eklendiginde, o noktanin ALTINDAKI butun Gecmis
' ve Rapor kayitlarinin satir numaralarini kaydirir.
Public Sub GecmisSatirlariKaydir(ByVal dosyaAdi As String, ByVal sayfaAdi As String, _
                                 ByVal esikSatir As Long, ByVal miktar As Long)

    Dim gec As Worksheet, rpr As Worksheet
    Dim son As Long, i As Long

    If miktar = 0 Then Exit Sub

    GecmisSayfasiHazirla
    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    son = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    For i = 2 To son
        If BH(Trim$(GuvenliMetin(gec.Cells(i, G_HEDEF_DOSYA).Value))) = BH(dosyaAdi) And _
           BH(Trim$(GuvenliMetin(gec.Cells(i, G_HEDEF_SAYFA).Value))) = BH(sayfaAdi) Then
            If GuvenliSayi(gec.Cells(i, G_ILK).Value) >= esikSatir Then
                gec.Cells(i, G_ILK).Value = GuvenliSayi(gec.Cells(i, G_ILK).Value) + miktar
                gec.Cells(i, G_SON).Value = GuvenliSayi(gec.Cells(i, G_SON).Value) + miktar
            End If
        End If
    Next i

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    For i = 2 To son
        If GuvenliMetin(rpr.Cells(i, R_OLAY).Value) = "AKTARIM" And _
           BH(Trim$(GuvenliMetin(rpr.Cells(i, R_HEDEF_DOSYA).Value))) = BH(dosyaAdi) And _
           BH(Trim$(GuvenliMetin(rpr.Cells(i, R_HEDEF_SAYFA).Value))) = BH(sayfaAdi) Then
            If GuvenliSayi(rpr.Cells(i, R_HEDEF_ILK).Value) >= esikSatir Then
                rpr.Cells(i, R_HEDEF_ILK).Value = GuvenliSayi(rpr.Cells(i, R_HEDEF_ILK).Value) + miktar
                rpr.Cells(i, R_HEDEF_SON).Value = GuvenliSayi(rpr.Cells(i, R_HEDEF_SON).Value) + miktar
            End If
        End If
    Next i

End Sub


' Bu kaynak dosya/sayfa, bu hedefe daha once aktarilmis mi? (tamamen
' geri alinmis aktarimlar SAYILMAZ)
Public Function OncekiAktarimlar(ByVal kaynakDosya As String, ByVal kaynakSayfa As String, _
                                 ByVal hedefDosya As String, ByVal hedefSayfa As String) As String

    Dim rpr As Worksheet
    Dim son As Long, i As Long, sayac As Long
    Dim metin As String
    Dim v As Variant
    Dim aktifler As Sozluk

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    If son < 2 Then Exit Function

    Set aktifler = AktifAktarimlar()
    v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))

    For i = UBound(v, 1) To 1 Step -1
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And _
           BH(Trim$(GuvenliMetin(v(i, R_KAYNAK_DOSYA)))) = BH(kaynakDosya) And _
           BH(Trim$(GuvenliMetin(v(i, R_KAYNAK_SAYFA)))) = BH(kaynakSayfa) And _
           BH(Trim$(GuvenliMetin(v(i, R_HEDEF_DOSYA)))) = BH(hedefDosya) And _
           BH(Trim$(GuvenliMetin(v(i, R_HEDEF_SAYFA)))) = BH(hedefSayfa) Then
            If aktifler.Exists(Trim$(GuvenliMetin(v(i, R_AKTARIM)))) Then
                sayac = sayac + 1
                If sayac <= 3 Then
                    metin = metin & "   " & TarihMetni(v(i, R_TARIH)) & "   " & _
                            GuvenliMetin(v(i, R_SATIR)) & " satir   (ana dosya " & _
                            GuvenliMetin(v(i, R_HEDEF_ILK)) & " - " & GuvenliMetin(v(i, R_HEDEF_SON)) & _
                            ", " & GuvenliMetin(v(i, R_KULLANICI)) & ")" & vbNewLine
                End If
            End If
        End If
    Next i

    If sayac > 3 Then metin = metin & "   ... ve " & (sayac - 3) & " kayit daha" & vbNewLine
    OncekiAktarimlar = metin

End Function


' Gecmis'te en az bir sutunu hala AKTIF olan aktarim numaralari
Public Function AktifAktarimlar() As Sozluk

    Dim gec As Worksheet
    Dim son As Long, i As Long
    Dim v As Variant
    Dim sozluk As Sozluk

    Set sozluk = New Sozluk
    GecmisSayfasiHazirla
    Set gec = ThisWorkbook.Worksheets(SAYFA_GECMIS)
    son = gec.Cells(gec.Rows.Count, G_PARTI).End(xlUp).Row
    If son >= 2 Then
        v = IkiBoyutlu(gec.Range(gec.Cells(2, 1), gec.Cells(son, G_DURUM)))
        For i = 1 To UBound(v, 1)
            If AktifMi(v(i, G_DURUM)) Then sozluk.Item(Trim$(GuvenliMetin(v(i, G_AKTARIM)))) = True
        Next i
    End If
    Set AktifAktarimlar = sozluk

End Function


' ============================================================
'  RAPOR  (olay gunlugu)
' ============================================================
Public Sub RaporSayfasiHazirla()

    Dim rpr As Worksheet

    Set rpr = SayfaOlustur(SAYFA_RAPOR)

    ' v30 duzeni (B1 = PARTI) bulunursa kayitlar yeni duzene tasinir
    If BH(Trim$(GuvenliMetin(rpr.Cells(1, 2).Value))) = "PARTI" Then
        RaporGoc rpr
        Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    End If

    If Trim$(GuvenliMetin(rpr.Cells(1, 1).Value)) = "" Then
        RaporBasligiYaz rpr
    ElseIf Trim$(GuvenliMetin(rpr.Cells(1, R_YAZILAN_TOPLAM).Value)) = "" Then
        ' v33 raporu: v34 mutabakat sutunlari eklenir
        rpr.Cells(1, R_YAZILAN_TOPLAM).Value = "HEDEFE YAZILAN TOPLAM"
        rpr.Cells(1, R_HEDEF_ONCEKI).Value = "HEDEF ONCEKI TOPLAM"
        rpr.Cells(1, R_HEDEF_YENI).Value = "HEDEF YENI TOPLAM"
        With rpr.Range(rpr.Cells(1, R_YAZILAN_TOPLAM), rpr.Cells(1, R_HEDEF_YENI))
            .Font.Bold = True
            .Interior.Color = Palet(prTabloBaslik)
            .ColumnWidth = 18
        End With
    End If

End Sub


Private Sub RaporBasligiYaz(ByVal rpr As Worksheet)

    Dim basliklar As Variant
    Dim i As Long

    basliklar = Array("TARIH", "KULLANICI", "OLAY", "PANEL", "PARTI", "AKTARIM NO", _
                      "ACIKLAMA", "KARAR / SONUC", "KAYNAK DOSYA", "KAYNAK SAYFA", _
                      "HEDEF DOSYA", "HEDEF SAYFA", "KAYNAK SON SATIR", "AKTARILAN SATIR", _
                      "SUTUN SAYISI", "HEDEF ILK SATIR", "HEDEF SON SATIR", "FILTRE", _
                      "SABIT DEGER", "TUTAR MUTABAKATI", "AYLIK DOSYA TOPLAMI", _
                      "TUTAR HEDEF SUTUN", "EXCEL KULLANICI ADI", "HEDEFE YAZILAN TOPLAM", _
                      "HEDEF ONCEKI TOPLAM", "HEDEF YENI TOPLAM")
    For i = 0 To UBound(basliklar)
        rpr.Cells(1, i + 1).Value = basliklar(i)
    Next i

    With rpr.Range(rpr.Cells(1, 1), rpr.Cells(1, R_SON))
        .Font.Bold = True
        .Interior.Color = Palet(prTabloBaslik)
        .VerticalAlignment = xlCenter
    End With

    rpr.Columns(R_TARIH).ColumnWidth = 17
    rpr.Columns(R_TARIH).NumberFormat = "dd.mm.yyyy hh:mm"
    rpr.Cells(1, R_TARIH).NumberFormat = "@"
    rpr.Columns(R_KULLANICI).ColumnWidth = 14
    rpr.Columns(R_OLAY).ColumnWidth = 14
    rpr.Columns(R_PANEL).ColumnWidth = 7
    rpr.Columns(R_PARTI).ColumnWidth = 7
    rpr.Columns(R_AKTARIM).ColumnWidth = 10
    rpr.Columns(R_ACIKLAMA).ColumnWidth = 70
    rpr.Columns(R_ACIKLAMA).WrapText = True
    rpr.Columns(R_KARAR).ColumnWidth = 28
    rpr.Columns(R_KARAR).WrapText = True
    rpr.Range(rpr.Columns(R_KAYNAK_DOSYA), rpr.Columns(R_SON)).ColumnWidth = 16
    rpr.Columns(R_MUTABAKAT).ColumnWidth = 40
    rpr.Columns(R_FILTRE).ColumnWidth = 30
    rpr.Rows(1).WrapText = False

End Sub


' v30 Rapor sayfasini (her satir bir aktarim, 19 sutun) olay gunlugu
' duzenine tasir. Once yeni sayfa tam yazilir, eski ancak sonra silinir.
Private Sub RaporGoc(ByVal eski As Worksheet)

    Dim yeni As Worksheet
    Dim son As Long, i As Long, n As Long
    Dim v As Variant, w() As Variant

    On Error GoTo Hata

    son = eski.Cells(eski.Rows.Count, 1).End(xlUp).Row
    Application.DisplayAlerts = False
    If SayfaVarMi("Rapor_gecis") Then ThisWorkbook.Worksheets("Rapor_gecis").Delete
    Set yeni = ThisWorkbook.Worksheets.Add(After:=eski)
    yeni.Name = "Rapor_gecis"
    RaporBasligiYaz yeni

    If son >= 2 Then
        n = son - 1
        v = IkiBoyutlu(eski.Range(eski.Cells(2, 1), eski.Cells(son, 19)))
        ReDim w(1 To n, 1 To R_SON)
        For i = 1 To n
            w(i, R_TARIH) = v(i, 1)
            w(i, R_KULLANICI) = v(i, 17)
            w(i, R_OLAY) = "AKTARIM"
            w(i, R_PANEL) = v(i, 4)
            w(i, R_PARTI) = v(i, 2)
            w(i, R_AKTARIM) = v(i, 3)
            w(i, R_ACIKLAMA) = GuvenliMetin(v(i, 10)) & " satir, " & GuvenliMetin(v(i, 11)) & _
                               " sutun -> ana dosya satir " & GuvenliMetin(v(i, 12)) & " - " & _
                               GuvenliMetin(v(i, 13)) & "   (v30 kaydindan tasindi)"
            w(i, R_KARAR) = HucreMetni(GuvenliMetin(v(i, 16)))
            w(i, R_KAYNAK_DOSYA) = v(i, 5)
            w(i, R_KAYNAK_SAYFA) = v(i, 6)
            w(i, R_HEDEF_DOSYA) = v(i, 7)
            w(i, R_HEDEF_SAYFA) = v(i, 8)
            w(i, R_KAYNAK_SON) = v(i, 9)
            w(i, R_SATIR) = v(i, 10)
            w(i, R_SUTUN) = v(i, 11)
            w(i, R_HEDEF_ILK) = v(i, 12)
            w(i, R_HEDEF_SON) = v(i, 13)
            w(i, R_FILTRE) = HucreMetni(GuvenliMetin(v(i, 14)))
            w(i, R_SABIT) = v(i, 15)
            w(i, R_MUTABAKAT) = HucreMetni(GuvenliMetin(v(i, 16)))
            w(i, R_AYLIK_TOPLAM) = v(i, 18)
            w(i, R_TUTAR_SUTUN) = v(i, 19)
            w(i, R_EXCEL_KULLANICI) = v(i, 17)
        Next i
        yeni.Range(yeni.Cells(2, 1), yeni.Cells(son, R_SON)).Value = w
        yeni.Range(yeni.Cells(2, R_OLAY), yeni.Cells(son, R_OLAY)).Interior.Color = OlayRengi("AKTARIM")
    End If

    eski.Delete
    yeni.Name = SAYFA_RAPOR
    Application.DisplayAlerts = True
    Exit Sub

Hata:
    Application.DisplayAlerts = True
    MsgBox "RAPOR SAYFASI YENI DUZENE TASINAMADI" & vbNewLine & vbNewLine & _
           "Hata : " & Err.Number & " - " & Err.Description & vbNewLine & vbNewLine & _
           "Eski Rapor sayfasi silinmedi. 'Rapor_gecis' adinda yarim" & vbNewLine & _
           "bir sayfa kalmis olabilir; onu silip tekrar dene.", vbCritical, "Rapor"

End Sub


Public Function OlayRengi(ByVal olay As String) As Long
    Select Case olay
        Case "AKTARIM":                         OlayRengi = RGB(220, 252, 231)
        Case "DIKKAT", "UYARI", "SATIR EKLEME": OlayRengi = RGB(254, 243, 199)
        Case "ENGEL", "HATA":                   OlayRengi = RGB(254, 202, 202)
        Case "GERI ALMA":                       OlayRengi = RGB(219, 234, 254)
        Case "KONTROL", "MUTABAKAT":            OlayRengi = RGB(237, 233, 254)
        Case "IPTAL":                           OlayRengi = RGB(241, 245, 249)
        Case Else:                              OlayRengi = RGB(255, 255, 255)
    End Select
End Function


' Rapor sayfasina bir olay satiri yazar, satir numarasini dondurur.
' Kayit yazilamasa bile asil islemi DURDURMAZ (kendi hata yakalayicisi var).
Public Function OlayYaz(ByVal olay As String, ByVal panelNo As Long, _
                        ByVal parti As Long, ByVal aktarim As Long, _
                        ByVal aciklama As String, ByVal karar As String, _
                        Optional ByVal kaynakDosya As String = "", _
                        Optional ByVal kaynakSayfa As String = "", _
                        Optional ByVal hedefDosya As String = "", _
                        Optional ByVal hedefSayfa As String = "") As Long

    Dim rpr As Worksheet
    Dim yaz As Long

    OlayYaz = 0
    On Error GoTo Cikis

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    yaz = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row + 1

    rpr.Cells(yaz, R_TARIH).Value = Now
    rpr.Cells(yaz, R_KULLANICI).Value = HucreMetni(KullaniciAdi())
    rpr.Cells(yaz, R_OLAY).Value = olay
    rpr.Cells(yaz, R_OLAY).Interior.Color = OlayRengi(olay)
    If panelNo > 0 Then rpr.Cells(yaz, R_PANEL).Value = panelNo
    If parti > 0 Then rpr.Cells(yaz, R_PARTI).Value = parti
    If aktarim > 0 Then rpr.Cells(yaz, R_AKTARIM).Value = aktarim
    rpr.Cells(yaz, R_ACIKLAMA).Value = HucreMetni(Left$(aciklama, 30000))
    rpr.Cells(yaz, R_KARAR).Value = HucreMetni(Left$(karar, 30000))
    If kaynakDosya <> "" Then rpr.Cells(yaz, R_KAYNAK_DOSYA).Value = HucreMetni(kaynakDosya)
    If kaynakSayfa <> "" Then rpr.Cells(yaz, R_KAYNAK_SAYFA).Value = "'" & kaynakSayfa
    If hedefDosya <> "" Then rpr.Cells(yaz, R_HEDEF_DOSYA).Value = HucreMetni(hedefDosya)
    If hedefSayfa <> "" Then rpr.Cells(yaz, R_HEDEF_SAYFA).Value = "'" & hedefSayfa
    rpr.Cells(yaz, R_EXCEL_KULLANICI).Value = HucreMetni(Application.UserName)

    OlayYaz = yaz

Cikis:

End Function


' Basarili bir aktarimi Rapor'a AKTARIM olayi olarak yazar
Public Sub RaporYaz(ByVal partiNo As Long, ByVal aktarimNo As Long, ByVal k As Long, _
                    ByVal kaynakDosya As String, ByVal kaynakSayfa As String, _
                    ByVal hedefDosya As String, ByVal hedefSayfa As String, _
                    ByVal kaynakSonSatir As Long, ByVal aktarilanSatir As Long, _
                    ByVal sutunSayisi As Long, ByVal hedefIlk As Long, ByVal hedefSon As Long, _
                    ByVal filtre As String, ByVal sabit As String, ByVal mutabakat As String, _
                    ByVal tutarVar As Boolean, ByVal kaynakToplam As Double, ByVal tutarHedefHarf As String, _
                    ByVal yazilanToplam As Double, ByVal hedefOnceki As Double, ByVal hedefYeni As Double)

    Dim rpr As Worksheet
    Dim yaz As Long

    yaz = OlayYaz("AKTARIM", k, partiNo, aktarimNo, _
                  SayiMetni(aktarilanSatir) & " satir, " & sutunSayisi & _
                  " sutun  ->  ana dosya satir " & SayiMetni(hedefIlk) & " - " & SayiMetni(hedefSon), _
                  mutabakat, kaynakDosya, kaynakSayfa, hedefDosya, hedefSayfa)
    If yaz = 0 Then Exit Sub

    On Error Resume Next
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    rpr.Cells(yaz, R_KAYNAK_SON).Value = kaynakSonSatir
    rpr.Cells(yaz, R_SATIR).Value = aktarilanSatir
    rpr.Cells(yaz, R_SUTUN).Value = sutunSayisi
    rpr.Cells(yaz, R_HEDEF_ILK).Value = hedefIlk
    rpr.Cells(yaz, R_HEDEF_SON).Value = hedefSon
    rpr.Cells(yaz, R_FILTRE).Value = HucreMetni(Trim$(filtre))
    rpr.Cells(yaz, R_SABIT).Value = HucreMetni(sabit)
    rpr.Cells(yaz, R_MUTABAKAT).Value = HucreMetni(mutabakat)
    If tutarVar Then
        rpr.Cells(yaz, R_AYLIK_TOPLAM).Value = kaynakToplam
        rpr.Cells(yaz, R_TUTAR_SUTUN).Value = tutarHedefHarf
        rpr.Cells(yaz, R_YAZILAN_TOPLAM).Value = yazilanToplam
        rpr.Cells(yaz, R_HEDEF_ONCEKI).Value = hedefOnceki
        rpr.Cells(yaz, R_HEDEF_YENI).Value = hedefYeni
    End If
    On Error GoTo 0

End Sub


' ============================================================
'  TUTAR MUTABAKATI (son aktif PARTI, simdiki hal)
'
'  Son aktif partideki (ayni AKTAR turu) HER aktarim icin Rapor'da
'  saklanan "aylik dosyada aktarilan toplam" ile ana dosyada o satir
'  araliginin SIMDIKI toplami karsilastirilir ve tek tabloda gosterilir.
'  Bulundugun panelin satiri ">>" ile isaretlenir. Aktarimdan sonra
'  elle yapilan degisiklikleri de yakalar.
' ============================================================
Public Sub TutarMutabakatiSon()

    Dim rpr As Worksheet, hedefWs As Worksheet
    Dim son As Long, i As Long, n As Long, aktifPanel As Long, panelNo As Long
    Dim parti As Long, sonParti As Long, ilkSatir As Long, sonSatirDeger As Long, hedefSutunNo As Long
    Dim hedefDosya As String, hedefSayfa As String, hedefSutunHarf As String
    Dim kaynakToplam As Double, simdikiToplam As Double, fark As Double
    Dim aktifler As Sozluk, gorulen As Sozluk
    Dim v As Variant
    Dim aktarimVar As Boolean, tutarliVar As Boolean, hepsiTutuyor As Boolean
    Dim tablo As String, satir As String, durum As String, no As String

    RaporSayfasiHazirla
    Set rpr = ThisWorkbook.Worksheets(SAYFA_RAPOR)
    son = rpr.Cells(rpr.Rows.Count, R_TARIH).End(xlUp).Row
    If son < 2 Then
        MsgBox "Rapor kaydi yok; once bir aktarim yapilmali.", vbInformation
        Exit Sub
    End If

    Set aktifler = AktifAktarimlar()
    v = IkiBoyutlu(rpr.Range(rpr.Cells(2, 1), rpr.Cells(son, R_SON)))

    ' --- Tutari kayitli ETKIN aktarimlarin en son partisi ---
    For i = UBound(v, 1) To 1 Step -1
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" Then
            aktarimVar = True
            If IsNumeric(v(i, R_AYLIK_TOPLAM)) And Trim$(GuvenliMetin(v(i, R_AYLIK_TOPLAM))) <> "" Then
                tutarliVar = True
                If aktifler.Exists(Trim$(GuvenliMetin(v(i, R_AKTARIM)))) Then
                    parti = GuvenliSayi(v(i, R_PARTI))
                    If parti > sonParti Then sonParti = parti
                End If
            End If
        End If
    Next i

    If sonParti = 0 Then
        If Not aktarimVar Then
            MsgBox "Henuz hic aktarim yapilmamis.", vbInformation, "Tutar Mutabakati"
        ElseIf Not tutarliVar Then
            MsgBox "Son aktarimlarda tutar sutunu secilmemis." & vbNewLine & _
                   "Panelde 'Tutar sutunu (mutabakat)' secip aktarim yaparsan" & vbNewLine & _
                   "bu kontrol calisir.", vbExclamation, "Tutar Mutabakati"
        Else
            MsgBox "Kontrol edilecek ETKIN bir aktarim yok (hepsi geri alinmis).", vbInformation, "Tutar Mutabakati"
        End If
        Exit Sub
    End If

    aktifPanel = AktifBlok()
    hepsiTutuyor = True
    Set gorulen = New Sozluk

    ' --- Partideki her aktarim (Rapor sirasiyla) ---
    For i = 1 To UBound(v, 1)
        no = Trim$(GuvenliMetin(v(i, R_AKTARIM)))
        If GuvenliMetin(v(i, R_OLAY)) = "AKTARIM" And GuvenliSayi(v(i, R_PARTI)) = sonParti And _
           aktifler.Exists(no) And Not gorulen.Exists(no) Then
            gorulen.Item(no) = True
            panelNo = GuvenliSayi(v(i, R_PANEL))
            satir = IIf(panelNo = aktifPanel, ">> ", "   ") & "PANEL " & panelNo & "  (aktarim " & no & ")  "

            If IsNumeric(v(i, R_AYLIK_TOPLAM)) And Trim$(GuvenliMetin(v(i, R_AYLIK_TOPLAM))) <> "" Then
                hedefDosya = GuvenliMetin(v(i, R_HEDEF_DOSYA))
                hedefSayfa = GuvenliMetin(v(i, R_HEDEF_SAYFA))
                ilkSatir = GuvenliSayi(v(i, R_HEDEF_ILK))
                sonSatirDeger = GuvenliSayi(v(i, R_HEDEF_SON))
                hedefSutunHarf = GuvenliMetin(v(i, R_TUTAR_SUTUN))
                hedefSutunNo = SutunNo(hedefSutunHarf)
                kaynakToplam = CDbl(v(i, R_AYLIK_TOPLAM))
                Set hedefWs = SayfaBul(DosyaBul(hedefDosya), hedefSayfa)

                If hedefWs Is Nothing Then
                    durum = "ANA DOSYA ACIK DEGIL"
                    hepsiTutuyor = False
                    satir = satir & hedefDosya & ": " & durum
                ElseIf hedefSutunNo = 0 Or ilkSatir < 2 Then
                    durum = "KAYIT OKUNAMADI"
                    hepsiTutuyor = False
                    satir = satir & durum
                Else
                    simdikiToplam = AralikToplami(hedefWs, hedefSutunNo, ilkSatir, sonSatirDeger)
                    fark = kaynakToplam - simdikiToplam
                    If Abs(fark) < 0.005 Then
                        durum = "TUTUYOR"
                    Else
                        durum = "FARK VAR"
                        hepsiTutuyor = False
                    End If
                    satir = satir & hedefSutunHarf & " " & ilkSatir & "-" & sonSatirDeger & ":  aylik " & _
                            TutarMetni(kaynakToplam) & "  simdi " & TutarMetni(simdikiToplam) & _
                            "  fark " & TutarMetni(fark) & "  " & durum
                End If
                OlayYaz "MUTABAKAT", panelNo, sonParti, GuvenliSayi(no), _
                        "Tutar kontrolu: aylik " & TutarMetni(kaynakToplam) & ", ana dosyada simdi " & _
                        TutarMetni(simdikiToplam) & " (sutun " & hedefSutunHarf & ", satir " & ilkSatir & "-" & sonSatirDeger & ")", _
                        durum, "", "", hedefDosya, hedefSayfa
            Else
                satir = satir & "tutar sutunu secilmemisti"
            End If
            tablo = tablo & satir & vbNewLine
            n = n + 1
        End If
    Next i

    MsgBox "SON PARTI (" & sonParti & ", " & n & " aktarim)  -  " & _
           IIf(hepsiTutuyor, "TUTARI TUTUYOR", "FARK VAR") & vbNewLine & vbNewLine & _
           Sigdir(tablo, 750) & vbNewLine & _
           IIf(hepsiTutuyor, "Butun aktarimlarda FARK = 0,00.", _
               "Fark olan satirlarda aktarimdan sonra degisiklik yapilmis ya da" & vbNewLine & _
               "ustlerine satir eklenip silinmis olabilir; kontrol et.") & vbNewLine & _
           "(>> : bulundugun panel)", IIf(hepsiTutuyor, vbInformation, vbExclamation), "Tutar Mutabakati"

End Sub
