Attribute VB_Name = "M02_Yardimci"
Option Explicit
Option Private Module

' ============================================================
' AKTARIM PANELI  v34  -  M02_Yardimci
'
' Uygulamadan bagimsiz yardimcilar. Bu modul hicbir sayfa duzenini
' bilmez (Panel'e ozgu geometri M03_Panel'dedir).
'
'   EKRAN KILIDI : EkranKilitle / EkranAc / HesapDurdur / HesapAc
'   METIN        : BH, Normalize, TR, GuvenliMetin, HucreMetni, Sigdir ...
'   SUTUN/HARF   : SutunNo, SutunHarfi, SutunAdi
'   DOSYA/SAYFA  : DosyaBul, DosyaAc, SayfaBul, SayfaOlustur ...
'   OKUMA        : SutunOku (tasma korumali), SonDoluSatir, SutunSonDolu
'   VERI TIPI    : VeriTipiKaynak / Hedef / Dizi, TipCatisir ...
'   BASLIK       : BaslikDogrula, AyniAdlar
'   RENK         : HucreRengi, YesilMi
'   MENU         : MenuKur, MenuKurSerbest
'   ALAN         : AlanTemizle (Cells.Clear yerine)
' ============================================================

Public Const PANEL_YAZI_TIPI As String = "Segoe UI"
Public Const PANEL_YAZI_BOYUT As Long = 9


' ============================================================
'  EKRAN KILIDI
'
'  Her giris makrosu:
'      EkranKilitle
'      On Error GoTo Cikis
'      ...
'  Cikis:
'      EkranAc
'  Ic ice cagrilarda (Aktar -> AktarCalistir, PanelKur -> Yenile)
'  derinlik sayilir; Excel yalnizca en distaki EkranAc'ta acilir.
'  EkranAc True (zorla) derinligi sifirlar: Tani ve Workbook_Open
'  Excel'i her durumda normale dondurmek icin kullanir.
' ============================================================
Public Sub EkranKilitle()

    g_KilitDerinlik = g_KilitDerinlik + 1
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo 0

End Sub


Public Sub EkranAc(Optional ByVal zorla As Boolean = False)

    If zorla Then
        g_KilitDerinlik = 0
    Else
        g_KilitDerinlik = g_KilitDerinlik - 1
        If g_KilitDerinlik > 0 Then Exit Sub
        g_KilitDerinlik = 0
    End If

    On Error Resume Next
    HesapAc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

End Sub


' Hesaplamayi elle moda alir (buyuk yazmalardan once). HesapAc veya en
' distaki EkranAc onceki moda dondurur.
Public Sub HesapDurdur()

    If g_HesapDurdu Then Exit Sub
    On Error Resume Next
    g_EskiHesap = Application.Calculation
    Application.Calculation = xlCalculationManual
    If Err.Number = 0 Then g_HesapDurdu = True
    Err.Clear
    On Error GoTo 0

End Sub


Public Sub HesapAc()

    If Not g_HesapDurdu Then Exit Sub
    On Error Resume Next
    Application.Calculation = g_EskiHesap
    Err.Clear
    On Error GoTo 0
    g_HesapDurdu = False

End Sub


' Kullaniciya bir mesaj gostermeden once ekrani gecici olarak acar
' (ScreenUpdating False iken MsgBox arkasinda eski goruntu kalabilir)
Public Sub EkraniTazele()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    On Error GoTo 0
End Sub


Public Sub DurumCubugu(ByVal metin As String)
    On Error Resume Next
    If metin = "" Then
        Application.StatusBar = False
    Else
        Application.StatusBar = metin
    End If
    On Error GoTo 0
End Sub


' ============================================================
'  TURKCE KARAKTER VE KARSILASTIRMA
' ============================================================

' Dil ayarindan bagimsiz BUYUK HARF + Turkce harfleri ASCII'ye katlama.
' Turkce Windows'ta UCase("i") = noktali I (ChrW 304) dondurur; o zaman
' "Geri alinan" ile "GERI ALINAN" sessizce tutmaz. Ayrica TR() ile
' yazilmis hucre metinleri ("Sutunun kendi sonu" Turkce harfli) ASCII
' sabitlerle karsilastirilabilsin diye u/s/g/o/c da katlanir.
Public Function BH(ByVal s As String) As String

    s = UCase$(s)
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(350), "S")
    s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(286), "G")
    s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(220), "U")
    s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(214), "O")
    s = Replace(s, ChrW(246), "O")
    s = Replace(s, ChrW(199), "C")
    s = Replace(s, ChrW(231), "C")
    s = Replace(s, ChrW(775), "")
    BH = s

End Function


' Sutun adi karsilastirmasi icin sadelestirir: Turkce harf, buyuk/kucuk,
' bosluk, alt cizgi, nokta, parantez, tire farki yok sayilir.
Public Function Normalize(ByVal metin As String) As String

    Dim s As String

    s = Trim$(metin)

    s = Replace(s, ChrW(304), "i")
    s = Replace(s, ChrW(305), "i")
    s = Replace(s, ChrW(350), "s")
    s = Replace(s, ChrW(351), "s")
    s = Replace(s, ChrW(286), "g")
    s = Replace(s, ChrW(287), "g")
    s = Replace(s, ChrW(220), "u")
    s = Replace(s, ChrW(252), "u")
    s = Replace(s, ChrW(214), "o")
    s = Replace(s, ChrW(246), "o")
    s = Replace(s, ChrW(199), "c")
    s = Replace(s, ChrW(231), "c")

    s = LCase$(s)
    s = Replace(s, ChrW(775), "")
    s = Replace(s, ChrW(305), "i")

    s = Replace(s, " ", "")
    s = Replace(s, ChrW(160), "")
    s = Replace(s, vbTab, "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    s = Replace(s, ".", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "/", "")
    s = Replace(s, "\", "")
    s = Replace(s, ":", "")
    s = Replace(s, "'", "")

    Normalize = s

End Function


' Hucre/sekil metni icin Turkce harf ureticisi. Kaynak kod ASCII kalir:
'   #i = noktasiz i   #I = noktali I   #s #S   #g #G   #u #U   #o #O   #c #C
'   #a #A = sapkali a   #- = uzun tire   #> = ok   #* = orta nokta   ## = #
Public Function TR(ByVal s As String) As String

    Dim i As Long, n As Long
    Dim ch As String, sonuc As String

    If InStr(s, "#") = 0 Then
        TR = s
        Exit Function
    End If

    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "#" And i < n Then
            sonuc = sonuc & TRHarf(Mid$(s, i + 1, 1))
            i = i + 2
        Else
            sonuc = sonuc & ch
            i = i + 1
        End If
    Loop

    TR = sonuc

End Function


' Select Case metin karsilastirmasi buyuk/kucuk harfe duyarlidir
' (Option Compare Binary); "i" ile "I" ayri yakalanir.
Private Function TRHarf(ByVal c As String) As String

    Select Case c
        Case "i": TRHarf = ChrW(305)
        Case "I": TRHarf = ChrW(304)
        Case "s": TRHarf = ChrW(351)
        Case "S": TRHarf = ChrW(350)
        Case "g": TRHarf = ChrW(287)
        Case "G": TRHarf = ChrW(286)
        Case "u": TRHarf = ChrW(252)
        Case "U": TRHarf = ChrW(220)
        Case "o": TRHarf = ChrW(246)
        Case "O": TRHarf = ChrW(214)
        Case "c": TRHarf = ChrW(231)
        Case "C": TRHarf = ChrW(199)
        Case "a": TRHarf = ChrW(226)
        Case "A": TRHarf = ChrW(194)
        Case "-": TRHarf = ChrW(8212)
        Case ">": TRHarf = ChrW(8594)
        Case "*": TRHarf = ChrW(183)
        Case "#": TRHarf = "#"
        Case Else: TRHarf = "#" & c
    End Select

End Function


Public Function KilavuzAdi() As String
    KilavuzAdi = "K" & ChrW(305) & "lavuz"
End Function


' ============================================================
'  GUVENLI METIN / SAYI
' ============================================================

' Hata degeri (#N/A vb.) iceren hucrede CStr cokmesin
Public Function GuvenliMetin(ByVal x As Variant) As String
    If IsError(x) Then
        GuvenliMetin = ""
    ElseIf IsNull(x) Then
        GuvenliMetin = ""
    ElseIf IsObject(x) Then
        GuvenliMetin = ""
    ElseIf IsArray(x) Then
        GuvenliMetin = ""
    Else
        GuvenliMetin = CStr(x)
    End If
End Function


' = + - @ ile baslayan metin hucreye yazilinca Excel onu formul sanar.
' Basina gorunmeyen tek tirnak konur.
Public Function HucreMetni(ByVal s As String) As String
    HucreMetni = s
    If Len(s) > 0 Then
        If InStr("=+-@", Left$(s, 1)) > 0 Then HucreMetni = "'" & s
    End If
End Function


' Cok satirli metni Rapor hucresi icin tek satira indirir
Public Function TekSatir(ByVal s As String) As String
    s = Replace(s, vbNewLine, " | ")
    s = Replace(s, vbLf, " | ")
    s = Replace(s, vbCr, " | ")
    s = Trim$(s)
    Do While Right$(s, 2) = " |"
        s = Trim$(Left$(s, Len(s) - 2))
    Loop
    TekSatir = s
End Function


' MsgBox ~1024 karakterden sonrasini sessizce keser
Public Function Sigdir(ByVal metin As String, ByVal sinir As Long) As String
    If Len(metin) <= sinir Then
        Sigdir = metin
    Else
        Sigdir = Left$(metin, sinir) & vbNewLine & "   ... (kisaltildi; tamami Rapor sayfasinda)"
    End If
End Function


' Madde listesini ("- madde" satirlari) MsgBox sinirina sigdirir.
'   - Her madde en az MADDE_EN_AZ karakter gorunur (sutun adlari ve ornek
'     satir numaralari kesilmesin); yer varsa daha fazlasi.
'   - En fazla MADDE_EN_FAZLA madde gosterilir; sigmayanlar tek satirda
'     "... ve X madde daha (tamami Rapor sayfasinda)" olarak ozetlenir.
' Maddelerin TAM hali her zaman Rapor'a yazilir (OnKontrol, OnayAl).
Public Function MaddeKisalt(ByVal liste As String, ByVal sinir As Long) As String

    Const MADDE_EN_AZ As Long = 110
    Const MADDE_EN_FAZLA As Long = 10
    Const OZET_PAYI As Long = 60

    Dim parcalar() As String, maddeler() As String
    Dim i As Long, n As Long, her As Long, gosterilen As Long
    Dim sonuc As String, satir As String

    If Len(liste) <= sinir Then
        MaddeKisalt = liste
        Exit Function
    End If

    parcalar = Split(liste, vbNewLine)
    ReDim maddeler(0 To UBound(parcalar))
    For i = 0 To UBound(parcalar)
        If Trim$(parcalar(i)) <> "" Then
            maddeler(n) = parcalar(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        MaddeKisalt = Sigdir(liste, sinir)
        Exit Function
    End If

    her = (sinir - OZET_PAYI) \ IIf(n < MADDE_EN_FAZLA, n, MADDE_EN_FAZLA) - 2
    If her < MADDE_EN_AZ Then her = MADDE_EN_AZ

    For i = 0 To n - 1
        satir = maddeler(i)
        If Len(satir) > her Then satir = Left$(satir, her - 3) & "..."
        If gosterilen >= MADDE_EN_FAZLA Then Exit For
        If gosterilen > 0 And Len(sonuc) + Len(satir) + 2 > sinir - OZET_PAYI Then Exit For
        sonuc = sonuc & satir & vbNewLine
        gosterilen = gosterilen + 1
    Next i

    If gosterilen < n Then
        sonuc = sonuc & "- ... ve " & (n - gosterilen) & " madde daha (tamami Rapor sayfasinda)" & vbNewLine
    End If
    MaddeKisalt = sonuc

End Function


' Tek hucreli aralik da olsa her zaman (satir, sutun) dizisi doner
Public Function IkiBoyutlu(ByVal rng As Range) As Variant
    Dim a() As Variant
    If rng.Cells.Count = 1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = rng.Value
        IkiBoyutlu = a
    Else
        IkiBoyutlu = rng.Value
    End If
End Function


' Windows oturum adi (Citrix'te herkes icin ayni cikabilen Excel
' kullanici adindan daha guvenilir)
Public Function KullaniciAdi() As String
    Dim s As String
    On Error Resume Next
    s = Environ$("USERNAME")
    If s = "" Then s = Environ$("USER")
    On Error GoTo 0
    If s = "" Then s = Application.UserName
    KullaniciAdi = s
End Function


Public Function TarihMetni(ByVal x As Variant) As String
    If IsDate(x) Then
        TarihMetni = Format$(x, "dd.mm.yyyy hh:nn")
    Else
        TarihMetni = GuvenliMetin(x)
    End If
End Function


Public Function GuvenliSayi(ByVal x As Variant) As Long
    GuvenliSayi = 0
    If IsError(x) Then Exit Function
    If IsNumeric(x) And Trim$(GuvenliMetin(x)) <> "" Then
        On Error Resume Next
        GuvenliSayi = CLng(x)
        On Error GoTo 0
    End If
End Function


' Hucre degeri sayi mi? Metin olarak yazilmis sayilari da kabul eder.
Public Function SayiDegeri(ByVal x As Variant, ByRef deger As Double) As Boolean

    SayiDegeri = False
    deger = 0
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function

    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            deger = CDbl(x)
            SayiDegeri = True
        Case vbDate
            deger = CDbl(x)
            SayiDegeri = True
        Case vbString
            If Trim$(x) <> "" And IsNumeric(Trim$(x)) Then
                On Error Resume Next
                deger = CDbl(Trim$(x))
                SayiDegeri = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0
            End If
    End Select

End Function


' Hucre degeri gecerli bir TARIH mi? (Excel'in ISNUMBER mantigi:
' gercek tarih veya tarih seri numarasi). Metin tarih SAYILMAZ.
Public Function TarihSayisi(ByVal x As Variant, ByRef deger As Double) As Boolean

    TarihSayisi = False
    deger = 0
    If IsError(x) Or IsEmpty(x) Then Exit Function

    Select Case VarType(x)
        Case vbDate
            deger = CDbl(x)
            TarihSayisi = True
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            deger = CDbl(x)
            TarihSayisi = (deger > 0 And deger < 2958466)
    End Select

End Function


' ENGEL/DIKKAT listesine bir madde ekler
Public Sub MaddeEkle(ByRef liste As String, ByRef sayi As Long, ByVal madde As String)
    liste = liste & "- " & madde & vbNewLine
    sayi = sayi + 1
End Sub


Public Function SayiMetni(ByVal x As Double) As String
    SayiMetni = Format$(x, "#,##0")
End Function


Public Function TutarMetni(ByVal x As Double) As String
    TutarMetni = Format$(x, "#,##0.00")
End Function


' ============================================================
'  SUTUN HARFI / NUMARASI
' ============================================================

' "C - police_no" -> "police_no"
Public Function SutunAdi(ByVal metin As Variant) As String

    Dim s As String
    Dim p As Long

    s = Trim$(GuvenliMetin(metin))
    p = InStr(s, " - ")
    If p > 0 Then
        SutunAdi = Trim$(Mid$(s, p + 3))
    Else
        SutunAdi = s
    End If

End Function


' "C - police_no" -> 3. Gecersiz metinde COKMEZ, 0 dondurur.
Public Function SutunNo(ByVal metin As Variant) As Long

    Dim harf As String, c As String
    Dim p As Long, i As Long, k As Long

    SutunNo = 0
    harf = BH(Trim$(GuvenliMetin(metin)))
    p = InStr(harf, " - ")
    If p > 0 Then harf = Trim$(Left$(harf, p - 1))

    If Len(harf) = 0 Or Len(harf) > 3 Then Exit Function
    For i = 1 To Len(harf)
        c = Mid$(harf, i, 1)
        If c < "A" Or c > "Z" Then Exit Function
    Next i

    k = 0
    For i = 1 To Len(harf)
        k = k * 26 + (Asc(Mid$(harf, i, 1)) - 64)
    Next i

    If k < 1 Or k > 16384 Then Exit Function
    SutunNo = k

End Function


' 3 -> "C"
Public Function SutunHarfi(ByVal no As Long) As String

    Dim k As Long, kalan As Long
    Dim s As String

    If no < 1 Or no > 16384 Then Exit Function
    k = no
    Do While k > 0
        kalan = (k - 1) Mod 26
        s = Chr$(65 + kalan) & s
        k = (k - 1 - kalan) \ 26
    Loop
    SutunHarfi = s

End Function


Public Function SutunEtiketi(ByVal no As Long, ByVal baslik As String) As String
    SutunEtiketi = SutunHarfi(no) & " - " & baslik
End Function


' ============================================================
'  DOSYA / SAYFA
' ============================================================

' Panelde yazan degeri dosyaya cevirir. Deger tam yol olabilir ya da
' acik bir dosyanin adi (uzantili/uzantisiz). Bulamazsa ve yol gibi
' gorunuyorsa dosyayi acmayi dener. Bulamazsa Nothing.
Public Function DosyaBul(ByVal ad As Variant) As Workbook

    Dim wb As Workbook
    Dim aranan As String, mevcut As String

    aranan = BH(YoluTemizle(ad))
    If aranan = "" Then Exit Function

    For Each wb In Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            If BH(Trim$(wb.FullName)) = aranan Then
                Set DosyaBul = wb
                Exit Function
            End If
            mevcut = BH(Trim$(wb.Name))
            If mevcut = aranan Then
                Set DosyaBul = wb
                Exit Function
            End If
            If UzantisiSil(mevcut) = UzantisiSil(aranan) Then
                Set DosyaBul = wb
                Exit Function
            End If
        End If
    Next wb

    If InStr(aranan, "\") > 0 Or InStr(aranan, "/") > 0 Then
        Set DosyaBul = DosyaAc(CStr(ad))
    End If

End Function


Public Function YolDurumu(ByVal deger As Variant) As String

    Dim s As String, varMi As String
    Dim wb As Workbook

    s = YoluTemizle(deger)
    If s = "" Then
        YolDurumu = "(bos)"
        Exit Function
    End If

    For Each wb In Workbooks
        If BH(Trim$(wb.FullName)) = BH(s) Or BH(Trim$(wb.Name)) = BH(s) Then
            YolDurumu = "ACIK" & IIf(wb.ReadOnly, " (SALT OKUNUR)", "")
            Exit Function
        End If
    Next wb

    If InStr(s, "\") > 0 Or InStr(s, "/") > 0 Then
        On Error Resume Next
        varMi = Dir(s)
        On Error GoTo 0
        If varMi <> "" Then
            YolDurumu = "yol gecerli, henuz acik degil"
        Else
            YolDurumu = "Dir ile gorunmuyor (SharePoint/uzun yol olabilir)"
        End If
    Else
        YolDurumu = "acik degil"
    End If

End Function


' Yapistirilan yolun bastaki/sondaki tirnak ve bosluklarini atar
Public Function YoluTemizle(ByVal deger As Variant) As String

    Dim s As String

    s = Trim$(GuvenliMetin(deger))
    Do While Len(s) > 0
        If Left$(s, 1) = Chr$(34) Or Left$(s, 1) = "'" Then
            s = Mid$(s, 2)
        ElseIf Right$(s, 1) = Chr$(34) Or Right$(s, 1) = "'" Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    YoluTemizle = Trim$(s)

End Function


' Verilen yoldaki dosyayi acar; acamazsa Nothing. VBA ile acilan
' dosyalar Korumali Gorunum'e takilmaz. SharePoint adresleri Dir ile
' gorunmez ama Excel acabilir; bu yuzden Dir'e bakilmadan denenir.
Public Function DosyaAc(ByVal yolHam As String) As Workbook

    Dim wb As Workbook
    Dim yol As String

    yol = YoluTemizle(yolHam)
    If yol = "" Then Exit Function

    DurumCubugu "Dosya aciliyor: " & yol
    Err.Clear
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=yol, UpdateLinks:=0)
    If wb Is Nothing Then
        g_SonAcmaHatasi = "Acilamadi: " & yol & vbNewLine & _
                          "Excel hatasi: " & Err.Number & " - " & Err.Description
    Else
        g_SonAcmaHatasi = ""
    End If
    Err.Clear
    On Error GoTo 0
    DurumCubugu ""

    Set DosyaAc = wb

End Function


Public Function UzantisiSil(ByVal ad As String) As String
    Dim p As Long
    p = InStrRev(ad, ".")
    If p > 1 Then
        UzantisiSil = Left$(ad, p - 1)
    Else
        UzantisiSil = ad
    End If
End Function


' Sayi/tarih gibi gorunen sayfa adlari da (orn. "2026", "09.2026")
' metin olarak karsilastirilir.
Public Function SayfaBul(ByVal wb As Workbook, ByVal ad As Variant) As Worksheet

    Dim ws As Worksheet
    Dim aranan As String

    If wb Is Nothing Then Exit Function
    aranan = BH(Trim$(GuvenliMetin(ad)))
    If aranan = "" Then Exit Function

    For Each ws In wb.Worksheets
        If BH(Trim$(CStr(ws.Name))) = aranan Then
            Set SayfaBul = ws
            Exit Function
        End If
    Next ws

End Function


Public Function DosyaAdiGuvenli(ByVal wb As Workbook) As String
    If wb Is Nothing Then
        DosyaAdiGuvenli = "BULUNAMADI"
    Else
        DosyaAdiGuvenli = wb.Name
    End If
End Function


Public Function SayfaAdiGuvenli(ByVal ws As Worksheet) As String
    If ws Is Nothing Then
        SayfaAdiGuvenli = "BULUNAMADI"
    Else
        SayfaAdiGuvenli = CStr(ws.Name)
    End If
End Function


Public Function SayfaVarMi(ByVal ad As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ad)
    On Error GoTo 0
    SayfaVarMi = Not (ws Is Nothing)
End Function


' Sayfa yoksa en sona ekler. Excel yeni sayfayi aktif yaptigi icin
' onceki aktif sayfaya geri donulur.
Public Function SayfaOlustur(ByVal ad As String) As Worksheet

    Dim ws As Worksheet
    Dim eskiSayfa As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set eskiSayfa = ActiveSheet
        On Error GoTo 0

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ad

        On Error Resume Next
        If Not eskiSayfa Is Nothing Then eskiSayfa.Activate
        On Error GoTo 0
    End If

    Set SayfaOlustur = ws

End Function


' Belirli bir araligi TAMAMEN temizler (deger, bicim, menu, not,
' birlestirme). ws.Cells.Clear YERINE kullanilir: menulu sayfada tum
' sayfayi Clear etmek kaydederken XML bozulmasina yol acabiliyor.
Public Sub AlanTemizle(ByVal rng As Range, Optional ByVal bicimDe As Boolean = True)

    On Error Resume Next
    rng.Validation.Delete
    rng.ClearComments
    If bicimDe Then rng.UnMerge
    rng.ClearContents
    If bicimDe Then
        rng.ClearFormats
        rng.Font.Name = PANEL_YAZI_TIPI
        rng.Font.Size = PANEL_YAZI_BOYUT
    End If
    Err.Clear
    On Error GoTo 0

End Sub


' ============================================================
'  OKUMA
' ============================================================

' Bir sutunu blok halinde okur. Tek satir olsa bile HER ZAMAN
' (1 To n, 1 To 1) dizisi dondurur (Excel tek hucrede skaler verir).
'
' TASMA (hata 6): tarih bicimli hucrede tarih araligina sigmayan sayi
' (hucrede #####) .Value okurken Overflow verir. O zaman sutun .Value2
' ile okunur ve g_OkumaUyarisi listesine eklenir (aktarimda DIKKAT).
Public Function SutunOku(ByVal ws As Worksheet, ByVal sutunNoDeger As Long, _
                         ByVal ilk As Long, ByVal sonDeger As Long, _
                         Optional ByVal formulMu As Boolean = False) As Variant

    Dim rng As Range
    Dim arr() As Variant
    Dim sonuc As Variant

    If sonDeger < ilk Then sonDeger = ilk
    Set rng = ws.Range(ws.Cells(ilk, sutunNoDeger), ws.Cells(sonDeger, sutunNoDeger))

    If formulMu Then
        ' R1C1: goreli satir referanslari hedefte kendiliginden kayar
        sonuc = rng.FormulaR1C1
    Else
        On Error Resume Next
        sonuc = rng.Value
        If Err.Number <> 0 Then
            Err.Clear
            sonuc = rng.Value2
            If InStr(g_OkumaUyarisi, ws.Name & "!" & SutunHarfi(sutunNoDeger) & " ") = 0 Then
                g_OkumaUyarisi = g_OkumaUyarisi & ws.Name & "!" & SutunHarfi(sutunNoDeger) & " "
            End If
        End If
        On Error GoTo 0
    End If

    If sonDeger > ilk Then
        SutunOku = sonuc
    Else
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = sonuc
        SutunOku = arr
    End If

End Function


' Sayfanin herhangi bir sutunundaki en alttaki dolu satir (bos ise 0)
Public Function SonDoluSatir(ByVal ws As Worksheet) As Long

    Dim bulunan As Range

    On Error Resume Next
    Set bulunan = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                LookIn:=xlFormulas, LookAt:=xlPart, _
                                SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If bulunan Is Nothing Then
        SonDoluSatir = 0
    Else
        SonDoluSatir = bulunan.Row
    End If

End Function


' Bir sutundaki en alttaki dolu satir (bos ise 0)
Public Function SutunSonDolu(ByVal ws As Worksheet, ByVal sutunNoDeger As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, sutunNoDeger).End(xlUp).Row
    If Trim$(GuvenliMetin(ws.Cells(r, sutunNoDeger).Value)) = "" And _
       Not ws.Cells(r, sutunNoDeger).HasFormula Then r = 0
    SutunSonDolu = r
End Function


Public Function SonSutun(ByVal ws As Worksheet, Optional ByVal sinir As Long = 0) As Long
    Dim n As Long
    n = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If n = 1 And Trim$(GuvenliMetin(ws.Cells(1, 1).Value)) = "" Then n = 0
    If sinir > 0 And n > sinir Then n = sinir
    SonSutun = n
End Function


' ============================================================
'  BENZERLIK
' ============================================================

' Iki sutun adinin benzerligi (0 = alakasiz, 1 = birebir)
Public Function Benzerlik(ByVal a As String, ByVal b As String) As Double

    Dim s1 As String, s2 As String
    Dim mesafe As Long, uzun As Long

    s1 = Normalize(a)
    s2 = Normalize(b)
    If s1 = "" Or s2 = "" Then Exit Function

    If s1 = s2 Then
        Benzerlik = 1
        Exit Function
    End If
    If InStr(s1, s2) > 0 Or InStr(s2, s1) > 0 Then
        Benzerlik = 0.85
        Exit Function
    End If

    mesafe = Levenshtein(s1, s2)
    uzun = Len(s1)
    If Len(s2) > uzun Then uzun = Len(s2)
    Benzerlik = 1 - (mesafe / uzun)

End Function


Public Function Levenshtein(ByVal a As String, ByVal b As String) As Long

    Dim d() As Long
    Dim i As Long, j As Long, maliyet As Long
    Dim n As Long, m As Long
    Dim ca() As Integer, cb() As Integer

    n = Len(a)
    m = Len(b)
    If n = 0 Then
        Levenshtein = m
        Exit Function
    End If
    If m = 0 Then
        Levenshtein = n
        Exit Function
    End If

    ReDim ca(1 To n)
    ReDim cb(1 To m)
    For i = 1 To n
        ca(i) = AscW(Mid$(a, i, 1))
    Next i
    For j = 1 To m
        cb(j) = AscW(Mid$(b, j, 1))
    Next j

    ReDim d(0 To n, 0 To m)
    For i = 0 To n
        d(i, 0) = i
    Next i
    For j = 0 To m
        d(0, j) = j
    Next j

    For i = 1 To n
        For j = 1 To m
            If ca(i) = cb(j) Then
                maliyet = 0
            Else
                maliyet = 1
            End If
            d(i, j) = EnKucuk(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + maliyet)
        Next j
    Next i

    Levenshtein = d(n, m)

End Function


Private Function EnKucuk(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    EnKucuk = a
    If b < EnKucuk Then EnKucuk = b
    If c < EnKucuk Then EnKucuk = c
End Function


' ============================================================
'  VERI TIPI
'
'  Bir sutunun ~300 dolu hucresine bakilir: TARIH / SAYI / METIN
'  (%60 ve ustu cogunluk), KARISIK veya BOS.
' ============================================================

Public Function VeriTipiKaynak(ByVal ws As Worksheet, ByVal no As Long) As String

    Dim son As Long

    VeriTipiKaynak = ""
    If ws Is Nothing Or no < 1 Then Exit Function

    son = ws.Cells(ws.Rows.Count, no).End(xlUp).Row
    If son < 2 Then
        VeriTipiKaynak = "BOS"
        Exit Function
    End If
    If son > TIP_ORNEK + 1 Then son = TIP_ORNEK + 1
    VeriTipiKaynak = VeriTipiDizi(SutunOku(ws, no, 2, son))

End Function


' Ana dosyada en SON 300 satira bakilir (en guncel veri)
Public Function VeriTipiHedef(ByVal ws As Worksheet, ByVal no As Long) As String

    Dim son As Long, ilk As Long

    VeriTipiHedef = ""
    If ws Is Nothing Or no < 1 Then Exit Function

    son = ws.Cells(ws.Rows.Count, no).End(xlUp).Row
    If son < 2 Then
        VeriTipiHedef = "BOS"
        Exit Function
    End If
    ilk = son - TIP_ORNEK + 1
    If ilk < 2 Then ilk = 2
    VeriTipiHedef = VeriTipiDizi(SutunOku(ws, no, ilk, son))

End Function


Public Function VeriTipiDizi(ByVal v As Variant, Optional ByVal enFazla As Long = 0) As String

    Dim tarihSay As Long, sayiSay As Long, metinSay As Long
    Dim ornek As String

    VeriTipiSay v, tarihSay, sayiSay, metinSay, ornek, enFazla
    VeriTipiDizi = TipKarari(tarihSay, sayiSay, metinSay)

End Function


' Dizideki dolu hucreleri tipine gore sayar; ilk dolu degeri ornek
' verir. enFazla > 0 ise o kadar satira bakilir.
Public Sub VeriTipiSay(ByVal v As Variant, ByRef tarihSay As Long, ByRef sayiSay As Long, _
                       ByRef metinSay As Long, ByRef ornek As String, _
                       Optional ByVal enFazla As Long = 0)

    Dim i As Long, son As Long
    Dim x As Variant
    Dim s As String

    tarihSay = 0
    sayiSay = 0
    metinSay = 0
    ornek = ""
    If Not IsArray(v) Then Exit Sub

    son = UBound(v, 1)
    If enFazla > 0 Then
        If LBound(v, 1) + enFazla - 1 < son Then son = LBound(v, 1) + enFazla - 1
    End If

    For i = LBound(v, 1) To son
        x = v(i, LBound(v, 2))
        If IsError(x) Then
            ' atla
        ElseIf IsEmpty(x) Then
            ' atla
        ElseIf VarType(x) = vbDate Then
            tarihSay = tarihSay + 1
            If ornek = "" Then ornek = Format$(x, "dd.mm.yyyy")
        ElseIf VarType(x) = vbString Then
            s = Trim$(x)
            If s <> "" Then
                If ornek = "" Then ornek = s
                If MetinTarihMi(s) Then
                    tarihSay = tarihSay + 1
                ElseIf KodGibiMi(s) Then
                    metinSay = metinSay + 1     ' "0004821", 16+ hane: sayi degil kod
                ElseIf IsNumeric(s) Then
                    sayiSay = sayiSay + 1
                Else
                    metinSay = metinSay + 1
                End If
            End If
        ElseIf VarType(x) = vbBoolean Then
            metinSay = metinSay + 1
            If ornek = "" Then ornek = CStr(x)
        ElseIf IsNumeric(x) Then
            sayiSay = sayiSay + 1
            If ornek = "" Then ornek = CStr(x)
        Else
            metinSay = metinSay + 1
        End If
    Next i

End Sub


' Metin olarak saklanmis numara KODU mu? (yalnizca rakam; basinda sifir
' veya 16+ hane). Bunlar sayi gibi gorunse de police/hasar numarasidir.
Public Function KodGibiMi(ByVal s As String) As Boolean
    KodGibiMi = False
    If Len(s) < 2 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    KodGibiMi = (Left$(s, 1) = "0" Or Len(s) >= 16)
End Function


Public Function TipKarari(ByVal tarihSay As Long, ByVal sayiSay As Long, ByVal metinSay As Long) As String

    Dim toplam As Long

    toplam = tarihSay + sayiSay + metinSay
    If toplam = 0 Then
        TipKarari = "BOS"
    ElseIf tarihSay / toplam >= 0.6 Then
        TipKarari = "TARIH"
    ElseIf sayiSay / toplam >= 0.6 Then
        TipKarari = "SAYI"
    ElseIf metinSay / toplam >= 0.6 Then
        TipKarari = "METIN"
    Else
        TipKarari = "KARISIK"
    End If

End Function


' "01.09.2026", "1/9/26", "2026-09-01", "01.09.2026 00:00" gibi metin tarihleri tanir
Public Function MetinTarihMi(ByVal s As String) As Boolean

    Dim p() As String
    Dim i As Long

    MetinTarihMi = False
    s = Trim$(s)
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    s = Replace(Replace(s, "/", "."), "-", ".")

    p = Split(s, ".")
    If UBound(p) <> 2 Then Exit Function

    For i = 0 To 2
        If Len(p(i)) = 0 Then Exit Function
        If p(i) Like "*[!0-9]*" Then Exit Function
    Next i

    If Len(p(0)) <= 2 And Len(p(1)) <= 2 And (Len(p(2)) = 2 Or Len(p(2)) = 4) Then
        MetinTarihMi = True
    ElseIf Len(p(0)) = 4 And Len(p(1)) <= 2 And Len(p(2)) <= 2 Then
        MetinTarihMi = True
    End If

End Function


' Iki tip kesin olarak farkli mi? (BOS, KARISIK, bilinmeyen = catismaz)
Public Function TipCatisir(ByVal a As String, ByVal b As String) As Boolean
    TipCatisir = False
    If a <> "TARIH" And a <> "SAYI" And a <> "METIN" Then Exit Function
    If b <> "TARIH" And b <> "SAYI" And b <> "METIN" Then Exit Function
    TipCatisir = (a <> b)
End Function


Public Function TipMetni(ByVal a As String, ByVal b As String) As String
    If a = "" Or b = "" Then
        TipMetni = "karsilastirilamadi"
    ElseIf a = b Then
        TipMetni = "ayni (" & a & ")"
    Else
        TipMetni = "aylik " & a & ", ana dosya " & b
    End If
End Function


' Veri tipini bir islem boyunca bir kez hesaplar
Public Function OnbellekTip(ByVal onb As Sozluk, ByVal ws As Worksheet, ByVal no As Long, _
                            ByVal kaynakMi As Boolean) As String

    Dim anahtar As String

    OnbellekTip = ""
    If ws Is Nothing Or no < 1 Then Exit Function

    If kaynakMi Then
        anahtar = "K" & no
    Else
        anahtar = "H" & no
    End If

    If Not onb.Exists(anahtar) Then
        If kaynakMi Then
            onb.Item(anahtar) = VeriTipiKaynak(ws, no)
        Else
            onb.Item(anahtar) = VeriTipiHedef(ws, no)
        End If
    End If
    OnbellekTip = onb.Item(anahtar)

End Function


' ============================================================
'  BASLIK DOGRULAMA  (aktarim ve geri almadan hemen once)
'
'  Panel "F - TUTAR" der; dosyada F'nin basligi hala TUTAR mi?
'  Degilse ayni adli TEK sutun aranir ve no guncellenir.
' ============================================================
Public Function BaslikDogrula(ByVal ws As Worksheet, ByRef no As Long, ByVal beklenen As String) As BaslikDurumu

    Dim hedefAd As String
    Dim sonSut As Long, i As Long, bulunan As Long, adet As Long
    Dim v As Variant

    hedefAd = Normalize(beklenen)
    If hedefAd = "" Then
        BaslikDogrula = bdYerinde
        Exit Function
    End If

    If no > 0 Then
        If Normalize(GuvenliMetin(ws.Cells(1, no).Value)) = hedefAd Then
            BaslikDogrula = bdYerinde
            Exit Function
        End If
    End If

    sonSut = SonSutun(ws, 2000)
    If sonSut < 1 Then
        BaslikDogrula = bdYok
        Exit Function
    End If
    v = IkiBoyutlu(ws.Range(ws.Cells(1, 1), ws.Cells(1, sonSut)))

    For i = 1 To UBound(v, 2)
        If Normalize(GuvenliMetin(v(1, i))) = hedefAd Then
            adet = adet + 1
            bulunan = i
        End If
    Next i

    If adet = 0 Then
        BaslikDogrula = bdYok
    ElseIf adet > 1 Then
        BaslikDogrula = bdBelirsiz
    Else
        no = bulunan
        BaslikDogrula = bdYerDegisti
    End If

End Function


' "C - Tutar", "K - Tutar" gibi ayni adli sutunlari bulur.
' Donus ornegi: "Tutar (F, K); Tarih (B, M)"
Public Function AyniAdlar(ByVal v As Variant) As String

    Dim i As Long, p As Long
    Dim harfler As Sozluk, gorunen As Sozluk
    Dim deger As String, ad As String, harf As String
    Dim anahtar As Variant
    Dim metin As String

    If Not IsArray(v) Then Exit Function
    Set harfler = New Sozluk
    Set gorunen = New Sozluk

    For i = LBound(v, 1) To UBound(v, 1)
        deger = GuvenliMetin(v(i, LBound(v, 2)))
        ad = Normalize(SutunAdi(deger))
        If ad <> "" Then
            p = InStr(deger, " - ")
            If p > 0 Then
                harf = Left$(deger, p - 1)
            Else
                harf = "?"
            End If
            If harfler.Exists(ad) Then
                harfler.Item(ad) = harfler.Item(ad) & ", " & harf
            Else
                harfler.Item(ad) = harf
                gorunen.Item(ad) = SutunAdi(deger)
            End If
        End If
    Next i

    For Each anahtar In harfler.Keys
        If InStr(harfler.Item(anahtar), ",") > 0 Then
            metin = metin & gorunen.Item(anahtar) & " (" & harfler.Item(anahtar) & "); "
        End If
    Next anahtar

    If Len(metin) > 2 Then metin = Left$(metin, Len(metin) - 2)
    AyniAdlar = metin

End Function


' ============================================================
'  RENK
' ============================================================

' Hucrenin EKRANDA GORUNEN dolgu rengi (kosullu bicim / tablo stili
' dahil, DisplayFormat). Dolgu yoksa veya beyazsa False.
Public Function HucreRengi(ByVal hucre As Range, ByRef renk As Long) As Boolean

    Dim ci As Variant

    HucreRengi = False
    renk = RGB(255, 255, 255)

    On Error Resume Next
    ci = hucre.DisplayFormat.Interior.ColorIndex
    If Err.Number = 0 Then
        If ci <> xlNone Then
            renk = hucre.DisplayFormat.Interior.Color
            HucreRengi = True
        End If
    End If
    Err.Clear

    If Not HucreRengi Then
        If hucre.Interior.ColorIndex <> xlNone Then
            renk = hucre.Interior.Color
            HucreRengi = True
        End If
    End If
    Err.Clear
    On Error GoTo 0

    If renk = RGB(255, 255, 255) Then HucreRengi = False

End Function


' Yesil kanali digerlerinden belirgin baskinsa yesil (zorunlu sutun).
Public Function YesilMi(ByVal renk As Long) As Boolean

    Dim r As Long, g As Long, b As Long

    r = renk Mod 256
    g = (renk \ 256) Mod 256
    b = (renk \ 65536) Mod 256
    YesilMi = (g > r + 12) And (g > b + 12)

End Function


' ============================================================
'  ACILIR MENU (Data Validation)
' ============================================================

' Listeler sayfasinin bir sutunundan kati menu kurar (listede olmayan
' deger yazilamaz).
Public Sub MenuKur(ByVal hedef As Range, ByVal lst As Worksheet, ByVal sutunIndex As Long)
    MenuKurOrtak hedef, lst, sutunIndex, True
End Sub


' Menu kurar ama listede olmayan degerin elle yazilmasina da izin verir.
Public Sub MenuKurSerbest(ByVal hedef As Range, ByVal lst As Worksheet, ByVal sutunIndex As Long)
    MenuKurOrtak hedef, lst, sutunIndex, False
End Sub


Private Sub MenuKurOrtak(ByVal hedef As Range, ByVal lst As Worksheet, ByVal sutunIndex As Long, _
                         ByVal kati As Boolean)

    Dim son As Long
    Dim adres As String

    On Error Resume Next
    hedef.Validation.Delete
    On Error GoTo 0

    If Trim$(GuvenliMetin(lst.Cells(1, sutunIndex).Value)) = "" Then Exit Sub
    son = lst.Cells(lst.Rows.Count, sutunIndex).End(xlUp).Row

    adres = "='" & lst.Name & "'!" & _
            lst.Range(lst.Cells(1, sutunIndex), lst.Cells(son, sutunIndex)).Address

    On Error Resume Next
    With hedef.Validation
        If kati Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=adres
        Else
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=adres
            .ShowError = False
        End If
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0

End Sub


' Bir diziyi Listeler sayfasinin bir sutununa tek seferde yazar
' (Application.Transpose KULLANILMAZ: 65.536 siniri ve tek satir hatasi)
'
' metinOlarak = True: sutun once METIN bicimine alinir; "0012", "2026",
' "01.09.2026" gibi degerler Excel tarafindan sayiya/tarihe CEVRILMEZ.
' Filtre ve sayfa menuleri boyle yazilir; aksi halde menuden secilen
' deger kaynaktaki degerle eslesmezdi.
Public Sub ListeYaz(ByVal lst As Worksheet, ByVal sutunIndex As Long, ByVal degerler As Variant, _
                    Optional ByVal metinOlarak As Boolean = False)

    Dim n As Long, i As Long
    Dim cikti() As Variant

    lst.Columns(sutunIndex).ClearContents
    If metinOlarak Then lst.Columns(sutunIndex).NumberFormat = "@"
    If Not IsArray(degerler) Then Exit Sub
    n = UBound(degerler) - LBound(degerler) + 1
    If n < 1 Then Exit Sub

    ReDim cikti(1 To n, 1 To 1)
    For i = 1 To n
        If metinOlarak Then
            cikti(i, 1) = GuvenliMetin(degerler(LBound(degerler) + i - 1))
        Else
            cikti(i, 1) = degerler(LBound(degerler) + i - 1)
            If VarType(cikti(i, 1)) = vbString Then cikti(i, 1) = HucreMetni(CStr(cikti(i, 1)))
        End If
    Next i
    lst.Cells(1, sutunIndex).Resize(n, 1).Value = cikti

End Sub
