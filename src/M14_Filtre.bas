Attribute VB_Name = "M14_Filtre"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M14_Filtre  (ORTAK FILTRE MOTORU)
'
' Panel FILTRE sutunu ve Analiz istek karti AYNI motoru kullanir
' (v1.0'da M06 icinde Private idi).
'
' FILTRE SOZDIZIMI
'   AXA            tam deger (buyuk/kucuk harf, Turkce harf farki yok)
'   AXA, ALLIANZ   birden fazla deger (virgul; deger virgul iceriyorsa
'                  ';' kullan). Metin, sutunda AYNEN varsa tek deger sayilir.
'   AX*  A?A       joker
'   >100 <=500 =0  sayi; >01.01.2026 tarih karsilastirma
'                  (v1.1: "1.5" ve "1,5" yerel ayardan bagimsiz 1,5)
'   <>IPTAL        esit olmayan (birden fazla <> hepsi saglanmali)
'   BOS / DOLU     bos ya da dolu hucre
' ============================================================

Public Const F_ESIT As Long = 1
Public Const F_DESEN As Long = 2
Public Const F_KARSILASTIR As Long = 3
Public Const F_BOS As Long = 4
Public Const F_DOLU As Long = 5

Public Type FiltreOge
    Tur As Long
    Negatif As Boolean
    Metin As String
    Esik As Double
    Op As String
End Type

Public Type FiltreSutun
    Sutun As Long
    Ad As String
    Ham As String
    Adet As Long
    Ogeler() As FiltreOge
End Type


' ============================================================
'  FILTRE MOTORU
' ============================================================
Public Function FiltreAyristir(ByVal ham As String) As FiltreSutun

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
    Dim d As Double

    TekKarsilastirmaMi = False
    op = KarsilastirmaOp(t, geri)
    If op = "" Then Exit Function
    If MetinSayisi(geri, d) Or IsNumeric(geri) Or (IsDate(geri) And InStr(geri, ",") = 0) Then TekKarsilastirmaMi = True

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
    Dim d As Double

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
        If op <> "" And MetinSayisi(geri, d) Then
            ' M3: "1.5" / "1,5" yerel ayardan bagimsiz 1,5 olur
            o.Tur = F_KARSILASTIR
            o.Op = op
            o.Esik = d
        ElseIf op <> "" And IsNumeric(geri) Then
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
Public Sub FiltreBelirsizligiGider(ByRef f As FiltreSutun, ByVal v As Variant)

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
Public Function FiltreUyar(ByRef f As FiltreSutun, ByVal x As Variant) As Boolean

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
