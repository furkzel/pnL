Attribute VB_Name = "M12_Butunluk"
Option Explicit

' ============================================================
' AKTARIM PANELI  v1.1  -  M12_Butunluk  (PARMAK IZI / KARMA)
'
' CRC-32/IEEE (zlib ile ayni polinom ve sonuc) - kazara degisiklik
' dedektoru, kriptografik imza degildir. Platform API'si gerekmez.
'
' v1.1 HIZLI KARMA (KarmaBaslat / KarmaDeger / KarmaBitir)
'   Her hucre bir tip etiketi + yuk olarak karmaya girer:
'     bos       : etiket 0
'     sayi/tarih: etiket 5 + IEEE-754 Double'in 8 bayti (LSet)
'     metin     : etiket 8 + uzunluk (2 x 16 bit) + UTF-16LE birimleri
'     mantiksal : etiket 11 + 1/0
'     hata      : etiket 10 + hata kodu (CVErr)
'   16 bitlik tablo (65.536 girdi) ile iki bayt bir adimda islenir;
'   sayilar metne cevrilmez. v1.0-dev'in hucre basina Str$ + karakter
'   basina Mid$ yontemine gore hucre basina adim sayisi ~5-10 kat azdir.
'   Sonuc, ayni bayt dizisinin zlib.crc32'siyle aynidir (test_enterprise).
'
' DiziKarmasi: v1.1-dev taslaginin "CRC32=" imzalari icin korunur
'   (yalnizca o imzalari dogrulamak icin; yeni imza uretmez).
' ============================================================

Private crcTablo(0 To 255) As Long
Private crcHazir As Boolean
Private tablo16() As Long
Private hazir16 As Boolean

Private Type CiftBayt
    d As Double
End Type

Private Type DortBirim
    a As Integer
    b As Integer
    c As Integer
    e As Integer
End Type


' ------------------------------------------------------------
'  HIZLI KARMA
' ------------------------------------------------------------
Private Sub Tablo16Kur()
    Dim i As Long, j As Long, x As Long
    If hazir16 Then Exit Sub
    ReDim tablo16(0 To 65535)
    For i = 0 To 65535
        x = i
        For j = 1 To 16
            If (x And 1) <> 0 Then
                If x < 0 Then
                    x = (((x And &H7FFFFFFF) \ 2) Or &H40000000) Xor &HEDB88320
                Else
                    x = (x \ 2) Xor &HEDB88320
                End If
            Else
                If x < 0 Then
                    x = ((x And &H7FFFFFFF) \ 2) Or &H40000000
                Else
                    x = x \ 2
                End If
            End If
        Next j
        tablo16(i) = x
    Next i
    hazir16 = True
End Sub


Public Function KarmaBaslat() As Long
    Tablo16Kur
    KarmaBaslat = &HFFFFFFFF
End Function


' Tek bir 16 bitlik birimi (0..65535) karmaya ekler
Private Sub Birim(ByRef crc As Long, ByVal u As Long)
    Dim x As Long
    x = crc Xor u
    crc = (((x And &H7FFFFFFF) \ 65536) Or ((x < 0) And &H8000&)) Xor tablo16(x And &HFFFF&)
End Sub


Public Sub KarmaDeger(ByRef crc As Long, ByVal x As Variant)

    Dim b() As Byte
    Dim cb As CiftBayt
    Dim db As DortBirim
    Dim j As Long, n As Long, y As Long, s As String

    Select Case VarType(x)
        Case vbEmpty, vbNull
            Birim crc, 0
        Case vbString
            s = x
            n = Len(s)
            Birim crc, 8
            Birim crc, n And &HFFFF&
            Birim crc, (n \ 65536) And &HFFFF&
            If n > 0 Then
                b = s   ' LO-atla
                For j = 0 To UBound(b) Step 2
                    y = crc Xor (b(j) Or (b(j + 1) * 256&))
                    crc = (((y And &H7FFFFFFF) \ 65536) Or ((y < 0) And &H8000&)) Xor tablo16(y And &HFFFF&)
                Next j
            End If
        Case vbBoolean
            Birim crc, 11
            If x Then
                Birim crc, 1
            Else
                Birim crc, 0
            End If
        Case vbError
            Birim crc, 10
            Birim crc, CLng(Val(Mid$(CStr(x), 7))) And &HFFFF&
        Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal, vbByte, vbDate
            cb.d = CDbl(x)
            LSet db = cb   ' LO-atla
            Birim crc, 5
            Birim crc, db.a And &HFFFF&
            Birim crc, db.b And &HFFFF&
            Birim crc, db.c And &HFFFF&
            Birim crc, db.e And &HFFFF&
        Case Else
            s = CStr(x)
            KarmaDeger crc, s
    End Select

End Sub


Public Function KarmaBitir(ByVal crc As Long) As String
    KarmaBitir = Right$("00000000" & Hex$(Not crc), 8)
End Function


' 2-B (n x 1) dizinin tamaminin karmasi
Public Function HizliKarma(ByVal v As Variant) As String
    Dim crc As Long, i As Long
    crc = KarmaBaslat()
    For i = LBound(v, 1) To UBound(v, 1)
        KarmaDeger crc, v(i, LBound(v, 2))
    Next i
    HizliKarma = KarmaBitir(crc)
End Function


' ------------------------------------------------------------
'  v1.1-dev UYUMLULUGU: "CRC32=" imzalarini dogrular
' ------------------------------------------------------------
Private Function SagaSekiz(ByVal x As Long) As Long
    SagaSekiz = (x And &H7FFFFFFF) \ 256
    If x < 0 Then SagaSekiz = SagaSekiz Or &H800000
End Function

Private Sub CrcBaslat()
    Dim i As Long, j As Long, x As Long
    If crcHazir Then Exit Sub
    For i = 0 To 255
        x = i
        For j = 1 To 8
            If (x And 1) <> 0 Then
                x = ((x And &H7FFFFFFF) \ 2 Or IIf(x < 0, &H40000000, 0)) Xor &HEDB88320
            Else
                x = (x And &H7FFFFFFF) \ 2 Or IIf(x < 0, &H40000000, 0)
            End If
        Next j
        crcTablo(i) = x
    Next i
    crcHazir = True
End Sub

Private Sub CrcMetin(ByRef crc As Long, ByVal s As String)
    Dim i As Long, kod As Long
    For i = 1 To Len(s)
        kod = AscW(Mid$(s, i, 1)) And &HFFFF&
        crc = SagaSekiz(crc) Xor crcTablo((crc Xor (kod And 255)) And 255)
        crc = SagaSekiz(crc) Xor crcTablo((crc Xor (kod \ 256)) And 255)
    Next i
End Sub

Public Function DiziKarmasi(ByVal v As Variant) As String
    Dim crc As Long, i As Long, x As Variant, s As String, tur As String
    CrcBaslat
    crc = &HFFFFFFFF
    For i = LBound(v, 1) To UBound(v, 1)
        x = v(i, 1)
        tur = CStr(VarType(x))
        If IsError(x) Then
            s = CStr(x)
        ElseIf IsEmpty(x) Or IsNull(x) Then
            s = ""
        ElseIf VarType(x) = vbString Then
            s = CStr(x)
        ElseIf VarType(x) = vbBoolean Then
            s = IIf(x, "1", "0")
        Else
            s = Trim$(Str$(CDbl(x)))
        End If
        CrcMetin crc, tur & ":" & CStr(Len(s)) & ":" & s & ";"
    Next i
    DiziKarmasi = Right$("00000000" & Hex$(Not crc), 8)
End Function
