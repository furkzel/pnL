Attribute VB_Name = "M12_Butunluk"
Option Explicit

' CRC32/IEEE of typed, length-delimited UTF-16LE values, in row order.
' Change detector, not a cryptographic signature. No platform API required.
Private crcTablo(0 To 255) As Long
Private crcHazir As Boolean

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
