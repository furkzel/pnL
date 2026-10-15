VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sozluk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Sozluk  -  Scripting.Dictionary yerine YEREL (native) sozluk
'
' Neden: kurumsal Citrix / Mac ortamlarinda "Scripting.Runtime"
' COM kutuphanesi engelli olabiliyor. Bu sinif yalnizca VBA'nin
' kendi Collection nesnesini ve dizileri kullanir.
'
' KULLANIM (Dictionary'ye cok benzer, ama varsayilan uye YOKTUR;
' kod elle yapistirildiginda Attribute satirlari kaybolabildigi icin
' her zaman acik yazim kullanilir):
'     Dim s As Sozluk
'     Set s = New Sozluk
'     s.Item("ANAHTAR") = 5          ' ekle / guncelle
'     If s.Exists("ANAHTAR") Then ...
'     x = s.Item("YOK")              ' olmayan anahtar -> Empty (EKLEMEZ)
'     s.Artir "ANAHTAR"              ' sayac (+1), yeni degeri dondurur
'     For Each a In s.Keys ...       ' ekleme sirasiyla
'
' ONEMLI: Anahtarlar BUYUK/KUCUK HARF AYIRMAZ (Collection gibi).
' Kod zaten anahtarlari BH() veya Normalize() ile sadelestirerek
' kullanir; ayrim gerekmez.
' ============================================================

Private m_Indeks As Collection      ' "k" & anahtar -> slot no
Private m_Anahtar() As Variant
Private m_Deger() As Variant
Private m_Dolu() As Boolean
Private m_Sayi As Long              ' canli kayit sayisi
Private m_Son As Long               ' kullanilan en buyuk slot


Private Sub Class_Initialize()
    RemoveAll
End Sub


Public Sub RemoveAll()
    Set m_Indeks = New Collection
    ReDim m_Anahtar(1 To 16)
    ReDim m_Deger(1 To 16)
    ReDim m_Dolu(1 To 16)
    m_Sayi = 0
    m_Son = 0
End Sub


Private Function Slot(ByVal anahtar As Variant) As Long
    Slot = 0
    On Error Resume Next
    Slot = m_Indeks.Item("k" & CStr(anahtar))
    Err.Clear
    On Error GoTo 0
End Function


Private Function YeniSlot(ByVal anahtar As Variant) As Long

    Dim yeniBoy As Long

    m_Son = m_Son + 1
    If m_Son > UBound(m_Anahtar) Then
        yeniBoy = UBound(m_Anahtar) * 2
        ReDim Preserve m_Anahtar(1 To yeniBoy)
        ReDim Preserve m_Deger(1 To yeniBoy)
        ReDim Preserve m_Dolu(1 To yeniBoy)
    End If

    m_Indeks.Add m_Son, "k" & CStr(anahtar)
    m_Anahtar(m_Son) = anahtar
    m_Dolu(m_Son) = True
    m_Sayi = m_Sayi + 1
    YeniSlot = m_Son

End Function


Public Function Exists(ByVal anahtar As Variant) As Boolean
    Exists = (Slot(anahtar) > 0)
End Function


Public Property Get Count() As Long
    Count = m_Sayi
End Property


' Olmayan anahtar icin Empty doner ve anahtari EKLEMEZ.
Public Property Get Item(ByVal anahtar As Variant) As Variant
    Dim i As Long
    i = Slot(anahtar)
    If i = 0 Then
        Item = Empty
    ElseIf IsObject(m_Deger(i)) Then
        Set Item = m_Deger(i)
    Else
        Item = m_Deger(i)
    End If
End Property


Public Property Let Item(ByVal anahtar As Variant, ByVal deger As Variant)
    Dim i As Long
    i = Slot(anahtar)
    If i = 0 Then i = YeniSlot(anahtar)
    m_Deger(i) = deger
End Property


Public Property Set Item(ByVal anahtar As Variant, ByVal deger As Variant)
    Dim i As Long
    i = Slot(anahtar)
    If i = 0 Then i = YeniSlot(anahtar)
    Set m_Deger(i) = deger
End Property


' Ekle (varsa gunceller; Dictionary.Add gibi hata VERMEZ)
Public Sub Add(ByVal anahtar As Variant, ByVal deger As Variant)
    Item(anahtar) = deger
End Sub


' Sayac: anahtarin degerini artirir (yoksa 0'dan baslar), yeni degeri dondurur
Public Function Artir(ByVal anahtar As Variant, Optional ByVal miktar As Double = 1) As Double
    Dim i As Long
    i = Slot(anahtar)
    If i = 0 Then
        i = YeniSlot(anahtar)
        m_Deger(i) = 0
    End If
    If IsEmpty(m_Deger(i)) Or Not IsNumeric(m_Deger(i)) Then m_Deger(i) = 0
    m_Deger(i) = m_Deger(i) + miktar
    Artir = m_Deger(i)
End Function


Public Sub Remove(ByVal anahtar As Variant)
    Dim i As Long
    i = Slot(anahtar)
    If i = 0 Then Exit Sub
    m_Indeks.Remove "k" & CStr(anahtar)
    m_Dolu(i) = False
    m_Anahtar(i) = Empty
    m_Deger(i) = Empty
    m_Sayi = m_Sayi - 1
End Sub


' Anahtarlar, ekleme sirasiyla. Bos sozlukte bos dizi (For Each calisir).
Public Function Keys() As Variant
    Dim a() As Variant
    Dim i As Long, n As Long
    If m_Sayi = 0 Then
        Keys = Array()
        Exit Function
    End If
    ReDim a(0 To m_Sayi - 1)
    For i = 1 To m_Son
        If m_Dolu(i) Then
            a(n) = m_Anahtar(i)
            n = n + 1
        End If
    Next i
    Keys = a
End Function


' Degerler, ekleme sirasiyla (nesne degerler desteklenmez).
Public Function Items() As Variant
    Dim a() As Variant
    Dim i As Long, n As Long
    If m_Sayi = 0 Then
        Items = Array()
        Exit Function
    End If
    ReDim a(0 To m_Sayi - 1)
    For i = 1 To m_Son
        If m_Dolu(i) Then
            If Not IsObject(m_Deger(i)) Then a(n) = m_Deger(i)
            n = n + 1
        End If
    Next i
    Items = a
End Function
