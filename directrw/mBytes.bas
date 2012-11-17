Attribute VB_Name = "mBytes"
Option Explicit

Public Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Public Function HiWord(ByVal dw As Long) As Integer
    HiWord = (dw And &HFFFF0000) \ 65536
End Function

Public Function MAKEWORD(ByVal bLo As Byte, ByVal bHi As Byte) As Integer
    If bHi And &H80 Then
        MAKEWORD = (((bHi And &H7F) * 256) + bLo) Or &H8000
    Else
        MAKEWORD = (bHi * 256) + bLo
    End If
End Function

Public Function MakeDWord(ByVal wLo As Integer, ByVal wHi As Integer) As Long
    MakeDWord = (wHi * 65536) + (wLo And &HFFFF&)
End Function

Public Function LongToUShort(ULong As Long) As Integer
   If ULong > &H7FFF Then
      LongToUShort = CInt(ULong - &H10000)
   Else
      LongToUShort = CInt(ULong)
   End If
End Function

Public Function vbIntToUShort(vbInt As Integer) As Long
   If vbInt < 0 Then
      vbIntToUShort = vbInt + &H10000
   Else
      vbIntToUShort = CLng(vbInt)
   End If
End Function

Public Function MakeFAT12(ByVal wLo As Integer, ByVal wHi As Integer) As Long
    MakeFAT12 = (wHi * &H1000& And &HFFF000) + (wLo And &HFFF&)
End Function

Public Function TrimNULL(ByVal str As String) As String
    If InStr(str, Chr$(0)) > 0& Then
        TrimNULL = Left$(str, InStr(str, Chr$(0)) - 1&)
    Else
        TrimNULL = str
    End If
End Function

