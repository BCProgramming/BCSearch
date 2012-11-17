Attribute VB_Name = "ModEndian"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type TwoInts
    Ints(1 To 2) As Integer
End Type
Private Type TwoBytes
    Bytes(1 To 2) As Byte
End Type


Public Function SwapEndianInt(ByVal InInt As Integer) As Integer
    Dim newint As Integer, tempbyte As Byte
    Dim structuse As TwoBytes
    CopyMemory structuse, InInt, Len(InInt)
    tempbyte = structuse.Bytes(1)
    structuse.Bytes(1) = structuse.Bytes(2)
    structuse.Bytes(2) = tempbyte
    CopyMemory newint, structuse, Len(newint)
    SwapEndianInt = newint

End Function
Public Function SwapEndianLong(ByVal dw As Long) As Long

  SwapEndianLong = _
      (((dw And &HFF000000) \ &H1000000) And &HFF&) Or _
      ((dw And &HFF0000) \ &H100&) Or _
      ((dw And &HFF00&) * &H100&) Or _
      ((dw And &H7F&) * &H1000000)
  If (dw And &H80&) Then SwapEndianLong = SwapEndianLong Or &H80000000
End Function


