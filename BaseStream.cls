VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IInputStream
Implements IOutputStream




Public Function ReadBytes(NumBytes As Long, Optional ByRef Numread As Long) As Byte()
    Err.Raise vbObjectError + NotImplemented, "readbytes", "BaseStream does not provide implementation."

End Function
public Function




Private Sub IInputStream_CloseStream()
'
End Sub

Private Function IInputStream_EOF() As Boolean
'
End Function

Private Function IInputStream_GetPos() As Double
'
End Function

Private Function IInputStream_ReadAllStr(Optional ByVal Stringmode As StringReadMode = -1&) As String
'
End Function

Private Function IInputStream_readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long) As Byte()
'
End Function

Private Function IInputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Double
'
End Function

Private Function IInputStream_Size() As Double
'
End Function

Private Function IInputStream_Valid() As Boolean
'
End Function
