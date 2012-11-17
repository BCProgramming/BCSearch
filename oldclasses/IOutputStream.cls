VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOutputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum StringReadMode
    StrRead_Default = -1
    StrRead_ANSI = 0
    StrRead_Unicode = 1
End Enum
Public Enum FileSeekBase
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2


End Enum


'OutputStream
Public Function GetPos() As Double
    'return current position in stream/File.
End Function

Public Function SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = FILE_BEGIN)
    'Seek to the specified position in the file.
End Function
Public Sub WriteString(Stringwrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
'
End Sub

Public Function WriteBytes(Bytes() As Byte) As Boolean
    'writes the given bytes to the stream. returns true for success, false otherwise.

End Function
Public Sub CloseStream()
'
End Sub

Public Function Valid() As Boolean
    'a confusing function. returns true if the target stream will work when accessed through this interface.
    
End Function
