VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IInputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IInputStream- specifies methods that must be implemented by all input streams.
'Note- ERR_NOTIMPLEMENTED can be returned for those methods that are inapplicable.

Public Function GetPos() As Double
    'return current position in stream/File.
End Function

Public Function SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = FILE_BEGIN) As Double
    'Seek to the specified position in the file. Returns old position.
End Function
Public Function ReadAllStr(Optional ByVal Stringmode As StringReadMode = StrRead_Default) As String
'
End Function
Public Function ReadBytes(ByVal NumBytes As Long, Optional ByRef numread As Long) As Byte()
    'Read Bytes- and... return them...

End Function
Public Function EOF() As Boolean
    'returns true if at or beyond end of file. False otherwise.
End Function
Public Function Size() As Double
    '
End Function
Public Sub CloseStream()
'
End Sub
Public Function Valid() As Boolean
    'a confusing function. returns true if the target stream will work when accessed through this interface.
    
End Function
