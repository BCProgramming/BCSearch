VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'IDataStream:
'Implemented by any "streamable" object; such as, for example, the filestream and Memstream classes.
'the Interface is extremely basic; but it allows any "stream" type object to access the powers of "CBinaryStream" and (soon to be(?)) CTextStream.
#If False Then
 Const Stream_Read = 1
 Const Stream_Write = 2
 Const Stream_Seek = 4
#End If

Public Enum ESupportedStreamOperations
    Stream_Read = 1
    Stream_Write = 2
    Stream_Seek = 4
    
        
End Enum
Public Enum StreamSeekBase
    STREAM_BEGIN = 0
    STREAM_CURRENT = 1
    STREAM_END = 2
    End Enum
            

Public Function ReadBytes(ByVal NumBytes As Long, Optional ByRef Numread As Long) As Byte()
'NumBytes: Number of Bytes to read.
'NumRead: Number of bytes read from stream.




End Function
Public Sub WriteBytes(BytesWrite() As Byte)
'write the bytes to the stream.

End Sub
Public Property Get Supported() As ESupportedStreamOperations
    'return supported values.
End Property
Public Function SeekTo(ByVal NewPosition As Variant, SeekBase As StreamSeekBase) As Variant
'Seeks to the given position
End Function
Public Function GetSeekPos() As Variant
    'returns current position
End Function

