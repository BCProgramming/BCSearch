VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormattedTextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Formatted Text Writer Class

Private mInputStream As BCFile.IInputStream



Public Sub Init(WithStream As BCFile.IInputStream)
    Set mInputStream = WithStream
    
End Sub
Public Property Get Row() As Double
    'returns the row position in the file.
    'count the number of linefeeds between the current position and the start of the file.
    Dim CurrPos As Long
    CurrPos = mInputStream.GetPos
End Property
Public Property Get Column() As Double
    'returns the current column position.
    'count the number of characters between the current position and the previous lf.
    Dim CurrPos As Long
    Dim ReadString As String
    CurrPos = mInputStream.GetPos
    'Seek to the start of the file...
    mInputStream.SeekTo 0, STREAM_BEGIN
    'read up to currpos...
    ReadString = ModStreamFunc.ReadString(mInputStream, CurrPos, StrRead_Default)
    


End Property

