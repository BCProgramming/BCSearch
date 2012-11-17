VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Represents a Stream in Memory.
Implements IOutputStream
Implements IInputStream



Private mBytes() As Byte 'our stream bytes.
Private mCurrLocation As Long 'current location in memory stream.

Private Function IInputStream_GetPos() As Double
    IInputStream_GetPos = mCurrLocation
End Function
