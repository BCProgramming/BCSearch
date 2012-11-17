VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStreamFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub FilterStream(inputstream As IInputStream, outputstream As IOutputStream, Optional ByVal flReverse As Boolean = False)


'flReverse: indicates to do reverse action; ie, if action is compression, perform decompression.
    'perform whatever filter routine is being performed- compression, decompression, etc.

End Sub
