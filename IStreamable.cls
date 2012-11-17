VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStreamable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IStreamable; implemented by objects, such as files, whose handles can be used to open streams (CBinaryStream for example)

'Size
'Fullpath
Public Property Get Size() As Long
'
End Property
Public Property Get Fullpath() As String
'
End Property
Public Property Get Name() As String
'
End Property
Public Function OpenOutputStream() As IOutputStream
'
End Function
Public Function OpenInputStream() As IInputStream
'
End Function
