VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFileSystemObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IFileSystemObject:
'Implemented by CFile,CDirectory and the soon to be CDrive classes.


Public Function Self() As Object
    'returns actual object.
End Function
Public Function DisplayName() As String 'return display name of this object.
    '
End Function
Public Function GetIcon(Optional ByVal IconSizeParams As IconSizeConstants = Icon_Shell) As Long
    'return an hIcon used for this object.
End Function
