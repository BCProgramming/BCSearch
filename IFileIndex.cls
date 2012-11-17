VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFileIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IFileIndex: implemented by objects that store/retrieve search indexes.

'improvements to this interface should go into IFileIndex2; this is a extremely basic implementation that essentially only allows for a few items to be stored.



Public Function GetIndexEntries(ByVal StrFileMask As String) As String()
'
End Function
