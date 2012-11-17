VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFileSearchCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IFileSearchCallback.
'notify when files are found, the search completes, etc.

Public Sub Found(Sender As Object, Found As String, Optional ByRef Cancel As Boolean, Optional FiltersFound As Variant)
'
'Found=Full path of found object. (Why not the object? takes time!)
'Note that unless being used with CFileSearchEx and filters that use a regular expression to search within the file, REMatchCol() will be empty.
'Otherwise, it will be an array of MatchCollections, where each MatchCollection represents the matches found with a single filter.
End Sub
Public Sub ExecuteComplete(Sender As Object)
'
End Sub
Public Function AllowRecurse(InDir As String) As Boolean
    'return wether to recurse into the passed directory.
End Function
Public Sub ProgressMessage(ByVal StrMessage As String)
    '
End Sub
Public Function Cancelled() As Boolean
    '
End Function


Public Sub SearchError(ErrCode As Long, ErrDesc As String, ByRef Cancel As Boolean)
    'set cancel to true to stop the search.
    
End Sub
