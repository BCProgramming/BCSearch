VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Files collection.

'If the files() collection is retrieved directly from a CDirectory Object, the DirConnect Property will be set.
Private mDirConnect As Directory
Private mCol As Collection 'either mDirConnect can be valid, or mCol can be valid- not both.
'Private mEnumerator As CFileWalker
Friend Sub ConnectDir(ToDir As Directory)
    Set mDirConnect = ToDir
End Sub
'Private mCol As Collection
Public Property Get Count() As Long
    'reimplement- whats the fastest way to get the count of files in a dir?
    If Not mCol Is Nothing Then
        Count = mCol.Count
    End If
End Property
Public Function Add(AddIt As Variant) As CFile
'current implementation:
'technically, Add should add a file to dirconnect- HOWEVER- we want the files collection to be valid
'for storing files from different directories. In order to facilitate this, it is necessary to cleverly determine, essentially, which "MODE" is being used
'adding a file immediately disables Directory walking abilities



If IsObject(AddIt) Then
    If TypeOf AddIt Is CFile Then
        If mCol Is Nothing Then Set mCol = New Collection
        mCol.Add AddIt, AddIt.Filename
        Set Add = AddIt
    Else
        Err.Raise 13, "Files.Add", "Only CFile objects can be added to the File Collection."
    End If
Else

    Debug.Assert False
End If




End Function
Public Function Item(ByVal VntIndexKey As Variant) As CFile
Attribute Item.VB_UserMemId = 0
'this implementation requires strict revision as well.
    Set Item = mCol.Item(VntIndexKey)
End Function
Public Function NewEnum() As stdole.IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
    If Not mCol Is Nothing Then
        Set NewEnum = mCol.[_NewEnum]
    End If
    
    
    Dim enumerator As New EnumVARIANT
    Dim tmp As New IVariantEnumerable
    Dim Walker As CFileWalker
    Set Walker = New CFileWalker
    
    Set tmp = Walker
    enumerator.Init ObjPtr(tmp)
    
    Set NewEnum = enumerator
    Set enumerator = Nothing
    Set tmp = Nothing
    
    
    
End Function
Public Function GetAsString() As String
    Dim LoopFile As CFile, retit As String, flfirst As Boolean
    For Each LoopFile In mCol
        If Not flfirst Then retit = retit & ","
        retit = retit & LoopFile.Filename
    Next
    GetAsString = Left$(retit, Len(retit) - 1)
End Function

Private Sub Class_Initialize()
'    Set mCol = New Collection
End Sub
