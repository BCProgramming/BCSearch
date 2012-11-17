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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Files collection.
Private WithEvents mWalker As CDirWalker
Attribute mWalker.VB_VarHelpID = -1
'Private mEnumerator As VBFCustomCollection.IEnumVARIANT
Private mDirConnect As String
'Private mCol As collection
Public Function GetAll(Optional ByVal FileSpec As String = "*") As FilesCollection
    Dim retcollection As FilesCollection, LoopFile As CFile, loopobj As Object
    Set retcollection = New FilesCollection
    'For Each LoopFile In Me
    Dim gotWalker As CDirWalker
    Set gotWalker = Me.GetWalker(FileSpec)
    
    Do Until gotWalker.GetNext(loopobj) Is Nothing
        Set LoopFile = loopobj
        retcollection.Add LoopFile
    
    'Next
    Loop
    Set GetAll = retcollection

End Function
Public Function ItemByIndex(ByVal Index As Long) As CFile
    Dim gotWalker As CDirWalker, I As Long, currloop As Object
    Set gotWalker = FileSystem.GetDirectory(mDirConnect).GetWalker("*", 0)
    For I = 1 To Index
        Set currloop = gotWalker.GetNext
        If currloop Is Nothing Then
            Err.Raise 9, "Files::ItemByIndex", "Index out of bounds"
        
        End If
    Next I
    Set ItemByIndex = currloop

End Function
Public Function Item(ByVal StrFileSpec As String) As CFile
    Dim gotWalker As CDirWalker
    Set gotWalker = FileSystem.GetDirectory(mDirConnect).GetWalker(StrFileSpec, 0)
    Set Item = gotWalker.GetNext
    Set gotWalker = Nothing
    
    




End Function
Public Function Items(ByVal StrFileSpec As String) As FilesCollection
    Dim makeCollection As FilesCollection
    Dim gotWalker As CDirWalker
    Dim loopobject As Object
    Set makeCollection = New FilesCollection
    Set gotWalker = FileSystem.GetDirectory(mDirConnect).GetWalker(StrFileSpec, 0)
    Do Until gotWalker.GetNext(loopobject) Is Nothing
        If TypeOf loopobject Is CFile Then
        
            makeCollection.Add loopobject
        End If
       
    Loop
    
    Set Items = makeCollection
    
End Function
Public Function GetWalker(Optional ByVal FileSpec As String = "*") As CDirWalker
'
    CDebug.Post "Files::GetWalker...""" & mDirConnect & """"
    Set GetWalker = FileSystem.GetDirectory(mDirConnect).GetWalker(FileSpec, 0, FILE_ATTRIBUTE_DIRECTORY)
End Function
Friend Sub Init(InDir As Directory, Walker As CDirWalker)
'
'must be given a parent dir and a Walker.

'Set mEnumerator = Walker.GetIEnum

Set mWalker = Walker
mDirConnect = InDir.Path
End Sub
Public Property Get count() As Long
'    Count = mCol.Count
'quick little routine...
'enumerate all the files in the given directory...
Dim hfind As Long
Dim finddata As WIN32_FIND_DATA
Dim retval As Long, currcount As Long
Debug.Print "finding " & mDirConnect & mWalker.FSpec
hfind = FindFirstFile(mDirConnect & mWalker.FSpec, finddata)

If hfind <= 0 Then
    count = 0
Else
    'currcount = 1
    Do
    
    retval = FindNextFile(hfind, finddata)
    'optimized: no zeromemory needed- we are only counting the files, so we just need to check the returned attributes to mask out directories.
    
    
    If Not ((finddata.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) And retval <> 0 Then
        'don't count directories...
        If Not ((Left$(finddata.cFileName, 1) = ".") Or (Left$(finddata.cFileName, 2) = "..")) Then
            currcount = currcount + 1
        End If
    End If
    
    Loop Until retval <= 0
    
    FindClose hfind
    
End If

count = currcount

End Property

''Public Function NewEnum() As IUnknown
''    Set NewEnum = mEnumerator
''End Function


'Public Function Add(AddIt As Variant) As CFile
'
'If IsObject(AddIt) Then
'    If TypeOf AddIt Is CFile Then
'        mCol.Add AddIt, AddIt.Name
'        Set Add = AddIt
'    Else
'        Err.Raise 13, "Files.Add", "Only CFile objects can be added to the File Collection."
'    End If
'Else
'
'    Debug.Assert False
'End If
'
'
'
'
'End Function
'Public Function Item(ByVal VntIndexKey As Variant) As CFile
'    Set Item = mCol.Item(VntIndexKey)
'End Function
'Public Function NewEnum() As IUnknown
'    Set NewEnum = mCol.[_NewEnum]
'End Function
'Public Function GetAsString() As String
'    Dim LoopFile As CFile, retit As String, flfirst As Boolean
'    For Each LoopFile In mCol
'        If Not flfirst Then retit = retit & ","
'        retit = retit & LoopFile.Name
'    Next
'    GetAsString = Left$(retit, Len(retit) - 1)
'End Function
'
'Private Sub Class_Initialize()
'    Set mCol = New collection
'End Sub

Private Sub Class_Initialize()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub

Private Sub mWalker_ItemFound(FoundItem As Object, Exclude As Boolean)
'
End Sub
