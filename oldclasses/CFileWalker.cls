VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'CFileWalker
'Directory Walker Class

'The Directory Walker Class can be used to walk all the objects in a Directory, just subdirectories, or just files. It also supports the use of File Specifications.
'will be used for implementation of custom IEnumVARIANT of Files and Directories collections.

Private mFindData As WIN32_FIND_DATA  'last find data
Private mFindHandle As Long
    
Private mDirConnect As Directory 'the directory we are Enumerating.

'Private data used during the "walk"
Private mUseAttr As FileAttributeConstants, FileSpec As String


'Public Sub InitWalker(ByVal InDirectory As Directory, Optional ByVal MustHaveAttr As FileAttributeConstants, Optional ByVal FileSpec As String = "*")
Public Sub InitWalker(ByVal InDirectory As Directory, Optional ByVal FileSpec As String = "*")

    'Initializes this walker object to begin searching.
    If Not mDirConnect Is Nothing Then
        If Not mFindHandle = 0 Then FindClose mFindHandle
        Set mDirConnect = Nothing
    End If
    'zero it out- erase any previous search.
    ZeroMemory mFindData, Len(mFindData)
    'mUseAttr = MustHaveAttr
    'FileSpec = FileSpec
    'mFindData.cFileName = InDirectory.Path & FileSpec & vbNullChar & vbNullChar
    'mFindData.dwFileAttributes = MustHaveAttr
    'FindFirstFile InDirectory.Path & FileSpec, mFindData
    
End Sub

Public Function GetNext() As Object
    'returns: a Directory or a CFile, could be either.
    Dim findnext As Long, flMeetsCondition
    Dim FullFileName As String
    ZeroMemory mFindData, Len(mFindData)
    If mFindHandle = 0 Then
        mFindHandle = FindFirstFile(mDirConnect.Path & FileSpec & vbNullChar & vbNullChar)
    Else
         findnext = FindNextFile(mFindHandle, mFindData)
    End If
    
    '///TASK:
    'Convert mfinddata structure data into appropriate filename/file path, and create a CFile or CDirectory Object, depending on it's actual type.
    FullFileName = mFindData.cFileName
    FullFileName = Trim$(mFindData.cFileName)
    FullFileName = Replace$(FullFileName, vbNullChar, "")
    FullFileName = mDirConnect.Path & FullFileName
    '///convert to appropriate object-
    'ise attributes to determine wether it is a directory or file.
    If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        'it is a directory, Convert to CDirectory.
        Set GetNext = FileSystem.GetDirectory(FullFileName)
    Else
        If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DEVICE) <> FILE_ATTRIBUTE_DEVICE Then
            'ignore device stuff here...
            'convert to CFile.
            Set GetNext = FileSystem.GetFile(FullFileName)
    
    
    
        End If
    End If
    
    


End Function

Public Function Clone() As CDirWalker
'
'Clone, hmm- that's a toughie in this instance.
Set Clone = Me
End Function
Public Sub Reset()
    FindClose mFindHandle
    mFindHandle = 0
End Sub
