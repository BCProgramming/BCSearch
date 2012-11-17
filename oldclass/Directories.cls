VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Directorys collection.
'completely revised. Now only returns needed data instead of requiring population, also supports For...Each construct.




'Private mEnumerator As VBFCustomCollection.IEnumVARIANT

Private mCol As Collection
Private mParentDir As String 'as string to prevent circular references.
Public Function GetWalker() As CDirWalker
    
    
Set GetWalker = FileSystem.GetDirectory(mParentDir).GetWalker("*", FILE_ATTRIBUTE_DIRECTORY, 0)

End Function
Friend Sub Init(InDir As Directory, Walker As CDirWalker)
'
'must be given a parent dir and a Walker.
'Set mEnumerator = Walker.GetIEnum
mParentDir = InDir.Path
End Sub
Public Function Count() As Long
    'use API to count directories in parent dir...
    Dim finddata As WIN32_FIND_DATA
    Dim FHandle As Long
    Dim CountRunner As Long, ret As Long
    ZeroMemory finddata, Len(finddata)
    
    FHandle = FindFirstFile(mParentDir & "*", finddata)
    If FHandle = INVALID_HANDLE_VALUE Then
    'odd....
        Count = 0
        RaiseAPIError Err.LastDllError, "Directories::Count"
    Else
        Do
        If Left$(finddata.cFileName, 1) <> "." Then
        If (finddata.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            CountRunner = CountRunner + 1
        
        End If
        End If
        
        'ZeroMemory finddata, Len(finddata)
        ret = FindNextFile(FHandle, finddata)
        
        
    
        Loop Until ret <= 0
    End If
    FindClose FHandle
    Count = CountRunner
End Function

'Public Function Item(ByVal VntIndexKey As Variant) As Directory

    'Set Item = mCol.Item(VntIndexKey)
'End Function

'commented out. it crashes after being compiled :(
'Public Function NewEnum() As IUnknown
''Change- actually, you know, handle the NewEnum ourself....
'    Set NewEnum = mEnumerator
'End Function
'Public Function GetAsString() As String
'    Dim LoopDirectory As Directory, retit As String
'    For Each LoopDirectory In mCol
'        retit = retit & "," & LoopDirectory.Path
'    Next
'    GetAsString = retit
'End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
    mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub
