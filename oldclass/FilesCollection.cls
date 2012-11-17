VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FilesCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'a collection of files- could be anywhere on a drive and on separate drives.


Private mcol As Collection



'Public Function CopyTo(ByVal HwndOwner As Long, ToDirectory As Directory, Optional ByVal ProgressCallback As IProgressCallback) As FilesCollection
'    'copy all files in this collection to the specified directory.
'
'    Dim Loopfile As CFile, newcollection As collection
'    Set newcollection = New collection
'    For Each Loopfile In mcol
'        Loopfile.Copy ToDirectory.Path, HwndOwner
'        If Not ProgressCallback Is Nothing Then
'            ProgressCallback.UpdateProgress Loopfile, ToDirectory, Loopfile.Size, 1
'        End If
'        newcollection.Add Loopfile
'
'    Next Loopfile
'
'        Set CopyTo = newcollection
'
'
'
'
'End Function
Public Function ToString() As String
    Dim retStr As String
    Dim loopfile As CFile
    For Each loopfile In mcol
        retStr = retStr & loopfile.Fullpath & ";"
    Next
ToString = retStr


End Function
Public Sub MoveTo(ByVal HwndOwner As Long, ToDirectory As Directory)
Attribute MoveTo.VB_Description = "Moves the Files in this collection to the specified Directory. The collection now reflects the new location of the files."
Dim loopfile As CFile, newcol As Collection

  Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    SHstruct.hwnd = HwndOwner
    SHstruct.fFlags = FOF_FILESONLY
    
'    If Not Showprompt Then
'        SHstruct.fFlags = SHstruct.fFlags + FOF_NOCONFIRMATION
'    End If

    'shstruct.
    SHstruct.wFunc = FO_MOVE


Set newcol = New Collection
For Each loopfile In mcol
    
  
    
    SHstruct.pFrom = SHstruct.pFrom & loopfile.Fullpath & vbNullChar
Next
SHstruct.pFrom = SHstruct.pFrom & vbNullChar
    SHstruct.pTo = ToDirectory.Path
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress

    SHFileOperation SHstruct
    'mvarFilename = SHstruct.pTo
    'refresh our info...





End Sub

Public Function Add(AddMe As CFile) As CFile
Attribute Add.VB_UserMemId = -4
    mcol.Add AddMe


End Function
Public Function NewEnum() As IUnknown
    Set NewEnum = mcol.[_NewEnum]
End Function

Public Function Item(VntIndexKey) As CFile
    Set Item = mcol.Item(VntIndexKey)
End Function
Private Sub Class_Initialize()
    Set mcol = New Collection
    mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub
