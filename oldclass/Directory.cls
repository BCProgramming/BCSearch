VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements IStreamable
'Directory class. represents- duh- a directory.
Private mDirInfo As BY_HANDLE_FILE_INFORMATION
Private mvarpath As String
Private mvarhDir As Long
'add methods to directly open streams on files in this dir.


Public Function CreateStream(ByVal filename As String) As FileStream

    'creates a stream in this dir.
    Dim absolutepath As String
    Dim madestream As FileStream
    absolutepath = MdlFileSystem.MakePathAbsolute(filename, Me.Path)
    Set CreateStream = FileSystem.CreateStream(absolutepath)
    


End Function



Public Function GetFileIcon(Optional ByVal IconSizeParams As IconSizeConstants = icon_shell) As Long
'TODO:// add size parameter.
    Dim finfo As SHFILEINFO
    Dim lIconType As Long
    
   
    ' be sure that there is the mbNormalIcon too
   
    ' retrieve the item's icon
    SHGetFileInfo Me.Path, 0, finfo, Len(finfo), SHGFI_ICON + SHGFI_SMALLICON
    'return hIcon- add to our imagelist, as well.
'    Select Case IconSizeParams
'    Case icon_small
'        LargeIcons.AddFromHandle Finfo.hIcon, IMAGE_ICON, Me.Name
'    Case icon_shell
'        ShellIcons.AddFromHandle Finfo.hIcon, IMAGE_ICON, Me.Name
'    Case icon_large
'        LargeIcons.AddFromHandle Finfo.hIcon, IMAGE_ICON, Me.Name
'    End Select
    GetFileIcon = finfo.hIcon
End Function


Public Sub Checkvalid()
Dim notfound As Boolean
CDebug.Post "Directory::CheckValid"
    If mvarhDir <= 0 Then
        If mvarpath <> "" Then
            'try to refresh the handle...
            On Error Resume Next
            SetDir mvarpath
            If Err <> 0 Then
                Err.Raise 53, "Directory::CheckValid", "Directory Object Invalidated"
            
            End If
        Else
            'damn shame- no, really-
            Err.Raise 53, "Directory::Checkvalid", "Directory Object invalidated."
        
        End If
    End If

End Sub
Public Property Get AlternateStreams(Optional ByVal ForceRefresh As Boolean = False) As CAlternateStreams
    Static AltStreams As CAlternateStreams
    Checkvalid
    If Not AltStreams Is Nothing And Not ForceRefresh Then
    
    Else
        
        Set AltStreams = ListStreams(Me.Path)

    End If
        Set AlternateStreams = AltStreams


End Property
Public Function Move(ByVal Destination As String, Optional ByVal HwndOwner As Long) As Directory
    'Copies this directory to Destination, and returns a CFile Object for that File.
    Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    'Dim Copied As Directory
    Checkvalid
    SHstruct.hWnd = HwndOwner
    'SHstruct.fFlags = FOF_CONFIRMMOUSE
    'shstruct.
    SHstruct.wFunc = FO_MOVE
    SHstruct.pFrom = Mid$(Me.Path, 1, Len(Me.Path) - 1) & vbNullChar & vbNullChar
    SHstruct.pTo = Destination & Me.name & vbNullChar & vbNullChar
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress
    
    'since we want to move it, we will need to close our hDir- this also invalidates this object- oh well.
    
    Class_Terminate
    SHFileOperation SHstruct
    
    If SHstruct.hNameMaps <> 0 Then
        SHFreeNameMappings SHstruct.hNameMaps
    End If
    Set Move = FileSystem.GetDirectory(SHstruct.pTo)
    'if we are referenced again, that's ok, we'll just err out anyway.
End Function

Public Function Copy(ByVal Destination As String, Optional ByVal HwndOwner As Long) As Directory
    'Copies this directory to Destination, and returns a CFile Object for that File.
    Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    Checkvalid
    'Dim Copied As Directory
    SHstruct.hWnd = HwndOwner
    'SHstruct.fFlags = FOF_CONFIRMMOUSE
    'shstruct.
    SHstruct.wFunc = FO_COPY
    SHstruct.pFrom = Mid$(Me.Path, 1, Len(Me.Path) - 1) & vbNullChar & vbNullChar
    SHstruct.pTo = Destination & Me.name & vbNullChar & vbNullChar
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress
    SHFileOperation SHstruct
    
    If SHstruct.hNameMaps <> 0 Then
        SHFreeNameMappings SHstruct.hNameMaps
    End If
    Set Copy = FileSystem.GetDirectory(SHstruct.pTo)
End Function
Public Function Execute(ownerHwnd As Long, Optional ByVal Verb As String = "Open") As Boolean
     Dim dummy As Long, executablestart$
     Dim FilePath As String
     FilePath = Me.Path
               'open the file using the default Editor or viewer.
     dummy = ShellExecute(ownerHwnd, Verb, FilePath & Chr$(0), Chr$(0), Left$(FilePath, InStr(FilePath, "\")), vbNormalFocus)
     
'     If dummy <> 33 Then
'        executablestart = FindExecutable(filepath)
'
'
'     End If
'
     
     
     
End Function

Public Function GetWalker(Optional ByVal pSpec As String = "*", Optional ByVal IncludeAttr As FileAttributeConstants = FILE_ATTRIBUTE_NORMAL, Optional ByVal ExcludeAttr As FileAttributeConstants = FILE_ATTRIBUTE_DIRECTORY) As CDirWalker
    Dim Walker As CDirWalker
    Set Walker = New CDirWalker
  Dim x As WIN32_FIND_DATA
  
  CDebug.Post "Directory::GetWalker, pspec=" & pSpec & " Includeattr=" & IncludeAttr
  Checkvalid

   
    Walker.InitWalker Me, pSpec, IncludeAttr, ExcludeAttr
    Set GetWalker = Walker

End Function
Public Function GetRelativeDirectory(ByVal RelativePath As String) As Directory
    Dim strrelative As String
    strrelative = MakePathAbsolute(RelativePath, Me.Path)
    
    Set GetRelativeDirectory = FileSystem.GetDirectory(strrelative)




End Function
Public Property Get Directories(Optional ByVal Spec As String = "*") As Directories
    Dim Walker As CDirWalker, dirs As Directories
    Checkvalid
    Set Walker = New CDirWalker
    Walker.InitWalker Me, Spec, FILE_ATTRIBUTE_DIRECTORY
    Set dirs = New Directories
    dirs.Init Me, Walker
    Set Directories = dirs

'    Dim MakeCol As Directories
'    Dim FData As WIN32_FIND_DATA
'    Dim hFind As Long
'    Dim MakeDir As Directory
'    Set MakeCol = New Directories
'    ZeroMemory FData, Len(FData)
'    hFind = FindFirstFile(Me.Path & "\" & Spec, FData)
'
'    Do
'        If (FData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'            'Add it.
'            Set MakeDir = New Directory
'            On Error Resume Next
'            MakeDir.SetDir Me.Path & Trim$(Replace$(FData.cFileName, vbNullChar, ""))
'            If Err <> 0 Then
'                Err.Clear
'            Else
'                MakeCol.Add MakeDir
'
'            End If
'
'        End If
'        ZeroMemory FData, Len(FData)
'        If FindNextFile(hFind, FData) = 0 Then
'            FindClose hFind
'            Exit Do
'        End If
'
'    Loop
'
'    Set Directories = MakeCol
'
'


End Property

Public Property Get Files(Optional ByVal Spec As String = "*") As Files


  Dim Walker As CDirWalker, Filecol As Files
  Checkvalid
    Set Walker = New CDirWalker
    'all files except directories....
    Walker.InitWalker Me, Spec, 0, FILE_ATTRIBUTE_DIRECTORY
    Set Filecol = New Files
    Filecol.Init Me, Walker
    Set Files = Filecol


'retrieve the files in this dir.
'    Dim FData As WIN32_FIND_DATA
'    Dim hFind As Long
'    Dim MakeCol As Files
'    Dim makefile As CFile
'    Set MakeCol = New Files
'     ZeroMemory FData, Len(FData)
'    hFind = FindFirstFile(Me.Path & "\" & Spec, FData)
'
'
'    Do
'       If Asc(FData.cFileName) = 0 Then Exit Do
'        If (FData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'
'        Else
'
'
'
'           Set makefile = New CFile
'           On Error Resume Next
'            makefile.SetFileName Me.Path & Trim$(Replace$(FData.cFileName, vbNullChar, ""))
'            If Err <> 0 Then
'                'CDebug.Post "failed to access: " + Fdata.cFileName & """ "
'                Err.Clear
'            Else
'                'the other info can be retrieved later if necessary.
'                MakeCol.Add makefile
'            End If
'
'
'
'        End If
'         ZeroMemory FData, Len(FData)
'        If FindNextFile(hFind, FData) = 0 Then
'            FindClose hFind
'            Exit Do
'
'
'
'        End If
'
'    Loop
'
'    Set Files = MakeCol
End Property
Friend Sub SetDir(ByVal DirName As String)
    Dim hf As Long, pfinddat As WIN32_FIND_DATA
   AppendSlash DirName
    mvarpath = DirName
    'must use FILE_FLAG_BACKUP_SEMANTICS for dirs.
    mvarhDir = CreateFile(DirName, GENERIC_READ, FILE_SHARE_READ + FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If mvarhDir <= 0 Then
        hf = FindFirstFile(DirName, pfinddat)
        If hf = 0 Then
            Call FindClose(hf)
        Else
            RaiseAPIError Err.LastDllError, "SetDir"
        End If
    End If
    RefreshInfo

End Sub
Public Property Get name() As String
    Dim DirParts() As String
    DirParts = Split(mvarpath, "\")
    name = DirParts(UBound(DirParts) - 1)
End Property
Public Property Get Extension() As String
    Dim mextension As String
    Dim mpath As String
    mpath = Me.Path
    Call ParsePathParts(Mid$(mpath, 1, Len(mpath) - 1), , , , mextension)
    Extension = mextension

End Property
Public Property Get Path() As String
    
    If (Right$(mvarpath, 1) <> "\" And Right$(mvarpath, 1) <> "/") Then mvarpath = mvarpath & "\"
    Path = mvarpath
End Property
Public Property Get Attributes() As FileAttributeConstants
    Checkvalid
        Attributes = mDirInfo.dwFileAttributes
    'acquire attributes....
    

End Property
Public Function size(ByVal IncludeSubfolders As Boolean) As Double
    'return the size of the files, and optionally the subdirs of this directory.
    
    
    Dim AccumSize As Double
    Dim LoopDir As Object
    Dim loopfile As Object
    With Me.Files.GetWalker
    'For Each loopfile In Me.Files
    Do Until .GetNext(loopfile) Is Nothing
        AccumSize = loopfile.size + AccumSize
    'Next loopfile
    Loop
    End With
    If IncludeSubfolders Then
        With Me.Directories.GetWalker
        'For Each LoopDir In Me.Directories
        Do Until .GetNext(LoopDir) Is Nothing
            'Debug.Print TypeName(LoopDir)
            AccumSize = AccumSize + LoopDir.size(True)
        
        Loop
        End With
    End If
    
    size = AccumSize
    
    





End Function
Public Property Let DateCreated(ByVal vdata As Date)
    SetFileTimes mvarpath, vdata, Me.DateLastAccessed, Me.DateModified
End Property
Public Property Let DateModified(ByVal vdata As Date)
    SetFileTimes mvarpath, DateCreated, DateLastAccessed, vdata
End Property
Public Property Let DateLastAccessed(ByVal vdata As Date)
    SetFileTimes mvarpath, Me.DateCreated, vdata, Me.DateModified
End Property


Public Property Get DateCreated() As Date
    Checkvalid
    DateCreated = FileTime2Date(mDirInfo.ftCreationTime)
End Property
Public Property Get DateModified() As Date
    Checkvalid
    DateModified = FileTime2Date(mDirInfo.ftLastWriteTime)
End Property
Public Property Get DateLastAccessed() As Date
    Checkvalid
    DateLastAccessed = FileTime2Date(mDirInfo.ftLastAccessTime)


End Property

Public Function ShowExplorerMenu(Optional ByVal HwndOwner As Long = 0, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, _
        Optional menucallback As IContextCallback = Nothing) As Long
    'current implementation: Shows the context menu popup at the mouse position.
    'This is currently my most treasured (or part of my most treasured) routines.
    'Not ShowExplorermenu, but rather the ShowShellMenu() routine it uses.
    Dim retVal As Long
    On Error GoTo ReturnFalse
    Checkvalid
    retVal = MdlFileSystem.ShowExplorerMenu(HwndOwner, Left$(Me.Path, Len(Me.Path) - 1))
    
    ShowExplorerMenu = retVal
    Exit Function
ReturnFalse:
    

End Function
Private Sub RefreshInfo()
    Checkvalid
    Dim ret As Long
    ret = GetFileInformationByHandle(mvarhDir, mDirInfo)
     
    'ret = GetFileTime(mvarhDir, mDirInfo.ftCreationTime, mDirInfo.ftLastAccessTime, mDirInfo.ftLastWriteTime)
    'ret = GetFileSizeEx(mvarhdir,
    'mDirInfo.nFileSizeLow = GetFileSize(mvarhDir, mDirInfo.nFileSizeHigh)
    
    

End Sub
Public Property Get Volume() As CVolume
    Dim gotvol As CVolume, volpart As String
    
    ParsePathParts Me.Path, volpart
    
    Set gotvol = FileSystem.GetVolume(volpart)


    Set Volume = gotvol
End Property
Private Function ISearchItem_Self() As Object
    Set ISearchItem_Self = Me
End Function


Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
    CloseHandle mvarhDir
    mvarhDir = 0
    mTotalObjectCount = mTotalObjectCount - 1
End Sub
