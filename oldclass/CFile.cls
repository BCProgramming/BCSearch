VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_PATH As Long = 260
Implements IProgressCallback
Implements IStreamable

'Private Type SHFILEINFO
'    hIcon As Long ' : icon
'    iIcon As Long ' : icondex
'    dwAttributes As Long ' : SFGAO_ flags
'    szDisplayName As String * MAX_PATH ' : display name (or path)
'    szTypeName As String * 80 ' : type name
'End Type
'Private Declare Function GetFileAttributesEx Lib "kernel32.dll" Alias "GetFileAttributesExA" (ByVal lpFileName As String, ByVal fInfoLevelId As Struct_MembersOf_GET_FILEEX_INFO_LEVELS, ByRef lpFileInformation As Any) As Long




Private Declare Function OpenFile Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Const OFS_MAXPATHNAME As Long = 128

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type


Private mSHInfo As SHFILEINFO
Private ShInit As Boolean
'CFile class: represents a file.
Private flNeedInfo As Boolean, CloseMe As Long
Private mFileInfo As BY_HANDLE_FILE_INFORMATION
Private mvarFileName As String

'File Manipulation functions- Copy, Move, Etc...


Public Sub Delete()
Attribute Delete.VB_Description = "Deletes this file. The CFile object will not be valid for future access."
    'delete with DeleteFile() API routine.
    DeleteFile Me.Fullpath

    'file is now no longer a valid reference...
 

End Sub
Public Function Rename(ByVal NewName As String, Optional ByVal HwndOwner As Long = 0, Optional ByVal Showprompt As Boolean = True)
Attribute Rename.VB_Description = "Renames this file to a new name. The current instance will point to the new name."
'rename: works for CFile...
    Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    SHstruct.hWnd = HwndOwner
    SHstruct.fFlags = FOF_FILESONLY
    
    If Not Showprompt Then
        SHstruct.fFlags = SHstruct.fFlags + FOF_NOCONFIRMATION
    End If

    'shstruct.
    SHstruct.wFunc = FO_RENAME
    
    SHstruct.pFrom = Me.Fullpath & vbNullChar & vbNullChar
    SHstruct.pTo = Me.Directory.Path & NewName
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress

    SHFileOperation SHstruct
    mvarFileName = SHstruct.pTo
    'refresh our info...
   GetInfo

    Set Rename = Me


End Function
Public Sub SHDelete(Optional ByVal HwndOwner As Long, Optional ByVal Showprompt As Boolean = False)
Attribute SHDelete.VB_Description = "Uses the Shell to delete this file."
'Delete this File.
Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    SHstruct.hWnd = HwndOwner
    SHstruct.fFlags = FOF_FILESONLY
    
    If Not Showprompt Then
        SHstruct.fFlags = SHstruct.fFlags + FOF_NOCONFIRMATION
    End If

    'shstruct.
    SHstruct.wFunc = FO_DELETE
    
    SHstruct.pFrom = Me.Fullpath & vbNullChar & vbNullChar
    'SHstruct.pTo = Destination & vbNullChar & vbNullChar
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress
    SHFileOperation SHstruct
    
   




End Sub
Public Function Copy(ByVal Destination As String, Optional ByVal HwndOwner As Long) As CFile
Attribute Copy.VB_Description = "Copies this file to the destination, returning the CFile object of the new file."
    'Copies this file to Destination, and returns a CFile Object for that File.
    Dim SHstruct As SHFILEOPSTRUCT
    Dim MapPtr As Long
    Dim Mapped As SHNAMEMAPPING
    SHstruct.hWnd = HwndOwner
    SHstruct.fFlags = FOF_FILESONLY
    'shstruct.
    SHstruct.wFunc = FO_COPY
    SHstruct.pFrom = Me.Fullpath & vbNullChar & vbNullChar
    SHstruct.pTo = Destination & vbNullChar & vbNullChar
    'SHstruct.hNameMaps =
    'SHstruct.fFlags
   ' SHstruct.sProgress = "Copying" & vbNullChar
    
    'ShStruct.sProgress
    SHFileOperation SHstruct
    
    If SHstruct.hNameMaps <> 0 Then
        SHFreeNameMappings SHstruct.hNameMaps
    End If

End Function

'copyEx... needs testing!
Public Function CopyEx(ByVal Destination, Optional progressSink As IProgressCallback) As CFile
    'Copies this file to the given destination.
    
    'Destination can be:
    
    'a Directory object
    'a Directory path
    'a File specification
    Dim destFileName As String
    
    
    Dim destStream As FileStream
    Dim SrcStream As FileStream
    Dim castFile As CFile
    Set SrcStream = Me.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_WRITE, OPEN_EXISTING)
    If IsObject(Destination) Then
        If TypeOf Destination Is Directory Then
            destFileName = Destination.Path & Me.filename
            Set destStream = FileSystem.CreateFile(destFileName).OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING)
        ElseIf TypeOf Destination Is CFile Then
            Set castFile = Destination
            Set destStream = castFile.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING)
        ElseIf varType(Destination) = vbString Then
            If IsFileName(Destination) Then
                'set deststream = filesystem.CreateFile(destination).OpenAsBinaryStream
                Set destStream = FileSystem.CreateFile(Destination).OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING)
            'elseif isDirectory
            ElseIf isDirectory(Destination) Then
            '
            End If
        Else
        
            Err.Raise 5, "CFile::CopyEx", "Unrecognized Destination argument."
        End If
    End If
    StreamCopy SrcStream, destStream, progressSink
    SrcStream.CloseStream
    destStream.CloseStream
End Function


'Public Function CopyEx(ByVal Destination, Optional ByVal HwndOwner As Long, Optional ProgressSink As IProgressCallback)
''
''Currently Broken :(
'Dim punk As olelib.IUnknown
'Dim Cancelled As Long, flags As Long
'Dim dest As String
'Dim SendMe As BCCOPYFILEDATA
'
'SendMe.DestinationFile = Destination & Me.Filename
'SendMe.SourceFile = Me.Fullpath
'dest = Destination
'Dim sinkptr As Long
'If Not ProgressSink Is Nothing Then
'
'        sinkptr = ObjPtr(ProgressSink)
'        'since the object might go out of scope between now and when the callback tries to reference it
'        'we need to addRef it
''        Set punk = ProgressSink
''        punk.AddRef
'    Set SendMe.BCCallback = ProgressSink
'
'Else
'    sinkptr = ObjPtr(Me)
'    Set SendMe.BCCallback = Me
'End If
'
'If sinkptr <> 0 Then
'    CopyFileEx Me.Fullpath, dest, AddressOf CopyProgressRoutine, VarPtr(SendMe), Cancelled, flags
'Else
'    CopyFileEx Me.Fullpath, dest, AddressOf CopyProgressRoutine, VarPtr(SendMe), Cancelled, flags
'End If
'
'
'Set SendMe.BCCallback = Nothing
'
'
'End Function

Public Property Get CompressedSize() As Double
Attribute CompressedSize.VB_Description = "Returns the size on disk of this file."
    Dim LongHigh As Long
    Dim LongLow As Long
    LongLow = GetCompressedFileSize(Me.Fullpath, LongHigh)
    
    CompressedSize = LI2DBL(LongLow, LongHigh)




End Property

Public Function DisplayName() As String
Attribute DisplayName.VB_Description = "Retrieves the Display Name of this file."
    DisplayName = FileSystem.GetFilenamePart(Me.Fullpath)



End Function

'Public Property Get Name() As String
'    Name = mvarFilename
'End Property
Public Property Get Fullpath() As String
Attribute Fullpath.VB_Description = "returns the full path of this file."
    Fullpath = mvarFileName
End Property
Public Function OpenAsBinaryStream(Optional ByVal dwDesiredAccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal dwShareMode As FileShareModeConstants = FILE_SHARE_READ + FILE_SHARE_WRITE, Optional ByVal dwCreationDisposition As CreationDispositionConstants = OPEN_EXISTING, Optional ByVal dwFlagsAndAttributes As FileAttributeConstants) As FileStream
Attribute OpenAsBinaryStream.VB_Description = "Opens this file as a Binary Stream with the given parameters and returns the FileStream Object."
 Dim retthis As FileStream
 Set retthis = New FileStream
 retthis.OpenStream Me.Fullpath, dwDesiredAccess, dwShareMode, dwCreationDisposition, dwFlagsAndAttributes
 Set OpenAsBinaryStream = retthis
 
 
End Function
'Public Function OpenAsBackupStream(Optional ByVal dwDesiredAccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal dwShareMode As FileShareModeConstants = FILE_SHARE_READ + FILE_SHARE_WRITE, Optional ByVal dwCreationDisposition As CreationDispositionConstants = OPEN_EXISTING) As CBackupStream
' Dim retthis As CBackupStream
' Set retthis = New CBackupStream
' retthis.OpenStream Me.Fullpath, dwDesiredAccess, dwShareMode, dwCreationDisposition
' Set OpenAsBackupStream = retthis
'
'
'End Function


Public Sub SetFileName(ByVal vdata As String)
    mvarFileName = vdata
    'GetInfo
    flNeedInfo = True
End Sub
Public Property Get filename() As String
Attribute filename.VB_Description = "returns the filename of this file."
    filename = FileSystem.GetFilenamePart(Me.Fullpath)


End Property
Public Property Get BaseName() As String
Attribute BaseName.VB_Description = "Returns the name of this file, without a path or extension."
    Dim mname As String, inrev As Long
    
    mname = Me.filename
    inrev = InStrRev(mname, ".") - 1
    If inrev = 0 Then inrev = Len(mname)
    
    BaseName = Mid$(mname, 1, inrev)


End Property
Friend Sub GetInfo()
    Dim hfile As Long, sAttrib As SECURITY_ATTRIBUTES
    Dim retVal As Long, oostruct As OFSTRUCT
    Dim finddata As WIN32_FIND_DATA
    Const OF_READ As Long = &H0
    Const OF_SHARE_DENY_NONE = &H40
    
    
    hfile = CreateFile(mvarFileName, GENERIC_DEVICE_QUERY, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS + FILE_FLAG_NO_BUFFERING, 0)
    If hfile <= 0 Then
        'try OpenFile First...
        hfile = OpenFile(mvarFileName, oostruct, OF_READ + OF_SHARE_DENY_NONE)
        
        
    
        If hfile <= 0 Then
        Dim hfind As Long
        'One last ditch effort- use FindFirstFile() and search for that file...
            ZeroMemory finddata, Len(finddata)
            hfind = FindFirstFile(mvarFileName, finddata)
            If hfind <> 0 Then
                mFileInfo.nFileSizeLow = finddata.nFileSizeLow
                mFileInfo.nFileSizeHigh = finddata.nFileSizeHigh
                mFileInfo.dwFileAttributes = finddata.dwFileAttributes
                mFileInfo.ftLastWriteTime = finddata.ftLastWriteTime
                'mfileinfo.dwVolumeSerialNumber=
                mFileInfo.ftCreationTime = finddata.ftCreationTime
                mFileInfo.ftLastAccessTime = finddata.ftLastAccessTime
                'rest can go away.
                CloseHandle hfind
                CloseHandle hfile
                Exit Sub
            Else
            RaiseAPIError Err.LastDllError, "CFile::GetInfo"  'either a honest to goodness access denied error or file not found.
            
            End If
            
            
    
            
        End If
    End If
    
    
    retVal = GetFileInformationByHandle(hfile, mFileInfo)
    
    



    CloseHandle hfile
End Sub
Public Property Get size() As Double
Attribute size.VB_Description = "returns the size of this file."
    If flNeedInfo Then GetInfo
    size = LI2DBL(mFileInfo.nFileSizeLow, mFileInfo.nFileSizeHigh)
End Property

Public Property Get DateCreated() As Date
Attribute DateCreated.VB_Description = "Returns the Date this file was Created."
    If flNeedInfo Then GetInfo
    DateCreated = FileTime2Date(mFileInfo.ftCreationTime)
End Property
Public Property Let DateCreated(ByVal vdata As Date)
    SetFileTimes mvarFileName, vdata, Me.DateLastAccessed, Me.DateModified
End Property
Public Property Let DateModified(ByVal vdata As Date)
    SetFileTimes mvarFileName, DateCreated, DateLastAccessed, vdata
End Property
Public Property Let DateLastAccessed(ByVal vdata As Date)
    SetFileTimes mvarFileName, Me.DateCreated, vdata, Me.DateModified
End Property
Public Property Get DateModified() As Date
Attribute DateModified.VB_Description = "returns the Modified date of this file."
    If flNeedInfo Then GetInfo
    DateModified = FileTime2Date(mFileInfo.ftLastWriteTime)
End Property

Public Property Get DateLastAccessed() As Date
Attribute DateLastAccessed.VB_Description = "Returns the Last accessed Time of this file."
    If flNeedInfo Then GetInfo
    DateLastAccessed = FileTime2Date(mFileInfo.ftLastAccessTime)


End Property
Public Property Get FileAttributes() As FileAttributeConstants
Attribute FileAttributes.VB_Description = "Sets/Returns the attribute mask of this file."
    If flNeedInfo Then GetInfo
    FileAttributes = mFileInfo.dwFileAttributes
End Property
Public Property Let FileAttributes(vdata As FileAttributeConstants)

    SetFileAttributes mvarFileName, vdata
    mFileInfo.dwFileAttributes = vdata
End Property
Public Property Get FileIndex() As Double
Attribute FileIndex.VB_Description = "Returns the FileIndex of this file."
    If flNeedInfo Then GetInfo
    FileIndex = LI2DBL(mFileInfo.nFileIndexLow, mFileInfo.nFileIndexHigh)

End Property


Public Property Get HardLinkCount() As Long
    If flNeedInfo Then GetInfo
    HardLinkCount = mFileInfo.nNumberOfLinks

End Property

'File spec manipulation functions.



Public Function ShowExplorerMenu(Optional ByVal HwndOwner As Long = 0, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, _
        Optional menucallback As IContextCallback = Nothing) As Long
Attribute ShowExplorerMenu.VB_Description = "Displays the Right-Click Explorer menu for the file at the specified position."
    'current implementation: Shows the context menu popup at the mouse position.
    'This is currently my most treasured (or part of my most treasured) routines.
    'Not ShowExplorermenu, but rather the ShowShellMenu() routine it uses.
    Dim retVal As Long
    On Error GoTo ReturnFalse
    retVal = MdlFileSystem.ShowExplorerMenu(HwndOwner, Me.Fullpath, x, y, menucallback)
    
    ShowExplorerMenu = retVal
    Exit Function
ReturnFalse:
    

End Function

Friend Sub OpenHandle(ByVal vdata As Long)
    'opens file handle. probably created before, etc.
    'CloseMe = Vdata
    GetFileInformationByHandle vdata, mFileInfo
    mvarFileName = GetFileNameFromHandle(vdata)
    'erm- TADA!
End Sub

Friend Sub Create(ByVal filename As String, flAllowOverwrite As Boolean)
    Dim hfile As Long
    '
    If FileExists(filename) Then
    End If
    
    
    
    hfile = CreateFile(filename, GENERIC_ALL, FileShareModeConstants.FILE_SHARE_READ, ByVal &O0, CREATE_ALWAYS, 0, 0)
    If hfile <> 0 Then
    
        mvarFileName = filename
        GetInfo
    End If
        
    CloseHandle hfile
    '
End Sub

Public Function GetFileIcon(Optional ByVal IconSizeParams As IconSizeConstants = icon_shell) As Long
Attribute GetFileIcon.VB_Description = "returns a handle to a Icon that is the icon used by the shell to represent this file."
'TODO:// add size parameter.
    Dim finfo As SHFILEINFO
    Dim lIconType As Long
    
   
    ' be sure that there is the mbNormalIcon too
   
    ' retrieve the item's icon
    If IconSizeParams = ICON_SMALL Then
        
        SHGetFileInfo Me.Fullpath, 0, finfo, Len(finfo), SHGFI_ICON + SHGFI_SMALLICON
    ElseIf IconSizeParams = icon_large Then
        SHGetFileInfo Me.Fullpath, 0, finfo, Len(finfo), SHGFI_ICON + SHGFI_LARGEICON
        
    
    End If
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


Private Sub Class_Initialize()
    mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
    If CloseMe <> 0 Then CloseHandle CloseMe
    mTotalObjectCount = mTotalObjectCount - 1
End Sub



Private Function IProgressCallback_UpdateProgress(Source As Object, Destination As Object, Optional FileSize As Double = -1#, Optional FileProgress As Double = -1#, Optional StreamSize As Double = -1#, Optional StreamProgress As Double = -1#) As Boolean
'
Debug.Print "progress update: filesize="; FileSize, " file progress = " & FileProgress
End Function

Private Function ISearchItem_Self() As Object
    Set ISearchItem_Self = Me
End Function
Public Function GetAttributeString() As String
Attribute GetAttributeString.VB_Description = "converts the attributes of this file into a readable string for display."
    'returns a comma separated list of the attributes of the file.
    'IE: R,H,S,T etc.
    Static Lookup() As Long, flInit As Boolean
    Static AttrStr() As String
'    FILE_ATTRIBUTE_ARCHIVE = &H20
'    FILE_ATTRIBUTE_COMPRESSED = &H800
'    FILE_ATTRIBUTE_DEVICE = &H40
'    FILE_ATTRIBUTE_DIRECTORY = &H10
'    FILE_ATTRIBUTE_ENCRYPTED = &H4000
'    FILE_ATTRIBUTE_HIDDEN = &H2
'    FILE_ATTRIBUTE_NORMAL = &H80
'    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
'    FILE_ATTRIBUTE_OFFLINE = &H1000
'    FILE_ATTRIBUTE_READONLY = &H1
'    FILE_ATTRIBUTE_REPARSE_POINT = &H400
'    FILE_ATTRIBUTE_SPARSE_FILE = &H200
'    FILE_ATTRIBUTE_SYSTEM = &H4
'    FILE_ATTRIBUTE_TEMPORARY = &H100
    Dim attrib As FileAttributeConstants, I As Long
    Dim buildstr As String
    If Not flInit Then
        flInit = True
        ReDim AttrLook(1 To 14)
        ReDim Lookup(1 To 14)
        Lookup(1) = FILE_ATTRIBUTE_ARCHIVE
        AttrLook(1) = "A"
        Lookup(2) = FILE_ATTRIBUTE_COMPRESSED
        AttrLook(2) = "C"
        Lookup(3) = FILE_ATTRIBUTE_DEVICE
        Lookup(4) = FILE_ATTRIBUTE_DIRECTORY
        Lookup(5) = FILE_ATTRIBUTE_ENCRYPTED
        AttrLook(3) = "E"
        Lookup(6) = FILE_ATTRIBUTE_HIDDEN
        AttrLook(4) = "H"
        Lookup(7) = FILE_ATTRIBUTE_NORMAL
        Lookup(8) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
        Lookup(9) = FILE_ATTRIBUTE_OFFLINE
        Lookup(10) = FILE_ATTRIBUTE_READONLY
        AttrLook(10) = "R"
        Lookup(11) = FILE_ATTRIBUTE_REPARSE_POINT
        Lookup(12) = FILE_ATTRIBUTE_SPARSE_FILE
        Lookup(13) = FILE_ATTRIBUTE_SYSTEM
        AttrLook(13) = "S"
        Lookup(14) = FILE_ATTRIBUTE_TEMPORARY
    End If
    For I = 1 To 14
        If (Me.FileAttributes And Lookup(I)) = Lookup(I) Then
        buildstr = buildstr & AttrLook(I) & ","
        
        End If
    
    Next I
    If Right$(buildstr, 1) = "," Then buildstr = Left$(buildstr, Len(buildstr) - 1)
    GetAttributeString = buildstr
End Function
Public Function FileType() As String
Attribute FileType.VB_Description = "returns the filetype of this file."
'

     Dim FileStruct As SHFILEINFO
     Call SHGetFileInfo(Me.Fullpath, Me.FileAttributes, FileStruct, Len(FileStruct), SHGFI_TYPENAME)
     FileType = Trim$(FileStruct.szTypeName)



End Function

Public Property Get Extension() As String
Attribute Extension.VB_Description = "Retrieves the extension of this file."
Dim mextension As String
Call ParsePathParts(Me.Fullpath, , , , mextension)
Extension = mextension

End Property
Public Function Directory() As Directory
Attribute Directory.VB_Description = "Returns the Directory that this file resides in."
    Dim DirGet As String, volget As String
    'DirGet = Mid$(Me.Fullpath, 1, Len(Me.Fullpath) - Len(Me.DisplayName))
    ParsePathParts Me.Fullpath, volget, DirGet
    DirGet = volget & DirGet
    Set Directory = FileSystem.GetDirectory(DirGet)



End Function
Public Function OpenWith(ByVal hwnduse As Long) As Boolean
Dim hIcon As Long
hIcon = Me.GetFileIcon
If MdlFileSystem.OpenWith(hwnduse, Me.Fullpath, hIcon) Then

End If

DestroyIcon hIcon

End Function
Public Function Execute(ownerHwnd As Long, Optional ByVal Verb As String = "Open") As Boolean
Attribute Execute.VB_Description = "Attempts to load this file, either as a Program or with the default program associated with it's file type."
     Dim dummy As Long, executablestart$
     Dim FilePath As String
     FilePath = Me.Fullpath
               'open the file using the default Editor or viewer.
     dummy = ShellExecute(ownerHwnd, Verb, FilePath & Chr$(0), Chr$(0), Left$(FilePath, InStr(FilePath, "\")), vbNormalFocus)
     
'     If dummy <> 33 Then
'        executablestart = FindExecutable(filepath)
'
'
'     End If
'
     
     
     
End Function
Public Property Get AlternateStreams(Optional ByVal ForceRefresh As Boolean = False) As CAlternateStreams
Attribute AlternateStreams.VB_Description = "retrieves a collection of alternate streams present in this file."
    Static AltStreams As CAlternateStreams
    
    If Not AltStreams Is Nothing And Not ForceRefresh Then
    
    Else
        
        'Set AltStreams = GetAlternateStreamsByPath(Me.Fullpath)
        Set AltStreams = ListStreams(Me.Fullpath)
        
    End If
    Set AlternateStreams = AltStreams


End Property
Private Property Get IStreamable_Fullpath() As String
IStreamable_Fullpath = Me.Fullpath
End Property

Private Property Get IStreamable_Size() As Long
    IStreamable_Size = Me.size
End Property
