VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCFSObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function SearchPath Lib "kernel32.dll" Alias "SearchPathA" (ByVal lpPath As Long, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long


Public Enum ParsePathPartsConstants
    Parse_Volume
    Parse_Path
    Parse_FName
    Parse_FExt
    Parse_Stream
    Parse_All
End Enum

Public Enum DesiredAccessFlags
    GENERIC_DEVICE_QUERY = 0
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
    ACCESS_DELETE = &H10000
    READ_CONTROL = &H20000
    SPECIFIC_RIGHTS_ALL = &HFFFF&
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_REQUIRED = &HF0000
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    SYNCHRONIZE = &H100000
    WRITE_DAC = &H40000
    WRITE_OWNER = &H80000
    ACCESS_SYSTEM_SECURITY = &H1000000
    GENERIC_ALL = &H10000000
    GENERIC_EXECUTE = &H20000000
    MAXIMUM_ALLOWED = &H2000000
End Enum
#If False Then
'kludge to force capitalization...
    Const GENERIC_DEVICE_QUERY = 0
    Const GENERIC_READ = &H80000000
    Const GENERIC_WRITE = &H40000000
    Const ACCESS_DELETE = &H10000
    Const READ_CONTROL = &H20000
    Const SPECIFIC_RIGHTS_ALL = &HFFFF&
    Const STANDARD_RIGHTS_ALL = &H1F0000
    Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    Const STANDARD_RIGHTS_READ = (READ_CONTROL)
    Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    Const SYNCHRONIZE = &H100000
    Const WRITE_DAC = &H40000
    Const WRITE_OWNER = &H80000
    Const ACCESS_SYSTEM_SECURITY = &H1000000
    Const GENERIC_ALL = &H10000000
    Const GENERIC_EXECUTE = &H20000000
    Const MAXIMUM_ALLOWED = &H2000000
    
    Const CREATE_ALWAYS = 2
    Const CREATE_NEW = 1
    Const OPEN_ALWAYS = 4
    Const OPEN_EXISTING = 3
    Const TRUNCATE_EXISTING = 5
    
    Const FILE_SHARE_DELETE = &H4
    Const FILE_SHARE_READ = &H1
    Const FILE_SHARE_WRITE = &H2
    
#End If
'Optional ByVal LongMetric As Boolean = False,
'Optional ByVal ShowThousandsSeparator As Boolean = True,
'Optional ByVal UseNewStandards As Boolean

Public Enum EFormatSizeFlags
    Format_ShowTotalInBrackets = 1
    Format_longmetric = 2
    Format_ShowThousandsSeparator = 4
    Format_UseNewStandard = 8
End Enum

Public Enum FileShareModeConstants
    FILE_SHARE_DELETE = &H4
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum
Public Enum CreationDispositionConstants
    CREATE_ALWAYS = 2
    CREATE_NEW = 1
    OPEN_ALWAYS = 4
    OPEN_EXISTING = 3
    TRUNCATE_EXISTING = 5
End Enum



Public Enum SpecialFolderConstants
 CSIDL_ADMINTOOLS = &H30
 CSIDL_ALTSTARTUP = &H1D
 CSIDL_APPDATA = &H1A
 CSIDL_BITBUCKET = &HA
 CSIDL_COMMON_ADMINTOOLS = &H2F
 CSIDL_COMMON_ALTSTARTUP = &H1E
 CSIDL_COMMON_APPDATA = &H23
 CSIDL_COMMON_DESKTOPDIRECTORY = &H19
 CSIDL_COMMON_DOCUMENTS = &H2E
 CSIDL_COMMON_FAVORITES = &H1F
 CSIDL_COMMON_PROGRAMS = &H17
 CSIDL_COMMON_STARTMENU = &H16
 CSIDL_COMMON_STARTUP = &H18
 CSIDL_COMMON_TEMPLATES = &H2D
 CSIDL_CONNECTIONS = &H31
 CSIDL_CONTROLS = &H3
 CSIDL_COOKIES = &H21
 CSIDL_DESKTOP = &H0
 CSIDL_DESKTOPDIRECTORY = &H10
 CSIDL_DRIVES = &H11
 CSIDL_FAVORITES = &H6
 CSIDL_FLAG_CREATE = &H8000
 CSIDL_FLAG_DONT_VERIFY = &H4000
 CSIDL_FLAG_MASK = &HFF00&
 CSIDL_FLAG_PFTI_TRACKTARGET = CSIDL_FLAG_DONT_VERIFY
 CSIDL_FONTS = &H14
 CSIDL_HISTORY = &H22
 CSIDL_INTERNET = &H1
 CSIDL_INTERNET_CACHE = &H20
 CSIDL_LOCAL_APPDATA = &H1C
 CSIDL_MYPICTURES = &H27
 CSIDL_NETHOOD = &H13
 CSIDL_NETWORK = &H12
 CSIDL_PERSONAL = &H5
 CSIDL_PRINTERS = &H4
 CSIDL_PRINTHOOD = &H1B
 CSIDL_PROFILE = &H28
 CSIDL_PROGRAM_FILES = &H26
 CSIDL_PROGRAM_FILES_COMMON = &H2B
 CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
 CSIDL_PROGRAM_FILESX86 = &H2A
 CSIDL_PROGRAMS = &H2
 CSIDL_RECENT = &H8
 CSIDL_SENDTO = &H9
 CSIDL_STARTMENU = &HB
 CSIDL_STARTUP = &H7
 CSIDL_SYSTEM = &H25
 CSIDL_SYSTEMX86 = &H29
 CSIDL_TEMPLATES = &H15
 CSIDL_WINDOWS = &H24
End Enum


Public Enum ShellFileInfoConstants
 BCF_SHGFI_ADDOVERLAYS = &H20
 BCF_SHGFI_ATTR_SPECIFIED& = &H20000
 BCF_SHGFI_ATTRIBUTES& = &H800
 BCF_SHGFI_DISPLAYNAME& = &H200
 BCF_SHGFI_EXETYPE& = &H2000
 BCF_SHGFI_ICON& = &H100
 BCF_SHGFI_ICONLOCATION& = &H1000
 BCF_SHGFI_LARGEICON& = &H0
 BCF_SHGFI_LINKOVERLAY& = &H8000
 BCF_SHGFI_OPENICON& = &H2
 BCF_SHGFI_OVERLAYINDEX& = &H40
 BCF_SHGFI_PIDL& = &H8
 BCF_SHGFI_SELECTED& = &H10000
 BCF_SHGFI_SHELLICONSIZE& = &H4
 BCF_SHGFI_SMALLICON& = &H1
 BCF_SHGFI_SYSICONINDEX& = &H4000
 BCF_SHGFI_TYPENAME& = &H400
 BCF_SHGFI_USEFILEATTRIBUTES& = &H10
End Enum
Public Enum FileAttributeConstants
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_COMPRESSED = &H800
    FILE_ATTRIBUTE_DEVICE = &H40
    FILE_ATTRIBUTE_DIRECTORY = &H10
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
    FILE_ATTRIBUTE_OFFLINE = &H1000
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_REPARSE_POINT = &H400
    FILE_ATTRIBUTE_SPARSE_FILE = &H200
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
    FILE_FLAG_BACKUP_SEMANTICS = &H2000000
    FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    FILE_FLAG_NO_BUFFERING = &H20000000
    FILE_FLAG_OPEN_NO_RECALL = &H100000
    FILE_FLAG_OPEN_REPARSE_POINT = &H200000
    FILE_FLAG_OVERLAPPED = &H40000000
    FILE_FLAG_POSIX_SEMANTICS = &H1000000
    FILE_FLAG_RANDOM_ACCESS = &H10000000
    FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    FILE_FLAG_WRITE_THROUGH = &H80000000
End Enum
Public Enum IconSizeConstants
    icon_large = SHGFI_ICON
    icon_shell = SHGFI_SHELLICONSIZE
    ICON_SMALL = SHGFI_SMALLICON
End Enum
Public Type OpenWithListItem
    StrName As String
    strCommandLine As String
    strParameters As String
End Type
'Public Const STD_INPUT_HANDLE As Long = -10&
'Public Const STD_OUTPUT_HANDLE As Long = -11&
'Public Const STD_ERROR_HANDLE As Long = -12&
Public Enum StandardIOStreamsEnum
    Standard_Input = -10&
    standard_Output = -11&
    standard_Error = -12&
End Enum


'declarations for the "FileOperation" routine:

Public Enum FileOperationErrors
    DE_SAMEFILE = &H71 'Source and Destination are the same.
    DE_MANYSRC1DEST = &H72 'Multiple file paths were specified in the source, but only one destination.
    DE_DIFFDIR = &H73 'Rename Operation was specified but destination path is a different directory. Use the move operation instead.
    DE_ROOTDIR = &H74 'The source is a root directory, which cannot be moved or renamed.
    DE_OPCANCELLED = &H75 'Operation was cancelled by the user or silently cancelled if appropriate flags were supplied.
    DE_DESTSUBTREE = &H76 'Destination is a subtree of the source.
    DE_ACCESSDENIEDSRC = &H78 'Security settings denied access to the source.
    DE_PATHTOODEEP = &H79 'Source or destination paths would exceed max_path
    DE_MANYDEST = &H7A 'Operation involved multiple destination maths, which can fail in the case of a move operation.
    DE_INVALIDFILES = &H7C 'the path in the source or destination is invalid.
    DE_DESTSAMETREE = &H7D 'Source and destination have the same parent folder.
    DE_FLDDESTISFILE = &H7E 'The destination path is an existing file.
    DE_FILEDESTISFLD = &H80 'The destination is an existing folder.
    DE_FILENAMETOOLONG = &H81 'Filename exceeds MAX_PATH,
    DE_DEST_IS_CDROM = &H82 'Destination is CD-ROM, possibly unformatted.
    DE_DEST_IS_DVD = &H83 'Destination is read-only DVD, possibly unformatted.
    DE_DEST_IS_CDRECORD = &H84 'destination is a writable CD-ROM, possibly unformatted.
    DE_FILE_TOO_LARGE = &H85 'the file involved in the operation is too large for the destination media or file system
    DE_SRC_IS_CDROM = &H86 'The source is a read-only CD-ROM, possibly unformatted.
    DE_SRC_IS_DVD = &H87 'the source is a read-only DVD, possibly unformatted.
    DE_ERROR_MAX = &HB7 'MAX_PATH was exceeded during the operation.
    ERRORONDEST = &H10000 'unspecified error occured on the destination.
    ROOTDIRRENAME = &H10074 'destination is a root directory and cannot be renamed.
End Enum
Public Enum FileOperationFlags
    FOF_ALLOWUNDO = &H40
    FOF_CONFIRMMOUSE = &H2
    FOF_FILESONLY = &H80
    FOF_MULTIDESTFILES = &H1
    FOF_NO_CONNECTED_ELEMENTS = &H2000
    FOF_NOCONFIRMATION = &H10
    FOF_NOCONFIRMMKDIR = &H200
    FOF_NOCOPYSECURITYATTRIBS = &H800
    FOF_NOERRORUI = &H400
    FOF_NORECURSION = &H1000
    FOF_RENAMEONCOLLISION = &H8
    FOF_SILENT = &H4
    FOF_SIMPLEPROGRESS = &H100
    FOF_WANTMAPPINGHANDLE = &H20
    FOF_WANTNUKEWARNING = &H4000

End Enum
Public Enum CopyFileExFlags
 COPY_FILE_FAIL_IF_EXISTS = &H1
 COPY_FILE_RESTARTABLE = &H2
 COPY_FILE_OPEN_SOURCE_FOR_WRITE = &H4
 COPY_FILE_ALLOW_DECRYPTED_DESTINATION = &H8
 COPY_FILE_COPY_SYMLINK = &H800
 COPY_FILE_NO_BUFFERING = &H1000
End Enum
Public Enum CopyProgressRoutineReturnConstants
 PROGRESS_CONTINUE = 0
 PROGRESS_CANCEL = 1
 PROGRESS_STOP = 2
 PROGRESS_QUIET = 3
End Enum
Public Enum CopyProgressRoutineStateChangeConstants
    CALLBACK_CHUNK_FINISHED = &H0
    CALLBACK_STREAM_SWITCH = &H1

End Enum




Private WithEvents Downloader As CDownload
Attribute Downloader.VB_VarHelpID = -1
Private mDownErr As Long, mDownDesc As String, mdowncomplete As Boolean

'Public Function PerformFilteronString(Byval












Property Get g_ForceANSI() As Long
    g_ForceANSI = ForceANSI
End Property
Public Property Let g_ForceANSI(ByVal Vdata As Long)
    ForceANSI = Vdata
End Property
Private Sub Downloader_Completed()
'
'Set Downloader = Nothing
mdowncomplete = True
End Sub

Private Sub Downloader_DataArrival(ByVal bytesTotal As Long)
'
Debug.Print "dataarrival, " & bytesTotal
End Sub

Private Sub Downloader_Error(ByVal Number As Integer, Description As String)
'

mdowncomplete = True

'    Set Downloader = Nothing
End Sub

Private Sub Downloader_Starting(ByVal FileSize As Long, ByVal Header As String)
'
End Sub
Public Function IsValidPath(ByVal strpath As String) As Boolean
    On Error Resume Next
    Dim openedStream As FileStream
    If Not GetDirectory(strpath) Is Nothing Then
        IsValidPath = True
    ElseIf Not GetFile(strpath) Is Nothing Then
        IsValidPath = True
    Else
    Set openedStream = OpenStream(strpath)
        If Not openedStream Is Nothing Then
            IsValidPath = True
            openedStream.CloseStream
        End If
    End If




End Function
Public Function GetSaveFileName(ByVal OwnerWnd As Long, ByVal Title As String, ByVal Filter As String, ByVal DefaultExtension As String, Optional ByVal Flags As OpenFileNameFlags = OFN_EXPLORER, Optional callbackobj As IDialogHook) As String
    'uses the cFileDialog class to acquire a filename to save to.
    Dim savedialog As CFileDialog
    Set savedialog = New CFileDialog
    With savedialog
    .DefExt = DefaultExtension
    .Filter = Filter
    Set .EventCallback = callbackobj
    .Flags = Flags
    .Caption = Title
    GetSaveFileName = .SelectSaveFile(OwnerWnd)
    End With
End Function

Public Function DownloadRemoteFile(ByVal URL As String, Progressobj As IProgressCallback) As CFile
    'downloads the file; saves to temp and returns the file.

    Dim tempFile As String
    tempFile = GetTempFileNameAndPathEx()
    'ftp is port 21.
    Set Downloader = New CDownload
    With Downloader
    Downloader.URL = URL
    mdowncomplete = False
    Set Downloader.ProgressObject = Progressobj
    Downloader.Download URL, tempFile
    Do Until mdowncomplete
        DoEvents
    Loop
    
    End With
    'Stop
    
 
    
    If mDownErr <> 0 Then
        Err.Raise mDownErr, , mDownDesc
    
    Else
        Set DownloadRemoteFile = Me.GetFile(tempFile)
    
    
    End If
    
    
    mDownErr = 0
    mDownDesc = ""




End Function



Public Sub testopenwith()

Dim ret() As OpenWithListItem
Dim lCount As Long



ret = GetOpenWithList("*", lCount)

 
'Stop
End Sub
Private Sub MergeOpenwith(ArrayMergewith() As OpenWithListItem, arraymerge() As OpenWithListItem)
    Dim I As Long, NewUB As Long, oldUB As Long
    'On Error GoTo Breakout
    On Error Resume Next
    NewUB = UBound(ArrayMergewith)
        If Err <> 0 Then
            ReDim ArrayMergewith(1 To 1)
        End If
        Err.Clear
        oldUB = NewUB
    NewUB = NewUB + UBound(arraymerge)
     
     If Err <> 0 Then
        Exit Sub
    End If
    ReDim Preserve ArrayMergewith(1 To NewUB)
    For I = LBound(arraymerge) To UBound(arraymerge)
        ArrayMergewith(oldUB + I) = arraymerge(I)
        
    
    Next I
            

Exit Sub
breakout:
Exit Sub
notinitialized:
    'ReDim ArrayMergewith(1 To 1)
    'Resume
    
End Sub
Public Function FormatOpenWithData(ByVal ForExtension As String) As String

    Dim gotdata() As OpenWithListItem
    Dim gotcount As Long, I As Long, ret As String
    gotdata = GetOpenWithList(ForExtension, gotcount)
    RemoveDuplicates gotdata
    If gotcount = 0 Then
        ret = "No Files returned :("
    Else
        For I = LBound(gotdata) To UBound(gotdata)
            With gotdata(I)
                ret = ret & "----------" & vbCrLf
                ret = ret & "Name:" & .StrName & " Commandline=""" & .strCommandLine & """"
                ret = ret & vbCrLf
            End With
    
        Next
    End If
    FormatOpenWithData = ret

End Function

Private Function GetOpenWithListHKCUData(ByVal Classkey As Long, ByVal Key As String, ByRef AddTo() As String) As Long

'HKCU\.fileext or HKCU\type list openwith data as subsections.
'OpenWithList is a bit less clearcut, first, open "MRUList" value from that key.
'Enumerate each Character, loading the value in the value with the name of the character (for example, "a")
'
'look in "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\, for a value matching each found executable in the list.
    If Right$(Key, 1) <> "\" Then Key = Key & "\"
    Key = Key & "OpenWithList"
    mreg.Classkey = Classkey
    mreg.SectionKey = Key
    Dim enumnames() As String
    Dim enumcount As Long
    Call mreg.EnumerateSections(enumnames, enumcount)




End Function
Private Function GetOpenWithListFileExtsData(ByVal Key As String, ByRef AddTo()) As Long
    If Right$(Key, 1) <> "\" Then Key = Key & "\"
    Key = Key & "OpenWithList"
    
    
End Function

Private Function GetOpenWithListN(ByVal ForExtension As String, Optional ByRef ListCount As Long = 0) As OpenWithListItem()

'Windows XP checks HKEY_CLASSES_ROOT\.ext (where .ext is the extension that defines the file type).
'If it finds an OpenWith subkey, the applications listed under that subkey are added to the Open With menu,
'and they also appear in the Open With dialog box in the Programs list's Recommended Programs branch.

'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\
Dim appexes() As String, sSect() As String, Scount As Long

'Check HKEY_CLASSES_ROOT\<Forextension>, If we find an OpenWithList or OpenWithProgIDs:

'with OpenWithList, Enumerate each sub-section of OpenWithList; each section is a executable name. Look up the executable name in
' and add it to "AppExes()"

'with OpenWithProgIDs, enumerate the values; each one needs to be used in a recursive call, as "forExtension".

'check HKEY_CLASSES_ROOT\.ext to see whether the file type has a PerceivedType Value.
'If so, it means the file type also has an associated perceived type; use that to look into:

'HKEY_CLASSES_ROOT\SystemFileAssociations\PerceivedType\OpenWithList

'and enumerate the sections for further executables to add to the growing list.


End Function

Public Function GetOpenWithList(ByVal ForExtension As String, Optional ByRef ListCount As Long = 0) As OpenWithListItem()

    Dim creg As cRegistry
'find open with commands/names for a extension.

'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\Extension
'For here, there will be a key,
Dim tmpreturn() As OpenWithListItem
Dim retArray() As OpenWithListItem
Dim baseExplorer As String
Dim values() As String, vcount As Long
Dim sect() As String, Scount As Long
Dim I As Long, retcount As Long, mruList As String
Dim appexes() As String, acount As Long
Dim lCount As Long
Dim Currpos As Long, CurrChar As String, userChoice As String

baseExplorer = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ForExtension
Set creg = New cRegistry
creg.Classkey = HHKEY_CURRENT_USER
creg.SectionKey = baseExplorer & "\OpenWithprogIDs"


'OpenWithprogIDs. Enumerate the values, and recursively call this routine, passing the value name as the Extension.
If creg.EnumerateValues(values(), vcount) Then
    For I = 1 To vcount
        retcount = 0
        If StrComp(values(I), ForExtension, vbTextCompare) <> 0 Then
        tmpreturn = GetOpenWithList(values(I), retcount)
            If retcount > 0 Then
                MergeOpenwith retArray(), tmpreturn()
                
            
            End If
        End If
    Next I
End If

Dim perceived As String

'get perceived type from the value: HKCR\<forextension>\PercievedType.

'check HKEY_CLASSES_ROOT\SystemFileAssociations\<PerceivedType>\OpenWithList;


perceived = creg.ValueEx(hhkey_classes_root, ForExtension, "PerceivedType", RREG_SZ, "")
If perceived <> "" Then
'Stop
    creg.Classkey = hhkey_classes_root
    creg.SectionKey = "SystemFileAssociations\" & perceived & "\OpenWithList"
    creg.ValueKey = ""
    If creg.EnumerateSections(sect(), Scount) Then
    
        For I = 1 To Scount
            acount = acount + 1
            ReDim Preserve appexes(1 To acount)
            appexes(I) = sect(I)
            
        
        Next I

    End If


End If



'OpenWithList is a bit less clearcut, first, open "MRUList" value from that key.
'Enumerate each Character, loading the value in the value with the name of the character (for example, "a")
'each entry is an application Executable, which we later add to our array- just add it to the appexes() array, for now.
creg.SectionKey = baseExplorer & "\OpenWithList"
    mruList = creg.ValueEx(HHKEY_CURRENT_USER, baseExplorer & "\OpenWithList", "MRUList", RREG_SZ, "")
    
    'iterate through each character...
    For Currpos = 1 To Len(mruList)
        CurrChar = Mid$(mruList, Currpos, 1)
        
        acount = acount + 1
        ReDim Preserve appexes(1 To acount)
        appexes(acount) = creg.ValueEx(HHKEY_CURRENT_USER, baseExplorer & "\OpenWithList", CurrChar, RREG_SZ, "")
        
        
        
    Next Currpos
    
    'whew, got the Application executables- for this one.
    



'HKEY_CLASSES_ROOT\Extension

'OpenWithProgIds is a group of values here, each one is the name of the class- again- recursively call this same routine and
'merge the results into our result array.


'add the default executable.

If creg.ValueEx(hhkey_classes_root, ForExtension & "\shell\Open\Command", "", RREG_SZ, "") <> "" Then
    acount = acount + 1
    ReDim Preserve appexes(1 To acount)
    appexes(acount) = creg.ValueEx(hhkey_classes_root, ForExtension & "\shell\Open\Command", "", RREG_SZ, "")
    If Left$(appexes(acount), 1) = """" Then
        appexes(acount) = Mid$(appexes(acount), 2, InStr(3, appexes(acount), """", vbTextCompare) - 2)
        'appexes(acount) = GetFilenamePart(appexes(acount))
    End If

End If

creg.Classkey = hhkey_classes_root

creg.SectionKey = ForExtension
If creg.value <> "" Then
        tmpreturn = GetOpenWithList(creg.value, lCount)
        If lCount > 0 Then
            MergeOpenwith retArray(), tmpreturn()
        End If

End If


creg.SectionKey = ForExtension & "\OpenWithProgids"
'HKEY_CLASSES_ROOT\.txt\OpenWithProgids
If creg.EnumerateValues(sect(), Scount) Then
    On Error Resume Next
    For I = 1 To UBound(sect)
        If Err <> 0 Then
            Err.Clear
            Exit For
        End If
        tmpreturn = GetOpenWithList(sect(I), lCount)
        If lCount > 0 Then
            MergeOpenwith retArray(), tmpreturn()
        End If
    Next I


End If
'whew....

'"Openwithlist" has the same format as that in the HKCU key described above.

'**wrong**
'within HKEY_CLASSES_ROOT, it will be sections as well.
    creg.Classkey = hhkey_classes_root
    creg.SectionKey = ForExtension & "\OpenWithList"
    If creg.EnumerateSections(sect, Scount) Then
        For I = 1 To Scount
            acount = acount + 1
            ReDim Preserve appexes(1 To acount)
            appexes(I) = sect(I)
            
        
        Next I

    End If
'   mruList = creg.ValueEx(HHKEY_CLASSES_ROOT, ForExtension & "\OpenWithList", "MRUList", RREG_SZ, "")
'
'    'iterate through each character...
'    For CurrPos = 1 To Len(mruList)
'        CurrChar = Mid$(mruList, CurrPos, 1)
'        If creg.ValueEx(HHKEY_CURRENT_USER, ForExtension & "\OpenWithList", CurrChar, RREG_SZ, "") <> "" Then
'            acount = acount + 1
'            ReDim Preserve AppExes(1 To acount)
'            AppExes(acount) = creg.ValueEx(HHKEY_CURRENT_USER, ForExtension & "\OpenWithList", CurrChar, RREG_SZ, "")
'        End If
'
'
'    Next CurrPos



'lastly, if the default value of HCR\extension is present as HCR\<value>, then call recursively, this routine with that value and merge the array.
Dim defvalue As String
defvalue = creg.ValueEx(hhkey_classes_root, ForExtension, "", RREG_SZ, "")
'change: also check "old Default" value.
If defvalue <> "" Then
    creg.SectionKey = defvalue
    If creg.KeyExists Then
    tmpreturn = GetOpenWithList(defvalue, lCount)
        If lCount > 0 Then
            MergeOpenwith retArray, tmpreturn
        End If
    End If
End If
Dim newExes() As String, newexecount As Long
Dim LoopOuter As Long, LoopInner As Long
': restartloop
Dim innerStr As String, outerStr As String, flFound As Boolean


For LoopOuter = 1 To UBound(appexes)



    appexes(I) = Trim$(appexes(I))
Next LoopOuter


For LoopOuter = 1 To UBound(appexes)
    If Err <> 0 Then Exit For
    Err.Clear
    outerStr = appexes(LoopOuter)
    flFound = False
    For LoopInner = 1 To UBound(appexes)
        If LoopOuter <> LoopInner Then
                If StrComp(appexes(LoopInner), outerStr, vbTextCompare) = 0 Then
                    flFound = True
                    Exit For
                End If
                
                
        
        
        End If
    Next
    If Not flFound Then
        newexecount = newexecount + 1
        ReDim Preserve newExes(1 To newexecount)
        newExes(newexecount) = outerStr
    End If
Next
If Err = 0 Then
appexes = newExes
acount = newexecount
End If
'iterate through appexes, retrieve executable names from:
'"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"

'use GetEXEFriendlyName(scmd) for friendly name.
On Error Resume Next
For I = 1 To UBound(appexes)
    If Err <> 0 Then Exit For
    If BCFile.FileExists(appexes(I)) Then
        defvalue = Trim$(appexes(I))
    
    
    Else
        defvalue = creg.ValueEx(HHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & appexes(I), "", RREG_SZ, "")
        
        If defvalue = "" Then
            Dim bufferpath As String
            bufferpath = Space$(255)
            Call SearchPath(0, appexes(I), ".exe", Len(bufferpath), bufferpath, "")
            'use searchpath()
            bufferpath = Left$(bufferpath, InStr(bufferpath, vbNullChar) - 1)
            defvalue = bufferpath
        End If
    End If
    defvalue = Trim$(defvalue)
    If defvalue <> "" Then
        'it is the full path-
        On Error Resume Next
        ReDim Preserve retArray(1 To UBound(retArray) + 1)
        If Err <> 0 Then ReDim Preserve retArray(1 To 1)
        With retArray(UBound(retArray))
            .strCommandLine = defvalue
            .strCommandLine = Replace$(.strCommandLine, vbNullChar, "")
            .strCommandLine = Replace$(.strCommandLine, """", "")
            .StrName = GetEXEFriendlyName(.strCommandLine)
            .StrName = Replace$(.StrName, vbNullChar, "")
            If .StrName = "" Then Stop
        
        End With
        
    
    End If
    
    
Err.Clear
Next I
On Error Resume Next

RemoveDuplicates retArray


ListCount = UBound(retArray)
GetOpenWithList = retArray
End Function
Private Sub RemoveDuplicates(dupsremove() As OpenWithListItem)

'remove all duplicated commands.
Dim LoopOuter As Long, LoopInner As Long, retthis() As OpenWithListItem, retcount As Long
Dim flDupFound As Long
Dim CompareA As String, CompareB As String
On Error Resume Next
For LoopOuter = LBound(dupsremove) To UBound(dupsremove)
    flDupFound = False
    If Err <> 0 Then Exit Sub
    Err.Clear
    For LoopInner = LBound(retthis) To UBound(retthis)
        If Err <> 0 Then Exit For
        If LoopOuter <> LoopInner Then
        'If InStr(dupsremove(LoopOuter).strCommandLine, "wmplayer") <> 0 Then Stop
            CompareA = dupsremove(LoopOuter).strCommandLine
            CompareB = retthis(LoopInner).strCommandLine
            CompareA = Trim$(Replace$(CompareA, """", ""))
            CompareB = Trim$(Replace$(CompareB, """", ""))
            
            If StrComp(CompareA, CompareB, vbTextCompare) = 0 Then
                flDupFound = True
                Exit For
            End If

        End If
    Next LoopInner
    Err.Clear
    If Not flDupFound Then
        retcount = retcount + 1
        ReDim Preserve retthis(1 To retcount)
        retthis(retcount) = dupsremove(LoopOuter)
    End If
    
Next LoopOuter
dupsremove = retthis
End Sub

Private Function GetOpenWithListInt(ByVal ForExtension As String, Optional ByRef ListCount As Long, Optional ByVal useclasskey As ERegistryClassConstants = hhkey_classes_root, Optional ByVal SectionKey As String = "") As OpenWithListItem()
'TODO//improve this routine.
 
'improvements:
'find the "other" locations of items that seem to be able to add themselves to open with without any sort of key (???)
'damn it- we need to do the SAME thing with the
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts
'as we do currently in HKEY_CLASSES_ROOT.

    Dim retArr() As OpenWithListItem, retcount As Long
    Dim tmpreturn() As OpenWithListItem, I As Long
    Dim reguse As cRegistry, lCount As Long
    Dim appexes() As String, currcopy As Long
    Dim sect() As String, Scount As Long, vals() As String, vcount As Long
    Dim defvalue As String
    Dim usekey As String
    Dim usebase As String, basesection As String
    basesection = SectionKey
    usebase = SectionKey & "\" & ForExtension
    usekey = usebase
    Set reguse = New cRegistry
    reguse.Classkey = useclasskey
    reguse.SectionKey = usekey
    If ForExtension = "" Then Exit Function
    
    
    
    
    If reguse.KeyExists Then
        'rules:
        'open .ForExtension\openWith key, enumerate sections.

    
         
            reguse.EnumerateSections sect, Scount
            If useclasskey <> hhkey_classes_root Then
                reguse.EnumerateValues vals, vcount
                ReDim Preserve sect(1 To UBound(sect) + vcount)
                For I = 1 To vcount
                    sect(Scount + vcount) = vals(vcount)
                    
                
                Next I
                    
                vcount = 0
            End If
        'Add to our current appExes() array
        'pah- just direct copy it.
        appexes = sect
        
        
       
       reguse.SectionKey = usebase & "\openWithprogids"
        'open .ForExtension\openWithProgids key, enumerate values.
        Call reguse.EnumerateValues(vals(), vcount)
        
            defvalue = reguse.ValueEx(useclasskey, usebase, "", RREG_SZ, "")
            If defvalue <> "" Then
                'If reguse.ValueEx(useclasskey, defvalue, "", RREG_SZ, "") <> "" Then
            
            
            
                     ReDim Preserve vals(1 To vcount + 1)
                     vals(UBound(vals)) = reguse.ValueEx(useclasskey, usekey, "", RREG_SZ, "")
                    vcount = vcount + 1
                'End If
            End If
            defvalue = reguse.ValueEx(useclasskey, usebase, "Old Default", RREG_SZ, "")
             If defvalue <> "" Then
                'If reguse.ValueEx(useclasskey, usekey, "", RREG_SZ, "") <> "" Then
                     ReDim Preserve vals(1 To vcount + 1)
                     vals(UBound(vals)) = reguse.ValueEx(useclasskey, usekey, "Old Default", RREG_SZ, "")
                    vcount = vcount + 1
                'End If
            End If
            
            
            
        If vcount Then
        
            For I = 1 To vcount
                tmpreturn = GetOpenWithListInt(vals(I), lCount, useclasskey, basesection)
                       'for each value
                'open that key in HKEY_CLASSES_ROOT- recursively by calling this same function with the value as the parameter.
                'append to our array the results from said function.
                'append this array to our retarr() array.
                For currcopy = 1 To lCount
                    retcount = retcount + 1
                    ReDim Preserve retArr(1 To retcount)
                    retArr(retcount) = tmpreturn(I)
                
                
                Next
                
                
                
            Next I
     
        
        End If
    End If
    
    'recurse, and add entries from:
    'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts
    If SectionKey <> "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts" Then
    tmpreturn = GetOpenWithListInt(ForExtension, lCount, HHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts")
    Debug.Print "explorer sees " & lCount & " entries." & " for " & ForExtension
    End If

'
'
'
'
'    End If
    
    'convert our appexes() array into proper entries in our return by converting the EXE names to paths via the sections here:
    'the default value is the full path to the exe.
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths

'and append it...
    On Error Resume Next
    For I = LBound(appexes) To UBound(appexes)
        If Err = 0 Then
            Dim scmd As String
            scmd = reguse.ValueEx(HHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & appexes(I), "", RREG_SZ, "")
            'TODO:// use "Searchpath" API to find files that return "" for scmd here.
            
            
            If scmd <> "" Then
            retcount = retcount + 1
            ReDim Preserve retArr(1 To retcount)
            
            'retArr(retcount).StrName =  AppExes(i)
            If Right$(scmd, 1) = vbNullChar Then
                scmd = Mid$(scmd, 1, Len(scmd) - 1)
            End If
            
            scmd = Replace$(scmd, """", "")
            retArr(retcount).StrName = GetEXEFriendlyName(scmd)
            
            End If
            retArr(retcount).strCommandLine = scmd
        
    
        End If
    Next I
    
    GetOpenWithListInt = retArr
    ListCount = retcount
    Debug.Print ListCount
End Function
'Public Function PerformFileOperation(StrFrom() As String, StrTo() As String, ByVal Operation As FILEOP, Optional ProgressCallback As IProgressCallback) As Boolean
'
'
'
'
'
'
'End Function

Public Function PerformFileOperation(ByVal StrFrom As String, StrTo As String, ByVal Operation As FILEOP, ByVal Flags As FILEOP_FLAGS, _
    Optional ByVal hWndOwner As Long, Optional ByVal Silent As Boolean, Optional ByVal ProgressMessage As String = "") As Long

    Dim OpStruct As SHFILEOPSTRUCT, ret As Long
    'strFrom and StrTo should be properly Null delimited...
    With OpStruct
        .hwnd = hWndOwner
        .pFrom = StrFrom & vbNullChar & vbNullChar
        .pTo = StrTo & vbNullChar & vbNullChar
        .sProgress = ProgressMessage
        .fFlags = Flags
        .wFunc = Operation
    End With
    ret = SHFileOperation(OpStruct)










    End Function
Public Sub ShellExec(ByVal OwnerWnd As Long, ByVal StrFilename As String, Optional ByVal Verb As String = "Open")


     Dim dummy As Long, executablestart$
     Dim FilePath As String
     FilePath = StrFilename
               'open the file using the default Editor or viewer.
     dummy = ShellExecute(OwnerWnd, Verb, FilePath & Chr$(0), Chr$(0), Left$(FilePath, InStr(FilePath, "\")), vbNormalFocus)
     

     
     
End Sub



Public Function GetStreamCRC16(ofStream As IInputStream) As Long
    Static crcobj As cCRC32
    If crcobj Is Nothing Then Set crcobj = New cCRC32
    
    GetStreamCRC16 = crcobj.GetStreamCRC16(ofStream)

End Function
Public Function GetStreamCRC32(ofStream As IInputStream) As Long
    Static crcobj As cCRC32
    If crcobj Is Nothing Then Set crcobj = New cCRC32
    
    GetStreamCRC32 = crcobj.GetStreamCRC32(ofStream)

End Function
Public Function GetStreamMD5sum(ofStream As IInputStream) As Variant
    Static msum As Cmd5sum
    Dim strread As String
    If msum Is Nothing Then Set msum = New Cmd5sum
    ofStream.SeekTo 0, STREAM_BEGIN
    strread = ofStream.ReadAllStr(StrRead_ANSI)
    GetStreamMD5sum = msum.MD5(strread)
   




End Function
Public Function GetStreamSHA256Sum(ofStream As IInputStream) As Variant
    Static msum As CSHa256
    Dim strread As String
    If msum Is Nothing Then Set msum = New CSHa256
    ofStream.SeekTo 0, STREAM_BEGIN
    strread = ofStream.ReadAllStr(StrRead_ANSI)
    GetStreamSHA256Sum = msum.SHA256(strread)




End Function
    
Public Property Get CurrentDirectory() As Directory
    Dim Curget As String
  
    Call GetCurrentDirectory(Curget)
    Set CurrentDirectory = GetDirectory(Curget)
End Property
Public Function CreateDirectory(ByVal PathSpec As String, Optional ByVal templateDir As Directory = Nothing, Optional ByVal CreateParentFolders As Boolean = False) As Directory
Attribute CreateDirectory.VB_Description = "Creates the specified Directory path and returns it."
    'Call CreateDirectory Routine...
    Dim retval As Long, hDir As Long
    If templateDir Is Nothing Then
    'Easy.
    
    'first- if we are to create parent folders described then do so.
    
    'retval = mdlfilesystem.CreateDirectoryEx(
    retval = MdlFileSystem.CreateDirectory(PathSpec, 0)
    If retval = 0 Then
        RaiseAPIError Err.LastDllError, "BCFSObject::CreateDirectory"
    Else
        Set CreateDirectory = GetDirectory(PathSpec)
    End If

    Else
      
            retval = CreateDirectoryEx(templateDir.Path, PathSpec, 0)
            If retval = 0 Then
                RaiseAPIError Err.LastDllError, "BCFSObject::CreateDirectory"
            Else
            'success.
            Set CreateDirectory = GetDirectory(PathSpec)
            
            End If
        
        
            CloseHandle hDir
    End If
    
    
End Function
Public Function CreateFile(ByVal PathSpec As String, Optional ByVal forceoverwrite As Boolean = False) As CFile
Attribute CreateFile.VB_Description = "Creates the specified file and returns a CFile object that represents it."
    'create the file, and return a reference to it as well.
    
    Dim ffile As CFile
    
    
    
    PathSpec = FixPath(PathSpec)
    Set ffile = New CFile
    ffile.Create PathSpec, forceoverwrite
    Set CreateFile = ffile

End Function

Friend Function GetScratchStream() As FileStream
    Set GetScratchStream = OpenTempFile()
End Function
Public Function OpenTempFile(Optional ByVal PrefixStr As String = "BC", Optional ByVal dwDesiredAccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal dwShareMode As FileShareModeConstants = FILE_SHARE_READ) As FileStream
    Dim tpath As String, tfile As String, FHandle As Long
    Dim retfile As CFile, retStream As FileStream
    'creates and returns a file designated as a temporary file.
    tpath = Space$(2048)
    Call GetTempPath(2047, tpath)
    tpath = Left$(tpath, InStr(tpath, vbNullChar) - 1)
    tfile = GetTempFileNameEx
    
    tfile = IIf(Right$(tpath, 1) <> "\", tpath & "\", tpath) & tfile
    'ok, create the file via a handle, as such.
    'FHandle = MdlFileSystem.CreateFile(tfile, GENERIC_ALL, FILE_SHARE_READ + FILE_SHARE_WRITE, ByVal &O0, CREATE_NEW, FILE_FLAG_DELETE_ON_CLOSE, 0)
    'Set retfile = New CFile
    'retfile.OpenHandle FHandle
    Set retStream = New FileStream
    retStream.OpenStream tfile, dwDesiredAccess, dwShareMode, CREATE_NEW, FILE_FLAG_DELETE_ON_CLOSE
    
    Set OpenTempFile = retStream
    'Set retStream = retfile.OpenAsBinaryStream(dwDesiredAccess, FILE_SHARE_READ, OPEN_EXISTING, FILE_FLAG_DELETE_ON_CLOSE)
    
    'CloseHandle FHandle
End Function
Public Function GetVolume(ByVal PathSpec As String, Optional ByVal RaiseError As Boolean = False) As CVolume
'GetVolume: Returns a Volume for a given Drive/Volume Name.
'If an Error occurs, Nothing is returned, and the Err Object is populated as appropriate.


    Dim newvol As CVolume
    Set newvol = New CVolume
    On Error GoTo ErrorGetVolume
    newvol.Init PathSpec
    Set GetVolume = newvol

    Exit Function
ErrorGetVolume:
    If RaiseError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        Set GetVolume = Nothing
    End If
    
    
    
End Function
Public Function ShortenPath(ByRef strpath As String, ByVal hDC As Long, ByVal MaxPixelWidth As Long) As String
    Dim shortened As String
    'PathCompactPathEx
    'PathCompactPathEx(shortened,strpath,
    shortened = Space$(MAX_PATH)
    Mid$(shortened, 1, Len(strpath & vbNullChar)) = strpath & vbNullChar
    PathCompactPath hDC, strpath, MaxPixelWidth
    ShortenPath = strpath

End Function
Public Function ShortenPathToCharLength(ByVal strpath As String, ByVal Desiredwidth As Long) As String
    Dim shortened As String
    shortened = Space$(MAX_PATH)
    PathCompactPathEx shortened, strpath, Desiredwidth, 0
    ShortenPathToCharLength = shortened

End Function
Public Function GetFile(ByVal PathSpec As String, Optional ByVal RaiseError As Boolean = False) As CFile

    PathSpec = FixPath(PathSpec)
    Dim newfile As CFile
    Dim getattributes As Long
    Set newfile = New CFile
    If FileSystem.Exists(PathSpec) Then
        If (GetFileAttributes(PathSpec) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            Err.Raise 53, "GetFile", "GetFile Called with a directory specification"
        Else
        Call newfile.SetFileName(PathSpec)
        Set GetFile = newfile
        End If
    Else
        If RaiseError Then
            Err.Raise 53, "BCFSObject::GetFile", "File Not Found, """ & PathSpec & """"
        Else
            Set GetFile = Nothing
        End If
        'Err.Number = 53
        'Err.Description = "File Not Found"
        
    
    End If



End Function
Public Function GetObjIcon(Objstr As String, iconsize As IconSizeConstants) As Long

    GetObjIcon = GetFileIcon(Objstr, iconsize)


End Function



Public Function GetDirectory(ByVal Path As String, Optional ByVal RaiseError As Boolean = False) As Directory
    Dim NewDir As Directory
    Set NewDir = New Directory
    
    On Error GoTo ReturnNothing
    
    NewDir.SetDir Path
    Set GetDirectory = NewDir
    Exit Function
    
ReturnNothing:
    If RaiseError Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Else
        Set GetDirectory = Nothing
    End If
    
    
End Function
Public Function GetSpecialFolder(ByVal FolderConst As SpecialFolderConstants) As Directory
    Dim Spath As String
    Spath = Space$(256)
    Call SHGetSpecialFolderPath(0, Spath, FolderConst, BOOL.APITRUE)
    
    Spath = Trim$(Replace$(Spath, vbNullChar, " "))
    If Spath = "" Then Exit Function
    Set GetSpecialFolder = GetDirectory(Spath)



End Function
Public Function ExpandSpecialFolderNames(ByVal Strexpand As String) As String
ExpandSpecialFolderNamesSub Strexpand
ExpandSpecialFolderNames = Strexpand
End Function
Private Sub ExpandSpecialFolderNamesSub(ByRef Strexpand As String)
    'creates a collection of special folders...
    
    Dim FolderIDs
    Dim FolderNames
    Dim FolderPaths
    ' CSIDL_FONTS = &H14
' CSIDL_HISTORY = &H22
' CSIDL_INTERNET = &H1
' CSIDL_INTERNET_CACHE = &H20
' CSIDL_LOCAL_APPDATA = &H1C
' CSIDL_MYPICTURES = &H27
' CSIDL_NETHOOD = &H13
' CSIDL_NETWORK = &H12
' CSIDL_PERSONAL = &H5
' CSIDL_PRINTERS = &H4
' CSIDL_PRINTHOOD = &H1B
' CSIDL_PROFILE = &H28
' CSIDL_PROGRAM_FILES = &H26
' CSIDL_PROGRAM_FILES_COMMON = &H2B
' CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
' CSIDL_PROGRAM_FILESX86 = &H2A
' CSIDL_PROGRAMS = &H2
' CSIDL_RECENT = &H8
' CSIDL_SENDTO = &H9
' CSIDL_STARTMENU = &HB
' CSIDL_STARTUP = &H7
' CSIDL_SYSTEM = &H25
' CSIDL_SYSTEMX86 = &H29
' CSIDL_TEMPLATES = &H15
' CSIDL_WINDOWS = &H24
'Bitbucket Is invalid...
'Connections Is invalid...
'Controls Is invalid...
'Drives Is invalid...
'Internet Is invalid...
'Network Is invalid...
'Printers Is invalid...
    FolderIDs = Array(CSIDL_ADMINTOOLS, CSIDL_ALTSTARTUP, CSIDL_APPDATA, CSIDL_COMMON_ADMINTOOLS, CSIDL_COMMON_ALTSTARTUP, CSIDL_COMMON_APPDATA, _
    CSIDL_COMMON_DESKTOPDIRECTORY, CSIDL_COMMON_DOCUMENTS, CSIDL_COMMON_FAVORITES, CSIDL_COMMON_PROGRAMS, CSIDL_COMMON_STARTMENU, CSIDL_COMMON_STARTUP, _
    CSIDL_COMMON_TEMPLATES, CSIDL_COOKIES, CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY, CSIDL_FAVORITES, CSIDL_FONTS, _
    CSIDL_HISTORY, CSIDL_INTERNET_CACHE, CSIDL_LOCAL_APPDATA, CSIDL_MYPICTURES, CSIDL_NETHOOD, CSIDL_PERSONAL, _
    CSIDL_PRINTHOOD, CSIDL_PROFILE, CSIDL_PROGRAM_FILES, CSIDL_PROGRAM_FILES_COMMON, CSIDL_PROGRAMS, CSIDL_PROGRAM_FILES_COMMONX86, CSIDL_PROGRAMS, CSIDL_RECENT, CSIDL_SENDTO, CSIDL_STARTMENU, CSIDL_STARTUP, _
    CSIDL_SYSTEM, CSIDL_SYSTEMX86, CSIDL_TEMPLATES, CSIDL_WINDOWS)
    FolderNames = Array("AdminTools", "AltStartup", "appdata", "Common_AdminTools", "Common_AltStartup", "Common_Appdata", _
    "Common_desktop", "Common_Documents", "Common_Favorites", "Common_Programs", "Common_Startmenu", "Common_Startup", _
    "Common_Templates", "Cookies", "CSIDL_Desktop", "Desktop", "Favorites", "Fonts", _
    "History", "InternetCache", "Local_Appdata", "MyPictures", "NetHood", "Personal", _
    "PrintHood", "Profile", "programfiles", "CommonProgramFiles", "programs", "Commonx86", "programs", "recent", "SendTo", "Startmenu", "Startup")
    
    
    ReDim FolderPaths(0 To UBound(FolderIDs))
    Dim I As Long, tempfolder As Directory
    For I = 0 To UBound(FolderPaths)
    Set tempfolder = GetSpecialFolder(FolderIDs(I))
    If Not tempfolder Is Nothing Then
    FolderPaths(I) = tempfolder.Path
    Else
        Debug.Print FolderNames(I); " Is invalid..."
    End If
    Next I
    
    For I = 0 To UBound(FolderNames)
       Strexpand = Replace$(Strexpand, "%" & FolderNames(I) & "%", FolderPaths(I), , , vbTextCompare)
    Next I
    

    
    
'CSIDL_ADMINTOOLS = &H30
' CSIDL_ALTSTARTUP = &H1D
' CSIDL_APPDATA = &H1A
' CSIDL_BITBUCKET = &HA
' CSIDL_COMMON_ADMINTOOLS = &H2F
' CSIDL_COMMON_ALTSTARTUP = &H1E
' CSIDL_COMMON_APPDATA = &H23
' CSIDL_COMMON_DESKTOPDIRECTORY = &H19
' CSIDL_COMMON_DOCUMENTS = &H2E
' CSIDL_COMMON_FAVORITES = &H1F
' CSIDL_COMMON_PROGRAMS = &H17
' CSIDL_COMMON_STARTMENU = &H16
' CSIDL_COMMON_STARTUP = &H18
' CSIDL_COMMON_TEMPLATES = &H2D
' CSIDL_CONNECTIONS = &H31
' CSIDL_CONTROLS = &H3
' CSIDL_COOKIES = &H21
' CSIDL_DESKTOP = &H0
' CSIDL_DESKTOPDIRECTORY = &H10
' CSIDL_DRIVES = &H11
' CSIDL_FAVORITES = &H6
' CSIDL_FLAG_CREATE = &H8000
' CSIDL_FLAG_DONT_VERIFY = &H4000
' CSIDL_FLAG_MASK = &HFF00&
' CSIDL_FLAG_PFTI_TRACKTARGET = CSIDL_FLAG_DONT_VERIFY

    
    
    
    
    
    
    
    
End Sub
Public Function ResolveShortcut(ByVal StrShortcut As String) As String
    Dim Linker As cShellLink
    Set Linker = New cShellLink
    On Error GoTo ReturnMain
    'if the right is "desklink", just return the desktop folder.
    If StrComp(Right$(StrShortcut, 8), "desklink", vbTextCompare) = 0 Then
        ResolveShortcut = GetSpecialFolder(CSIDL_DESKTOPDIRECTORY).Path
        Exit Function
    
    
    End If
    
    
    Linker.Resolve StrShortcut, SLR_ANY_MATCH, 0

    ResolveShortcut = Linker.Path
    Exit Function
ReturnMain:
    ResolveShortcut = StrShortcut

End Function
'Public Function GetFileIcon(ByVal Filename As String, ByVal IconSize As ShellFileInfoConstants) As Long
'    Dim Finfo As SHFILEINFO
'    SHGetFileInfo Filename, 0, Finfo, Len(Finfo), IconSize
'
'
'    GetFileIcon = Finfo.hIcon
'
'
'
'
'
'
'End Function




Public Function ShowShellMenu(ByVal hwnd As Long, ByVal FilePath As String, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1) As Long
    ShowExplorerMenu hwnd, FilePath
End Function
Public Function ShowExplorerMenuMulti(ByVal hWndOwner As Long, ByVal pszPath As String, StrFiles() As String, Optional X As Long = -1, Optional Y As Long = -1, Optional CallbackObject As IContextCallback = Nothing) As Long

ShowExplorerMenuMulti = MdlFileSystem.ShowExplorerMenuMulti(hWndOwner, pszPath, StrFiles(), X, Y, CallbackObject)


End Function



Public Function FormatSize(ByVal Bytes As Double, Optional ByVal FlagsUse As EFormatSizeFlags = Format_ShowThousandsSeparator) As String
    Static LookupPrefix() As String, LongPrefix() As String, flInit As Boolean
    Static Proper() As String, ProperPrefix() As String
    
    Dim ShowtotalInBrackets As Boolean, LongMetric As Boolean, ShowThousandsSeparator As Boolean, UseNewStandards As Boolean
    ShowtotalInBrackets = (FlagsUse And Format_ShowTotalInBrackets) = Format_ShowTotalInBrackets
    LongMetric = (FlagsUse And Format_longmetric) = Format_longmetric
    ShowThousandsSeparator = (FlagsUse And Format_ShowThousandsSeparator) = Format_ShowThousandsSeparator
    UseNewStandards = (FlagsUse And Format_UseNewStandard) = Format_UseNewStandard
    Dim currval As Double, countdiv As Long, buildstr As String
    'the index into the loopupprefix is the number of times it needed to be divided by 1024 to get below 1024.
    If Not flInit Then
        flInit = True
        ReDim LookupPrefix(0 To 8)
        LookupPrefix(0) = " Bytes"
        LookupPrefix(1) = "KB"
        LookupPrefix(2) = "MB"
        LookupPrefix(3) = "GB"
        LookupPrefix(4) = "TB"
        LookupPrefix(5) = "PB"
        LookupPrefix(6) = "EB"
        LookupPrefix(7) = "ZB"
        LookupPrefix(8) = "YB"
        
        
        ReDim LongPrefix(0 To 8)
        LongPrefix(0) = " Bytes"
        LongPrefix(1) = " Kilobytes"
        LongPrefix(2) = " Megabytes"
        LongPrefix(3) = " Gigabytes"
        LongPrefix(4) = " Terabytes"
        LongPrefix(5) = " Petabytes"
        LongPrefix(6) = " Exabytes"
        LongPrefix(7) = " Zetabytes"
        LongPrefix(8) = " Yottabytes"
        
        ReDim Proper(0 To 8)
        
'        1024    Ki  kibi    K   kilo
'10242   Mi  mebi    M   mega
'10243   Gi  gibi    G   giga
'10244   Ti  tebi    T   tera
'10245   pi pebi
'10246   Ei exbi
'10247   Zi zebi
        '10248   Yi yobi
        Proper(0) = " Bytes"
        Proper(1) = " Kibibytes"
        Proper(2) = " Mebibytes"
        Proper(3) = " Gibibytes"
        Proper(4) = " Tebibytes"
        Proper(5) = " Pebibytes"
        Proper(6) = " Exbibytes"
        Proper(7) = " Zebibytes"
        Proper(8) = " Yobibytes"
        
        ReDim ProperPrefix(0 To 8)
        ProperPrefix(0) = " Bytes"
        ProperPrefix(1) = "Ki"
        ProperPrefix(2) = "Mi"
        ProperPrefix(3) = "Gi"
        ProperPrefix(4) = "Ti"
        ProperPrefix(5) = "pi"
        ProperPrefix(6) = "Ei"
        ProperPrefix(7) = "Zi"
        ProperPrefix(8) = "Yi"
        
        
        
        
    End If
    currval = Bytes
    Do Until currval < 1024
        currval = currval / 1024
        countdiv = countdiv + 1
    Loop
    Dim usearray() As String
    If UseNewStandards Then
        If LongMetric Then
            usearray = Proper
        Else
            usearray = ProperPrefix
        End If
    Else
        If LongMetric Then
            usearray = LongPrefix()
        Else
            usearray = LookupPrefix()
        End If
 
    
    End If
    
    'First, always start with the number.
    buildstr = Trim$(str(Round(currval, 2)))
    buildstr = buildstr & usearray(countdiv)
    
    
'    buildstr = Round(currval, 2) & LookupPrefix(countdiv)
'
    If ShowtotalInBrackets And countdiv > 0 Then
        If ShowThousandsSeparator Then
            buildstr = buildstr & "(" & FormatNumber(Bytes, 0, True, vbFalse, vbTrue) & " Bytes)"
        Else
            buildstr = buildstr & "(" & FormatNumber(Bytes, 0, True, vbFalse, vbFalse) & " Bytes)"
        
        End If
    End If



    FormatSize = buildstr



End Function
'Private Function CreateSparseFile(ByVal Path As String, ByVal SparseSize As Long) As CFile
'Dim hfile As Long
'   hfile = MdlFileSystem.CreateFile(Path, GENERIC_ALL, FILE_SHARE_READ, 0, CREATE_ALWAYS, 0, 0)
'
'   DeviceIoControl hfile, FSC
'
'
'
'
'
'End Function
Public Function GetAttributeString(ByVal ForAttribute As FileAttributeConstants, Optional ByVal Longform As Boolean = False)

'Public Function GetAttributeString(Optional ByVal longform As Boolean = False) As String
GetAttributeString = MdlFileSystem.GetAttributeString(ForAttribute, Longform)



End Function


Public Function GetExtension(ByVal Path As String) As String
    'returns the extension portion of the path.
    Dim gotext As String
    Call ParsePathParts(Path, , , , gotext, , Parse_FExt)
GetExtension = gotext
End Function

Public Function GetFilenamePart(ByVal Path As String) As String
Dim totallen As Long
totallen = Len(MdlFileSystem.GetVolume(Path)) + Len(GetPathPart(Path))
GetFilenamePart = Mid$(Path, totallen)


End Function

Public Function GetPathPart(ByVal Path As String) As String
    Dim vollength As Long, retme As String
    Path = Replace$(Path, "\", "/")
    vollength = Len(MdlFileSystem.GetVolume(Path))
    
    retme = Mid$(Path, vollength, InStrRev(Path, "/") - vollength + 1)
    If vollength = 3 Then
        retme = Replace$(retme, "/", "\")
    End If
    GetPathPart = retme


End Function
Public Function Exists(ByVal Path As String)
  Dim fattr As Long
    
fattr = GetFileAttributes(Path)
If fattr = -1 Then
    Exists = False
Else
    Exists = True

End If



End Function

Public Function Getvolumes() As Volumes
    Static mvolumes As Volumes
    If mvolumes Is Nothing Then Set mvolumes = New Volumes
    Set Getvolumes = mvolumes


End Function
Public Function CreateStream(ByVal PathSpec As String) As FileStream

    Set CreateStream = Me.CreateFile(PathSpec).OpenAsBinaryStream(GENERIC_ALL, FILE_SHARE_DELETE, OPEN_EXISTING)



End Function
Public Function CreateMemoryStream() As CMemStream

    Dim Memstream As CMemStream
    Set Memstream = New CMemStream
    Set CreateMemoryStream = Memstream



End Function
Public Function FileStreamToMemStream(FromStream As FileStream) As CMemStream
    Dim newMem As CMemStream
    'Dim castoutput As IOutputStream
    Set newMem = CreateMemoryStream
    'Set castoutput = newMem
    newMem.WriteStream FromStream
    Set FileStreamToMemStream = newMem




End Function
Public Function OpenStream(ByVal PathSpec As String) As FileStream
    Dim ret As FileStream
    Set ret = Me.GetFile(PathSpec).OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_DELETE, OPEN_EXISTING)
    Set OpenStream = ret
End Function
'Friend Function OpenBackupStream(ByVal PathSpec As String) As CBackupStream
'    Dim ret As CBackupStream
'    Set ret = Me.GetFile(PathSpec).OpenAsBackupStream(GENERIC_READ, FILE_SHARE_DELETE + FILE_SHARE_READ + FILE_SHARE_WRITE, OPEN_EXISTING)
'    Set OpenBackupStream = ret
'End Function
Public Function ParsePathParts(ByVal StrInput As String, Optional ByRef Volume As String, Optional ByRef Path As String, _
Optional ByRef Filename As String, Optional ByRef Extension As String, Optional ByRef StreamName As String, Optional ByVal ParseLevel As ParsePathPartsConstants = Parse_All)

ParsePathParts = MdlFileSystem.ParsePathParts(StrInput, Volume, Path, Filename, Extension, StreamName, ParseLevel)


End Function
Public Function SplitPath(ByVal pathSplit As String) As String()
    SplitPath = MdlFileSystem.SplitPath(pathSplit)
    
End Function
Public Function ParseFilename(ByVal pathparse As String, Optional ByVal includeextension As Boolean = True) As String
'
ParseFilename = MdlFileSystem.ParseFilename(pathparse, includeextension)
End Function
'Public Function ParsePathParts(ByVal StrInput As String, Optional ByRef Volume As String, Optional ByRef Path As String, _
'Optional ByRef Filename As String, Optional ByRef Extension As String, Optional ByRef StreamName As String, Optional ByVal ParseLevel As ParsePathPartsConstants = Parse_All)
'   ParsePathParts = MdlFileSystem.ParsePathParts(StrInput, Volume, Path, Filename, Extension, StreamName, ParseLevel)
'End Function
Public Function ParseExtension(ByVal pathparse As String) As String
    ParseExtension = MdlFileSystem.ParseExtension(pathparse)
End Function

Private Sub Class_Initialize()
    MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
    
    
    
End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub
Public Sub testexe()

Dim X As String
'x = Redirect("Cmd /c dir /w | sort /+16")
Stop



End Sub
'Public Function GetThumbnail(ByVal fileStr As String, Optional DesiredWidth As Long = 100, Optional DesiredHeight As Long = 100) As Long
'    'returns... an hbitmap. sorry people...
'
'
'
'
'
'End Function


'Function to get the SystemImageList Index of a file object.

Public Function SysIMLIndex(ByVal StrFilename As String, Optional ByVal GetOverlay As Boolean = False, Optional ByVal GetOpenIcon As Boolean = False) As Long

    Dim shitem As SHFILEINFO
    
    Dim useflags As Long
    useflags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX
    If GetOverlay Then useflags = useflags Or SHGFI_OVERLAYINDEX
    If GetOpenIcon Then useflags = useflags Or SHGFI_OPENICON
    
    
    SHGetFileInfo ByVal StrFilename, FILE_ATTRIBUTE_NORMAL, shitem, LenB(shitem), useflags

    SysIMLIndex = shitem.iIcon
    

    



End Function


Public Function GetFileTypeFromExtension(ByVal Extension As String) As String

 Dim FirstValue As String, secondvalue As String
FirstValue = mreg.ValueEx(hhkey_classes_root, "." & Trim$(Extension), "", RREG_SZ, "")

If FirstValue <> "" Then
    secondvalue = mreg.ValueEx(hhkey_classes_root, FirstValue, "", RREG_SZ, "")
    If secondvalue <> "" Then
        GetFileTypeFromExtension = secondvalue
    Else
        GetFileTypeFromExtension = FirstValue
    End If
End If





End Function
Public Function isDirectory(ByVal strpath As String) As Boolean

    'Dim hFile As Long, fileattribs As Long
    
    Dim attributeget As FileAttributeConstants
    attributeget = GetFileAttributes(strpath)
    If (attributeget And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        isDirectory = True
    End If
    
    
    
    'Old code follows:
    
    
    
'    hFile = MdlFileSystem.CreateFile(strpath, 0, FILE_SHARE_READ + FILE_SHARE_DELETE + FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
'    If hFile = -1 Then
'        'could be a drive...
'        If GetVolume(strpath).RootFolder.Path = strpath Then
'            isDirectory = True
'            Exit Function
'        Else
'        isDirectory = False
'        End If
'    Else
'        fileattribs = GetFileAttributes(hFile)
'        If (fileattribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'            isDirectory = True
'        Else
'            isDirectory = False
'        End If
'        CloseHandle hFile
'    End If
    


End Function
Public Function GetEXEFriendlyName(strpath As String) As String
GetEXEFriendlyName = modFileVersion.GetEXEFriendlyName(strpath)

End Function
Public Sub GetBCFileVersion(ByRef MajorVersion As Long, ByRef MinorVersion As Long, ByRef Revision As Long, ByRef DebugBuild As Boolean)

#If NODEBUGLOG Then
    DebugBuild = False
#Else
    DebugBuild = True
#End If
MajorVersion = App.Major
MinorVersion = App.Minor
Revision = App.Revision



End Sub
Public Function GetStandardStream(ByVal stdStream As StandardIOStreamsEnum) As FileStream


'err... yeah I know it says it's a "FileStream"... but... err... yeah......
Dim hHandle As Long, readmodesel As DesiredAccessFlags


hHandle = GetStdHandle(stdStream)

If stdStream = standard_Error Then readmodesel = GENERIC_WRITE
If stdStream = standard_Output Then readmodesel = GENERIC_WRITE
If stdStream = Standard_Input Then readmodesel = GENERIC_READ

Dim retStream As FileStream
Set retStream = New FileStream
retStream.SetStream hHandle, readmodesel



Set GetStandardStream = retStream


End Function
Public Function GetFileVerInfo(ByVal Filename As String, ByRef CompanyName As String, ByRef Filedescription As String, _
    ByRef FileVersion As String, ByRef InternalName As String, ByRef LegalCopyright As String, ByRef OriginalFilename As String, ByRef ProductName As String, ByRef ProductVersion As String)


GetFileVerInfoDirect Filename, CompanyName, Filedescription, FileVersion, InternalName, LegalCopyright, OriginalFilename, ProductName, ProductVersion
End Function
Public Function GetFileVersion(ByVal Filename As String) As String
    Dim ret As String
    GetFileVerInfoDirect Filename, , , ret
    GetFileVersion = ret
End Function
Public Sub testlocale()
    Dim X As CLocaleInfo
    Set X = New CLocaleInfo
    X.test
End Sub
Public Function StrToMemStream(ByVal StringUse As String, Optional Stringmode As StringReadMode = StrRead_unicode) As CMemStream
    Dim newstream As CMemStream
    Set newstream = New CMemStream
    newstream.WriteString StringUse, Stringmode
    newstream.SeekTo 0, STREAM_BEGIN
    Set StrToMemStream = newstream

End Function
Public Sub TestUnicode(ByVal StrCreate As String)
Dim newstream As FileStream
Set newstream = CreateFile(StrCreate, True).OpenAsBinaryStream(GENERIC_ALL, FILE_SHARE_READ + FILE_SHARE_WRITE, OPEN_EXISTING)
newstream.WriteString "This is Unicode text.", StrRead_unicode
newstream.CloseStream
Set newstream = Nothing




End Sub
Public Function EncryptBytes(bEncrypt() As Byte, ByVal password As String, Optional ByVal mhashAlgorithm As HashAlgorithmEnum = HashAlgorithmEnum.MD5, _
                    Optional ByVal mEncAlgorithm As EncAlgorithmEnum = EncAlgorithmEnum.RC4, Optional ByVal flReverse As Boolean = False) As Byte()
        
        Dim TempInStream As CMemStream
        Dim tempoutstream As CMemStream
        Dim usefilter As IStreamFilter, ccore As CCoreFilters
        'create the temporary stream out of the string to encrypt...
        Set TempInStream = New CMemStream
        TempInStream.WriteBytes bEncrypt()
        TempInStream.SeekTo 0, STREAM_BEGIN
        Set tempoutstream = New CMemStream
        
        'now... encrypt...
        Set ccore = New CCoreFilters
        ccore.encAlgorithm = mEncAlgorithm
        ccore.hashAlgorithm = mhashAlgorithm
        ccore.password = password
        Set usefilter = ccore
        usefilter.FilterStream TempInStream, tempoutstream, flReverse
        tempoutstream.SeekTo 0, STREAM_BEGIN
        
        EncryptBytes = tempoutstream.readbytes(tempoutstream.Size)
                    
                    
                    
End Function


'Public Sub TestCTask(OwnerWnd As Long)
'    Dim testdialog As CTaskDialog
'    Set testdialog = New CTaskDialog
'    testdialog.AddButton IDYES, "Why not?", False, True
'    testdialog.AddButton IDNO, "No thanks.", False, False
'    testdialog.AddButton IDABORT, "ABORT!", True
'    testdialog.AddButton IDIGNORE, "IGNORE!", True
'    testdialog.WindowTitle = "This is a window title."
'    testdialog.Content = "This is content."
'    testdialog.CollapsedControlText = "CollapsedControlText"
'    testdialog.Expanded = "this is expanded."
'    testdialog.ExpandedControlText = "This is expandedControl Text."
'    testdialog.MainInstruction = "This is the main instruction."
'    testdialog.VerificationText = "This is verification text."
'    testdialog.Verify = "This is verify"
'    testdialog.Footer = "this is a footer"
'    testdialog.MainIcon = testdialog.MAKEINTRESOURCE(TD_SHIELD_ICON)
'    testdialog.FooterIcon = testdialog.MAKEINTRESOURCE(TD_INFORMATION_ICON)
'    testdialog.Show OwnerWnd
'End Sub
 Public Function GetAllFileAttributes() As FileAttributeConstants()
    GetAllFileAttributes = MdlFileSystem.GetAllFileAttributes()
 End Function

Public Sub SplitStream(StreamSplit As IInputStream, pChunkSize As Long, OutStreams() As IInputStream)
'SplitStream: splits the given input stream into "pchunksize" pieces, returning an array of IInputStreams() that can be used to read each split piece to a file or something.

'internally:
'use a set of temporary files and return those in the array.

Dim numStreams As Long

Dim I As Long, usestreamtemp As FileStream
Dim tempreadbytes() As Byte, sizeread As Long
numStreams = StreamSplit.Size / pChunkSize

numStreams = numStreams + 1 'last one will only read to eof anyway.

ReDim OutStreams(0 To numStreams - 1)

For I = 0 To UBound(OutStreams)
    Set usestreamtemp = FileSystem.OpenTempFile()
    tempreadbytes = StreamSplit.readbytes(pChunkSize, sizeread)
    usestreamtemp.WriteBytes tempreadbytes
    usestreamtemp.SeekTo 0, STREAM_BEGIN
    Set OutStreams(I) = usestreamtemp
    
    
Next I



End Sub

Public Function GetFileVerbs(ByVal FilePath As String) As Variant
' Returns string array of verb names
Dim sh As Object, folder1 As Object, file1 As Object, verbcol As Object, I As Integer
Set sh = CreateObject("Shell.Application")

I = InStrRev(FilePath, "\")
Set folder1 = sh.NameSpace(Left(FilePath, I - 1))
Set file1 = folder1.ParseName(Mid(FilePath, I + 1))
Set verbcol = file1.Verbs
Dim verbStrings() As Variant, v As Object
ReDim verbStrings(verbcol.count - 1)
For I = 0 To verbcol.count - 1
    Set v = verbcol.Item(I + 1)
    verbStrings(I) = v.Name
Next
GetFileVerbs = verbStrings

End Function
Public Function GetRefCount(ByRef obj As IUnknown) As Long
    If obj Is Nothing Then Exit Function
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    GetRefCount = GetRefCount - 3
End Function
