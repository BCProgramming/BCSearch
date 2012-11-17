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
    strCommand As String

End Type
Private WithEvents Downloader As CDownload
Attribute Downloader.VB_VarHelpID = -1
Private mDownErr As Long, mDownDesc As String, mdowncomplete As Boolean
Public Property Get g_ForceANSI() As Long
    g_ForceANSI = ForceANSI
End Property
Public Property Let g_ForceANSI(ByVal vdata As Long)
    ForceANSI = vdata
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



Public Function DownloadRemoteFile(ByVal URL As String) As CFile
    'downloads the file; saves to temp and returns the file.

    Dim tempFile As String
    tempFile = GetTempFileNameAndPathEx()
    'ftp is port 21.
    Set Downloader = New CDownload
    With Downloader
    Downloader.URL = URL
    mdowncomplete = False
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
Dim lcount As Long



ret = GetOpenWithList("*", lcount)

 
'Stop
End Sub
Private Sub MergeOpenwith(ArrayMergewith() As OpenWithListItem, arraymerge() As OpenWithListItem)
    Dim I As Long, NewUB As Long
    'On Error GoTo Breakout
    On Error Resume Next
    NewUB = UBound(ArrayMergewith)
        If Err <> 0 Then
            ReDim ArrayMergewith(1 To 1)
        End If
        
    NewUB = NewUB + UBound(arraymerge)
     
     If Err <> 0 Then
        Exit Sub
    End If
    ReDim Preserve ArrayMergewith(1 To UBound(ArrayMergewith) + UBound(arraymerge))
    For I = LBound(arraymerge) To UBound(arraymerge)
        ArrayMergewith(UBound(ArrayMergewith) + I) = arraymerge(I)
        
    
    Next I
            

Exit Sub
breakout:
Exit Sub
notinitialized:
    ReDim ArrayMergewith(1 To 1)
    Resume
    
End Sub
Public Function GetOpenWithList(ByVal ForExtension As String, ByRef ListCount As Long) As OpenWithListItem()

    Dim creg As cRegistry
'find open with commands/names for a extension.

'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\Extension
'For here, there will be a key,
Dim tmpreturn() As OpenWithListItem
Dim retArray() As OpenWithListItem
Dim baseExplorer As String
Dim values() As String, vcount As Long
Dim sect() As String, scount As Long
Dim I As Long, retcount As Long, mruList As String
Dim AppExes() As String, acount As Long
Dim lcount As Long
Dim CurrPos As Long, CurrChar As String

baseExplorer = "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & ForExtension
Set creg = New cRegistry
creg.Classkey = HHKEY_CURRENT_USER
creg.SectionKey = baseExplorer & "\OpenWithprogIDs"


'OpenWithprogIDs. Enumerate the values, and recursively call this routine, passing the value name as the Extension.
If creg.EnumerateValues(values(), vcount) Then
    For I = 1 To vcount
        retcount = 0
        tmpreturn = GetOpenWithList(values(I), retcount)
        If retcount > 0 Then
            MergeOpenwith retArray(), tmpreturn()
            
        
        End If

    Next I
End If


'OpenWithList is a bit less clearcut, first, open "MRUList" value from that key.
'Enumerate each Character, loading the value in the value with the name of the character (for example, "a")
'each entry is an application Executable, which we later add to our array- just add it to the appexes() array, for now.
creg.SectionKey = baseExplorer & "\OpenWithList"
    mruList = creg.ValueEx(HHKEY_CURRENT_USER, baseExplorer & "\OpenWithList", "MRUList", RREG_SZ, "")
    
    'iterate through each character...
    For CurrPos = 1 To Len(mruList)
        CurrChar = Mid$(mruList, CurrPos, 1)
        
        acount = acount + 1
        ReDim Preserve AppExes(1 To acount)
        AppExes(acount) = creg.ValueEx(HHKEY_CURRENT_USER, baseExplorer & "\OpenWithList", CurrChar, RREG_SZ, "")
        
        
        
    Next CurrPos
    
    'whew, got the Application executables- for this one.
    



'HKEY_CLASSES_ROOT\Extension

'OpenWithProgIds is a group of values here, each one is the name of the class- again- recursively call this same routine and
'merge the results into our result array.

creg.Classkey = HHKEY_CLASSES_ROOT
creg.SectionKey = ForExtension & "\OpenWithProgids"
'HKEY_CLASSES_ROOT\.txt\OpenWithProgids
If creg.EnumerateValues(sect(), scount) Then
    On Error Resume Next
    For I = 1 To UBound(sect)
        If Err <> 0 Then
            Err.Clear
            Exit For
        End If
        tmpreturn = GetOpenWithList(sect(I), lcount)
        If lcount > 0 Then
            MergeOpenwith retArray(), tmpreturn()
        End If
    Next I


End If
'whew....

'"Openwithlist" has the same format as that in the HKCU key described above.

'**wrong**
'within HKEY_CLASSES_ROOT, it will be sections as well.
    creg.Classkey = HHKEY_CLASSES_ROOT
    creg.SectionKey = ForExtension & "\OpenWithList"
    If creg.EnumerateSections(sect, scount) Then
        For I = 1 To scount
            acount = acount + 1
            ReDim Preserve AppExes(1 To acount)
            AppExes(I) = sect(I)
            
        
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
defvalue = creg.ValueEx(HHKEY_CLASSES_ROOT, ForExtension, "", RREG_SZ, "")
'change: also check "old Default" value.
If defvalue <> "" Then
    creg.SectionKey = defvalue
    If creg.KeyExists Then
    tmpreturn = GetOpenWithList(defvalue, lcount)
        If lcount > 0 Then
            MergeOpenwith retArray, tmpreturn
        End If
    End If
End If


'iterate through appexes, retrieve executable names from:
'"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"

'use GetEXEFriendlyName(scmd) for friendly name.
On Error Resume Next
For I = 1 To UBound(AppExes)
    If Err <> 0 Then Exit For
    defvalue = creg.ValueEx(HHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & AppExes(I), "", RREG_SZ, "")
    
    If defvalue = "" Then
        Dim bufferpath As String
        bufferpath = Space$(255)
        Call SearchPath(0, AppExes(I), ".exe", Len(bufferpath), bufferpath, "")
        'use searchpath()
        bufferpath = Left$(bufferpath, InStr(bufferpath, vbNullChar) - 1)
        defvalue = bufferpath
    End If
    defvalue = Trim$(defvalue)
    If defvalue <> "" Then
        'it is the full path-
        On Error Resume Next
        ReDim Preserve retArray(1 To UBound(retArray) + 1)
        If Err <> 0 Then ReDim Preserve retArray(1 To 1)
        With retArray(UBound(retArray))
            .strCommand = defvalue
            .strCommand = Replace$(.strCommand, vbNullChar, "")
            .strCommand = Replace$(.strCommand, """", "")
            .StrName = GetEXEFriendlyName(.strCommand)
            .StrName = Replace$(.StrName, vbNullChar, "")
            If .StrName = "" Then Stop
        
        End With
        
    
    End If
    
    
Err.Clear
Next I
On Error Resume Next
ListCount = UBound(retArray)
GetOpenWithList = retArray
End Function

Private Function GetOpenWithListInt(ByVal ForExtension As String, Optional ByRef ListCount As Long, Optional ByVal useclasskey As ERegistryClassConstants = HHKEY_CLASSES_ROOT, Optional ByVal SectionKey As String = "") As OpenWithListItem()
'TODO//improve this routine.
 
'improvements:
'find the "other" locations of items that seem to be able to add themselves to open with without any sort of key (???)
'damn it- we need to do the SAME thing with the
'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts
'as we do currently in HKEY_CLASSES_ROOT.

    Dim retArr() As OpenWithListItem, retcount As Long
    Dim tmpreturn() As OpenWithListItem, I As Long
    Dim reguse As cRegistry, lcount As Long
    Dim AppExes() As String, currcopy As Long
    Dim sect() As String, scount As Long, vals() As String, vcount As Long
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

    
         
            reguse.EnumerateSections sect, scount
            If useclasskey <> HHKEY_CLASSES_ROOT Then
                reguse.EnumerateValues vals, vcount
                ReDim Preserve sect(1 To UBound(sect) + vcount)
                For I = 1 To vcount
                    sect(scount + vcount) = vals(vcount)
                    
                
                Next I
                    
                vcount = 0
            End If
        'Add to our current appExes() array
        'pah- just direct copy it.
        AppExes = sect
        
        
       
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
                tmpreturn = GetOpenWithListInt(vals(I), lcount, useclasskey, basesection)
                       'for each value
                'open that key in HKEY_CLASSES_ROOT- recursively by calling this same function with the value as the parameter.
                'append to our array the results from said function.
                'append this array to our retarr() array.
                For currcopy = 1 To lcount
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
    tmpreturn = GetOpenWithListInt(ForExtension, lcount, HHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts")
    Debug.Print "explorer sees " & lcount & " entries." & " for " & ForExtension
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
    For I = LBound(AppExes) To UBound(AppExes)
        If Err = 0 Then
            Dim scmd As String
            scmd = reguse.ValueEx(HHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & AppExes(I), "", RREG_SZ, "")
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
            retArr(retcount).strCommand = scmd
        
    
        End If
    Next I
    
    GetOpenWithListInt = retArr
    ListCount = retcount
    Debug.Print ListCount
End Function
Friend Function PerformFileOperation(ByVal StrFrom As String, StrTo As String, ByVal Operation As FILEOP, ByVal flags As FILEOP_FLAGS, _
    Optional ByVal HwndOwner As Long, Optional ByVal Silent As Boolean, Optional ByVal ProgressMessage As String = "") As Long
    
    Dim OpStruct As SHFILEOPSTRUCT, ret As Long
    'strFrom and StrTo should be properly Null delimited...
    With OpStruct
        .hWnd = HwndOwner
        .pFrom = StrFrom
        .pTo = StrTo
        .sProgress = ProgressMessage
        .fFlags = flags
        .wFunc = Operation
    End With
    ret = SHFileOperation(OpStruct)
    
    
    
    
    
    
    
    
    
    
    End Function
    

Public Function CreateDirectory(ByVal PathSpec As String, Optional ByVal templateDir As Directory = Nothing, Optional ByVal CreateParentFolders As Boolean = False) As Directory
Attribute CreateDirectory.VB_Description = "Creates the specified Directory path and returns it."
    'Call CreateDirectory Routine...
    Dim retVal As Long, hDir As Long
    If templateDir Is Nothing Then
    'Easy.
    
    'first- if we are to create parent folders described then do so.
    
    'retval = mdlfilesystem.CreateDirectoryEx(
    retVal = MdlFileSystem.CreateDirectory(PathSpec, 0)
    If retVal = 0 Then
        RaiseAPIError Err.LastDllError, "BCFSObject::CreateDirectory"
    Else
        Set CreateDirectory = GetDirectory(PathSpec)
    End If

    Else
      
            retVal = CreateDirectoryEx(templateDir.Path, PathSpec, 0)
            If retVal = 0 Then
                RaiseAPIError Err.LastDllError, "BCFSObject::CreateDirectory"
            Else
            'success.
            Set CreateDirectory = GetDirectory(PathSpec)
            
            End If
        
        
            CloseHandle hDir
    End If
    
    
End Function
Public Function CreateFile(ByVal PathSpec As String) As CFile
Attribute CreateFile.VB_Description = "Creates the specified file and returns a CFile object that represents it."
    'create the file, and return a reference to it as well.
    
    Dim ffile As CFile
    
    
    
    PathSpec = FixPath(PathSpec)
    Set ffile = New CFile
    ffile.Create PathSpec, False
    Set CreateFile = ffile

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

Public Function GetVolume(ByVal PathSpec As String) As CVolume
    Dim newvol As CVolume
    Set newvol = New CVolume
    newvol.Init PathSpec
    Set GetVolume = newvol


End Function

Public Function GetFile(ByVal PathSpec As String) As CFile

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
        Err.Raise 53, "BCFSObject::GetFile", "File Not Found, """ & PathSpec & """"
    
    End If



End Function
Public Function GetObjIcon(Objstr As String, iconsize As IconSizeConstants) As Long

    GetObjIcon = GetFileIcon(Objstr, iconsize)


End Function
Public Function GetDirectory(ByVal Path As String) As Directory
    Dim NewDir As Directory
    Set NewDir = New Directory
    
    On Error GoTo ReturnNothing
    
    NewDir.SetDir Path
    Set GetDirectory = NewDir
    Exit Function
    
ReturnNothing:
    Set GetDirectory = Nothing
    
    
End Function
Public Function GetSpecialFolder(ByVal FolderConst As SpecialFolderConstants) As Directory
    Dim Spath As String
    Spath = Space$(256)
    Call SHGetSpecialFolderPath(0, Spath, FolderConst, BOOL.APITRUE)
    
    Spath = Trim$(Replace$(Spath, vbNullChar, " "))
    Set GetSpecialFolder = GetDirectory(Spath)



End Function
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

Public Function ShowShellMenu(ByVal hWnd As Long, ByVal FilePath As String, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1) As Long
    ShowExplorerMenu hWnd, FilePath
End Function
Public Function ShowExplorerMenuMulti(ByVal HwndOwner As Long, ByVal pszPath As String, StrFiles() As String, Optional x As Long = -1, Optional y As Long = -1, Optional CallbackObject As IContextCallback = Nothing) As Long

ShowExplorerMenuMulti = MdlFileSystem.ShowExplorerMenuMulti(HwndOwner, pszPath, StrFiles(), x, y, CallbackObject)


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
        ProperPrefix(2) = "Gi"
        ProperPrefix(3) = "Ti"
        ProperPrefix(4) = "pi"
        ProperPrefix(5) = "Ei"
        ProperPrefix(6) = "Zi"
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
Public Function GetAttributeString(ByVal ForAttribute As FileAttributeConstants, Optional ByVal Shortform As Boolean = True)
    Static AttributeValues() As String
    Static AttributeLongNames() As String
    Static AttributeShortNames() As String
    Static FlInitialized As Boolean
    Dim I As Long
    If Not FlInitialized Then
        ReDim AttributeValues(13)
        ReDim AttributeLongNames(13)
        ReDim AttributeShortNames(13)
    
        AttributeValues(0) = FILE_ATTRIBUTE_ARCHIVE
        AttributeLongNames(0) = "FILE_ATTRIBUTE_ARCHIVE"
        AttributeShortNames(0) = "A"
        
        AttributeValues(1) = FILE_ATTRIBUTE_COMPRESSED
        AttributeLongNames(1) = "FILE_ATTRIBUTE_COMPRESSED"
        AttributeShortNames(1) = "C"
        
        AttributeValues(2) = FILE_ATTRIBUTE_DEVICE
        AttributeLongNames(2) = "FILE_ATTRIBUTE_DEVICE"
        AttributeShortNames(2) = "D"
        
        AttributeValues(3) = FILE_ATTRIBUTE_DIRECTORY
        AttributeLongNames(3) = "FILE_ATTRIBUTE_DIRECTORY"
        AttributeShortNames(3) = "DIR"
        
        AttributeValues(4) = FILE_ATTRIBUTE_ENCRYPTED
        AttributeLongNames(4) = "FILE_ATTRIBUTE_ENCRYPTED"
        AttributeShortNames(4) = "E"
        
        AttributeValues(5) = FILE_ATTRIBUTE_HIDDEN
        AttributeLongNames(5) = "FILE_ATTRIBUTE_HIDDEN"
        AttributeShortNames(5) = "H"
        
        AttributeValues(6) = FILE_ATTRIBUTE_NORMAL
        AttributeLongNames(6) = "FILE_ATTRIBUTE_NORMAL"
        AttributeShortNames(6) = "N"
        
        AttributeValues(7) = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
        AttributeLongNames(7) = "FILE_ATTRIBUTE_NOT_CONTENT_INDEXED"
        
        AttributeValues(8) = FILE_ATTRIBUTE_OFFLINE
        AttributeLongNames(8) = "FILE_ATTRIBUTE_OFFLINE"
        AttributeShortNames(8) = "OFF"
        
        
        AttributeValues(9) = FILE_ATTRIBUTE_READONLY
        AttributeLongNames(9) = "FILE_ATTRIBUTE_READONLY"
        AttributeShortNames(9) = "R"

        AttributeValues(10) = FILE_ATTRIBUTE_REPARSE_POINT
        AttributeLongNames(10) = "FILE_ATTRIBUTE_REPARSE_POINT"
        AttributeShortNames(10) = "REPARSE"

        AttributeValues(11) = FILE_ATTRIBUTE_SPARSE_FILE
        AttributeLongNames(11) = "FILE_ATTRIBUTE_SPARSE_FILE"
        AttributeShortNames(11) = "SPARSE"

        AttributeValues(12) = FILE_ATTRIBUTE_SYSTEM
        AttributeLongNames(12) = "FILE_ATTRIBUTE_SYSTEM"
        AttributeShortNames(12) = "S"

        AttributeValues(13) = FILE_ATTRIBUTE_TEMPORARY
        AttributeLongNames(13) = "FILE_ATTRIBUTE_TEMPORARY"
        AttributeShortNames(13) = "T"
        

   
    
    
    End If
    Dim StrBuild As String
    For I = 0 To UBound(AttributeValues)
        If (ForAttribute And AttributeValues(I)) = AttributeValues(I) Then
        
            If Shortform Then
            StrBuild = StrBuild & AttributeShortNames(I)
            
            Else
            StrBuild = StrBuild & AttributeLongNames(I)
            
            End If
            
            StrBuild = StrBuild & ","
        End If
    
    Next I

   If StrBuild <> "" Then StrBuild = Mid$(StrBuild, 1, Len(StrBuild) - 1)


    GetAttributeString = StrBuild



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

Public Function ParseFilename(ByVal pathparse As String, Optional ByVal includeextension As Boolean = True) As String
'
ParseFilename = MdlFileSystem.ParseFilename(pathparse, includeextension)
End Function
Public Function ParseExtension(ByVal pathparse As String) As String
    ParseExtension = MdlFileSystem.ParseExtension(pathparse)
End Function

Private Sub Class_Initialize()
    mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub
Public Sub testexe()

Dim x As String
x = Redirect("Cmd /c dir /w | sort /+16")
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
FirstValue = mreg.ValueEx(HHKEY_CLASSES_ROOT, "." & Trim$(Extension), "", RREG_SZ, "")

If FirstValue <> "" Then
    secondvalue = mreg.ValueEx(HHKEY_CLASSES_ROOT, FirstValue, "", RREG_SZ, "")
    If secondvalue <> "" Then
        GetFileTypeFromExtension = secondvalue
    Else
        GetFileTypeFromExtension = FirstValue
    End If
End If





End Function
Public Function isDirectory(ByVal StrPath As String) As Boolean

    Dim hfile As Long, fileattribs As Long
    'open the file....
    hfile = MdlFileSystem.CreateFile(StrPath, 0, FILE_SHARE_READ + FILE_SHARE_DELETE + FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If hfile = -1 Then
        isDirectory = False
    Else
        fileattribs = GetFileAttributes(hfile)
        If (fileattribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            isDirectory = True
        Else
            isDirectory = False
        End If
        CloseHandle hfile
    End If
    


End Function
Public Function GetEXEFriendlyName(StrPath As String) As String
GetEXEFriendlyName = modFileVersion.GetEXEFriendlyName(StrPath)

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
