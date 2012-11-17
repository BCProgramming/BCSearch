Attribute VB_Name = "MdlFileSystem"
Option Explicit
Public FileSystem As New BCFSObject
'Public Type FILETIME
'    dwLowDateTime As Long
'    dwHighDateTime As Long
'End Type
'Public Type PointAPI
'    x As Long
'    y As Long
'End Type
Public mreg As New cRegistry
Public Declare Function DestroyIcon Lib "User32.dll" (ByVal hIcon As Long) As Long

Public Enum ParsePathPartsConstants
    Parse_Volume
    Parse_Path
    Parse_FName
    Parse_FExt
    Parse_Stream
    Parse_All
End Enum



Public Type BCCOPYFILEDATA
    BCCallback As IProgressCallback     'the callback.
    SourceFile As String        'used- the handles given in the callback don't have enough info....
    DestinationFile As String

End Type
Private Type WIN32_STREAM_ID
    dwStreamID As Long
    dwStreamAttributes As Long
    dwStreamSizeLow As Long
    dwStreamSizeHigh As Long
    dwStreamNameSize As Long
    'cStreamName As Byte
    'cStreamName() will IMMEDIATELY follow after reading this structure, then the stream data- which we should seek through, I suppose.
    
End Type
'FindStreamData.... For Windows Vista/Server 2003 Stream Enumeration functions.

'typedef struct _WIN32_FIND_STREAM_DATA {
'  LARGE_INTEGER StreamSize;
'  WCHAR         cStreamName[MAX_PATH + 36];
'}WIN32_FIND_STREAM_DATA, *PWIN32_FIND_STREAM_DATA;

Public Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type


Public Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    sAcl As ACL
    dacl As ACL
End Type


Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

'Vista/Server 2003 Stream Enumeration...
'HANDLE WINAPI FindFirstStreamW(
'  __in        LPCWSTR lpFileName,
'  __in        STREAM_INFO_LEVELS InfoLevel,
'  __out       LPVOID lpFindStreamData,
'  __reserved  DWORD dwFlags
');
'infolevel will be zero, for now. no other valid enumerations.
'Private Type LARGE_INTEGER
'    LoPart As Long
'    HiPart As Long
'End Type
Private Type WIN32_FIND_STREAM_DATA
    StreamSize As LARGE_INTEGER
    cStreamName As String * 296
End Type


Private Const ERROR_HANDLE_EOF As Long = 38&

Private Declare Function FindFirstStreamW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal Infolevel As Long, lpFindStreamData As WIN32_FIND_STREAM_DATA, ByVal dwflags As Long) As Long
'BOOL WINAPI FindNextStreamW(
'  __in   HANDLE hFindStream,
'  __out  LPVOID lpFindStreamData
');
Private Declare Function FindNextStreamW Lib "kernel32.dll" (ByVal hFindStream As Long, lpFindStreamData As WIN32_FIND_STREAM_DATA) As Long

'HANDLE WINAPI FindFirstFileNameW(
'  __in     LPCWSTR lpFileName,
'  __in     DWORD dwFlags,
'  __inout  LPDWORD StringLength,
'  __inout  PWCHAR LinkName
');

Private Declare Function FindFirstFileNameW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwflags As Long, ByVal StrLen As Long, ByVal LinkName As Long) As Long


'Public Declare Function GetFileAttributesEx Lib "kernel32.dll" Alias "GetFileAttributesExA" (ByVal lpFileName As String, ByVal fInfoLevelId As Struct_MembersOf_GET_FILEEX_INFO_LEVELS, ByRef lpFileInformation As Any) As Long
Private Declare Function CreateDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String) As Long
Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long


Private Declare Function CreateDirectoryW Lib "kernel32.dll" (ByVal lpPathName As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function RemoveDirectoryW Lib "kernel32.dll" (ByVal lpPathName As Long) As Long


Public Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Public Declare Function GetFileType Lib "kernel32.dll" (ByVal hfile As Long) As Long

Public Declare Function GetFileAttributesA Lib "kernel32.dll" (ByVal lpFileName As String) As Long
Public Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long

Private Declare Function BackupRead Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Any) As Long
Private Declare Function BackupWrite Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Private Declare Function BackupSeek Lib "kernel32.dll" (ByVal hfile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef lpContext As Any) As Long

Public Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


' These WIDE/ANSI version are private. Their wrapper is made public.
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'
'Public Declare Function CreateDirectoryA Lib "kernel32" (ByVal lpPathName As String, lpSecurityAttributes As Any) As Long
'Public Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
'
'Private Declare Function GetFileAttributesA Lib "kernel32" (ByVal lpFileName As String) As Long
'Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
'
'Private Declare Function SetFileAttributesA Lib "kernel32" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
'
'Private Declare Function MoveFileA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
'Private Declare Function MoveFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
'
'Private Declare Function MoveFileExA Lib "kernel32" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
'Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
'
'Private Declare Function DeleteFileA Lib "kernel32" (ByVal lpFileName As String) As Long
'Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
'
'Private Declare Function CreateDirectoryExA Lib "kernel32" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, lpSECURITY_ATTRIBUTES As Any) As Long
'Private Declare Function CreateDirectoryExW Lib "kernel32" (ByVal lpTemplateDirectory As Long, ByVal lpNewDirectory As Long, lpSECURITY_ATTRIBUTES As Any) As Long
'
'Private Declare Function RemoveDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long
'Private Declare Function RemoveDirectoryW Lib "kernel32" (ByVal lpPathName As Long) As Long
'
'Private Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, ByVal lpWIN32_FIND_DATA As Any) As Long
'Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'
'Private Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpWIN32_FIND_DATA As Any) As Long
'
'Private Declare Function GetTempFileNameW Lib "kernel32" (ByVal lpszPath As Long, ByVal lpPrefixString As Long, ByVal wUnique As Long, ByVal lpTempFileName As Long) As Long
'Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'
'Private Declare Function GetTempPathW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
'Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'
'Private Declare Function WNetGetConnectionA Lib "mpr.dll" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
'Private Declare Function WNetGetConnectionW Lib "mpr.dll" (ByVal lpszLocalName As Long, ByVal lpszRemoteName As Long, cbRemoteName As Long) As Long
'
'Private Declare Function GetVolumeInformationA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function GetVolumeInformationW Lib "kernel32" (ByVal lpRootPathName As Long, ByVal lpVolumeNameBuffer As Long, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As Long, ByVal nFileSystemNameSize As Long) As Long
'
'Private Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetLogicalDriveStringsW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long

Private Declare Function SHGetSpecialFolderPathA Lib "shell32.dll" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function SHGetSpecialFolderPathW Lib "shell32.dll" (ByVal hWnd As Long, ByVal pszPath As Long, ByVal csidl As Long, ByVal fCreate As Long) As Long

Private Declare Function SetFileAttributesA Lib "kernel32.dll" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long

Private Declare Function CreateFileMappingA Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function CreateFileMappingALong Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function CreateFileMappingW Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpstrname As Long) As Long

Public Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByRef lpBaseAddress As Any) As Long



'The following require API wrappers:
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetMappedFileName Lib "psapi.dll" Alias "GetMappedFileNameA" (ByVal hProcess As Long, ByRef lpv As Any, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long



Public Declare Function GetCursorPos Lib "user32" (Point As POINTAPI) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileSizeHigh As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long


'Public Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long
Private Const PROGRESS_CONTINUE As Long = 0
Private Const PROGRESS_CANCEL As Long = 1

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Const MAX_PATH = 255

Public Type BCF_WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Type WIN32_FIND_DATAW
 
  dwFileAttributes        As Long
  ftCreationTime          As FILETIME
  ftLastAccessTime        As FILETIME
  ftLastWriteTime         As FILETIME

  nFileSizeHigh           As Long
  nFileSizeLow            As Long
  dwReserved0             As Long
  dwReserved1             As Long
  buffer(1 To 240) As Byte
End Type
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFileA Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindFirstFileW Lib "kernel32.dll" (ByVal lpFileName As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFileA Lib "kernel32.dll" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByRef lpWIN32_FIND_DATA As WIN32_FIND_DATA) As Long
'
Public Const ERROR_NO_MORE_FILES As Long = 18&
Private Declare Function GetDesktopWindow Lib "User32.dll" () As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwflags As Long, ByVal pszPath As String) As Long
Private Declare Sub IIDFromString Lib "ole32.dll" (ByVal lpsz As String, ByVal lpiid As Long)

Public CDebug As New CDebug

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As Guid)
Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Public Declare Function SetErrorMode Lib "kernel32.dll" (ByVal wMode As Long) As Long
'Public Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByRef lpProgressRoutine As Long, ByRef lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long


Declare Function CopyFileEx Lib "kernel32.dll" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long

'Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Declare Function GetFileInformationByHandle Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long

'Public Type SYSTEMTIME
'    wYear As Integer
'    wMonth As Integer
'    wDayOfWeek As Integer
'    wDay As Integer
'    wHour As Integer
'    wMinute As Integer
'    wSecond As Integer
'    wMilliseconds As Integer
'End Type

Public Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Long
'Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long


Public Type OVERLAPPED
    internal As Long
    internalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Public Declare Function ReadFileEx Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpOverlapped As Long, ByVal lpCompletionRoutine As Long) As Long
Public Declare Function WriteFileEx Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hfile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32.dll" (ByVal hfile As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

'Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long


'Private Const MAX_PATH As Long = 260

Public Type BCF_SHFILEINFO
    hIcon As Long ' : icon
    iIcon As Long ' : icondex
    dwAttributes As Long ' : SFGAO_ flags
    szDisplayName As String * MAX_PATH ' : display name (or path)
    szTypeName As String * 80 ' : type name
End Type

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    sProgress As String
End Type

Private Type ULARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

'Public Type SHFILEINFO
'    hIcon As Long ' : icon
'    iIcon As Long ' : icondex
'    dwAttributes As Long ' : SFGAO_ flags
'    szDisplayName As String * MAX_PATH ' : display name (or path)
'    szTypeName As String * 80 ' : type name
'End Type
'Public Type SHELLEXECUTEINFO
'    cbSize As Long
'    fMask As Long
'    hWnd As Long
'    lpVerb As String
'    lpFile As String
'    lpParameters As String
'    lpDirectory As String
'    nShow As Long
'    hInstApp As Long
'    ' fields
'    lpIDList As Long
'    lpClass As String
'    hkeyClass As Long
'    dwHotKey As Long
'    hIcon As Long
'    hProcess As Long
'End Type


'Public Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Public Declare Sub SHEmptyRecycleBinA Lib "shell32.dll" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwflags As Long)
Public Declare Sub SHEmptyRecycleBinW Lib "shell32.dll" (ByVal hWnd As Long, ByVal pszRootPath As Long, ByVal dwflags As Long)



Public Declare Function CreateDirectoryExW Lib "kernel32.dll" (ByVal lpTemplateDirectory As Long, ByVal lpNewDirectory As Long, ByVal lpSecurityAttributes As Long) As Long
Public Declare Function CreateDirectoryExA Lib "kernel32.dll" (ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, ByVal lpSecurityAttributes As Long) As Long


Private Declare Function GetCompressedFileSizeA Lib "kernel32.dll" (ByVal lpFileName As String, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function GetCompressedFileSizeW Lib "kernel32.dll" (ByVal lpFileName As Long, ByRef lpFileSizeHigh As Long) As Long


'Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
         Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long

Public Declare Function SHGetDiskFreeSpaceEx Lib "shell32.dll" Alias "SHGetDiskFreeSpaceExA" (ByVal pszDirectoryName As String, ByRef pulFreeBytesAvailableToCaller As ULARGE_INTEGER, ByRef pulTotalNumberOfBytes As ULARGE_INTEGER, ByRef pulTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
'Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'Error Code Value Meaning
Public Enum SHFileOperationErrors
    DE_SAMEFILE = &H71
'DE_SAMEFILE 0x71 The source and destination files are the same file.
    DE_MANYSRC1DEST = &H72
'DE_MANYSRC1DEST 0x72 Multiple file paths were specified in the source buffer, but only one destination file path.
    DE_DIFFDIR = &H73
'DE_DIFFDIR 0x73 Rename operation was specified but the destination path is a different directory. Use the move operation instead.
 DE_ROOTDIR = &H74
'DE_ROOTDIR 0x74 The source is a root directory, which cannot be moved or renamed.
 DE_OPCANCELLED = &H75
'DE_OPCANCELLED 0x75 The operation was cancelled by the user, or silently cancelled if the appropriate flags were supplied to SHFileOperation.
DE_DESTSUBTREE = &H76
'DE_DESTSUBTREE 0x76 The destination is a subtree of the source.
DE_ACCESSDENIEDSRC = &H78
'DE_ACCESSDENIEDSRC 0x78 Security settings denied access to the source.
DE_PATHTOODEEP = &H79
'DE_PATHTOODEEP 0x79 The source or destination path exceeded or would exceed MAX_PATH.
DE_MANYDEST = &H7A
'DE_MANYDEST 0x7A The operation involved multiple destination paths, which can fail in the case of a move operation.
DE_INVALIDFILES = &H7C
'DE_INVALIDFILES 0x7C The path in the source or destination or both was invalid.
DE_DESTSAMETREE = &H7D
'DE_DESTSAMETREE 0x7D The source and destination have the same parent folder.
DE_FLDDESTISFILE = &H7E
'DE_FLDDESTISFILE 0x7E The destination path is an existing file.
DE_FILEDESTISFLD = &H80
'DE_FILEDESTISFLD 0x80 The destination path is an existing folder.
DE_FILENAMETOOLONG = &H81
'DE_FILENAMETOOLONG 0x81 The name of the file exceeds MAX_PATH.
DE_DEST_IS_CDROM = &H82
'DE_DEST_IS_CDROM 0x82 The destination is a read-only CD-ROM, possibly unformatted.
DE_DEST_IS_DVD = &H83
'DE_DEST_IS_DVD 0x83 The destination is a read-only DVD, possibly unformatted.
DE_DEST_IS_CDRECORD = &H84
'DE_DEST_IS_CDRECORD 0x84 The destination is a writable CD-ROM, possibly unformatted.
DE_FILE_TOO_LARGE = &H85
'DE_FILE_TOO_LARGE 0x85 The file involved in the operation is too large for the destination media or file system.
DE_SRC_IS_CDROM = &H86
'DE_SRC_IS_CDROM 0x86 The source is a read-only CD-ROM, possibly unformatted.
DE_SRC_IS_DVD = &H87
'DE_SRC_IS_DVD 0x87 The source is a read-only DVD, possibly unformatted.
DE_SRC_IS_CDRECORD = &H88
'DE_SRC_IS_CDRECORD 0x88 The source is a writable CD-ROM, possibly unformatted.
DE_ERROR_MAX = &HB7

'DE_ERROR_MAX 0xB7 MAX_PATH was exceeded during the operation.

DE_UNKNOWN = &H402

ERRORONDEST = &H10000
' 0x402 An unknown error occurred. This is typically due to an invalid path in the source or destination. This error does not occur on Windows Vista and later.
'ERRORONDEST 0x10000 An unspecified error occurred on the destination.
'DE_ROOTDIR | ERRORONDEST 0x10074 Destination is a root directory and cannot be renamed.

End Enum

Public Const BCFileErrorBase = vbObjectError + 512 * 2 + 128 + 64
Public LargeIcons As cVBALImageList 'cache
Public ShellIcons As cVBALImageList 'cache
Public SmallIcons As cVBALImageList 'cache
Public Winmetrics As New SystemMetrics
'struct HANDLETOMAPPINGS
'{
'    UINT              uNumberOfMappings;  // Number of mappings in the array.
'    LPSHNAMEMAPPING   lpSHNameMapping;    // Pointer to the array of mappings.
'};
Public Declare Sub SHFreeNameMappings Lib "shell32.dll" (ByVal hNameMappings As Long)

Private Declare Sub SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWnd As Long, ByVal csidl As SpecialFolderConstants, ByVal ppidl As Long)

Public Type HANDLETOMAPPINGS
    uNumberOfMappings As Long
    LpFirstMapping As Long

End Type
Public Type SHNAMEMAPPING
    pszOldPath As String
    pszNewPath As String
    cchOldPath As Long
    cchNewPath As Long
End Type


Private Type MungeLong
    LongA As Long
    LongB As Long
End Type
Private Type MungeCurr
    CurrA As Currency
End Type
'
'Private ForceANSI As Boolean

Global ForceANSI As Boolean

Private Type IO_STATUS_BLOCK
IoStatus                As Long
Information             As Long
End Type

Private Const DATA_1 As String = "::$DATA"
Private Const DATA_2 As String = ":encryptable:$DATA"
Private Type FILE_STREAM_INFORMATION
    NextEntryOffset         As Long
    StreamNameLength        As Long
    StreamSize              As Long
    StreamSizeHi            As Long
    StreamAllocationSize    As Long
    StreamAllocationSizeHi  As Long
    StreamName(259)         As Byte
End Type

Private Const FileStreamInformation As Long = 22   ' from Enum FILE_INFORMATION_CLASS


Public mTotalObjectCount As Long


'typedef struct tagOPENASINFO {
'    LPCWSTR pcszFile;
'    LPCWSTR pcszClass;
'    OPEN_AS_INFO_FLAGS oaifInFlags;
'} OPENASINFO;
'enum tagOPEN_AS_INFO_FLAGS {
'    OAIF_ALLOW_REGISTRATION = 0x00000001,     // enable the "always use this file" checkbox (NOTE if you don't pass this, it will be disabled)
'    OAIF_REGISTER_EXT       = 0x00000002,     // do the registration after the user hits "ok"
'    OAIF_EXEC               = 0x00000004,     // execute file after registering
'    OAIF_FORCE_REGISTRATION = 0x00000008,     // force the "always use this file" checkbox to be checked (normally, you won't use the OAIF_ALLOW_REGISTRATION when you pass this)
'#if (NTDDI_VERSION >= NTDDI_LONGHORN)
'    OAIF_HIDE_REGISTRATION  = 0x00000020,     // hide the "always use this file" checkbox
'    OAIF_URL_PROTOCOL       = 0x00000040,     // the "extension" passed is actually a protocol, and open with should show apps registered as capable of handling that protocol
'#End If
'};
'typedef int OPEN_AS_INFO_FLAGS;

Public Enum OPEN_AS_INFO_FLAGS
    OAIF_ALLOW_REGISTRATION = &H1
    OAIF_REGISTER_EXT = &H2
    OAIF_EXEC = &H4
    OAIF_FORCE_REGISTRATION = &H8
    'only in vista...
    OAIF_VISTA_HIDE_REGISTRATION = &H20
    OAIF_VISTA_URL_PROTOCOL = &H40
End Enum
Private Type OPENASINFO
    pcszFile As Long 'wide string.
    pcszClass As Long 'wide string.
    oaidInFlags As OPEN_AS_INFO_FLAGS
End Type
'
'HRESULT SHOpenWithDialog(
'    HWND hwndParent,
'    const OPENASINFO *poainfo
');
'shite- only supported in Vista...
Public Declare Function SHOpenWithDialog Lib "shell32" (ByVal hWndParent As Long, poaInfo As OPENASINFO) As Long

Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function NtQueryInformationFile Lib "NTDLL.DLL" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal Length As Long, ByVal FileInformationClass As Long) As Long

Public Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hfile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long

Private Declare Function QueryDosDeviceW Lib "kernel32.dll" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long


Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, ByVal lpData As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (ByRef pBlock As Any, ByVal lpSubBlock As String, ByVal lplpBuffer As Long, ByRef puLen As Long) As Long

Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFOA) As Long
Private Type SHELLEXECUTEINFOA
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwflags As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByRef hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Public Enum SetErrorModeConstants
 SEM_FAILCRITICALERRORS = &H1
 SEM_NOALIGNMENTFAULTEXCEPT = &H4
 SEM_NOGPFAULTERRORBOX = &H2
 SEM_NOOPENFILEERRORBOX = &H8000
End Enum
Private mWaitingAsync As collection
'ReadFileEx Completion Routine:

Public Function TranslateAPIErrorCode(ByVal APIError As Long) As Long
    'Translates a Windows API Error Code to a Visual Basic Error.
    
    Static mLookup As Scripting.Dictionary
    If mLookup Is Nothing Then
        Set mLookup = New Dictionary
        'add values.
        
        
        
    End If
    
    
    



End Function


Public Function GetFileAttributes(ByVal StrFilename As String) As Long

    If MakeWideCalls Then
        If Not IsUNCPath(StrFilename) Then
            StrFilename = "//?/" & StrFilename
        End If
        GetFileAttributes = GetFileAttributesW(StrPtr(StrFilename))
    Else
        GetFileAttributes = GetFileAttributesA(StrFilename)
    
    End If


End Function
'VOID CALLBACK FileIOCompletionRoutine(
'  __in  DWORD dwErrorCode,
'  __in  DWORD dwNumberOfBytesTransfered,
'  __in  LPOVERLAPPED lpOverlapped
');

'routines used to add/remove streams that are waiting for a FileIOCompletionRoutine.
Public Function AddAsyncStream(ByVal obj As Object) As Long
    'adds to the collection. returns index (current count)
    
    If mWaitingAsync Is Nothing Then
        Set mWaitingAsync = New collection
    End If
    mWaitingAsync.Add obj
    AddAsyncStream = mWaitingAsync.Count + 1
    
    
End Function
Public Function removeAsyncStream(obj As Long) As Long
    mWaitingAsync.Remove obj



End Function
Public Function GetAsyncStream(ObjIndex As Long) As Object
    Set GetAsyncStream = mWaitingAsync.Item(ObjIndex)
End Function
Public Sub FileIOCompletionRoutine(ByVal dwErrorCode As Long, ByVal dwBytesTransferred As Long, OverlappedVar As OVERLAPPED)
'since we will only be using this in the context of WriteFileEx and ReadFileEx, we can use the "hevent" member to store the index into our array of pending file operations.
    Dim getobj As Object
    Dim casted As IAsyncProcess
    Debug.Print "FileIOCompletionRoutine"
    Set getobj = GetAsyncStream(OverlappedVar.hEvent)
    Set casted = getobj
    'casted.ExecAsync
    






End Sub



Public Function GetErrorMode() As SetErrorModeConstants
    Dim tmp As Long
    tmp = SetErrorMode(0)
    GetErrorMode = tmp
    SetErrorMode tmp
    
End Function
Public Function Random() As String

Const contextname As String = "Microsoft Enhanced Cryptographic Provider v1.0"



End Function



'Public Function GetFriendlyEXEName(ByVal StrEXE As String) As String
'Dim lpData As Long
'Dim lpSize As Long, stralloc As String
'Dim lpreturnstr As String, lpretlen As Long
'
'lpSize = GetFileVersionInfoSize(StrEXE, lpData)
'Dim ret As Long
'stralloc = Space$(lpSize)
'lpData = VarPtr(lpSize)
'ret = GetFileVersionInfo(StrEXE, 0, lpSize, ByVal lpData)
'lpreturnstr = Space$(255)
'VerQueryValue ByVal lpData, "FileDescription" & vbNullChar, lpreturnstr, lpretlen
'GetFriendlyEXEName = lpreturnstr
'
'End Function



'API WRAPPERS:
Public Sub SHEmptyRecycleBin(ByVal hWnd As Long, ByVal pszRootPath As Long, ByVal dwflags As Long)
    If MakeWideCalls Then
        SHEmptyRecycleBinW hWnd, StrPtr(pszRootPath), dwflags
    Else
        SHEmptyRecycleBinA hWnd, pszRootPath, dwflags
    End If



End Sub


Public Function GetCompressedFileSize(ByVal lpFileName As String, ByRef lpFileSizeHigh As Long) As Long
    
    If MakeWideCalls Then
        GetCompressedFileSize = GetCompressedFileSizeW(StrPtr(lpFileName), lpFileSizeHigh)
    Else
        GetCompressedFileSize = GetCompressedFileSizeA(lpFileName, lpFileSizeHigh)
    End If
    
End Function
Public Function SetFileAttributes(ByVal strFile As String, ByVal Attributes As FileAttributeConstants)

If MakeWideCalls Then
    SetFileAttributesW StrPtr(strFile), Attributes
Else
    SetFileAttributesA strFile, Attributes
End If

End Function
Public Sub SetFileTimes(ByVal mvarFileName As String, ByVal DateCreated As Date, ByVal lastAccess As Date, ByVal LastWrite As Date)
    Dim hfile As Long
    
    Dim ftdatecreated As FILETIME, ftlastaccess As FILETIME, ftlastwrite As FILETIME
    
    ftdatecreated = Date2FILETIME(DateCreated)
    ftlastaccess = Date2FILETIME(lastAccess)
    ftlastwrite = Date2FILETIME(LastWrite)
    'step one: open the file.
    hfile = CreateFile(mvarFileName, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0)
    
    If hfile > 0 Then
        'Success!
        
        
        SetFileTime hfile, ftdatecreated, ftlastaccess, ftlastwrite
        
        
        CloseHandle hfile
    Else
    
    
        RaiseAPIError Err.LastDllError, "MdlFileSystem::SetFileTimes()"
    
    End If



End Sub

'Private Const MAX_PATH = 260
'Private Declare Function CreateDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
'Private Declare Function RemoveDirectoryA Lib "kernel32.dll" (ByVal lpPathName As String) As Long
Public Function CreateDirectoryEx(ByVal lpTemplateDirectory As String, ByVal lpNewDirectory As String, ByVal lpSecurityAttributes As Long) As Long
If MakeWideCalls Then
    CreateDirectoryEx = CreateDirectoryExW(StrPtr(lpTemplateDirectory), StrPtr(lpNewDirectory), lpSecurityAttributes)
Else
    CreateDirectoryEx = CreateDirectoryExA(lpTemplateDirectory, lpNewDirectory, lpSecurityAttributes)
End If


End Function

Public Function CreateDirectory(ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long

If MakeWideCalls Then
    CreateDirectory = CreateDirectoryW(StrPtr(lpPathName), lpSecurityAttributes)
Else
    CreateDirectory = CreateDirectoryA(lpPathName, lpSecurityAttributes)
End If

End Function
Public Function RemoveDirectory(ByVal lpPathName As String) As Long

If MakeWideCalls Then
    RemoveDirectory = RemoveDirectoryW(StrPtr(lpPathName))
Else
    RemoveDirectory = RemoveDirectoryA(lpPathName)
End If

End Function

Public Function SHGetSpecialFolderPath(ByVal hWnd As Long, ByRef FolderLocation As String, FolderConst As SpecialFolderConstants, ByVal fCreate As Long) As Long
    
'    If MakeWideCalls Then
'        FolderLocation = Space$(MAX_PATH * 2)
'        SHGetSpecialFolderPath = SHGetSpecialFolderPathW(hwnd, StrPtr(FolderLocation), FolderConst, fCreate)
'        'folderlocation
'    Else
        FolderLocation = Space$(MAX_PATH)
        SHGetSpecialFolderPath = SHGetSpecialFolderPathA(hWnd, FolderLocation, FolderConst, fCreate)
'    End If




End Function

Public Function ListStreams_Vista(ByVal Spath As String) As CAlternateStreams
    Dim newstreams As CAlternateStreams
    Dim createalt As CAlternateStream
    Dim useBuff As WIN32_FIND_STREAM_DATA
    Dim FHandle As Long, usename As String
    Set newstreams = New CAlternateStreams
    FHandle = FindFirstStreamW(StrPtr(Spath), 0, useBuff, 0)
    Do
        Set createalt = New CAlternateStream
        usename = Replace$(Mid$(Trim$(StrConv(useBuff.cStreamName, vbFromUnicode)), 2), vbNullChar, "")
        'if the stream name ends with ":$DATA" then strip that off... but only if the name is longer  then 7 characters...
        
        If Len(usename) > 7 Then
            If Right$(usename, 6) = ":$DATA" Then
            usename = Mid$(usename, 1, Len(usename) - 6)
            
            End If
        
        End If
        
        If usename <> "" Then
            With createalt
                .Init Spath, usename, useBuff.StreamSize.lowpart, useBuff.StreamSize.highpart, 0
            End With
            newstreams.Add createalt
        End If
    
        ZeroMemory useBuff, Len(useBuff)
        If FindNextStreamW(FHandle, useBuff) = 0 Then
            Exit Do
        End If
        
    
    Loop
    FindClose FHandle

    Set ListStreams_Vista = newstreams
End Function
Public Function HardLinks_Vista(ByVal Spath As String, Optional ByRef Count As Long) As String()
'returns a list of files with hardlinks to this one.





End Function
Public Function HardLinks_PreVista(ByVal Spath As String, Optional ByRef Count As Long) As String()
    'Awful... just... awful...
    
    
    'anyway.... get the file index, and then literally search on all files on that drive for the links...
    'returns all hardlinks OTHER then this file.
    
    Dim FolderStack() As Directory, StackTop As Long
    Dim gotfile As CFile
    Dim findex As Double
    Dim ret() As String
    Dim linkcount As Long
    Dim startfolder As Directory
    Dim currfile As CFile
    Dim currdir As Directory
    Set gotfile = FileSystem.GetFile(Spath)
    'retrieve the file index...
    findex = gotfile.FileIndex
    If gotfile.HardLinkCount = 1 Then
        'only one link... this one.
        Count = 0
        
    Else
        Set startfolder = gotfile.Directory.Volume.RootFolder
        'find all files, then all folders.
        StackTop = 1
        ReDim FolderStack(1 To 1)
        Set FolderStack(1) = startfolder
        Do Until StackTop = 0
            'Grab the topmost item....
            Dim topItem As Directory
            Set topItem = FolderStack(StackTop)
            StackTop = StackTop - 1
            If StackTop > 0 Then
            ReDim Preserve FolderStack(1 To StackTop)
            End If
            '"remove" this item...
            'If StrComp(Left$(topItem.Path, 9), "D:\VBPROJ", vbTextCompare) = 0 Then Stop
            Dim loopfile As CFile, LoopDir As Directory
            'now, loop through all files...
            With topItem.Files.GetWalker
            Do Until .GetNext(loopfile) Is Nothing
            If StrComp(loopfile.Fullpath, "D:\vbproj\vb\testhl.txt", vbTextCompare) = 0 Then Stop
                If loopfile.FileIndex = findex Then
                    'add to our return array...
                    If loopfile.Fullpath <> gotfile.Fullpath Then
                        linkcount = linkcount + 1
                        ReDim Preserve ret(1 To linkcount)
                        ret(linkcount) = loopfile.Fullpath
                        If linkcount = (gotfile.HardLinkCount - 1) Then
                            'all links found...
                            'break out...
                            HardLinks_PreVista = ret
                            Count = linkcount
                            Exit Function
                        End If
                    End If
                   ' FindFirstStreamW
                End If
            Loop
            End With
            'OK, now loop through the directories...
            With topItem.Directories.GetWalker
                Do Until .GetNext(LoopDir) Is Nothing
                    'push it into the stack...
                    StackTop = StackTop + 1
                    ReDim Preserve FolderStack(1 To StackTop)
                    Set FolderStack(StackTop) = LoopDir
                    'If InStr(1, LoopDir.Path, "vbproj", vbTextCompare) > 0 Then Stop
                    Debug.Print "stacktop=" & StackTop & " Folder " & LoopDir.Path
                Loop
            
            End With
            
        
        Loop
    
    End If


End Function
Public Function ListStreams(Spath As String) As CAlternateStreams
    'Purpose: call ListStreams_Vista on Vista- Call ListStreams_NT for other OS's.
    
    'Non-NT platforms don't have Streams...
    If Not IsWinNt Then
        Set ListStreams = Nothing
        Exit Function
    End If
    
    
    If IsVistaOrLater Then
        Set ListStreams = ListStreams_Vista(Spath)
    Else
        Set ListStreams = ListStreams_NT(Spath)
    End If



End Function
Public Function ListStreams_NT(Spath As String) As CAlternateStreams
    Dim IOS As IO_STATUS_BLOCK
    Dim BBuf() As Byte
    Dim FSInfo As FILE_STREAM_INFORMATION
    Dim LBuf As Long, LInfo As Long, lRet As Long
    Dim LErr As Long
    Dim sName As String, SNames As String
    
    Dim newstreams As CAlternateStreams
    Set newstreams = New CAlternateStreams
    
    newstreams.Owner = Spath
    On Error Resume Next
    'ListStreams = ""
    lRet = CreateFile(Spath, DesiredAccessFlags.STANDARD_RIGHTS_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, _
    FILE_FLAG_BACKUP_SEMANTICS, 0&)
    If (lRet = -1) Then Exit Function
    
    LBuf = 4096
    LErr = 234
    ReDim BBuf(1 To LBuf)
    
    Do While LErr = 234
    
        LErr = NtQueryInformationFile(lRet, IOS, ByVal VarPtr(BBuf(1)), LBuf, _
        ByVal FileStreamInformation)
        If (LErr = 234) Then
        LBuf = LBuf + 4096
        ReDim BBuf(1 To LBuf)
        End If
        
    Loop
    
    LInfo = VarPtr(BBuf(1))
    Dim newstream As CAlternateStream
    Do
    
        CopyMemory ByVal VarPtr(FSInfo.NextEntryOffset), ByVal LInfo, Len(FSInfo)
        'CopyMemory ByVal VarPtr(FSInfo.StreamName(0)), ByVal LInfo + 24, _
'
        'FSInfo.StreamNameLength
        sName = Left$(FSInfo.StreamName, FSInfo.StreamNameLength / 2)
        
        If (InStr(1, sName, DATA_1, 1) = 0) And (InStr(1, sName, DATA_2, 1) = 0) _
        And (sName <> "") Then
            'SNames = SNames & Mid$(sName, 2, Len(sName) - 7) & " * " & _
            CStr(FSInfo.StreamSize) & "|"
          
            Set newstream = New CAlternateStream
            newstream.Init Spath, Mid$(sName, 2, Len(sName) - 7), _
            FSInfo.StreamSize, FSInfo.StreamSizeHi
            newstreams.Add newstream
            
        End If
        
        If FSInfo.NextEntryOffset Then
            LInfo = LInfo + FSInfo.NextEntryOffset
        Else
            Exit Do
        End If
        
    Loop
    CloseHandle lRet
    If (Len(SNames) > 0) Then SNames = Left$(SNames, (Len(SNames) - 1))
    Set ListStreams_NT = newstreams
    'Stop
End Function






Public Function GetDriveForNtDeviceName(ByVal sDeviceName As String) As String
Dim sFoundDrive As String
Dim strdrives As String
Dim DriveStr() As String
Dim vDrive As String, I As Long, ret As Long
strdrives = Space$(256)
ret = GetLogicalDriveStrings(255, strdrives)
strdrives = Trim$(Replace$(strdrives, vbNullChar, " "))
DriveStr = Split(strdrives, " ")
   'For Each vDrive In GetDrives()
   For I = 0 To UBound(DriveStr)
    vDrive = DriveStr(I)
      If StrComp(GetNtDeviceNameForDrive(vDrive), sDeviceName, vbTextCompare) = 0 Then
         sFoundDrive = vDrive
         Exit For
      End If
   Next I
   
   GetDriveForNtDeviceName = sFoundDrive
   
End Function

Public Function GetNtDeviceNameForDrive( _
   ByVal sDrive As String) As String
Dim bDrive() As Byte
Dim bresult() As Byte
Dim lR As Long
Dim sDeviceName As String

   If Right(sDrive, 1) = "\" Then
      If Len(sDrive) > 1 Then
         sDrive = Left(sDrive, Len(sDrive) - 1)
      End If
   End If
   bDrive = sDrive
   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
   ReDim bresult(0 To MAX_PATH * 2 + 1) As Byte
   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bresult(0)), MAX_PATH)
   If (lR > 2) Then
      sDeviceName = bresult
      sDeviceName = Left(sDeviceName, lR - 2)
      GetNtDeviceNameForDrive = sDeviceName
   End If
   
End Function



Public Function GetSpecialFolder(hWnd As Long, FolderConst As SpecialFolderConstants) As String

Dim FolderLocation As String, ret As Long
FolderLocation = Space$(2048)
ret = SHGetSpecialFolderPath(hWnd, FolderLocation, FolderConst, 0)

GetSpecialFolder = Trim$(Replace$(FolderLocation, vbNullChar, ""))



End Function
Public Function GetSpecialFolderPidl(ByVal hWnd As Long, ByVal folder As SpecialFolderConstants) As Long


Dim ret As Long

SHGetSpecialFolderLocation hWnd, folder, ret
GetSpecialFolderPidl = ret


End Function
Public Function MakeWideCalls() As Boolean

  'MakeWideCalls = (m_IsWinNt And (m_WideCallSupport <> AnsiVersion))
  MakeWideCalls = IsWinNt And Not ForceANSI
  
End Function

Public Function SizeOfString() As Long
  If MakeWideCalls Then
    SizeOfString = 2
  Else
    SizeOfString = 1
  End If
End Function

'ANSI/WIDE WRAPPERS


Public Function CreateFileMapping(ByVal hfile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Variant) As Long
    Dim lpstrname As String
    If Not IsNull(lpName) Then
    lpstrname = str(lpName)
    End If
    If MakeWideCalls Then
        If IsNull(lpName) Or IsEmpty(lpName) Or lpName = vbNullString Then
        CreateFileMapping = CreateFileMappingW(hfile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, ByVal 0&)
        Else
            CreateFileMapping = CreateFileMappingW(hfile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, StrPtr(lpstrname))
        End If
    Else
        If IsNull(lpName) Or IsEmpty(lpName) Or lpName = vbNullString Then
            CreateFileMapping = CreateFileMappingALong(hfile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, ByVal 0&)
        Else
            CreateFileMapping = CreateFileMappingA(hfile, lpFileMappingAttributes, flProtect, dwMaximumSizeHigh, dwMaximumSizeLow, lpstrname)
        End If
    End If

End Function
Public Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Dim lpfileret As Long
If Not IsUNCPath(lpFileName) Then

lpFileName = "\\?\" & lpFileName
End If

    If MakeWideCalls() Then
        lpfileret = CreateFileW(StrPtr(lpFileName), dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    
        
    Else
        lpfileret = CreateFileA(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)
    
    End If
    CreateFile = lpfileret
    If lpfileret = -1 Then
        Debug.Print "Error accessing " & lpFileName & ":" & GetAPIErrStr(Err.LastDllError)
    
    End If


End Function











Sub Main()
    Set Winmetrics = New SystemMetrics
    Set LargeIcons = New cVBALImageList
    LargeIcons.ColourDepth = ILC_COLOR32
    SmallIcons.ColourDepth = ILC_COLOR32
    LargeIcons.IconSizeY = Winmetrics.LargeIconSize
    LargeIcons.IconSizeX = Winmetrics.LargeIconSize
    LargeIcons.Create
    Set SmallIcons = New cVBALImageList
    SmallIcons.IconSizeX = Winmetrics.SmallIconSize
    SmallIcons.IconSizeX = Winmetrics.SmallIconSize
    SmallIcons.Create
    Set ShellIcons = New cVBALImageList
    
    Dim quickdllstream As FileStream
    Dim resbytes() As Byte, quickdllpath
    'expand DLLs needed.
    quickdllpath = FileSystem.GetSpecialFolder(CSIDL_SYSTEMX86).Path & "quick32.dll"
    
    If Not FileSystem.Exists(quickdllpath) Then
    resbytes = LoadResData("QUICK", "DLL")
    Set quickdllstream = FileSystem.CreateStream(quickdllpath)
    quickdllstream.WriteBytes resbytes
    quickdllstream.CloseStream
    End If
'    Dim TESTIT As String
'    Dim pidlRel As Long, pidlfile As Long
'    Dim relfolder As olelib.IShellFolder
    
End Sub
Public Sub DBL2LI(ByVal Dbl As Double, ByRef LoPart As Long, ByRef hipart As Long)
    Dim mungec As MungeCurr
    Dim mungeli As MungeLong
    mungec.CurrA = (Dbl / 10000#)
    LSet mungeli = mungec
    LoPart = mungeli.LongA
    hipart = mungeli.LongB
End Sub

Public Function LI2DBL(LoPart As Long, hipart As Long) As Double
    Dim mungel As MungeLong
    Dim mungec As MungeCurr
    mungel.LongA = LoPart
    mungel.LongB = hipart
    LSet mungec = mungel
    LI2DBL = mungec.CurrA * 10000#




End Function
Public Function FileTime2Date(FTIME As FILETIME) As Date
    Dim sTime As SYSTEMTIME
    Dim createdate As Date
    FileTimeToSystemTime FTIME, sTime
    With sTime
    createdate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    
    End With
    
    
    FileTime2Date = createdate
    
    
End Function
Public Function Date2FILETIME(Dateuse As Date) As FILETIME
    Dim retStruct As FILETIME
    Dim FromSTime As SYSTEMTIME
    With FromSTime
    .wSecond = Second(Dateuse)
    .wMinute = Minute(Dateuse)
    .wHour = Hour(Dateuse)
    .wDay = Day(Dateuse)
    
    .wMonth = Month(Dateuse)
    .wYear = Year(Dateuse)
    .wDayOfWeek = Weekday(Dateuse)
    
    End With
    SystemTimeToFileTime FromSTime, retStruct
    Date2FILETIME = retStruct
        
    
    
End Function
Public Sub AppendSlash(ByRef Path As String)
    If (Right$(Path, 1) <> "\" And Right$(Path, 1) <> "/") Then Path = Path & "\"

End Sub
Public Function FixPath(ByVal Path As String) As String
    'reformats a path to use understandable slashes, and other things.
    Dim Startreplace As Long
    'proper UNC form is
    '//SERVER/SHARE
    If IsUNCPath(Path) Then
        'Startreplace = InStr(3, Path, "/") + 1
        FixPath = Replace$(Path, "\", "/")
    Else
        FixPath = Replace$(Path, "/", "\")
        'Startreplace = 1
        
    
    End If
    


End Function
Public Sub RaiseAPIError(ByVal ErrCode As Long, ByVal ErrSource As String)

Dim MessageStr As String
MessageStr = GetAPIError(ErrCode)
If MessageStr = "" Then MessageStr = "Unexpected Error in " & ErrSource

Err.Raise BCFileErrorBase + ErrCode, ErrSource, MessageStr



End Sub
Public Function GetAPIError(ByVal ErrCode As Long) As String
    'raises a Windows API error.
 '   FormatMessage(
 ' FORMAT_MESSAGE_ALLOCATE_BUFFER |
 ' FORMAT_MESSAGE_FROM_SYSTEM,
 ' NULL,
 ' Err.LastDLLError(),
 ' MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), //The user default language
 ' (LPTSTR) &lpMessageBuffer,
 ' 0,
 ' NULL );
Dim lpBuffer As String
lpBuffer = Space$(128)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrCode, 0, lpBuffer, Len(lpBuffer), ByVal 0&
GetAPIError = Replace$(Trim$(lpBuffer), vbNullChar, "")


End Function
Public Function ParseVolume(ByVal StrFrom As String) As String

Dim ret As String
ParsePathParts StrFrom, ret, , , , , Parse_Volume
ParseVolume = ret


End Function
Public Function ParseStreamName(ByVal StrFrom As String) As String
    Dim retStr As String
    ParsePathParts StrFrom, , , , , retStr, Parse_Stream
    
    ParseStreamName = retStr



End Function
Public Function ParseExtension(ByVal StrFrom As String) As String
    Dim ext As String
    ParsePathParts StrFrom, , , , ext
    ParseExtension = ext
End Function
Public Function ParseFilename(ByVal StrFrom As String, Optional ByVal WithExtension As Boolean = True) As String


    Dim FName As String, Fextension As String
    ParsePathParts StrFrom, , , FName, Fextension
    If WithExtension Then
        ParseFilename = FName & "." & Fextension
    Else
        ParseFilename = FName
    End If



End Function
Public Function ParsePath(ByVal StrFrom As String) As String
    'parse the path from a specification
    Dim retVal As String
    ParsePathParts StrFrom, , retVal
    
    ParsePath = retVal


End Function

'Public Function TrimNull$(ByVal szString As String)
'    TrimNull = Trim$(Replace$(szString, vbNullChar, ""))
'End Function

'SH stuff.
Public Function DisplayName(ByVal lpszPath As String, Optional ByVal AssumeExist As Boolean) As String
    Dim finfo As SHFILEINFO
    Dim ValFlag As ShellFileInfoConstants
    ValFlag = SHGFI_DISPLAYNAME
    FixAssume ValFlag, AssumeExist
    SHGetFileInfo lpszPath, 0, finfo, Len(finfo), SHGFI_DISPLAYNAME
    DisplayName = TrimNull$(finfo.szDisplayName)
    


End Function
Public Function FileTypeName(ByVal strFile As String, Optional ByVal AssumeExist As Boolean) As String
    Dim finfo As SHFILEINFO
    Dim ValFlag As ShellFileInfoConstants
    ValFlag = SHGFI_TYPENAME
    FixAssume ValFlag, AssumeExist
    SHGetFileInfo strFile, 0, finfo, Len(finfo), SHGFI_TYPENAME
    FileTypeName = TrimNull$(finfo.szTypeName)
End Function








Private Sub FixAssume(ByRef ValFix As ShellFileInfoConstants, ByVal Assume As Boolean)
    If Assume Then
        ValFix = ValFix Or SHGFI_USEFILEATTRIBUTES
    End If
End Sub
Function PidlFromPath(Spath As String) As Long
    Dim pidl As Long, f As Long
    f = SHGetPathFromIDList(pidl, Spath)
    If f Then PidlFromPath = pidl
End Function
Public Function ShowExplorerMenu(ByVal HwndOwner As Long, ByVal pszPath As String, Optional x As Long = -1, Optional y As Long = -1, Optional menucallback As IContextCallback = Nothing, Optional CMFFlags As QueryContextMenuFlags = CMF_EXPLORE) As Long
    'displays the Explorer menu.
    
    'MFoldTool.ContextPopMenu hwndOwner, pszPath, x, y
    
    
    Dim pidlRel As Long, pidlpath As Long, pidlfile As Long
    Dim deskfolder As olelib.IShellFolder
    Dim parentfolder As olelib.IShellFolder
    Dim Pointuse As POINTAPI
    'Always relative to desktop.
    SHGetDesktopFolder deskfolder

    'PidlPath = SHSimpleIDListFromPath(pszPath)
    'PidlPath = PidlFromPath(pszPath)
     If x = -1 And y = -1 Then
            GetCursorPos Pointuse
        Else
            Pointuse.x = x
            Pointuse.y = y
        End If
    
    If Len(pszPath) <= 3 Then
        
        'why, it's a drive spec.
        deskfolder.ParseDisplayName HwndOwner, 0, StrPtr(pszPath), 0, pidlfile, 0
        Set parentfolder = deskfolder
    
    Else
    
        Set parentfolder = FolderFromItem(HwndOwner, pszPath, pidlpath)
       
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, 0, Pointuse)
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, PidlPath, Pointuse)
        
        
        'Current Work: this line errs when a Drive name is specified.
        
        
        
        parentfolder.ParseDisplayName HwndOwner, 0, StrPtr(Mid$(pszPath, InStrRev(pszPath, "\") + 1)), 0, pidlfile, 0
        
    End If
    
    'To allow for multiple files:
    
    'they would all need to be in the same dir, BTW-
    
    'retrieve common parentfolder
    'acquire pidl for each Item for which a context menu is to be shown-
    '(ParseDisplayName method)
    
    'with the new array of Pidls, call the ShowShellContextMenu function, with an appropriate count and the first item in the array.
    ShowShellContextMenu HwndOwner, parentfolder, 1, pidlfile, Pointuse, menucallback, , CMFFlags
    'Call ShowShellContextMenu(hwndOwner, PidlPath, 1, 0, Pointuse)





End Function

Public Function ShowExplorerMenuMulti(ByVal HwndOwner As Long, ByVal pszPath As String, StrFiles() As String, Optional x As Long = -1, Optional y As Long = -1, Optional CallbackObject As IContextCallback = Nothing) As Long
    'displays the Explorer menu.
    
    'MFoldTool.ContextPopMenu hwndOwner, pszPath, x, y
    
    
    Dim pidlRel As Long, pidlpath As Long, pidlfile() As Long
    Dim deskfolder As olelib.IShellFolder
    Dim parentfolder As olelib.IShellFolder
    Dim Pointuse As POINTAPI, I As Long
    'Always relative to desktop.
    SHGetDesktopFolder deskfolder

    'PidlPath = SHSimpleIDListFromPath(pszPath)
    'PidlPath = PidlFromPath(pszPath)
     If x = -1 And y = -1 Then
            GetCursorPos Pointuse
        Else
            Pointuse.x = x
            Pointuse.y = y
        End If
    
    'If Len(pszPath) <= 3 Then
        
        'why, it's a drive spec.
        'DeskFolder.ParseDisplayName hwndOwner, 0, StrPtr(pszPath), 0, pidlfile, 0
   '     Set ParentFolder = DeskFolder
    'Else
    'Else
    
        Set parentfolder = FolderFromItem(HwndOwner, pszPath, pidlpath)
    'End If
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, 0, Pointuse)
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, PidlPath, Pointuse)
        
        
        'Current Work: this line errs when a Drive name is specified.
          'retrieve common parentfolder
        'acquire pidl for each Item for which a context menu is to be shown-
        '(ParseDisplayName method)
        'with the new array of Pidls, call the ShowShellContextMenu function, with an appropriate count and the first item in the array.
        ReDim pidlfile(UBound(StrFiles))
        For I = 0 To UBound(StrFiles)
            parentfolder.ParseDisplayName HwndOwner, 0, StrPtr(pszPath & StrFiles(I)), 0, pidlfile(I), 0
        Next
        'ParentFolder.ParseDisplayName hwndOwner, 0, StrPtr(Mid$(pszPath, InStrRev(pszPath, "\") + 1)), 0, pidlfile, 0
        
   ' End If
    
    'To allow for multiple files:
    
    'they would all need to be in the same dir, BTW-
    
  
    
    
    ShowShellContextMenu HwndOwner, parentfolder, UBound(pidlfile) + 1, pidlfile(0), Pointuse, CallbackObject
    'Call ShowShellContextMenu(hwndOwner, PidlPath, 1, 0, Pointuse)





End Function
Public Function IsFileName(ByVal Spec As String) As Boolean
    'returns wether Spec specifies a Filename.
    'for example:
    
    'C:\ would return false.
    'C:\a would return true, unless a folder currently exists in C called "a"
    
    Dim Attribs As FileAttributeConstants
    Attribs = GetFileAttributes(Spec)
    
    If Attribs > 0 Then
        If (Attribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Or _
        (Attribs And FILE_ATTRIBUTE_DEVICE) = FILE_ATTRIBUTE_DEVICE Then
            IsFileName = False
        Else
        
            IsFileName = True
        End If
    
    Else
        IsFileName = False 'not found...
    End If



End Function
Public Function isDirectory(ByVal Spec As String) As Boolean
    Dim Attribs As FileAttributeConstants
    Attribs = GetFileAttributes(Spec)
    isDirectory = ((Attribs And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
End Function
Public Function GetName(ByVal Path As String) As String
    'returns the Name portion of the path.
    'for example:
    
    '"TEST.DLL"
    'would return TEST
    'and C:\Windows\System32.old\
    'should return System32
    
    
End Function
Public Function GetStream(ByVal OfPath As String) As String
    'returns the Stream name of the file.
    'look backwards through the path, if a ":" is closer then a \, then grab the value between the colon and the end of the string.
    Dim LastColon As Long, lastSlash As Long
    OfPath = Replace$(OfPath, "/", "\")
    LastColon = InStrRev(OfPath, ":")
    lastSlash = InStrRev(OfPath, "\")
    If LastColon > lastSlash Then
        'stream name present....
        GetStream = Mid$(OfPath, LastColon)
    Else
        'no stream name.
        GetStream = ""
        End If
    
    
End Function
'Code to extract portions from a file path.
Public Function GetVolume(ByVal Path As String) As String
'Precondition: Path is an absolute path.

'supported "Volume" syntaxes:

'<Driveletter>:

'//VolumeName/
Path = Replace$(Path, "/", "\")
If Left$(Path, 2) = "\\" Then

'UNC
    GetVolume = Replace$(Mid$(Path, 1, InStr(3, Path, "\")), "\", "/")
ElseIf Mid$(Path, 2, 1) = ":" Then
    'Drive spec.
    GetVolume = Mid$(Path, 1, 3)
End If


End Function





' Display a context menu from a folder
' Based on C code by Jeff Procise in PC Magazine
' Destroys any pidl passed to it, so pass duplicate if necessary
'Function ContextPopMenu(ByVal hwnd As Long, vItem As Variant, _
'                        ByVal x As Long, ByVal y As Long) As Boolean
'    InitIf  ' Initialize if in standard modue
'
'    Dim folder As IShellFolder, pidlMenu As Long
'    Dim menu As IContextMenu3, ici As CMINVOKECOMMANDINFO
'    Dim iCmd As Long, f As Boolean, hMenu As Long
'
'    ' Get folder and pidl from path, pidl, or special item
'    Set folder = FolderFromItem(vItem, pidlMenu)
'    If folder Is Nothing Then Exit Function
'
'    ' Get an IContextMenu object
'    On Error GoTo ContextPopMenuFail
'    folder.GetUIObjectOf hwnd, 1, pidlMenu, iidContextMenu, 0, menu
'
'    ' Create an empty popup menu and initialize it with QueryContextMenu
'    hMenu = CreatePopupMenu
'    On Error GoTo ContextPopMenuFail2
'    menu.QueryContextMenu hMenu, 0, 1, &H7FFF, CMF_EXPLORE
'
'    ' Convert x and y to client coordinates
'    ClientToScreenXY hwnd, x, y
'
'    ' Display the context menu
'    Const afMenu = TPM_LEFTALIGN Or TPM_LEFTBUTTON Or _
'                   TPM_RIGHTBUTTON Or TPM_RETURNCMD
'    iCmd = TrackPopupMenu(hMenu, afMenu, x, y, 0, hwnd, ByVal hNull)
'
'    ' If a command was selected from the menu, execute it.
'    If iCmd Then
'        ici.cbSize = LenB(ici)
'        ici.fMask = 0
'        ici.hwnd = hwnd
'        ici.lpVerb = iCmd - 1
'        ici.lpParameters = pNull
'        ici.lpDirectory = pNull
'        ici.nShow = SW_SHOWNORMAL
'        ici.dwHotKey = 0
'        ici.hIcon = hNull
'        menu.InvokeCommand ici
'        ContextPopMenu = True
'    End If
'
'ContextPopMenuFail2:
'    DestroyMenu hMenu
'
'ContextPopMenuFail:
'    ' Menu pidl is freed, so client had better not pass only copy
'    Allocator.Free pidlMenu
'    BugMessage Err.Description
'
'End Function
'DWORD CALLBACK CopyProgressRoutine(
'  __in      LARGE_INTEGER TotalFileSize,
'  __in      LARGE_INTEGER TotalBytesTransferred,
'  __in      LARGE_INTEGER StreamSize,
'  __in      LARGE_INTEGER StreamBytesTransferred,
'  __in      DWORD dwStreamNumber,
'  __in      DWORD dwCallbackReason,
'  __in      HANDLE hSourceFile,
'  __in      HANDLE hDestinationFile,
'  __in_opt  LPVOID lpData
');
'ugh, haven't written callbacks for a while- can't remember wether to use byval or not...
'#
Public Function CopyProgressRoutine(ByVal TotalFileSize As Currency, ByVal TotalBytesTransferred As Currency, _
ByVal StreamSize As Currency, ByVal StreamBytesTransferred As Currency, ByVal dwStreamNumber As Long, _
ByVal dwCallbackReason As Long, ByVal hSourceFile As Long, ByVal hDestinationFile As Long, ByVal lpData As Long) As Long
                                    
            'step one: convert large integer arguments to doubles.
            Dim dFileSize As Double, dBytesTransferred As Double, dStreamSize As Double, dStreamTransferred As Double
            Dim mCallback As IProgressCallback, SourceFile As CFile, DestFile As CFile
            Dim sSource As String, SDest As String
            Dim gotdata As BCCOPYFILEDATA
            Debug.Print "copyprogressroutine!"
            'Stop
            'FileSize = LI2DBL(TotalFileSizeHigh, TotalFileSizeLow)
            'BytesTransferred = LI2DBL(TotalBytesTransferredHigh, TotalBytesTransferredLow)
            'StreamSize = LI2DBL(StreamSizeHigh, StreamSizeLow)
            'StreamTransferred = LI2DBL(StreamBytesTransferredHigh, StreamBytesTransferredLo)
            'CDebug.Post "copyprogressroutine"
            dFileSize = TotalFileSize * 10000
            dBytesTransferred = TotalBytesTransferred * 10000
            dStreamSize = StreamSize * 10000
            dStreamTransferred = StreamBytesTransferred * 10000
            'whew.
            CDebug.Post "CopyProgressRoutine " & dFileSize & "," & dBytesTransferred & "," & dStreamSize & "," & dStreamTransferred
            'dwData will contain address of the object that invoked the filecopy.
            sSource = GetFileNameFromHandle(hSourceFile)
            SDest = GetFileNameFromHandle(hDestinationFile)
            Set SourceFile = FileSystem.GetFile(sSource)
            If SDest <> "" Then
            Set DestFile = FileSystem.GetFile(SDest)
            End If
          'copy it to a IProgressCallback object, and invoke...
          'CopyMemory mCallback, ByVal lpData, 4
          'copymemory
          If Not mCallback Is Nothing Then
          mCallback.UpdateProgress SourceFile, DestFile, dFileSize, dBytesTransferred, dStreamSize, dStreamTransferred
          
          CopyMemory mCallback, 0, 4
          End If
                                    
                                    
                                    
                                    
        
End Function
Public Function FileExists(ByVal PathSpec As String) As Boolean
    'returns wether a file exists.
'    Dim hFile As Long
'    hFile = CreateFile(PathSpec, GENERIC_DEVICE_QUERY, 0, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
'    If hFile > 0 Then
'        FileExists = True
'        CloseHandle hFile
'    Else
'
'
'
'        FileExists = False
'    End If
    'use getfileAttributes...
    Dim Attribs As Long
    
    Attribs = GetFileAttributes(PathSpec)
    If Attribs > 0 Then
        If (Attribs And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
        FileExists = True
        Else
            FileExists = False
        End If
    End If
    
End Function
Public Function GetTempFileNameAndPathEx() As String
    Dim tpath As String, tfile As String
    tpath = Space$(2048)
    Call GetTempPath(2047, tpath)
    tpath = Left$(tpath, InStr(tpath, vbNullChar) - 1)
    tfile = GetTempFileNameEx
    
    tfile = IIf(Right$(tpath, 1) <> "\", tpath & "\", tpath) & tfile
    GetTempFileNameAndPathEx = tfile
End Function

Public Function GetTempFileNameEx() As String
    'creates a temporary file name based on GUIDs.
    Dim FName As String
    Dim thebytes(1 To 16) As Byte
    Dim pg As Guid, I As Long
    CoCreateGuid pg
    CopyMemory thebytes(1), pg, 16
    'fname = fname & pg.Data1 / 2
    'fname = fname & pg.Data2 / 3
    'fname = fname & pg.Data3 / 4
    For I = 1 To 16
    FName = FName & Chr$((thebytes(I) Mod 24) + 65)
    Next I
    'fname = fname & pg.Data4(
    GetTempFileNameEx = FName & ".TMP"
End Function

'#define BUFSIZE 512
'
Public Function TrimNull(ByVal Strtrim As String) As String
If InStr(Strtrim, vbNullChar) > 0 Then
TrimNull = Mid$(Strtrim, 1, InStrRev(Strtrim, vbNullChar) - 1)
Else
    TrimNull = Strtrim
End If
End Function
Public Function GetFileNameFromHandle(ByVal FileHandle As Long) As String
'BOOL GetFileNameFromHandle(HANDLE hFile)
'{
'  BOOL bSuccess = FALSE;
'  TCHAR pszFilename[MAX_PATH+1];
'  HANDLE hFileMap;
'
'  // Get the file size.
'  DWORD dwFileSizeHi = 0;
'  DWORD dwFileSizeLo = GetFileSize(hFile, &dwFileSizeHi);
'
'  if( dwFileSizeLo == 0 && dwFileSizeHi == 0 )
'  {
'     printf("Cannot map a file with a length of zero.\n");
'     return FALSE;
'  }
'
'  // Create a file mapping object.
'  hFileMap = CreateFileMapping(hFile,
'                    NULL,
'                    PAGE_READONLY,
'                    0,
'                    1,
'                    NULL);
'

    Dim hFileMap As Long, sFileName As String
    Dim dwFileSizeHi As Long
    Dim dwFileSizeLo As Long, pmem As Long
    Dim breturn As Boolean, nbytes As Long
    Dim stemp As String, drives() As String, I As Long
    Const PAGE_READONLY As Long = &H2
    Const SECTION_MAP_READ As Long = &H4
    Const FILE_MAP_READ As Long = SECTION_MAP_READ
    
trysize:
    dwFileSizeLo = GetFileSize(FileHandle, dwFileSizeHi)
    If dwFileSizeLo = 0 And dwFileSizeHi = 0 Then
        'can't map zero-length file.
        'So... write to it >:)
        breturn = WriteFile(FileHandle, ".", 1, nbytes, ByVal &H0)
        If breturn = 0 Then
            'Epic FAIL.
            GetFileNameFromHandle = ""
        Else
            GoTo trysize
        End If
        
    Else
    
    
    
    
        hFileMap = CreateFileMapping(FileHandle, ByVal 0, PAGE_READONLY, 0, 1, "")
        If hFileMap Then
            pmem = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 1)
            If pmem Then
                sFileName = Space$(2048)
                Dim sztemp As String
                If GetMappedFileName(GetCurrentProcess, ByVal pmem, sFileName, 2047) Then
                    'translate path with device name to drive letters.
                    stemp = vbNullChar & Space(2047)
                    If GetLogicalDriveStrings(2057, stemp) Then
                        drives = Split(stemp, vbNullChar)
                        For I = 0 To UBound(drives)
                            If Trim$(drives(I)) <> "" Then
                                stemp = vbNullChar & Space(2047)
                                QueryDosDevice Left$(drives(I), 2), stemp, 2048
                                stemp = Left$(stemp, InStr(stemp, vbNullChar) - 1)
                                If Len(stemp) > 0 Then
                                
                                    If StrComp(stemp, Left$(sFileName, Len(stemp)), vbTextCompare) = 0 Then
                                        GetFileNameFromHandle = drives(I) & Mid$(sFileName, Len(stemp) + 2)
                                        Exit For
                                        
                                    End If
                                    
                                End If
                            End If
                        Next I
                    
                    
                    
                        
                    
                           'TCHAR szName[MAX_PATH];
'          TCHAR szDrive[3] = TEXT(" :");
'          BOOL bFound = FALSE;
'          TCHAR* p = szTemp;
'
'          Do
'          {
'            // Copy the drive letter to the template string
'            *szDrive = *p;
'
'            // Look up each device name
'            if (QueryDosDevice(szDrive, szName, MAX_PATH))
'            {
'              UINT uNameLen = _tcslen(szName);
'
'              if (uNameLen < MAX_PATH)
'              {
'                bFound = _tcsnicmp(pszFilename, szName,
'                    uNameLen) == 0;
'
'                if (bFound)
'                {
'                  // Reconstruct pszFilename using szTempFile
'                  // Replace device path with DOS path
'                  TCHAR szTempFile[MAX_PATH];
'                  StringCchPrintf(szTempFile,
'                            MAX_PATH,
'                            TEXT("%s%s"),
'                            szDrive,
'                            pszFilename+uNameLen);
'                  StringCchCopyN(pszFilename, MAX_PATH+1, szTempFile, _tcslen(szTempFile));
'                }
'              }
'            }
'
'            // Go to the next NULL character.
'            while (*p++);
'          } while (!bFound && *p); // end of string
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    End If
                
                End If
            
            End If
        End If
        
    End If
    If pmem <> 0 Then UnmapViewOfFile pmem
    If hFileMap <> 0 Then CloseHandle hFileMap
End Function
'  if (hFileMap)
'  {
'    // Create a file mapping to get the file name.
'    void* pMem = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 1);
'
'    if (pMem)
'    {
'      if (GetMappedFileName (GetCurrentProcess(),
'                             pMem,
'                             pszFilename,
'                             MAX_PATH))
'      {
'
'        // Translate path with device name to drive letters.
'        TCHAR szTemp[BUFSIZE];
'        szTemp[0] = '\0';
'
'        if (GetLogicalDriveStrings(BUFSIZE-1, szTemp))
'        {
'
'        }
'      }
'      bSuccess = TRUE;
'      UnmapViewOfFile(pMem);
'    }
'
'    CloseHandle(hFileMap);
'  }
'  _tprintf(TEXT("File name is %s\n"), pszFilename);
'  return(bSuccess);
'}



Public Function GetOpenFileString(ByVal Extension As String)
Dim creg As cRegistry, Filemask As String, defvalue As String
'Input: something such as "*.jpg" or "*.txt"
If Left$(Extension, 1) <> "." Then Extension = "." & Extension
'Registry operations will be used to determine the file type-
'for example- let's take, *.EXE

'look at key, ".EXE"
Filemask = Replace$(Extension, "*", "")
'if the default value of that key is also present in HKEY_CLASSES_ROOT, then grab the default value from that key- otherwise return the default value of this key.
Set creg = New cRegistry
defvalue = creg.ValueEx(HHKEY_CLASSES_ROOT, Extension, "", RREG_SZ, "")
If defvalue = "" Then

End If
End Function
Function GetFileIcon(ByVal Spath As String, ByVal iconsize As IconSizeConstants) As Long
    Dim finfo As SHFILEINFO
    Dim lIconType As Long
    Dim attruse As FileAttributeConstants
    Dim flags As Long
    
    If SmallIcons Is Nothing Then
        Set SmallIcons = New cVBALImageList
        
        SmallIcons.Create
    End If
    If LargeIcons Is Nothing Then
        Set LargeIcons = New cVBALImageList
            LargeIcons.Create
            
        End If
    ' be sure that there is the mbNormalIcon too
   
    ' retrieve the item's icon
    flags = SHGFI_ATTR_SPECIFIED
    If iconsize = icon_large Then
        flags = flags + SHGFI_ICON
    ElseIf iconsize = ICON_SMALL Then
        flags = flags + SHGFI_SMALLICON + SHGFI_ICON
    ElseIf iconsize = icon_shell Then
        flags = flags + SHGFI_SHELLICONSIZE + SHGFI_ICON
    End If

    SHGetFileInfo Spath, attruse, finfo, Len(finfo), flags
    ' convert the handle to a StdPicture
    GetFileIcon = finfo.hIcon
End Function
Public Function GetPathDepth(ByVal ppath As String) As Long
    'returns the depth of the specified file/folder.
    
    
    Dim Spath As String
    'parse the path...
    
    
    ParsePathParts ppath, , Spath, , , , Parse_Path
    
    'now, count the slashes. remove trailing slash if present.
    
    If Right$(Spath, 1) = "\" Then Spath = Mid$(Spath, 1, Len(Spath) - 1)
    GetPathDepth = Len(Spath) - Len(Replace$(Spath, "\", "")) + 2



End Function
Public Sub unittest()

    Dim str() As String
    ReDim str(0 To 1)
    str(0) = "CDRID.JPG"
    str(1) = "CDRID.gif"
    ShowExplorerMenuMulti FrmDebug.hWnd, "C:\", str()



End Sub
Public Sub TestVolumes()
    Dim currvolume As CVolume
    Dim vols As Volumes
    Set vols = FileSystem.Getvolumes
    Set currvolume = vols.GetNext
    Do
        If currvolume.IsReady Then
        Debug.Print "VOL:" & currvolume.RootFolder.Path
        End If
        
        Set currvolume = vols.GetNext
    
    Loop Until currvolume Is Nothing



End Sub
Public Sub testFiledialog()

    Dim grabfile As CFile
    Dim useopen As CFileDialog
    Set useopen = New CFileDialog
    Set grabfile = useopen.GetFileDirect(FrmDebug.hWnd, , OFN_EXPLORER + OFN_DONTADDTORECENT + OFN_ENABLEHOOK)
    


End Sub
Public Function FindNextFile(ByVal hfile As Long, Win32Data As WIN32_FIND_DATA) As Long
'Static wStruct As WIN32_FIND_DATAW
'Debug.Print "FindNextFile"
If MakeWideCalls Then
    ZeroMemory Win32Data, Len(Win32Data)
    'Debug.Print "Calling:"
    FindNextFile = FindNextFileW(hfile, Win32Data)
    
    'CopyMemory Win32Data, wStruct, Len(Win32Data)
   
    'LSet Win32Data = wStruct
    'BUG: the two buffers run into each other for files with short names...
       Win32Data.cFileName = StrConv(Win32Data.cFileName, vbFromUnicode)
    Win32Data.cFileName = Left$(Win32Data.cFileName, InStr(Win32Data.cFileName, vbNullChar))
    'Debug.Print "FindNextFile Found " & Trim$(Win32Data.cFileName)
Else

    FindNextFile = FindNextFileA(hfile, Win32Data)



End If




End Function

Public Function FindFirstFile(ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
'Dim wStruct As WIN32_FIND_DATAW
'Debug.Print "FindFirstFile..."
Debug.Print "FindFirstFile"
If MakeWideCalls Then
    'ReDim wStruct.Buffer(1 To 16768 + 1) 'TODO:// special handling via //?/ or whatever that damn prefix is.
    If Not IsUNCPath(lpFileName) Then
        'If InStr(lpFileName, "*") = 0 Then
        lpFileName = "\\?\" & lpFileName
        'End If
    End If
    FindFirstFile = FindFirstFileW(StrPtr(lpFileName), lpFindFileData)
'\\?\
    'lpFindFileData.dwFileAttributes = wStruct.dwFileAttributes
  
    'LSet lpFindFileData = wStruct
    'plop on some null characters, since that will be what the caller expects.
    'lpFindFileData.cFileName = Replace$(wStruct.Buffer, " ", vbNullChar)
    lpFindFileData.cFileName = StrConv(lpFindFileData.cFileName, vbFromUnicode)
    lpFindFileData.cFileName = Left$(lpFindFileData.cFileName, InStr(lpFindFileData.cFileName, vbNullChar))
    
Else
    FindFirstFile = FindFirstFileA(lpFileName, lpFindFileData)

End If




End Function
Public Sub Testalternate(Openme As String)
'D:\outtext.txt
Dim hfile As Long, streams() As String, StrJoin As String, I As Long

Dim testA As CAlternateStreams
Dim testb  As CAlternateStreams
Set testA = GetAlternateStreamsByPath(Openme)
Set testb = ListStreams(Openme)
Stop
'hFile = CreateFile( m_sFile.GetBuffer(0),
'                        GENERIC_READ,
'                        FILE_SHARE_READ,
'                        NULL,
'                        OPEN_EXISTING,
'                        FILE_FLAG_BACKUP_SEMANTICS,
'                        NULL );
'hFile = CreateFile(Openme, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
'streams() = GetAlternateStreamsByPath(hFile)
'For I = 0 To UBound(streams)
'    strjoin = strjoin & vbCrLf & streams(I)
'Next I
'MsgBox "Streams in File """ & Openme & """:" & vbCrLf & strjoin
'CloseHandle hFile
End Sub
Public Function GetAlternateStreamsByPath(ByVal OfFileName As String) As CAlternateStreams

    Dim streamStruct As WIN32_STREAM_ID, pContext As Long
    Dim OfFile As Long
    Dim StreamsRet As CAlternateStreams, newstream As CAlternateStream
    Dim lowpart As Long, highpart As Long, lowseeked As Long, highseeked As Long
    
    Dim dwBytesToRead As Long, dwBytesRead As Long, lpbytes() As Byte
    Dim bresult As Boolean, badata(4096) As Byte, callresult As Long, SStream() As String, streamCount As Long
    Set StreamsRet = New CAlternateStreams
    
    OfFile = CreateFile(OfFileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If OfFile = -1 Then
        RaiseAPIError Err.LastDllError, "GetAlternateStreamByPath "
        'Exit Function
    End If
    StreamsRet.Owner = OfFileName
    'dwBytesToRead = Len(streamStruct)
    'ReDim badata(Len(streamStruct) * 2 + 2)
    'bresult = True
    Do
'    if(!BackupRead( hFile,
'                        baData,
'                        dwBytesToRead,
'                        &dwBytesRead,
'                        FALSE, // am I done?
'                        FALSE,
'                        &pContext ) )
'*** "streams" appear to be at the end of files, not the beginning where the header is...
'dwBytesToRead = Len(streamStruct)
'seek to the end of the file- this is the file header+the file length.
'file header is 21 bytes(?)...
    'lowpart = GetFileSize(ofFile, highpart)
    'lowpart = lowpart
    'callresult = BackupSeek(ofFile, lowpart, highpart, lowseeked, highseeked, pContext)
    dwBytesToRead = Len(streamStruct)
    dwBytesRead = 0
        callresult = BackupRead(OfFile, badata(0), dwBytesToRead, dwBytesRead, False, False, pContext)
            If Not CBool(callresult) Then
                bresult = False
            Else
                If dwBytesRead = 0 Then
                    'all done...
                    bresult = False
                
                Else
                
                
                    bresult = True
                    CopyMemory streamStruct, badata(0), Len(streamStruct)
                    'we read the stream header successfully.
                    '... now we need to Stream name:
                    If LI2DBL(streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh) > 0 Then
                        'size is not zero- read that data into a string...
                        ReDim Preserve SStream(streamCount)
                        ReDim lpbytes(0 To streamStruct.dwStreamNameSize)
                        'read name...
                        If streamStruct.dwStreamNameSize > 0 Then
                            BackupRead OfFile, lpbytes(0), streamStruct.dwStreamNameSize, dwBytesRead, False, False, pContext
                    
                            '****
                            SStream(streamCount) = lpbytes() 'grab the stream name, and store it in the return array.
                            '**NOTES**
                            SStream(streamCount) = Trim$(SStream(streamCount))
                            
                            Set newstream = New CAlternateStream
                            newstream.Init OfFileName, SStream(streamCount), streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh, streamStruct.dwStreamAttributes
                            'create a collection and Class, "AlternateStreams" and "CAlternateStreams" or something. Here we would initialize the new
                            'CAlternateStreams Object as appropriate; note that "CAlternateStream" should contain methods similar to those of a file-
                            'path would be the entire path to the stream, etc.
                            'It' might even make sense to encapsulate the same functionality into CFile, except that might make it messy....
                            StreamsRet.Add newstream
                            streamCount = streamCount + 1
                        End If
                
                    End If
                    'seek to the end of the stream...
                    BackupSeek OfFile, streamStruct.dwStreamSizeLow, streamStruct.dwStreamSizeHigh, lowseeked, highseeked, pContext
                 End If
                
            End If
        
    Loop While bresult
    'allow backupread to deallocate...
    BackupRead OfFile, 0, 0, 0, 1, 0, pContext
    '    GetAlternateStreamsByHandle = SStream
    Set GetAlternateStreamsByPath = StreamsRet
End Function
Public Sub testpathparts(pathtest As String)

Dim vol As String, pth As String, FName As String, Extension As String, StreamName As String

ParsePathParts pathtest, vol, pth, FName, Extension, StreamName

MsgBox "parts of " & pathtest & vbCrLf & _
    "Volume:" & vol & vbCrLf & _
    "Path:" & pth & vbCrLf & _
    "Filename:" & FName & vbCrLf & _
    "Extension:" & Extension & vbCrLf & _
    "StreamName:" & StreamName

End Sub
'Path Parsing Routines.
Public Function IsUNCPath(pathtest As String) As Boolean
    IsUNCPath = (Left$(pathtest, 2) = "//")
End Function

'Public Function ParseProtocolSpec(ByVal StrInput As String, Optional ByRef Protocol As String, Optional ByRef Site As String, _
'    Optional ByRef Path As String, Optional ByVal Filename As String)
'
''    "parses a protocol string/ URL. For example in:
'
''"http://www.google.ca/search?hl=en&client=firefox-a&channel=s&rls=org.mozilla%3Aen-US%3Aofficial&hs=r0G&q=Protocol+Parsing&btnG=Search&meta="
'
''the protocol would be "HTTP://"
'
''site would be "www.google.ca"
'
''file would be the rest.
'
'
'StrInput = Replace$(StrInput, "\", "/")
'Dim ColonFound As Long
'Dim nextSlash As Long
'
'ColonFound = InStr(StrInput, "://")
'
'Protocol = Mid$(StrInput, 1, ColonFound + 2)
'
''Site will be everything between ColonFound+3 and the Next "/"
'nextSlash = InStr(Len(Protocol) + 1, StrInput, "/", vbTextCompare)
'Site = Mid$(StrInput, ColonFound + 3, nextSlash - (ColonFound + 3))
'
'
'
'
'
'
'
'    End Function
Public Function MakeAbsolutePath(ByVal StrPath As String) As String

'An absolute path has the following:
'a Drive specification (or UNC share spec)
'a Path.
'a Filename.

'Step one: determine if we have a volume name. if there is a Colon as the second character
'or two slashes




End Function
Public Function ParsePathParts(ByVal StrInput As String, Optional ByRef Volume As String, Optional ByRef Path As String, _
Optional ByRef filename As String, Optional ByRef Extension As String, Optional ByRef StreamName As String, Optional ByVal ParseLevel As ParsePathPartsConstants = Parse_All)
    Dim flpathUNC As Boolean, Countslash As Long, CurrPos As Long
    Dim lastSlash As Long, IsProtocol As Boolean
    flpathUNC = IsUNCPath(StrInput)
    StrInput = Replace$(StrInput, "/", "\")
    'Parses a path specification into it's constituent parts.
    
    
    
    
    'Retrieve the volume.
    'Volume could be a drive letter:
    
    'C:\
    
    'or a UNC volume:
    
    '//servername/share/
    
    
    'also, could have a protocol:
    '<protocol>://
    
'    'if we find a Colon followed by two slashes, we will assume that everything before it is the protocol- and everything until the first slash is the "volume"
'    If InStr(1, StrInput, ":\\", vbTextCompare) <> 0 Then
'        'Call my parseProtocol routine...
'    End If
'
'
    
    
    'Note that SHARE is technically part of the drive specification and so should be part of the returned volume name.
    
    'Simple: if it is a UNC path, return the string up to the fourth slash, otherwise the first three characters comprise the volume portion of the path.
    CurrPos = 0
    Countslash = 0
    If flpathUNC Then
        Do
            CurrPos = InStr(CurrPos + 1, StrInput, "\", vbTextCompare)
            Countslash = Countslash + 1
            
        Loop Until Countslash >= 4
        
        Volume = Mid$(StrInput, 1, CurrPos)
    
    'ElseIf IsProtocol Then
    'a Protocol-
    
    
    
    
    Else 'neither a protocol or a UNC path.
        'first three...
        Volume = Mid$(StrInput, 1, 3)
        CurrPos = 3
    End If
    'ADDED: logic to save time when caller doesn't want certain values.
    If ParseLevel >= Parse_Volume Then
        
        'Path... starts at currpos, lasts until last slash.
        Path = Mid$(StrInput, Len(Volume) + 1, InStrRev(StrInput, "\") - Len(Volume))
        
        
        If ParseLevel >= Parse_Path Then
            
            Dim LastColon As Long
            Dim lastDot As Long
            'Filename and stream
            lastSlash = InStrRev(StrInput, "\")
            lastDot = InStrRev(StrInput, ".")
            If lastDot = 0 Then
                filename = Mid$(StrInput, lastSlash + 1)
                Extension = ""
                StreamName = ""
            Else
                LastColon = InStrRev(StrInput, ":")
                If LastColon > lastSlash Then
                    filename = Mid$(StrInput, lastSlash + 1, lastDot - lastSlash - 1)
                    
                    Extension = Mid$(StrInput, lastDot + 1, LastColon - lastDot - 1)
                    
                    StreamName = Mid$(StrInput, LastColon + 1)
                Else
                    filename = Mid$(StrInput, lastSlash + 1, lastDot - lastSlash - 1)
                    
                    Extension = Mid$(StrInput, lastDot + 1)
                End If
                
            
            
                
            End If
        Else
            'ParseLevel >=ParsePath
        
        End If
    Else
        'Parselevel>=volume...
    End If
    'if we were a UNC path, replace back...
    If flpathUNC Then
     Path = Replace$(Path, "\", "/")
     Volume = Replace$(Volume, "\", "/")
    
    End If

End Function

Public Function SumAscii(ByVal Strsum As String)
    Dim I As Long
    Dim Currsum As Double
    For I = 1 To Len(Strsum)
        Currsum = Currsum + AscW(Mid$(Strsum, I, 1))
        Currsum = Currsum Mod 32768
    Next I
    SumAscii = Currsum



End Function
'FileChangeNotify Thread routine.
Public Function ThreadProcNotify(ByVal lpParameter As Long) As Long

Dim mCopyTo As CFileChangeNotify, ret As Long
'Set mCopyTo = New CFileChangeNotify
'four bytes for object.
CopyMemory mCopyTo, lpParameter, 4
'can't call library functions here. tssk tssk to me.
'just waitforsingleObject our parameter- and set the flag and get the F out of here.
ret = WaitForSingleObject(mCopyTo.ThreadhEvent, &HFFFFFF)


Call ZeroMemory(mCopyTo, 4)

End Function
'DWORD WINAPI ThreadProc(
'  __in  LPVOID lpParameter
');

Public Sub unittest2()
Dim ffile As CFile, fstream As FileStream
Dim ads As CAlternateStream
Set ffile = FileSystem.CreateFile("D:\blob5.bin")
Set ads = ffile.AlternateStreams(False).CreateStream("testit")
    Set fstream = ads.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_ALWAYS, 0)
    fstream.WriteString "THIS IS A STRING", StrRead_ANSI
    fstream.Flush
    fstream.CloseStream
    FileSystem.GetFile("D:\blob3.bin").AlternateStreams(True).Count


End Sub
Public Function JoinStr(StrJoin() As String, Optional ByVal Delimiter As String = ",") As String
    Dim builder As cStringBuilder
    Dim currItem As Long
    Set builder = New cStringBuilder
    On Error GoTo JoinFail
    For currItem = LBound(StrJoin) To UBound(StrJoin)
    
        builder.Append StrJoin(currItem)
        If currItem < UBound(StrJoin) Then
            builder.Append Delimiter
        End If
    Next currItem

    JoinStr = builder.ToString
    Set builder = Nothing
    Exit Function
JoinFail:
    JoinStr = ""
End Function
Public Sub testoperation()
    Dim srcfiles() As String
    Dim destfiles() As String
    ReDim srcfiles(1 To 2)
    ReDim destfiles(1 To 1)
    srcfiles(1) = "D:\XPtray.jpg"
    srcfiles(2) = "D:\VistaTray.jpg"
    destfiles(1) = "D:\VBProj\"
    
    FileOperation srcfiles, destfiles, FO_COPY, FOF_ALLOWUNDO + FOF_CONFIRMMOUSE, 0, "progress."

End Sub
Public Function OpenWith(ByVal HwndOwner As String, ByVal FilePath As String, Optional ByVal Icon As Long = 0) As Long

    Dim sei As SHELLEXECUTEINFOA
    Dim verba As String
    verba = "openas"
    sei.hIcon = Icon
    sei.hProcess = GetCurrentProcess()
    sei.hInstApp = App.hInstance
    
    sei.lpVerb = verba
    sei.lpFile = FilePath
    sei.nShow = vbNormalFocus
    sei.cbSize = Len(sei)
    
    OpenWith = ShellExecuteEx(sei)


    'Stop

End Function
Public Function FileOperation(SourceFiles() As String, destfiles() As String, SHop As olelib.FILEOP, flags As olelib.FILEOP_FLAGS, Optional ByVal OwnerWnd As Long = 0, Optional ByVal ProgressTitle As String = "")

Dim fbuf() As Byte
Dim Foperation As SHFILEOPSTRUCT
Dim srcStr As String, destStr As String
Dim ret As Long
'first, create null-delimited lists for sourcefiles() and destFiles()...

srcStr = JoinStr(SourceFiles(), vbNullChar) & vbNullChar & vbNullChar
destStr = JoinStr(destfiles(), vbNullChar) & vbNullChar & vbNullChar

'alright- we have source and dest...
With Foperation
    .pFrom = srcStr
    .pTo = destStr
    .hWnd = OwnerWnd
    .wFunc = SHop
    .fFlags = flags
End With

ReDim fbuf(1 To Len(Foperation) + 2)
        ' Now we need to copy the structure into a byte array
    'Call CopyMemory(fbuf(1), Foperation, Len(Foperation))

            ' Next we move the last 12 bytes by 2 to byte align the data
    'Call CopyMemory(fbuf(19), fbuf(21), 12)
    'last- call routine...
    
    
    ret = SHFileOperation(Foperation)
    'copy last 12 bytes back...
    'CopyMemory fbuf(21), fbuf(19), 12
    'copy into structure...
    'CopyMemory Foperation, fbuf(1), UBound(fbuf)
    Stop
    If ret <> 0 Then
        RaiseAPIError Err.LastDllError, "MdlFileSystem::FileOperation"
        
    Else
        '
    
    End If
'            If result <> 0 Then  ' Operation failed
'               MsgBox Err.LastDllError 'Show the error returned from
'                                       'the API.
'               Else
'               If FILEOP.fAnyOperationsAborted <> 0 Then
'                  MsgBox "Operation Failed"
'               End If
'            End If




End Function


'more utility functions- mostly dealing with Paths and Shell path handling routines.

Public Function MakePathAbsolute(ByVal RelativePath As String, Optional ByVal RelativeTo As String = "") As String
'
    Dim isunc As Boolean
   If IsUNCPath(RelativePath) Then isunc = True
    Dim BuildArray() As String
    Dim arraySize As Long
    Dim returnString As String, getvol As String
   Dim splPathParts() As String
   Dim I As Long
   
   If Mid$(RelativePath, 2, 1) = ":" Or _
    Left$(RelativePath, 2) = "//" Then
    'it's a volume name, so return the relative string unabated.
    getvol = MdlFileSystem.GetVolume(RelativePath)
    MakePathAbsolute = MakePathAbsolute(Mid$(RelativePath, Len(getvol)), getvol)
    Exit Function
  End If
   
   
    RelativePath = FixPath(RelativePath)
   ' RelativeTo = Replace$(RelativeTo, "/", "\")
   'so- the inevitable question is what needs to be done?
   'First, parse the path parts of both given paths. the RelativePath might have a filename/stream info, so we need to preserve that.
   
   'first, split our relativepath into it's specific path portions:
   ' relative = Split(RelativePath, "\")
   splPathParts = Split(RelativePath, "\")
   
   
   'we now have two arrays.
   Dim numseps As Long
   
   BuildArray = Split(RelativeTo, "\")
   arraySize = UBound(BuildArray)
   If BuildArray(arraySize) = "" Then
   arraySize = arraySize - 1
    ReDim Preserve BuildArray(arraySize)
    
    End If
   For I = 0 To UBound(splPathParts)
    If StrComp(splPathParts(I), "..") = 0 Then
        'remove the last item from Buildarray...
        If arraySize > 1 Then
            arraySize = arraySize - 1
            ReDim Preserve BuildArray(arraySize)
        End If
    ElseIf StrComp(splPathParts(I), ".") = 0 Then
        'no change...
    ElseIf splPathParts(I) = "" Then
    Else
        'add this to Buildarray...
        arraySize = arraySize + 1
        ReDim Preserve BuildArray(arraySize)
        BuildArray(arraySize) = splPathParts(I)
    
    End If
   
   
   
   
   Next I
   
   
    'now- reiterate and rebuild a path from buildarray().
    For I = 0 To UBound(BuildArray)
        returnString = returnString & BuildArray(I)
        If I < UBound(BuildArray) Then
        returnString = returnString & "\"
        End If
    
    Next I
    MakePathAbsolute = returnString


End Function

Public Sub CopyEntireStream(SourceStream As IInputStream, destStream As IOutputStream, Optional ByVal Chunksize As Long = 32768)

    If SourceStream.size < Chunksize Then
    Debug.Print "setting chunk size to " & SourceStream.size
        Chunksize = SourceStream.size
    End If
    Do Until SourceStream.EOF
        destStream.WriteBytes SourceStream.ReadBytes(Chunksize)

    '
    Loop

End Sub
Public Sub testunit()
'Unit test.

    Dim persistit As Object
    Dim fstream As FileStream
    Set persistit = CreateObject("Test.Persistor")
    
    Set fstream = FileSystem.CreateFile("D:\persistence2.dat").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_WRITE, OPEN_EXISTING)
    fstream.WriteObject persistit
    
    fstream.CloseStream
    Set persistit = Nothing
    
    Set fstream = FileSystem.GetFile("D:\persistence2.dat").OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING)
    Set persistit = fstream.ReadObject



End Sub

'Public Sub pleasenocrash()



'End Sub
Private Function Modulus(ByVal Dividend As Variant, ByVal Divisor) As Variant
    'returns the modulus of the two numbers.
    'The modulus is the remainder after dividing Divisor and dividend.
    Dim Quotient As Variant
    'It will, of course, be less then divisor.
    'IE:
    '10 mod 3.3 should be
    Quotient = CDec(Dividend / Divisor)
    'the floating point portion will be
    'the percentage of the divisor that fit at the end.
    Modulus = CDec(Quotient - Int(Quotient)) * CDec(Divisor)
    




End Function



Public Sub testcopy()
    Dim testin As FileStream, testout As FileStream
    
    Set testin = FileSystem.GetFile("D:\nsmrec3_converted.mp4").OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_WRITE, OPEN_EXISTING)
    Set testout = FileSystem.CreateFile("D:\nsm2.mp4").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_WRITE, OPEN_EXISTING)

    StreamCopy testin, testout, Nothing
    testout.CloseStream
End Sub
Public Function StreamCopy(inputstream As IInputStream, outputstream As IOutputStream, Optional Callback As ICopyMoveCallback) As Long

    Dim Chunksize As Long
    Dim CurrChunk As Long, TotalChunks As Long, ChunKRemainder As Long
    Dim inputsize As Double
    Dim thisChunkSize As Double
    Dim BytesTransfer() As Byte
    Chunksize = 32& * 1024&
    If Not inputstream.Valid Then
        Err.Raise 5, "MdlFileSystem::StreamCopy", "Passed InputStream Not Valid."
    ElseIf outputstream.Valid Then
        Err.Raise 5, "MdlFileSystem::StreamCopy", "Passed OutputStream not valid."
    
    End If
    If Callback Is Nothing Then Set Callback = New ICopyMoveCallback
    Callback.InitCopy inputstream, outputstream, Chunksize
    
    'Alright- calculate the number of chunks and the remainder chunk...
    inputsize = inputstream.size
    TotalChunks = Fix(inputsize / Chunksize)
    
    ChunKRemainder = Round(Modulus(inputsize, Chunksize), 0)
    
    'Loop from currchunk to TotalChunks+1....
    thisChunkSize = Chunksize
    For CurrChunk = 1 To TotalChunks + 1
        If CurrChunk > TotalChunks Then thisChunkSize = ChunKRemainder
        BytesTransfer = inputstream.ReadBytes(thisChunkSize)
        outputstream.WriteBytes BytesTransfer
        Callback.StreamProgress inputstream, outputstream, Chunksize, CurrChunk, CLng(inputsize)
        
        
    
    
        
    
    Next CurrChunk
    
    
    




End Function




Public Sub TestNetwork()
Call FileSystem.CreateFile("//Satellite/Shared/File.txt").OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, OPEN_EXISTING).WriteString(Space$(65536))

End Sub
Public Sub testcompressor()
Dim lzwin As FileStream, lzwout As FileStream
Dim filter As IStreamFilter, newfilt As CCoreFilters
Set lzwin = FileSystem.OpenStream("D:\testfiles\test8.txt")
Set lzwout = FileSystem.CreateStream("D:\huffmanout2.txt")


Set newfilt = New CCoreFilters
newfilt.FilterType = Huffman_Compress
Set filter = newfilt
filter.FilterStream lzwin, lzwout

lzwin.CloseStream
lzwout.CloseStream
Set newfilt = New CCoreFilters
newfilt.FilterType = Huffman_Expand
Set filter = newfilt

Set lzwin = FileSystem.OpenStream("D:\huffmanout2.txt")
Set lzwout = FileSystem.CreateStream("D:\huffman_exp2.txt")
filter.FilterStream lzwin, lzwout
lzwin.CloseStream
lzwout.CloseStream
End Sub



Public Sub test()
Dim sects() As String, scount As Long
'    mreg.ClassKey = HHKEY_USERS
'    mreg.Machine = "SATELLITE"
'    mreg.SectionKey = ""
'    mreg.EnumerateSections sects(), scount
'    Stop
'  mreg.ClassKey = HHKEY_USERS
    mreg.Machine = "SATELLITE"
'    mreg.SectionKey = ".DEFAULT\Control Panel\Colors"
'    'mreg.EnumerateSections sects(), scount
'    mreg.EnumerateValues sects(), scount
'    Stop
mreg.Classkey = HHKEY_USERS
mreg.SectionKey = ".DEFAULT\marksie"
mreg.ValueKey = "ARCIE"
mreg.CreateKey
End Sub
Public Function GetCountStr(ByVal StrCountIn As String, ByVal FindStr As String, Optional Comparemode As VbCompareMethod = vbBinaryCompare) As Long

    Dim CurrPos As Long, countof As Long
    CurrPos = 1
    countof = 1
    Do
        CurrPos = InStr(CurrPos + 1, StrCountIn, FindStr, Comparemode)
        If CurrPos > 0 Then
            countof = countof + 1
        Else
            If countof = 1 Then
                GetCountStr = 0
                Exit Function
            End If
        End If
    
    
    Loop Until CurrPos = 0
    
    
    
GetCountStr = countof



End Function
Public Function GetReducedPath(ByVal StrPath As String, ByVal TargetLength As Long) As String





'
'Dim parsedFileName As String, ParsedExtension As String, ParsedStream As String
'Dim ParsedPath As String
'Dim parsedvolume As String
'Dim filename As String
'Dim workString As String
'Dim PrevPath As String
'Dim SplitPath() As String
If Len(StrPath) <= TargetLength Then
    GetReducedPath = StrPath
    Exit Function
Else
    GetReducedPath = Left$(StrPath, TargetLength \ 2) & "..." & Right$(StrPath, TargetLength \ 2)


End If




'Dim I As Long, Numiterate As Long
'ParsePathParts StrPath, parsedvolume, ParsedPath, parsedFileName, ParsedExtension, ParsedStream, Parse_All
'If Right$(ParsedPath, 1) = "\" Then ParsedPath = Mid$(ParsedPath, 1, Len(ParsedPath) - 1)
'SplitPath = Split(ParsedPath, "\")
'If ParsedExtension <> "" Or parsedFileName <> "" Then
'    filename = parsedFileName & "." & ParsedExtension
'End If
'workString = parsedvolume
'Dim strStart As String, StrEnd As String
'strStart = parsedvolume
'StrEnd = SplitPath(UBound(SplitPath)) & "\" & filename
'
'For I = 0 To UBound(SplitPath)
'    workString = workString & SplitPath(I) & "\"
'
'Next I




'GetReducedPath = PrevPath

End Function

Public Sub TestPackage()
    Dim testpack As CSimpleFilePackage
    Const testpackfile As String = "D:\test.pakker"
    Set testpack = New CSimpleFilePackage
    testpack.AddFileToPackage "C:\usage.log"
    testpack.AddFileToPackage "C:\install.ini"
    testpack.AddFileToPackage "D:\image2.jpg"
    
    
    'testpack.WritePackage testpackfile
    testpack.WritePackage_Huffman testpackfile
    Set testpack = Nothing
    Set testpack = New CSimpleFilePackage
    testpack.ReadPackage_Huffman testpackfile
    Dim testextraction As FileStream
    Set testextraction = FileSystem.CreateStream("D:\testextract5.dat")
    testpack.ExtractFile "install.ini", testextraction
    testextraction.CloseStream
End Sub
