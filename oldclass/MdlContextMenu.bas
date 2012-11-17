Attribute VB_Name = "MdlContextMenu"


' General purpose shell definitons found in Shlobj.h


' shell32.dll string resource IDs, common to all file versions
Public Const IDS_SHELL32_EXPLORE = 8502     ' < Win2K: "&Explore", > Win2K: "E&xplore"
Public Const IDS_SHELL32_NAME = 8976           ' "Name"
Public Const IDS_SHELL32_SIZE = 8978             ' "Size"
Public Const IDS_SHELL32_TYPE = 8979            ' "Type"
Public Const IDS_SHELL32_MODIFIED = 8980    ' "Modified"

Public Const S_OK = 0           ' indicates success
Public Const S_FALSE = 1&   ' special HRESULT value

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example,
' if the location specified by the pidl parameter is not part of the file system.
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As Any) As Long


' ==============================================================
' SHBrowseForFolder

Public Type BROWSEINFO
  HwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String ' Return display name of item selected.
  lpszTitle As String              ' text to go in the banner over the tree.
  ulFlags As Long                 ' Flags that control the return stuff
  lpfn As Long
  lParam As Long      ' extra info that's passed back in callbacks
  iImage As Long      ' output var: where to return the Image index.
End Type

Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

' typedef int (CALLBACK* BFFCALLBACK)(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) as long

Public Enum BF_Flags
' Browsing for directory.
  BIF_RETURNONLYFSDIRS = &H1      ' For finding a folder to start document searching
  BIF_DONTGOBELOWDOMAIN = &H2     ' For starting the Find Computer
  ' Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
  ' this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
  ' rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
  ' all three lines of text.
  BIF_STATUSTEXT = &H4
  BIF_RETURNFSANCESTORS = &H8

#If (WIN32_IE >= &H400) Then
  BIF_EDITBOX = &H10               ' Add an editbox to the dialog.  Always on with BIF_USENEWUI
  BIF_VALIDATE = &H20              ' insist on valid result (or CANCEL)
  BIF_USENEWUI = &H40              ' Use the new dialog layout with the ability to resize.
#End If  ' // WIN32_IE >= &H400

  BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
  BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
  BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything
End Enum


' message from browser
Public Enum BFFM_FromDlg
  BFFM_INITIALIZED = 1
  BFFM_SELCHANGED = 2

#If (WIN32_IE >= &H400) Then
' If the user types an invalid name into the edit box, the browse dialog will call the
' application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
' This flag is ignored if BIF_EDITBOX is not specified.
  BFFM_VALIDATEFAILEDA = 3     ' lParam:szPath ret:1(cont),0(EndDialog)
  BFFM_VALIDATEFAILEDW = 4     ' lParam:wzPath ret:1(cont),0(EndDialog)
#End If  ' // WIN32_IE >= &H400
End Enum



' messages to browser
Private Const WM_USER = &H400
Public Enum BFFM_ToDlg
  BFFM_SETSTATUSTEXTA = (WM_USER + 100)
  BFFM_ENABLEOK = (WM_USER + 101)
  BFFM_SETSELECTIONA = (WM_USER + 102)
  BFFM_SETSELECTIONW = (WM_USER + 103)
  BFFM_SETSTATUSTEXTW = (WM_USER + 104)
End Enum

' ==============================================================
' SHGetFileInfo

Public Const MAX_PATH = 260

Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare Sub prvIIDFromString Lib "ole32.dll" (ByVal lpsz As String, lpiid As olelib.UUID)
' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As Any, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Enum SHGFI_flags
    SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
    SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
    SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
    SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
    SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
    ' Indicates that the function should not attempt to access the file specified by pszPath.
    ' Rather, it should act as if the file specified by pszPath exists with the file attributes
    ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
    ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
    SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
    SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
    SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
    SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
    SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
    SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
    SHGFI_OVERLAYINDEX = &H40
    
    SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
    SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
    SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
    SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
    SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
'



'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
'
' ==============================================================
' A fairly comprehensive wrapping of the IShellFolder and IEnumIDList interfaces with
' some IUnknown thrown in. Also will do about anything that can be done with a pidl...
'
' Note that "IShellFolder Extended Type Library v1.1" (ISHF_Ex.tlb) included with this
' project, must be present and correctly registered on your system, and referenced by
' this project to allow use of these interfaces.
' ==============================================================
'
' Procedure responsibility of pidl memory, unless specified otherwise:
' - Calling procedures are solely responsible for freeing pidls they create,
'   or receive as a return value from a called procedure.
' - Called procedures always copy pidls received in their params, and
'   *never* free pidl params.

' Global IContextMenu2 interface variable filled in ShowShellContextMenu on
' treeview and listview item right click. Used for menu messages in FrmWndProc.
Public ICtxMenu2 As IContextMenu2
' defined in mWindowDefs
'Private Const WM_USER = &H400
' Retrieves a pointer to the shell's IMalloc interface.
' Returns NOERROR if successful or or E_FAIL otherwise.
Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long

' Retrieves the IShellFolder interface for the desktop folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As olelib.IShellFolder) As Long
'Private Declare Sub SHGetDesktopFolder Lib "shell32.dll" (ByRef ppshf As Long)


' Frees memory allocated by the shell
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' GetItemID item ID retrieval constants
Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1
'
'  ' ====================================================
'  ' item ID (pidl) structs, just for reference
'  '
'  ' item identifier (relative pidl), allocated by the shell
'  Public Type SHITEMID
'    cb As Integer        ' size of struct, including cb itself
'    abID() As Byte    ' variable length item identifier
'  End Type
'
'  ' fully qualified pidl
'  Public Type ITEMIDLIST
'    mkid As SHITEMID  ' list of item identifers, packed into SHITEMID.abID
'  End Type
'




'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font

' ============================================================================
' common control definitions

Public Const NM_FIRST = -0&   ' (0U-  0U)       ' // generic to all controls
Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_RETURN = (NM_FIRST - 4)
Public Const NM_RCLICK = (NM_FIRST - 5)

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

' Callback constants

' TV/LV_ITEM.pszText
Public Const LPSTR_TEXTCALLBACK = (-1)

' TVITEM.iImage/iSelectedImage, LVITEM.iImage
Public Const I_IMAGECALLBACK = (-1)

' OCM_NOTIFY is WM_NOTIFY reflected to a C++ created ActiveX control.
' http://msdn.microsoft.com/library/devprods/vs6/visualc/vccore/_core_activex_controls.3a_.subclassing_a_windows_control.htm
Public Const WM_NOTIFY = &H4E
Public Const OCM__BASE = (WM_USER + &H1C00)
Public Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)

' ============================================================================
' window messages

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long   ' <---

Public Const WM_DESTROY = &H2
Public Const WM_CANCELMODE = &H1F
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_INITMENUPOPUP = &H117

' ============================================================
' imagelist definitions

Declare Function FileIconInit Lib "shell32.dll" Alias "#660" (ByVal cmd As Boolean) As Boolean

' transparent color (the imagelist will use each icon's mask)
Public Const CLR_NONE = &HFFFFFFFF
Declare Function ImageList_SetBkColor Lib "COMCTL32.DLL" (ByVal himl As Long, ByVal clrBk As Long) As Long
Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal himl As Long) As Long

' ============================================================================
' general window definitions

Public Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

Public Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Declare Function GetFocus Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As CBoolean) As CBoolean

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' ============================================================================
' string/kernel32 definitions

' Converts a Unicode str to a ANSII str.
' Specify -1 for cchWideChar and 0 for cchMultiByte to rtn str len.
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwflags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
' CodePage
Public Const CP_ACP = 0        ' ANSI code page
Public Const CP_OEMCP = 1   ' OEM code page

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwflags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
' dwFlags
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

' dwLanguageId
Public Const LANG_USER_DEFAULT = &H400&

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

' Loads a string resource from the executable file associated with a specified module
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Declare Function lstrcmpiA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function lstrcmpiW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)

' =================================================================
' FindFirstFile definitions

'Public Const MAX_PATH = 260

Public Type FILETIME   ' ft
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA   ' wfd
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * 500
  'cAlternateFileName As String * 14
End Type

' nFileSizeHigh: Specifies the high-order DWORD value of the file size, in bytes.
' This value is zero unless the file size is greater than MAXDWORD. The size of
' the file is equal to (nFileSizeHigh * MAXDWORD) + nFileSizeLow.
Public Const MAXDWORD = (2 ^ 32) - 1   ' 0xFFFFFFFF

'Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean

'FindFirstFile error rtn value
Public Const INVALID_HANDLE_VALUE = -1

' =================================================================
' file/time definitions

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Private Const CSIDL_DESKTOP As Long = &H0

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwflags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwflags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

' Local IDs
Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400

' Date Flag for GetDateFormat, Time Flag for GetTimeFormat
Public Const LOCALE_NOUSEROVERRIDE = &H80000000    ' do not use user overrides

' Date Flags for GetDateFormat
Public Const DATE_SHORTDATE = &H1                  ' use short date picture
Public Const DATE_LONGDATE = &H2                     ' use long date picture
Public Const DATE_USE_ALT_CALENDAR = &H4   ' use alternate calendar (if any)

' Time Flags for GetTimeFormat
Public Const TIME_NOMINUTESORSECONDS = &H1  ' do not use minutes or seconds
Public Const TIME_NOSECONDS = &H2                        ' do not use seconds
Public Const TIME_NOTIMEMARKER = &H4                 ' do not use time marker, i.e AM/PM
Public Const TIME_FORCE24HOURFORMAT = &H8     ' always use 24 hour format

' ============================================================================
' menu definitions
Private Declare Sub SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWnd As Long, ByVal csidl As Long, ByRef ppidl As Long)

Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As TPM_wFlags, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As Any) As Long
Public Enum TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40            ' Vert alignment matters more
  TPM_NONOTIFY = &H80           ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
End Enum
Public Declare Function lstrlenUPtr Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Public Declare Function lstrlenAPtr Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function ShellExecuteEx Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFO) As Long
Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Public Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As Long   ' String
  lpFile As Long   ' String
  lpParameters As Long   ' String
  lpDirectory As Long   ' String
  nShow As Long
  hInstApp As Long
  '  Optional fields
  lpIDList As Long
  lpClass As Long   ' String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

Public Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, lpiid As Any) As Long
'
'Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByRef pv As Any)
Private Declare Function CopyMemoryLPToStr Lib "kernel32.dll" Alias "rtlMoveMemory" (pszDest As String, pvSrc As Long, cbCopy As Long) As Long
' SHELLEXECUTEINFO fMask
Public Const SEE_MASK_INVOKEIDLIST = &HC

' SHELLEXECUTEINFO nShow
Public Const SW_SHOWNORMAL = 1
'


Function UPointerToString(p As Long) As String
    Dim c As Long
    ' Get length of Unicode string to first null
    c = lstrlenUPtr(p)
    ' Allocate a string of that length
    UPointerToString = String$(c, 0)
    ' Copy the pointer data to the string
    CopyMemory ByVal StrPtr(UPointerToString), ByVal p, c * 2
End Function

Function APointerToString(p As Long) As String
    Dim c As Long
    ' Get length of Unicode string to first null
    c = lstrlenAPtr(p)
    ' Allocate a string of that length
    APointerToString = String$(c, 0)
    ' Copy the pointer data to the string
    CopyMemoryLPToStr APointerToString, ByVal p, c
End Function



'Function IIDToString(iid As olelib.UUID) As String
'    Dim pStr As Long
'    ' Allocate a string, fill it with GUID, and return a pointer to it
'    StringFromIID iid, pStr
'    ' Copy characters from pointer to return string
'    IIDToString = PointerToString(pStr)
'    ' Free the allocated string
'    CoTaskMemFree pStr
'End Function

' ==============================================================
' SHGetFileInfo calls

' If successful returns the specified file's typename,
' returns an empty string otherwise.
'   pidl  - file's absolute pidl

Public Function GetFileTypeNamePIDL(pidl As Long) As String
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_TYPENAME) Then
    GetFileTypeNamePIDL = GetStrFromBufferA(sfi.szTypeName)
  End If
End Function

' Returns a file's small or large icon index within the system imagelist.
'   pidl       - file's absolute pidl
'   uType  - either SHGFI_SMALLICON or SHGFI_LARGEICON, and SHGFI_OPENICON

Public Function GetFileIconIndexPIDL(pidl As Long, uType As Long) As Long
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_SYSICONINDEX Or uType) Then
    GetFileIconIndexPIDL = sfi.iIcon
  End If
End Function

' Returns the handle of the small or large icon system imagelist.
'   uSize - either SHGFI_SMALLICON or SHGFI_LARGEICON

Public Function GetSystemImagelist(uSize As Long) As Long
  Dim sfi As SHFILEINFO
  ' Any valid file system path can be used to retrieve system image list handles.
  GetSystemImagelist = SHGetFileInfo("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uSize)
End Function

' ==============================================================
' SHBrowseForFolder

Public Function BrowseDialog(hWnd As Long, _
                                                sPrompt As String, _
                                                ulFlags As BF_Flags, _
                                                Optional pidlRoot As Long = 0, _
                                                Optional pidlPreSel As Long = 0) As Long
  Dim bi As BROWSEINFO
  
  With bi
    .HwndOwner = hWnd
    .pidlRoot = pidlRoot
    .lpszTitle = sPrompt
    .ulFlags = ulFlags
    .lParam = pidlPreSel
    .lpfn = FARPROC(AddressOf BrowseCallbackProc)
  End With
  
  BrowseDialog = SHBrowseForFolder(bi)
  
End Function

'Public Function BrowseCallbackProc(ByVal Hwnd As Long, _
'                                                            ByVal uMsg As Long, _
'                                                            ByVal Lparam As Long, _
'                                                            ByVal lpdata As Long) As Long
''  Dim sPath As String * MAX_PATH
'
'  Select Case uMsg
'
'    Case BFFM_INITIALIZED
'      ' Set the dialog's pre-selected folder from the pidl we set
'      ' bi.lParam to above (passed in the lpData param).
'      Call SendMessage(Hwnd, BFFM_SETSELECTIONA, ByVal CFalse, ByVal lpdata)
'
''    Case BFFM_SELCHANGED
''      If SHGetPathFromIDList(lParam, sPath) Then
''        ' Return the path
''        Debug.Print Left$(sPath, InStr(sPath, vbNullChar) - 1)
''      End If
'
'  End Select
'
'End Function



' Returns the low 16-bit integer from a 32-bit long integer

Public Function loWord(dwValue As Long) As Integer
  MoveMemory loWord, dwValue, 2
End Function

Public Function loByte(wValue As Integer) As Byte
 MoveMemory loByte, wValue, 1

End Function
' Returns the low 16-bit integer from a 32-bit long integer
Public Function HiByte(wValue As Integer) As Byte
    MoveMemory HiByte, VarPtr(wValue) + 1, 1
End Function
Public Function hiWord(dwValue As Long) As Integer
  MoveMemory hiWord, ByVal VarPtr(dwValue) + 2, 2
End Function


' Returns the system-defined description of an API error code

Public Function GetAPIErrStr(dwErrCode As Long) As String
  Dim sErrDesc As String * 256   ' max string resource len
  Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS Or _
                                 FORMAT_MESSAGE_MAX_WIDTH_MASK, _
                                 ByVal 0&, dwErrCode, LANG_USER_DEFAULT, _
                                 ByVal sErrDesc, 256, 0)
  GetAPIErrStr = GetStrFromBufferA(sErrDesc)
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Returns an ANSII string from a pointer to an ANSII string.

Public Function GetStrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  GetStrFromPtrA = sRtn
End Function

' Returns an ANSI string from a pointer to a Unicode string.

Public Function GetStrFromPtrW(lpszW As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
'  sRtn = String$(WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, 0, 0, 0, 0), 0)
  Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
  GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function

' Fills a GUID

Public Sub DEFINE_GUID(name As olelib.UUID, l As Long, w1 As Integer, w2 As Integer, _
                                          b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, _
                                          b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With name
    .Data1 = l
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = b0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = b3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub

' Fills an OLE GUID, the Data4 member always is "C000-000000046"

Public Sub DEFINE_OLEGUID(name As olelib.UUID, l As Long, w1 As Integer, w2 As Integer)
  DEFINE_GUID name, l, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub

' Provides a generic test for success on any status value.
' Non-negative numbers indicate success.

' If we incur any error situation from any API or interface member
' function's call to this proc, we'll let the user know that sometime's
' not right. What happens when execution continues after the error
' is indeternimate. and could possibly lead to a GPF...

Public Function SUCCEEDED(hr As Long) As Boolean   ' hr = HRESULT
  If (hr >= S_OK) Then
    SUCCEEDED = True
  Else
    If IsIDE Then
      If (MsgBox("Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr) & vbCrLf & vbCrLf & _
                        "View offending code?", vbExclamation Or vbYesNo) = vbYes) Then Stop
      ' hit Ctrl+L to view the call stack...
    Else
      MsgBox "Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr), vbExclamation
    End If
  End If
End Function

Public Function IsIDE() As Boolean
  On Error GoTo out
  Debug.Print 1 / 0
out:
  IsIDE = Err
End Function

' A dummy procedure that receives and returns the result
' of the AddressOf operator

Public Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

' Returns the top level parent window from the specified window handle.

Public Function GetTopLevelParent(hWnd As Long) As Long
  Dim hWndParent As Long
  Dim hwndTmp As Long
  
  hWndParent = hWnd
  Do
    hwndTmp = GetParent(hWndParent)
    If hwndTmp Then hWndParent = hwndTmp
  Loop While hwndTmp

  GetTopLevelParent = hWndParent

End Function

' rtns date/time string as "m/d/yy h:m AM/PM"

Public Static Function GetFileDateTimeStr(ftFile As FILETIME) As String
  Dim ftLocal As FILETIME
  Dim st As SYSTEMTIME

  Call FileTimeToLocalFileTime(ftFile, ftLocal)
  Call FileTimeToSystemTime(ftLocal, st)
  GetFileDateTimeStr = GetFileDateStr(st) & " " & GetFileTimeStr(st)

End Function

Public Static Function GetFileDateStr(st As SYSTEMTIME) As String
  Dim sDate As String * 32
  Dim wLen As Integer
  
  wLen = GetDateFormat(LOCALE_USER_DEFAULT, _
                                        LOCALE_NOUSEROVERRIDE Or DATE_SHORTDATE, _
                                        st, vbNullString, sDate, 64)
  
  If wLen Then GetFileDateStr = Left$(sDate, wLen - 1)
  
End Function

Public Static Function GetFileTimeStr(st As SYSTEMTIME) As String
  Dim sTime As String * 32
  Dim wLen As Integer
  
  wLen = GetTimeFormat(LOCALE_USER_DEFAULT, _
                                        LOCALE_NOUSEROVERRIDE Or TIME_NOSECONDS, _
                                        st, vbNullString, sTime, 64)
  
  If wLen Then GetFileTimeStr = Left$(sTime, wLen - 1)
  
End Function

' Returns the string resource contained within the specifed module
' from the specified string resource ID.

Public Function GetResourceString(sModule As String, idString As Long) As String
  Dim hModule As Long
  Dim sBuf As String * MAX_PATH
  Dim nChars As Long
  
  hModule = LoadLibrary(sModule)
  If hModule Then
    nChars = LoadString(hModule, idString, sBuf, MAX_PATH)
    If nChars Then GetResourceString = Left$(sBuf, nChars)
    Call FreeLibrary(hModule)
  End If
  
End Function



' ================================================================
' interface procs

' Returns a reference to the IMalloc interface.


Public Function isMalloc() As IMalloc
  Static im As IMalloc
  If (im Is Nothing) Then Call SUCCEEDED(SHGetMalloc(im))
  Set isMalloc = im
End Function

' Returns a reference to the desktop folder's IShellFolder interface.

Public Function isfDesktop() As IShellFolder
  Static isf As IShellFolder
  If (isf Is Nothing) Then Call SUCCEEDED(SHGetDesktopFolder(isf))
  Set isfDesktop = isf
End Function

' Returns the IShellFolder interface ID, {000214E6-0000-0000-C000-000000046}

Public Function IID_IShellFolder() As olelib.UUID
  Static iid As olelib.UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E6, 0, 0)
  IID_IShellFolder = iid
End Function
Public Function IID_IShellFolder2() As olelib.UUID
    Static iid As olelib.UUID
    If iid.Data1 = 0 Then Call DEFINE_GUID(IID_IShellFolder2, &H93F2F68C, &H1D1B, &H11D3, &HA3, &HE, &H0, &HC0, &H4F, &H79, &HAB, &HD1)
    IID_IShellFolder2 = iid
    
    End Function



' Returns the IShellDetails interface ID, {000214EC-0000-0000-C000-000000000046}

Public Function IID_IShellDetails() As olelib.UUID
Static iid As olelib.UUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214EC, 0, 0)
  IID_IShellDetails = iid
End Function

' Rtns the reference count for the specified object...

Public Function ObjRefCount(obj As olelib.IUnknown) As Long
  On Error GoTo IsNothing
  
  ' AddRef the object, will err if Nothing, returning 0
  obj.AddRef
  
  ' Release the object and return the original reference count,
  ' less one reference for the proc local "obj" variable.
  ObjRefCount = obj.Release - 1

IsNothing:
End Function

' ================================================================
' pidl utility procs

' Determines if the specified pidl is the desktop folder's pidl.
' Returns True if the pidl is the desktop's pidl, returns False otherwise.

' The desktop pidl is only a single item ID whose value is 0 (the 2 byte
' zero-terminator, i.e. SHITEMID.abID is empty). Direct descendents of
' the desktop (My Computer, Network Neighborhood) are absolute pidls
' (relative to the desktop) also with a single item ID, but contain values
' (SHITEMID.abID > 0). Drive folders have 2 item IDs, children of drive
' folders have 3 item IDs, etc. All other single item ID pidls are relative to
' the shell folder in which they reside (just like a relative path).

Public Function IsDesktopPIDL(pidl As Long) As Boolean
  
  ' The GetItemIDSize() call will also return 0 if pidl = 0
  If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)

End Function

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  
  ' If we try to access memory at address 0 (NULL), then it's bye-bye...
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2

End Function

' Returns the count of item IDs in a pidl.

Public Function GetItemIDCount(ByVal pidl As Long) As Integer
  Dim nItems As Integer
  
  ' If the size of an item ID is 0, then it's the zero
  ' value terminating item ID at the end of the pidl.
  Do While GetItemIDSize(pidl)
    pidl = GetNextItemID(pidl)
    nItems = nItems + 1
  Loop
  
  GetItemIDCount = nItems

End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHITEMID.cb, 2 bytes
  
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb

End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most item IDs are roughly 20 bytes in size, and since an item ID represents
  ' a folder, a pidl can never exceed 260 folders, or 5200 bytes).
  On Error GoTo out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminating item ID
    GetPIDLSize = cb + 2
  End If
  
out:
End Function

' Copies and returns the specified item ID from a complex pidl
'   pidl -    pointer to an item ID list from which to copy
'   nItem - 1-based position in the pidl of the item ID to copy

' If successful, returns a new item ID (single-element pidl)
' from the specified element positon. Returns 0 on failure.
' If nItem exceeds the number of item IDs in the pidl,
' the last item ID is returned.

' (calling proc is responsible for freeing the new pidl)

Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
  Dim nCount As Integer
  Dim I As Integer
  Dim cb As Integer
  Dim pidlNew As Long
  
  nCount = GetItemIDCount(pidl)
  If (nItem > nCount) Or (nItem = GIID_LAST) Then nItem = nCount
  
  ' GetNextItemID returns the 2nd item ID
  For I = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
    
  ' Get the size of the specified item identifier.
  ' If cb = 0 (the zero terminator), the we'll return a desktop pidl, proceed
  cb = GetItemIDSize(pidl)
  
  ' Allocate a new item identifier list.
  pidlNew = isMalloc.alloc(cb + 2)
  If pidlNew Then
    
    ' Copy the specified item identifier.
    ' and append the zero terminator.
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    MoveMemory ByVal pidlNew + cb, 0, 2
    
    GetItemID = pidlNew
  End If
  
End Function

' Creates a new pidl of the given size

' (calling proc is responsible for freeing the new pidl)

Public Function CreatePIDL(cb As Long) As Long
  Dim pidl As Long
  
  pidl = isMalloc.alloc(cb)
  If pidl Then
    FillMemory ByVal pidl, cb, 0 ' initialize to zero, set by caller
    CreatePIDL = pidl
  End If

End Function

' Returns a copy of a relative or absolute pidl

' (calling proc is responsible for freeing the new pidl)

Public Function CopyPIDL(pidl As Long) As Long
  Dim cb As Long
  Dim pidlNew As Long
  
  cb = GetPIDLSize(pidl)
  If cb Then
    pidlNew = CreatePIDL(cb)
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    CopyPIDL = pidlNew
  End If

End Function

' Frees the specified pidl and zeros it

Public Sub FreePIDL(pidl As Long)
  On Error GoTo out
  
  ' Free the pidl and zero it's *value* only
  ' (not what it points to!, i.e. ZeroMemory = FE...)
  If pidl Then isMalloc.Free ByVal pidl

out:
  If Err And (pidl <> 0) Then
    Call CoTaskMemFree(pidl)
  End If
  
  pidl = 0
  
End Sub

' Copies and returns all but the last item ID from the specified absolute pidl.

'   pidl                - pointer to the pidl from which to copy
'   fFreeOldPidl  - optional flag specifying whether to free and zero the passed pidl

'    ' If successful, returns a new absolute pid (relative to the desktop)
'    ' If either a valid single item ID pidl is passed to this proc (either the
'    ' desktop's pidl or a relative pidl), or an invalid pidl is passed, 0 is returned.

' If successful, returns a new absolute pid (relative to the desktop)
' If either a valid single item ID pidl is passed to this proc (either the
' desktop's pidl or a relative pidl), or an invalid pidl is passed, the
' desktop's pidl is returned.

' (calling proc is responsible for freeing the new pidl)

Public Function GetPIDLParent(pidl As Long, _
                                                  Optional fReturnDesktop As Boolean = False, _
                                                  Optional fFreeOldPidl As Boolean = False) As Long
  Dim nCount As Integer
  Dim pidl1 As Long
  Dim I As Integer
  Dim cb As Integer
  Dim pidlNew As Long
  
  nCount = GetItemIDCount(pidl)
  If (nCount = 0) And (fReturnDesktop = False) Then Exit Function
  
  ' Get the size of all but the pidl's last item ID and zero terminator.
  ' (maintain the value of the original pidl, it's passed ByRef !!)
  pidl1 = pidl
  For I = 1 To nCount - 1
    cb = cb + GetItemIDSize(pidl1)
    pidl1 = GetNextItemID(pidl1)
  Next
  
  ' Allocate a new item ID list with a new terminating 2 bytes.
  pidlNew = isMalloc.alloc(cb + 2)
  
  ' If the memory was allocated...
  If pidlNew Then
    ' Copy all but the last item ID from the original pidl
    ' to the new pidl and zero the terminating 2 bytes.
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    FillMemory ByVal pidlNew + cb, 2, 0
    
    If fFreeOldPidl Then Call FreePIDL(pidl)
    GetPIDLParent = pidlNew
  
  End If
  
End Function

' Creates a new pidl by prepending pidl2 to pidl1 (i.e pidlNew = pidl1pidl2)

' (calling proc is responsible for freeing the new pidl, the
' two passed pidls are still valid and are not freed unless specified)

Public Function CombinePIDLs(pidl1 As Long, _
                                                  pidl2 As Long, _
                                                  Optional fFreePidl1 As Boolean = False, _
                                                  Optional fFreePidl2 As Boolean = False) As Long
  Dim cb1 As Integer
  Dim cb2 As Integer
  Dim pidlNew As Long

  ' If pidl1 is non-zero...
  If pidl1 Then
    ' Get it's size
    cb1 = GetPIDLSize(pidl1)
    ' If pidl1 is valid (has a size), subtract the size of the zero terminator
    If cb1 Then cb1 = cb1 - 2
  End If
  
  ' If pidl2 is non-zero...
  If pidl2 Then
    ' Get it's size
    cb2 = GetPIDLSize(pidl2)
    ' If pidl2 is valid (has a size), subtract the size of the zero terminator
    If cb2 Then cb2 = cb2 - 2
  End If

  ' Create a new pidl sized to hold both pidl1, pidl2 and the zero terminator
  pidlNew = CreatePIDL(cb1 + cb2 + 2)
  If (pidlNew) Then
    
    ' If pidl1 is valid, put it's id list at the beginning of our new pidl
    If cb1 Then MoveMemory ByVal pidlNew, ByVal pidl1, cb1
    
    ' If pidl2 is valid, prepend it's id list to the end of the new pidl
    If cb2 Then MoveMemory ByVal pidlNew + cb1, ByVal pidl2, cb2
      
    ' Zero the terminating 2 bytes
    FillMemory ByVal pidlNew + cb1 + cb2, 2, 0
      
    ' Finally, free the pidls as specified
    If (pidl1 And fFreePidl1) Then isMalloc.Free ByVal pidl1
    If (pidl2 And fFreePidl2) Then isMalloc.Free ByVal pidl2
    
  End If
  
  CombinePIDLs = pidlNew

End Function

' Returns an absolute pidl's path only (doesn't rtn display names!)

Public Function GetPathFromPIDL(pidl As Long) As String
  Dim Spath As String * MAX_PATH   ' 260
  If SHGetPathFromIDList(pidl, Spath) Then
    GetPathFromPIDL = GetStrFromBufferA(Spath)
  End If
End Function

' ================================================================
' IShellFolder procs

' Returns a shell item's displayname

'   isfParent - item's parent folder IShellFolder
'   pidlRel    - item's pidl, relative to isfParent
'   uFlags    - specifies the type of name to retrieve

Public Function GetFolderDisplayName(isfParent As olelib.IShellFolder, _
                                                              pidlRel As Long, _
                                                              uFlags As SHGNO_Flags) As String
  Dim lpStr As STRRET   ' struct filled
  
  Call isfParent.GetDisplayNameOf(pidlRel, uFlags, lpStr)
    GetFolderDisplayName = GetStrRet(lpStr, pidlRel)
  

End Function

' Returns information from the STRRET struct (identical to the new IE5 StrRetToStr API).

Public Function GetStrRet(lpStr As STRRET, pidlRel As Long) As String
  Dim lpsz As Long         ' string pointer
  Dim uOffset As Long    ' offset to the string pointer
  
  Select Case (lpStr.uType)
  
    ' The 1st UINT (Long) of the array points to a Unicode
    ' str which *should* be allocated & freed.
    Case STRRET_WSTR
      MoveMemory lpsz, lpStr.CStr(0), 4
      GetStrRet = GetStrFromPtrW(lpsz)
      Call CoTaskMemFree(lpsz)
    
    ' The 1st UINT (Long) of the array points to the location
    ' (uOffset bytes) to the ANSI str in the pidl.
    Case STRRET_OFFSET
      MoveMemory uOffset, lpStr.CStr(0), 4
      GetStrRet = GetStrFromPtrA(pidlRel + uOffset)
    
    ' The display name is returned in cStr.
    Case STRRET_CSTR
      GetStrRet = GetStrFromPtrA(VarPtr(lpStr.CStr(0)))
  
  End Select

End Function

' Returns the IShellFolder for the specified relative pidl

'   isfParent - pidl's parent folder IShellFolder
'   pidlRel    - child folder's relative pidl we're returning the IShellFolder of.

' If an error occurs, the desktop's IShellFolder is returned.

Public Function GetIShellFolder(isfParent As olelib.IShellFolder, pidlRel As Long) As IShellFolder
  Dim isf As IShellFolder
  Dim isfptr As Long
  On Error GoTo out
  
  Call isfParent.BindToObject(pidlRel, 0, IID_IShellFolder, isfptr)
  CopyMemory isf, isfptr, 4

out:
  If Err Or (isf Is Nothing) Then
    Set GetIShellFolder = isfDesktop
  Else
    Set GetIShellFolder = isf
  End If

End Function

' Returns a reference to the parent IShellFolder of the last item ID in the specified
' fully qualified pidl (identical to the new Win2K SHBindToParent function).

' If pidlFQ is zero, or a relative (single item) pidl, then the desktop's IShellFolder
' is returned. If an unexpected error occurs, the object value Nothing is returned.

Public Function GetIShellFolderParent(ByVal pidlFQ As Long, _
                                                            Optional fRtnDesktop As Boolean = True) As IShellFolder
  Dim pidlParent As Long

  pidlParent = GetPIDLParent(pidlFQ, fRtnDesktop)
  If pidlParent Then
    Set GetIShellFolderParent = GetIShellFolder(isfDesktop, pidlParent)
    isMalloc.Free ByVal pidlParent
  End If

End Function
Function ToPidl(ByVal I As Long) As Long
    ' Set of imaginable special folder constant
    If I >= CSIDL_DESKTOP And (I <= 32767) Then
        ToPidl = PidlFromSpecialFolder(I)
    Else
        ToPidl = I
    End If
End Function
Property Get Allocator() As IMalloc
    Static alloc As IMalloc
    If alloc Is Nothing Then SHGetMalloc alloc
    Set Allocator = alloc
End Property
Function PidlFromSpecialFolder( _
                Optional ByVal csidl As Long, _
                Optional ByVal hWnd As Long = 0) As Long
   ' InitIf  ' Initialize if in standard module
    On Error Resume Next
    Dim pidl As Long
    
    SHGetSpecialFolderLocation hWnd, csidl, pidl
    If Err = 0 Then PidlFromSpecialFolder = pidl
End Function


Function DuplicateItemID(pidl As Long) As Long
    Dim c As Integer, pidlNew As Long, iZero As Integer ' = 0
    If pidl = 0 Then Exit Function
    ' Get the size
    c = ItemIDSize(pidl)
    If c = 0 Then Exit Function
    ' Allocate space plus two for zero terminator
    On Error Resume Next
    pidlNew = Allocator.alloc(c + 2)
    If pidlNew = 0 Then Exit Function
    
    ' Copy the pidl data
    CopyMemory ByVal pidlNew, ByVal pidl, c
    ' Terminating zero
    CopyMemory ByVal pidlNew + c, iZero, 2
    DuplicateItemID = pidlNew
End Function
Function ItemIDSize(ByVal pidl As Long) As Integer
    If pidl Then CopyMemory ItemIDSize, ByVal pidl, 2
End Function
' Get the next item ID in an item ID list
Function NextItemID(ByVal pidl As Long) As Long
    Dim c As Integer
    If pidl = 0 Then Exit Function
    c = ItemIDSize(pidl)
    If c = 0 Then Exit Function
    NextItemID = pidl + c
End Function
Function PidlCount(ByVal pidl As Long) As Long
    Dim cItem As Long
    If pidl = 0 Then Exit Function
    Do While ItemIDSize(pidl)
        pidl = NextItemID(pidl)
        cItem = cItem + 1
    Loop
    PidlCount = cItem
End Function
'Function GetFullPath(sFileName As String, _
'                     Optional FilePart As Long, _
'                     Optional ExtPart As Long, _
'                     Optional DirPart As Long) As String
'
'    Dim c As Long, p As Long, sRet As String
'    If sFileName = sEmpty Then Exit Function
'
'    ' Get the path size, then create string of that size
'    sRet = String(cMaxPath, 0)
'    c = GetFullPathName(sFileName, cMaxPath, sRet, p)
'    If c = 0 Then ApiRaise Err.LastDllError
'    BugAssert c <= cMaxPath
'    sRet = Left$(sRet, c)
'
'    ' Get the directory, file, and extension positions
'    GetDirExt sRet, FilePart, DirPart, ExtPart
'    GetFullPath = sRet
'
'End Function

Function FolderFromItem(HwndOwner As Long, vItem As Variant, _
                        Optional pidl As Long = -1) As olelib.IShellFolder
   ' InitIf  ' Initialize if in standard modue
    
    Dim folder As olelib.IShellFolder, folderNext As olelib.IShellFolder
    Dim pidlItem As Long, pidlTmp As Long, cItem As Long
   Dim iidShellFolder As olelib.UUID
   Dim retVal As Long, foldpointer As Long
   
    IIDFromString StrPtr("{000214E6-0000-0000-C000-000000000046}"), iidShellFolder
    On Error GoTo FolderFromItemFail
    'SHGetDesktopFolder folder
    Set folder = isfDesktop
    'CopyMemory folder, foldpointer, 4
    'folder.AddRef
    If varType(vItem) = vbString Then
        ' Make sure the file name is fully qualified
        'vItem = GetFullPath(CStr(vItem))
    
        ' Convert path name to pointer to an item ID list (pidl)
        Dim cParsed As Long, afItem As Long
        folder.ParseDisplayName HwndOwner, 0, StrPtr(CStr(vItem)), _
                                cParsed, pidlItem, 0
    Else
        ' If necessary, convert special folder to pidl
        pidlItem = ToPidl(vItem)
    End If

    ' Walk the list of item IDs and bind to each subfolder in list
    ' to find the folder containing the specified pidl

    cItem = PidlCount(pidlItem)
    ' If caller requests a pidl return, adjust to return pidl of parent
    If pidl <> -1 Then cItem = cItem - 1
    Do While cItem

        ' Create a one-item ID list for the next item in pidlMain
        pidlTmp = DuplicateItemID(pidlItem)
        If pidlTmp = 0 Then GoTo FolderFromItemFail

        'Debug.Print GetFolderName(folder, pidlTmp, SHGDN_NORMAL)
        
        ' Bind to the folder specified in the new igtem ID list
        Dim nextPtr As Long
        folder.BindToObject pidlTmp, 0, _
                            iidShellFolder, nextPtr
        CopyMemory folderNext, nextPtr, 4
        
        cItem = cItem - 1
        
        ' Release parent folder and reference current child
        Set folder = folderNext
        'delete foldernext object manually, to prevent crash.
        If Not folderNext Is Nothing Then
           ' CopyMemory folderNext, 0, 4
           Set folderNext = Nothing
        End If
        ' Free temporary pidl
        'Allocator.Free pidlTmp
        CoTaskMemFree pidlTmp
        pidlTmp = 0
        ' Point to next item (if any)
        If cItem Then pidlItem = NextItemID(pidlItem)
    Loop
    Set FolderFromItem = folder
    If pidl = -1 Then
        ' Free temporary pidl if user doen't request it
        'Allocator.Free pidlItem
        CoTaskMemFree pidlItem
    Else
        ' User who asked for pidl must free it
        pidl = pidlItem
    End If
    pidl = pidlItem
   ' CopyMemory folder, 0, 4
   'Set folder = Nothing
    Exit Function
    
FolderFromItemFail:
    pidl = 0
    If pidlTmp <> 0 Then Allocator.Free pidlTmp
    If pidlTmp <> 0 Then CoTaskMemFree pidlTmp
                   
End Function
Public Function ContextMenu3IID() As olelib.UUID
    Dim ret As olelib.UUID
    'uuid(BCFCE0A0-EC17-11d0-8D10-00A0C90F2719),
    CLSIDFromString "{BCFCE0A0-EC17-11d0-8D10-00A0C90F2719}", ret
    ContextMenu3IID = ret
End Function

' Displays the specified items' shell context menu.
'
'    hwndOwner  - window handle that owns context menu and any err msgboxes
'    isfParent       - pointer to the items' parent shell folder
'    cPidls            - count of pidls at, and after, pidlRel
'    pidlRel          - the first item's pidl, relative to isfParent
'    pt                  - location of the context menu, in screen coords
'
' Returns True if a context menu command was selected, False otherwise.

Public Function ShowShellContextMenu(HwndOwner As Long, _
                                     isfParent As olelib.IShellFolder, _
                                     cPidls As Long, _
                                     pidlRel As Long, _
                                     pt As POINTAPI, _
                                     Optional CallbackObject As IContextCallback = Nothing, _
                                     Optional ByRef hMenuReturn As Long = 0, _
                                     Optional ByVal CMFFlags As QueryContextMenuFlags) As Boolean
                                                                
    'NOTES: if hmenureturn is NOT zero, then the hmenu will be populated and the function cleanly exited.
    'IE: use it to get a handle to the hMenu of the context menu for a given item.
    Const useAPImenufunction = 0
    Dim IID_IContextMenu  As olelib.UUID

    Dim IID_IContextMenu2 As olelib.UUID

    Dim IID_IContextMenu3 As olelib.UUID

    Dim icm               As IContextMenu

    Dim Ictx3             As IShellFolderEx_TLB.IContextMenu3

    Dim Ictx2             As olelib.IContextMenu2

    Dim hr                As Long   ' HRESULT

    Dim hMenu             As Long

    Dim idCmd             As Long, temppunk As olelib.IUnknown

    Dim cmi               As CMINVOKECOMMANDINFO

    Dim cmi3              As IShellFolderEx_TLB.CMINVOKECOMMANDINFO

    Dim icmPtr            As Long

    Dim ICtx2ptr          As Long

    Dim IctxMenu3ptr      As Long

    ' Fill the IContextMenu interface ID, {000214E4-000-000-C000-000000046}
    'old code used IShellFolder library, but didn't work properly.
    'Edanmo's lib doesn't include a few things, and uses long pointers instead of objects.
    'but we can work around that.
    Call DEFINE_OLEGUID(IID_IContextMenu, &H214E4, 0, 0)
    'retrieve IcontextMenu3 IID...
    IID_IContextMenu3 = ContextMenu3IID

    ' Get a refernce to the item's IContextMenu interface
    icmPtr = isfParent.GetUIObjectOf(HwndOwner, cPidls, pidlRel, IID_IContextMenu, 0)
    'copy icmPtr into the object for use.
    CopyMemory icm, icmPtr, 4

    'gawd I can't believe I'm actually successful with a lot of this.
    If SUCCEEDED(hr) Then

        ' Fill the IContextMenu2 interface ID, {000214F4-000-000-C000-000000046}
        ' and get the folder's IContextMenu2. Is needed so the "Send To" and "Open
        ' With" submenus get filled from the HandleMenuMsg call in the Subclassing class.

        Call DEFINE_OLEGUID(IID_IContextMenu2, &H214F4, 0, 0)
        'can't use query interface on icm- use temporary Iunknown.
        Set temppunk = icm

        'CHANGE://
        'query for IcontextMenu3 first.
        'Call temppunk.QueryInterface(IID_IContextMenu3, IctxMenu3ptr)
        If IctxMenu3ptr <> 0 Then
            'IContextMenu3 works.
            Debug.Print "IContextMenu3 is supported!"
            CopyMemory Ictx3, IctxMenu3ptr, 4
            'woopee!

        Else
            Debug.Print "IContextMenu3 is NOT supported... trying IContextMenu2..."
            Call temppunk.QueryInterface(IID_IContextMenu2, ICtx2ptr)

            If ICtx2ptr <> 0 Then
                CopyMemory Ictx2, ICtx2ptr, 4
            End If
        End If

        ' Create a new popup menu...

        Dim CreateMenu As cPopupMenu

'        If Not useAPImenufunction Then
            Set CreateMenu = CreateObject("PopupMenu6.cPopupMenu")
'            'CreateMenu.Header = True
            CreateMenu.HeaderStyle = ecnmHeaderCaptionBar
            CreateMenu.OfficeXpStyle = True
            CreateMenu.HwndOwner = HwndOwner
            CreateMenu.CreateSubClass HwndOwner
            CreateMenu.ActiveMenuForeColor = vbRed
            CreateMenu.ButtonHighlight = True
            
'            CreateMenu.AddItem "Context Menu.", , , , , , , "POPUP"
'            hMenu = CreateMenu.hMenu(1)
'        Else

            hMenu = CreatePopupMenu()
'        End If

        If hMenu Then

            ' Add the item's shell commands to the popup menu.

            'Three possibilities- one of the three ContextMenu interfaces.
            If (Ictx3 Is Nothing) = False Then

                Call Ictx3.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)

            ElseIf (Ictx2 Is Nothing) = False Then

                Call Ictx2.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)

            Else
                'if no IContextMenu2 (probably Win95, or NT4.)
                Call icm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)

            End If

            If SUCCEEDED(hr) Then

                If hMenuReturn = 0 Then 'if they don't want the hmenu handle- we show the popup.

                    Dim ObjCast     As Object

                    Dim MenuClasser As CContextSubClasser

                    ' Show the item's context menu-
                    'BUT FIRST-
                    'create the CContextSubClasser class, and initialize it. It should handle events so that such things as the Send To Menu are populated correctly.

                    If Not Ictx2 Is Nothing Or Not Ictx3 Is Nothing Then
                        Set MenuClasser = New CContextSubClasser
                        'IMPORTANT:
                        'insert Callback code to allow for modification of the hMenu we are going to show; IE: to add menu items.
    
                        Debug.Print "Initializing MenuClasser...."

                        If Not Ictx3 Is Nothing Then
                            Debug.Print "passing Ictx3..."

                            On Error Resume Next

                            Set ObjCast = Ictx3

                            If Err <> 0 Then
                                Debug.Print "error casting!" & Error$
                            End If

                            Err.Clear
                            Debug.Print "menuclasser is nothing=" & MenuClasser Is Nothing
                            MenuClasser.Init HwndOwner, ObjCast, CallbackObject

                            If Err <> 0 Then
        
                                Debug.Print "error initializing: " & Error$
                            End If

                        Else
                            Debug.Print "passing Ictx2..."
                            'Set ObjCast = Ictx2
                            MenuClasser.Init HwndOwner, Ictx2, CallbackObject
                        End If
                    End If

                    Dim fCancel As Boolean

                    If Not CallbackObject Is Nothing Then
                        CallbackObject.BeforeShowMenu hMenu, fCancel
    
                    End If

                    If Not fCancel Then
'                        If Not useAPImenufunction Then
'                            idCmd = CreateMenu.ShowPopupMenu(0, 0)
'
'                        Else
                            idCmd = TrackPopupMenu(hMenu, TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_HORIZONTAL Or TPM_RETURNCMD, pt.x, pt.y, 0, HwndOwner, 0)
'                        End If

                        Set MenuClasser = Nothing

                        ' If a menu command is selected...
                        If Not CallbackObject Is Nothing Then
                            CallbackObject.AfterShowMenu idCmd
                        End If
                    End If

                    If idCmd Then
  
                        ' Fill the struct with the selected command's information.
                        With cmi
                            .cbSize = Len(cmi)
                            .hWnd = HwndOwner
                            .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
                            .nShow = SW_SHOWNORMAL
                        End With

                        ' Invoke the shell's context menu command. The call itself does
                        ' not err if the pidlRel item is invalid, but depending on the selected
                        ' command, Explorer *may* raise an err. We don't need the return
                        ' val, which should always be NOERROR anyway...
                        If (Ictx3 Is Nothing) = False Then
                            LSet cmi3 = cmi
                            Call Ictx3.InvokeCommand(cmi3)
                        ElseIf (Ictx2 Is Nothing) = False Then
                            Call Ictx2.InvokeCommand(cmi)
                        Else
                            Call icm.InvokeCommand(cmi)
                        End If
  
                    End If   ' idCmd

                Else 'return hmenu.
                    hMenuReturn = hMenu
                    hMenu = 0
                    'of course- now the caller is responsible for erasing the menu. good luck to them, I say.
                End If   ' hr >= NOERROR (QueryContextMenu)

                If hMenu <> 0 Then
                    CreateMenu.DestroySubClass
                    Call DestroyMenu(hMenu)
                End If

            End If   ' hMenu
        End If   ' hr >= NOERROR (GetUIObjectOf)

        ' Release the folder's IContextMenu2 from the global variable.
        Set Ictx2 = Nothing
        Set Ictx3 = Nothing
        ' Returns True if a menu command was selected
        ' (letting us know to explicitly select the right clicked object, if needed)
        ShowShellContextMenu = CBool(idCmd)
    End If
    End Function
    '
    '' Returns the list of displaynames for each relative pidl
    '' (item ID) in the specified fully qualified pidl (item ID list).
    '
    '' called from nowhere, a debugging proc.
    '
    'Public Function GetPIDLNames(pidlFQ As Long) As String
    '  Dim nItems As Integer
    '  Dim isfParent As IShellFolder
    '  Dim i As Integer
    '  Dim pidlRel As Long
    '  Dim sNames As String
    '
    '  ' Get the count of item ID's in the item ID list.
    '  nItems = GetItemIDCount(pidlFQ)
    '  If nItems Then
    '
    '    ' Start with the desktop's shell folder.
    '    Set isfParent = isfDesktop
    '
    '    ' Walk through the each item ID in the item ID list.
    '    For i = 1 To nItems '- 1
    '
    '      ' Get the current relative pidl (item ID) from the
    '      ' fully qualified pidl (item ID list)
    '      pidlRel = GetItemID(pidlFQ, i)
    '      If pidlRel Then
    '
    '        ' Append each item ID's displayname to the output string.
    '        sNames = sNames & GetFolderDisplayName(isfParent, _
    '                                                                              pidlRel, _
    '                                                                              SHGDN_INFOLDER) & vbCrLf
    '        ' Bind to the current item ID's shell folder,
    '        ' setting it as the new parent shell folder
    '        If SUCCEEDED(isfParent.BindToObject(pidlRel, 0, IID_IShellFolder, isfParent)) = False Then
    '          Exit For
    '        End If
    '
    '        ' Free the relative pidl we just got and
    '        ' set it to 0 so we know it's freed.
    '        isMalloc.Free ByVal pidlRel
    '        pidlRel = 0
    '
    '      End If   ' pidlRel
    '    Next
    '
    '  End If   ' nItems
    '
    '  ' If the BindToObject call failed above and we exited
    '  ' the For loop without freeing the relative pidl, free it now.
    '  If pidlRel Then isMalloc.Free ByVal pidlRel
    '
    '  ' Return the item ID list's displaynames
    '  GetPIDLNames = sNames
    '
    'End Function


Public Sub TestShellDetails(ByVal mHwnd As Long, ByVal strPath As String)
Dim deskfolder As olelib.IShellFolder2
Dim boundto  As olelib.IShellFolder2, pidlfile As Long
Dim IShellDetailsPtr As Long
Dim newptr As Long
Dim IshellDetCast As IShellDetails
Dim pidlpath As Long, parentfolder As IShellFolder2
Dim Spath As String, sFile As String
Spath = Left$(strPath, InStrRev(strPath, "\"))



Set deskfolder = isfDesktop
    If Len(Spath) <= 3 Then
        
        'why, it's a drive spec.
        deskfolder.ParseDisplayName HwndOwner, 0, StrPtr(strPath), 0, pidlfile, 0
        Set parentfolder = deskfolder
    
    Else
    
        Set parentfolder = FolderFromItem(mHwnd, Left$(strPath, InStrRev(strPath, "\") - 1), pidlpath)
       
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, 0, Pointuse)
        'Call ShowShellContextMenu(hwndOwner, DeskFolder, 1, PidlPath, Pointuse)
        
        
        'Current Work: this line errs when a Drive name is specified.
        
        
        
        'deskfolder.ParseDisplayName HwndOwner, 0, StrPtr(Mid$(strPath, InStrRev(strPath, "\") + 1)), 0, pidlfile, 0
        'Call deskfolder.GetUIObjectOf(HwndOwner, CSIDL_DESKTOP, pidlfile, IID_IShellFolder2, newptr)
        'IShellDetailsPtr = deskfolder.CreateViewObject(mHwnd, IID_IShellDetails)
        'CopyMemory IshellDetCast, IShellDetailsPtr, 4
        'CopyMemory boundto, newptr, 4
        Dim pstest As SHELLDETAILS, SHColtest As SHColInfo
        
        Call parentfolder.ParseDisplayName(mHwnd, 0, StrPtr(Mid$(strPath, InStrRev(strPath, "\") + 1)), 0, pidlfile, 0)
        
        IshellDetCast.GetDetailsOf pidlfile, 0, SHColtest
        parentfolder.GetDetailsOf pidlpath, 0, pstest
        
        
        
    End If


End Sub
Public Sub TestDetails2(ByVal mpath As String)
' Objeto Shell actual
Dim m_objCurrentApp  As shell32.Shell

' Descriptor de Ficheros Actual
Dim m_objCurrentFolderItem As shell32.FolderItem
' Descriptor directorio actual
Dim m_objCurrentFolder As shell32.folder
Dim PathPart As String, FilePart As String
Dim testcast As olelib.IShellFolder2

PathPart = Mid$(mpath, 1, InStrRev(mpath, "\"))
FilePart = FileSystem.GetFilenamePart(mpath)

If m_objCurrentApp Is Nothing Then
  
  Set m_objCurrentApp = New Shell
End If

Set m_objCurrentFolder = m_objCurrentApp.NameSpace(PathPart)
If Not (m_objCurrentFolder Is Nothing) Then
  Set m_objCurrentFolderItem = m_objCurrentFolder.ParseName(FilePart)
End If

'Position 10 is "Title" property.

If Not (m_objCurrentFolderItem Is Nothing) Then




  ExtendedProperties = m_objCurrentFolder.GetDetailsOf(m_objCurrentFolderItem, 2)
End If
Set testcast = m_objCurrentFolderItem




MsgBox ExtendedProperties 'I obtain "Owner" in place of " Title.Only it happens in Windows Vista"
End Sub
