VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrintDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Public Enum PrinterDialogFlagConstants
    PD_ALLPAGES = 0
    PD_COLLATE = 16
    PD_CURRENTPAGE = 4194304
    PD_DISABLEPRINTTOFILE = 524288
    PD_ENABLEPRINTTEMPLATE = 16384
    PD_ENABLEPRINTTEMPLATEHANDLE = 65536
    PD_EXCLUSIONFLAGS = 16777216
    PD_HIDEPRINTTOFILE = 1048576
    PD_NOCURRENTPAGE = 8388608
    PD_NOPAGENUMS = 8
    PD_NOSELECTION = 4
    PD_NOWARNING = 128
    PD_PAGENUMS = 2
    PD_PRINTTOFILE = 32
    PD_RETURNDC = 256
    PD_RETURNIC = 512
    PD_PD_SELECTION = 1
    PD_USEDEVMODECOPIES = 262144
    PD_USEDEVMODECOPIESANDCOLLATE = 262144
    PD_USELARGETEMPLATE = 268435456
    PD_EXCLCOPIESANDCOLLATE = 33024
End Enum
'typedef struct tagPRINTPAGERANGE {
'    DWORD nFromPage;
'    DWORD nToPage;
'} PRINTPAGERANGE, *LPPRINTPAGERANGE;
Private Type PRINTERPAGERANGE
    nFromPage As Long
    nToPage As Long
End Type
'typedef struct tagPDEX {
'    DWORD lStructSize;
'    HWND hwndOwner;
'    HGLOBAL hDevMode;
'    HGLOBAL hDevNames;
'    HDC hDC;
'    DWORD Flags;
'    DWORD Flags2;

Private Type PRINTDLG_T
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As PrinterDialogFlagConstants
    nFromPage As Integer
    nToPage As Integer
    nMinpage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long 'Ptr to String.
    lpSetupTemplateName As Long 'Ptr to string.
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type PRINTDLGEX_T
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long 'HGLOBAL...
    hDevNames As Long
    hDC As Long
    Flags As PrinterDialogFlagConstants
    Flags2 As Long
    ExclusionFlags As Long
    mPageRanges As Long
    mMaxPageRanges As Long
    lpPageRanges As Long  'Varptr() to first item in array...
    nMinpage As Long
    nMaxPage As Long
    nCopies As Long
    hInstance As Long
    lpPrintTemplateName As Long '(pointer...)
    lpCallback As Long
    nPropertyPages As Long
    lphPropertyPages As Long '(should be a pointer to a the first item.
    nStartPage As Long
    dwResultAction As Long
End Type
'    DWORD ExclusionFlags;
'    DWORD nPageRanges;
'    DWORD nMaxPageRanges;
'    LPPRINTPAGERANGE lpPageRanges;
'    DWORD nMinPage;
'    DWORD nMaxPage;
'    DWORD nCopies;
'    HINSTANCE hInstance;
'    LPCTSTR lpPrintTemplateName;
'    LPUNKNOWN lpCallback;
'    DWORD nPropertyPages;
'    HPROPSHEETPAGE *lphPropertyPages;
'    DWORD nStartPage;
'    DWORD dwResultAction;
'} PRINTDLGEX, *LPPRINTDLGEX;


Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Declare Function PRINTDLG_api Lib "comdlg32.dll" Alias "PrintDlgW" (ByRef pPrintDlg As PRINTDLG_T) As Long
Private Declare Function PrintDlgEx Lib "comdlg32.dll" Alias "PrintDlgExW" (ByRef pPrintDlg As PRINTDLGEX_T) As Long


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private mCachedVersion As OSVERSIONINFO, mvercached As Boolean
Private Function IsXPOrLater() As Boolean
    If Not mvercached Then
        mvercached = True
        mCachedVersion.dwOSVersionInfoSize = Len(mCachedVersion)
        GetVersionEx mCachedVersion
    End If
    With mCachedVersion
    IsXPOrLater = (.dwMajorVersion >= 5) Or (.dwMajorVersion = 5 And .dwMinorVersion >= 1)
    End With




End Function
'Private Sub ppExTostandard(EXStruct As PRINTDLGEX_T, std As PRINTDLG_T)
'
'    With EXStruct
'        std.Flags = .Flags
'        std.hDC = .hDC
'        std.hDevMode = .hDevMode
'        std.hDevNames = .hDevNames
'        std.hInstance = .hInstance
'        std.hPrintTemplate = .lpPrintTemplateName
'        std.hSetupTemplate = 0
'        std.hWndOwner = .hWndOwner
'        std.lCustData = 0
'        std.lpfnPrintHook = .lpfnpr
'
'
'End Sub

'Private Function PrintDlg(pPrintDlg As PRINTDLGEX_T) As Long
'    Dim useTemp As PRINTDLG_T
'    If IsXPOrLater() Then
'        PrintDlg = PrintDlgEx(pPrintDlg)
'
'    Else
'
'    End If
'
'
'
'
'
'End Function
