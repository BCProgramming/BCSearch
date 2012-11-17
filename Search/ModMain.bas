Attribute VB_Name = "ModMain"
Option Explicit
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Const DEFAULT_GUI_FONT As Long = 17
'Private Declare Sub CoCreateInstance Lib "ole32.dll" (rclsid As olelib.UUID, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Any)
Public Declare Function Beep Lib "kernel32.dll" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Type tagInitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type



Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32.dll" (ByVal lpBuffer As String, ByRef nSize As Long) As Long





Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const ICC_BAR_CLASSES = &H4
Private Const ICC_COOL_CLASSES = &H400
Global CurrApp As Application
'Private Declare Function LoadResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
   
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1
Private Type ChooseColorUDT
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public CDebug As CDebug
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SETICON = &H80

'typedef struct tagLVFINDINFO {
'    UINT flags;
'    LPCTSTR psz;
'    LPARAM lParam;
'    POINT pt;
'    UINT vkDirection;
'} LVFINDINFO, *LPFINDINFO;
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type AV_SCANNER_INFO
    Name As String
    CommandLine As String  'Include %1
End Type

Public Type ChooseFont
     lStructSize As Long
     hwndOwner As Long                                                            '  caller's window handle
     hDC As Long                                                                  '  printer DC/IC or NULL
     lpLogFont As Long                                                             ' LOGFONT          '  ptr. to a LOGFONT struct
     iPointSize As Long                                                           '  10 * size in points of selected font
     Flags As Long                                                                '  enum. type flags
     rgbColors As Long                                                            '  returned text Color
     lCustData As Long                                                            '  data passed to hook fn.
     lpfnHook As Long                                                             '  ptr. to hook function
     lpTemplateName As String                                                       '  custom template name
     hInstance As Long                                                            '  instance handle of.EXE that
               '    contains cust. dlg. template
     lpszStyle As String                                                            '  return the style field here
               '  must be LF_FACESIZE or bigger
     nFontType As Integer                                                            '  same value reported to the EnumFonts
               '    call back with the extra FONTTYPE_
               '    bits added
     MISSING_ALIGNMENT As Integer
     nSizeMin As Long                                                             '  minimum pt size allowed &
     nSizeMax As Long                                                             '  max pt size allowed if
               '    CF_LIMITSIZE is used
End Type




Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pColorChoice As ChooseColorUDT) As Long
Public Enum FontEnum
     CF_SCREENFONTS = &H1
     CF_PRINTERFONTS = &H2
     CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
     CF_SHOWHELP = &H4&
     CF_ENABLEHOOK = &H8&
     CF_ENABLETEMPLATE = &H10&
     CF_ENABLETEMPLATEHANDLE = &H20&
     CF_INITTOLOGFONTSTRUCT = &H40&
     CF_USESTYLE = &H80&
     CF_EFFECTS = &H100&
     CF_APPLY = &H200&
     CF_ANSIONLY = &H400&
     CF_SCRIPTSONLY = CF_ANSIONLY
     CF_NOVECTORFONTS = &H800&
     CF_NOOEMFONTS = CF_NOVECTORFONTS
     CF_NOSIMULATIONS = &H1000&
     CF_LIMITSIZE = &H2000&
     CF_FIXEDPITCHONLY = &H4000&
     CF_WYSIWYG = &H8000                                                   '  must also have CF_SCREENFONTS CF_PRINTERFONTS
     CF_FORCEFONTEXIST = &H10000
     CF_SCALABLEONLY = &H20000
     CF_TTONLY = &H40000
     CF_NOFACESEL = &H80000
     CF_NOSTYLESEL = &H100000
     CF_NOSIZESEL = &H200000
     CF_SELECTSCRIPT = &H400000
     CF_NOSCRIPTSEL = &H800000
     CF_NOVERTFONTS = &H1000000
End Enum

Public Const LF_FACESIZE = 32
Public Type LOGFONT
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName(LF_FACESIZE) As Byte
End Type
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4
Private Type APIRECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Global Const GlobalUpdateID = 1 'UpdateID for BCSearch

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As APIRECT) As Long
Private mPrevInstancePtr As Long
Private mNextInstancePtr As Long
Private mROTObject As ROTSupport

Global gWordInstalled As Boolean
Global gExcelInstalled As Boolean

Const LVM_FIRST = &H1000&
Const LVM_GETHEADER = (LVM_FIRST + 31)

Private Const S_OK = &H0
Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type

Public cmdParser As CommandLineParser
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DllGetVersion Lib "COMCTL32" (pdvi As DLLVERSIONINFO) As Long
Public Function GetComCtlVersion() As String
    Dim lmaj As Long, lMin As Long, lbuild As Long
    ComCtlVersion lmaj, lMin, lbuild
    GetComCtlVersion = Trim$(lmaj) & "." & Trim$(lMin) & "." & Trim$(lbuild)
End Function
Public Function ComCtlVersion( _
        ByRef lMajor As Long, _
        ByRef lMinor As Long, _
        Optional ByRef lbuild As Long _
    ) As Boolean
Dim hMod As Long
Dim lR As Long
Dim lptrDLLVersion As Long
Dim tDVI As DLLVERSIONINFO

    lMajor = 0: lMinor = 0: lbuild = 0

    hMod = LoadLibrary("comctl32.dll")
    If (hMod <> 0) Then
        lR = S_OK
        '/*
        ' You must get this function explicitly because earlier versions of the DLL
        ' don't implement this function. That makes the lack of implementation of the
        ' function a version marker in itself. */
        lptrDLLVersion = GetProcAddress(hMod, "DllGetVersion")
        If (lptrDLLVersion <> 0) Then
            tDVI.cbSize = Len(tDVI)
            lR = DllGetVersion(tDVI)
            If (lR = S_OK) Then
                lMajor = tDVI.dwMajor
                lMinor = tDVI.dwMinor
                lbuild = tDVI.dwBuildNumber
            End If
        Else
            'If GetProcAddress failed, then the DLL is a version previous to the one
            'shipped with IE 3.x.
            lMajor = 4
        End If
        FreeLibrary hMod
        ComCtlVersion = True
    
    End If

End Function



Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cX As Long
Dim cY As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hwnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cX = GetSystemMetrics(SM_CXICON)
   cY = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cX, cY, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cX = GetSystemMetrics(SM_CXSMICON)
   cY = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cX, cY, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub



Public Sub SetInstance(ByVal Setnext As Boolean, Instance As Application)
    Dim ppointer As Long, newptr As Long
    If Instance Is Nothing Then
        newptr = 0
        Exit Sub
    End If
    CopyMemory newptr, ObjPtr(Instance), Len(newptr)
    
    'mPrevInstancePtr = ppointer
    If Setnext Then
        mNextInstancePtr = newptr
    Else
        mPrevInstancePtr = newptr
    End If
End Sub
Public Function GetInstance(ByVal GetNextInstance As Boolean) As Application
    Dim ppointer As Long, retobj As Object
    CopyMemory ppointer, mPrevInstancePtr, Len(ppointer)
    If ppointer <> 0 Then
        'copy it into the object...
        CopyMemory retobj, ppointer, Len(ppointer)
        If TypeOf retobj Is Application Then
            Set GetInstance = retobj
        Else
            Set GetInstance = Nothing
        End If
    
    Else
        Set GetInstance = Nothing
    End If
End Function

'Public Function MakeLong(LowPart As Integer, Hipart As Integer)
'    Dim ret As Long
'    CopyMemory ByVal VarPtr(ret), ByVal VarPtr(LowPart), 2
'    CopyMemory ByVal VarPtr(ret) + 2, ByVal VarPtr(Hipart), 2
'    MakeLong = ret
'
'
'
'End Function
Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long

    MakeLong = (HiWord * &H10000) Or (LoWord And &HFFFF&)

End Function
Public Function GetListViewColumnByPosition(LvwObject As vbalListViewCtl, ByVal lPosition As Long) As cColumn
    Dim LoopI As Long
    For LoopI = 1 To LvwObject.columns.Count
        If LvwObject.columns.Item(LoopI).position = lPosition Then
            Set GetListViewColumnByPosition = LvwObject.columns.Item(LoopI)
            Exit Function
        End If
    Next
    
    

    Set GetListViewColumnByPosition = Nothing

End Function


Public Function GetListViewHeaderHeight(ByVal hwndOfYourListView As Long) As Long

    Dim hwndHeader As Long
    Dim rcHeader As APIRECT
    
    hwndHeader = SendMessage(hwndOfYourListView, LVM_GETHEADER, 0&, ByVal 0&)
    If hwndHeader Then
    Call GetWindowRect(hwndHeader, rcHeader)
    GetListViewHeaderHeight = rcHeader.Bottom - rcHeader.Top
    End If



End Function

Public Function GetComputerName() As String
    Dim cpname As String
    cpname = Space$(255)
    Call GetComputerNameA(cpname, Len(cpname))
    cpname = Trim$(Replace$(cpname, vbNullChar, ""))
    GetComputerName = cpname


End Function
Public Function GetUserName() As String
    Dim username As String
    username = Space$(255)
    Call GetUserNameA(username, 255)
    username = Trim$(Replace$(username, vbNullChar, ""))
    GetUserName = username
End Function
Public Function ShowColor(Optional Flags As Integer) As Long
               'displays a Dialog Box that allows to user to choose a Color.
               'If they press cancel, it returns negative One.(-1)
     Dim udtClrPick As ChooseColorUDT
     Dim lReturn As Long
     Dim intCnt As Integer
     Dim bytColors() As Byte
     Static isshown As Boolean
     If isshown = True Then Exit Function
     isshown = True
     udtClrPick.lStructSize = Len(udtClrPick)
     udtClrPick.Flags = Flags
     ReDim bytColors(0 To 16 * 4 - 1) As Byte
     For intCnt = LBound(bytColors) To UBound(bytColors)
          bytColors(intCnt) = 0
     Next
     udtClrPick.lpCustColors = bytColors()
     
     udtClrPick.Flags = Flags
     If ChooseColor(udtClrPick) <> 0 Then
          ShowColor = udtClrPick.rgbResult
          bytColors = StrConv(udtClrPick.lpCustColors, vbFromUnicode)
     Else
          ShowColor = 0
     End If
     isshown = False
End Function

Public Function SelectFont(Optional initfont As StdFont = Nothing, Optional ByVal hwndOwner As Long = 0, Optional Flags As FontEnum = CF_BOTH + CF_EFFECTS) As StdFont
     Dim CF As ChooseFont, hMem As Long, lf As LOGFONT, aFontName As String
     Dim retfont As StdFont
     Set retfont = New StdFont
     hMem = GlobalAlloc(0, Len(lf))
     CF.hInstance = App.hInstance
     CF.hwndOwner = hwndOwner
     CF.lpLogFont = hMem
     CF.lStructSize = Len(CF)
     CF.Flags = Flags
     If Not initfont Is Nothing Then
        With lf
            '.lfFaceName = initfont.Name
            .lfItalic = initfont.Italic
            .lfWeight = initfont.Weight
            .lfHeight = initfont.size
            .lfStrikeOut = initfont.Strikethrough
     
        End With
     End If
     If ChooseFont(CF) Then
          CopyMemory lf, ByVal hMem, Len(lf)
          aFontName = Space$(LF_FACESIZE)
          CopyMemory ByVal aFontName, lf.lfFaceName(0), LF_FACESIZE
          With retfont
               .Name = CStr(aFontName)
               .Bold = lf.lfWeight
               .Italic = lf.lfItalic
               .size = CF.iPointSize / 10
               .Underline = lf.lfUnderline
               .Charset = lf.lfCharSet
               .Strikethrough = lf.lfStrikeOut
               
               End With
       
          
     End If
     GlobalFree hMem
     Set SelectFont = retfont
End Function

Public Function GetDefaultUIFont() As String
'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\Current
'Version\FontSubstitutes

Static mreg As cRegistry, deffont As String
If mreg Is Nothing Then
    Set mreg = New cRegistry
    deffont = mreg.ValueEx(HHKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes", "MS Shell Dlg 2", RREG_SZ, "Tahoma")
''    Open "test.log" For Append As #8
''    Print #8, deffont
''    Close #8
End If
    GetDefaultUIFont = deffont
  

End Function
Public Function GetScannerData(Optional ByRef ScannerCount As Long, Optional ByVal ForceRefresh As Boolean = False) As AV_SCANNER_INFO
Static flInitialized As Boolean, ScannerData() As AV_SCANNER_INFO, infocount As Long
Dim I As Long
Static creg As cRegistry
'enumerate Software\Microsoft\Windows\CurrentVersion\UnInstall,
If Not flInitialized Or ForceRefresh Then
    Dim UninstallClsids() As String, CLSIDcount As Long
    Dim lookvalues() As String, valCount As Long
    Erase ScannerData
    infocount = 0
    Set creg = New cRegistry
    creg.Classkey = HHKEY_LOCAL_MACHINE
    creg.SectionKey = "Software\Microsoft\Windows\CurrentVersion\UnInstall"
    
    If creg.EnumerateSections(UninstallClsids(), CLSIDcount) Then
        'ok, look at each one- the items of interest are the values:
        'DisplayName
        '
        
        For I = 1 To CLSIDcount
            valCount = 0
            Erase lookvalues
            creg.SectionKey = creg.SectionKey & "\" & UninstallClsids(I)
            creg.EnumerateValues lookvalues(), valCount
        
        Next I
    End If
    
    
End If




End Function
Public Function GetVolumelabel(ByVal ofDrive As String)
    Dim obj As Object
    'Set Obj = CreateObject("Scripting.FileSystemObject")
    Set obj = New Scripting.FileSystemObject
    On Error Resume Next
    GetVolumelabel = obj.Drives(Left$(ofDrive, 2)).VolumeName

    
End Function




Public Sub Main()
    Dim tICCEx As tagInitCommonControlsEx
    Set CDebug = New CDebug
   On Error Resume Next
 
   
   
   LoadLibrary "Shell32.dll"
   With tICCEx
       .dwSize = LenB(tICCEx)
       .dwICC = ICC_BAR_CLASSES
   End With
   InitCommonControlsEx tICCEx
   On Error GoTo report



    'CHANGE:
    
    'before creating OUR application object, we will check for previous instances.
    Dim Previnstance As Object
    On Error Resume Next
    Set Previnstance = GetObject(, "BASeSearch.Application")
    If Err <> 0 Then
    

    Else
        Debug.Print "previous running object found...."
        
        If Not CBool(Previnstance.Settings.ReadProfileSetting("BCSearch", "MultiInstance", "1")) Then
            'multiple instances not allowed. Send the command line to the previous instance and quit.
            Previnstance.SendCmdLine Command$
            Set Previnstance = Nothing
            Exit Sub
        Else
            'record previous instance. but make sure to only store a pointer.
            
        End If
    End If
    'CommandLine = ModCommandLine.ParseCommandLine(Command$)
    Set cmdParser = New BCCParser.CommandLineParser
    cmdParser.ParseArguments Command$
    Debug.Print "arguments:" + cmdParser.Arguments.Count & " and " & cmdParser.Switches.Count & " switches."
    
    Dim currarg As Long, currswitch As Long
    Set CurrApp = New Application
    If cmdParser.Switches.Exists("hidden") Then
        'keep it hidden...
    Else
    
        CurrApp.Show
    End If
    
    
    'look for the following switches:
    'searchin: location to search. assign this to the cboSearchin combo.
    'searchfor: search mask.
    
    
    If cmdParser.Switches.Exists("searchin") And cmdParser.Switches.Exists("searchfor") Then
        CurrApp.MainForm.cboLookin.Text = cmdParser.GetSwitch("searchin", vbTextCompare).Arguments.Item(1).ArgString
        Dim newfilter As CSearchFilter
        Set newfilter = New CSearchFilter
        newfilter.FileSpec = cmdParser.GetSwitch("searchfor", vbTextCompare).Arguments.Item(1).ArgString
        Set newfilter.Tag = New CExtraFilterData
        
        CurrApp.SearchForm.mFileSearch.Filters.Add newfilter
    
    
    End If
    If Previnstance Is Nothing Then
        'if no previous instances, register the Application object in the ROT.
        Set mROTObject = New ROTSupport
        mROTObject.ExposeObject CurrApp
    
    End If
    
    
    Dim HelpFileFindPaths, AppPath As String, I As Long
    If Right$(App.Path, 1) = "\" Then AppPath = App.Path Else AppPath = App.Path & "\"
    HelpFileFindPaths = Array(AppPath, GetSpecialFolder(CSIDL_APPDATA).Path & "BCSearch\")
    For I = 0 To UBound(HelpFileFindPaths)
        If Dir$(HelpFileFindPaths(I) & "BCSearch.chm") <> "" Then
            App.HelpFile = HelpFileFindPaths(I) & "BCSearch.chm"
        End If
    
    
    Next I
    Dim mreg As cRegistry
    Set mreg = New cRegistry
    mreg.Classkey = hhkey_classes_root
    mreg.SectionKey = "Word.Application"
    gWordInstalled = mreg.KeyExists
    mreg.SectionKey = "Excel.Application"
    gExcelInstalled = mreg.KeyExists
    
    
    Exit Sub
report:
    MsgBox "error #" & Err.Number & " """ & Err.Description & """."
End Sub
Public Sub ChangeToDefaultFont(ObjForm As Object)
    Static mDefaultFont As StdFont

    Dim loopcontrol As Object

    If mDefaultFont Is Nothing Then
'        initialize font object.
        Set mDefaultFont = New StdFont
        mDefaultFont.Name = GetDefaultUIFont
        CDebug.Post "Default font is " + FontToString(mDefaultFont)

    End If
    On Error GoTo Nocollection
    For Each loopcontrol In ObjForm.Controls
        If StrComp(loopcontrol.Tag, "NoPersist", vbTextCompare) <> 0 Then
            On Error Resume Next
    
            Set loopcontrol.Font = mDefaultFont
        End If
    Next
    loopcontrol
'    done.
Nocollection:
End Sub
Public Function validateEmail(ByVal emailvalid As String) As Boolean
    Static objregexp As RegExp
    Static flcache As Boolean
    If Not flcache Then
        flcache = True
        Set objregexp = New RegExp
        'set to email regexp....
        objregexp.Pattern = "^[-!#$%&'*+/0-9=?A-Z^_a-z{|}~](\.?[-!#$%&'*+/0-9=?A-Z^_a-z{|}~])*@[a-zA-Z](-?[a-zA-Z0-9])*(\.[a-zA-Z](-?[a-zA-Z0-9])*)+$"
        
    End If
    objregexp.Global = True
    objregexp.IgnoreCase = True
    
    validateEmail = objregexp.test(emailvalid)


End Function
Public Function FontToString(FontFrom As StdFont) As String
    Dim strreturn As String
    With FontFrom
    strreturn = .Name & ","
    If .Bold Then strreturn = strreturn & " Bold,"
    If .Italic Then strreturn = strreturn & " Italicized,"
    If .Underline Then strreturn = strreturn & " Underlined,"
    If .Strikethrough Then strreturn = strreturn & "Strikethrough,"
    strreturn = strreturn & " at " & .size & " Pt."
    
    
    End With
    

FontToString = strreturn



End Function
Public Function StringToFont(ByVal StrFontDesc As String) As StdFont
    Dim strsplit() As String
    Dim newfont As StdFont
    On Error Resume Next
    Set newfont = New StdFont
    'Fontname,PointSize,Bold,Italic,Underline,Strikethrough
    strsplit = Split(StrFontDesc, ",")
    newfont.Name = strsplit(0)
    newfont.size = Val(strsplit(1))
    
    newfont.Bold = CBool(strsplit(2))
    newfont.Italic = CBool(strsplit(3))
    newfont.Underline = CBool(strsplit(4))
    newfont.Strikethrough = CBool(strsplit(5))
    Set StringToFont = newfont



End Function
Public Function CommandBarThemeFromStr(StrFrom As String) As vbalCmdBar6.EToolBarStyle
    Select Case UCase$(StrFrom)
        Case "EMONEY", "MSMONEY", "MONEY"
        CommandBarThemeFromStr = eMoney
        Case "ECOMCTL", "ECOMCTL32", "COMCTL", "COMCTL32"
        CommandBarThemeFromStr = eComCtl32
        Case "OFFICEXP", "EOFFICEXP"
        CommandBarThemeFromStr = eOfficeXP
        Case "OFFICE03", "OFFICE2003", "EOFFICE03", "EOFFICE2003"
        CommandBarThemeFromStr = eOffice2003
    End Select

End Function
Public Function CommandBarStyleConv(Value As Variant, ToString As Boolean) As Variant
Dim X As vbalCmdBar6.ECustomColors
Dim I As Long
Static lookup(1 To 30) As String, flinit As Boolean
If Not flinit Then
    flinit = True
    lookup(1) = "ButtonTextColor"
    lookup(2) = "ButtonTextHotColor"
    lookup(3) = "ButtonTextDisabledColor"
    lookup(4) = "ButtonBackgroundColorStart"
    lookup(5) = "ButtonBackgroundColorEnd"
    lookup(6) = "ButtonHotBackgroundColorStart"
    lookup(7) = "ButtonHotBackgroundColorEnd"
    lookup(8) = "ButtonCheckedBackgroundColorStart"
    lookup(9) = "ButtonCheckedBackgroundColorEnd"
    lookup(10) = "ButtonCheckedHotBackgroundColorStart"
    lookup(11) = "ButtonCheckedHotBackgroundColorEnd"
    lookup(12) = "MenuShadowColor"
    lookup(13) = "MenuBorderColor"
    lookup(14) = "MenuTextColor"
    lookup(15) = "MenuTextHotColor"
    lookup(16) = "MenuTextDisabledColor"
    lookup(17) = "MenuBackgroundColorStart"
    lookup(18) = "MenuBackgroundColorEnd"
    lookup(19) = "MenuHotBackgroundColorStart"
    lookup(20) = "MenuHotBackgroundColorEnd"
    lookup(21) = "MenuHotBorderColor"
    lookup(22) = "MenuCheckedBackgroundColorStart"
    lookup(23) = "MenuCheckedBackgroundColorEnd"
    lookup(24) = "MenuCheckedHotBackgroundColorStart"
    lookup(25) = "MenuCheckedHotBackgroundColorEnd"
    lookup(27) = "IconDisabledColor"
    lookup(27) = "LightColor"
    lookup(28) = "DarkColor"
    lookup(29) = "GradientColorStart"
    lookup(30) = "GradientColorEnd"
End If

If ToString Then
    CommandBarStyleConv = lookup(Value)
Else
    For I = LBound(lookup) To UBound(lookup)
        If StrComp(lookup(I), Value, vbTextCompare) = 0 Then
            CommandBarStyleConv = I
            Exit For
        End If
    Next I


End If


End Function
Public Sub DEFINE_GUID(Name As olelib.UUID, L As Long, w1 As Integer, w2 As Integer, _
                                          b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, _
                                          b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With Name
    .Data1 = L
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


Public Function GetColumnProviders(Optional ByRef providerCount As Long) As IColumnProvider()

    Dim retArray() As IColumnProvider
    Dim columnCLSIDs() As String, CLSIDcount As Long
    Static mRegistry As cRegistry
    Dim IID_IColumnProvider As olelib.UUID
    Dim newptr As Long, currUUID As olelib.UUID
    Dim I As Long
    Dim addrefpunk As olelib.IUnknown
    Dim currprovider As olelib.IUnknown, castprovider As IColumnProvider
    If mRegistry Is Nothing Then Set mRegistry = New cRegistry
    'enumerate HKEY_CLASSES_ROOT\Folder\shellex\ColumnHandlers
    'each item is a CLSID.
    mRegistry.Classkey = hhkey_classes_root
    mRegistry.SectionKey = "Folder\ShellEx\ColumnHandlers"
    mRegistry.EnumerateSections columnCLSIDs(), CLSIDcount
    ReDim retArray(1 To CLSIDcount)
    Call DEFINE_GUID(IID_IColumnProvider, &HE8025004, &H1C42, &H11D2, &HBE, &H2C, &H0, &HA0, &HC9, &HA8, &H3D, &HA1)
    For I = 1 To CLSIDcount
        
        CLSIDFromString columnCLSIDs(I), currUUID
        CoCreateInstance currUUID, Nothing, CLSCTX_INPROC_SERVER, IID_IColumnProvider, castprovider
        'If Not currprovider Is Nothing Then
                
        
        'End If
        Set retArray(I) = castprovider
        Set addrefpunk = castprovider
        addrefpunk.AddRef

    


    Next I

    GetColumnProviders = retArray
    providerCount = CLSIDcount
End Function

Public Sub TestColumns()


    Dim retArray() As IColumnProvider
    Dim columnCLSIDs() As String, CLSIDcount As Long
    Static mRegistry As cRegistry
    Dim IID_IColumnProvider As olelib.UUID
    Dim newptr As Long, currUUID As olelib.UUID
    Dim I As Long
    Dim addrefpunk As olelib.IUnknown
    Dim currprovider As olelib.IUnknown, castprovider As IColumnProvider
    If mRegistry Is Nothing Then Set mRegistry = New cRegistry
    'enumerate HKEY_CLASSES_ROOT\Folder\shellex\ColumnHandlers
    'each item is a CLSID.
    mRegistry.Classkey = hhkey_classes_root
    mRegistry.SectionKey = "Folder\ShellEx\ColumnHandlers"
    mRegistry.EnumerateSections columnCLSIDs(), CLSIDcount
    ReDim retArray(1 To CLSIDcount)
    Call DEFINE_GUID(IID_IColumnProvider, &HE8025004, &H1C42, &H11D2, &HBE, &H2C, &H0, &HA0, &HC9, &HA8, &H3D, &HA1)
    For I = 1 To CLSIDcount
        
        CLSIDFromString columnCLSIDs(I), currUUID
        CoCreateInstance currUUID, Nothing, CLSCTX_INPROC_SERVER, IID_IColumnProvider, castprovider
        'If Not currprovider Is Nothing Then
                
        
        'End If
        Set retArray(I) = castprovider
        Set addrefpunk = castprovider
        addrefpunk.AddRef

    


    Next I

    'now, iterate through each retarray item...
    Dim shinit As SHCOLUMNINIT
    Dim sinitfolder As String, Bytes() As Byte
    Dim ColumnInfo As SHCOLUMNINFO
    
    Dim spaced() As Byte
    Dim spcstr As String
    sinitfolder = "C:\"
    shinit.dwFlags = 0
    shinit.dwReserved = 0
    Bytes = sinitfolder
    spcstr = Space$(32768)
    spaced = spcstr
    Dim k As Long
    Call CopyMemory(shinit.wszFolder(0), Bytes(0), UBound(Bytes))
    For I = 1 To UBound(retArray)
        retArray(I).Initialize shinit
        'Do
           ' Columninfo.wszDescription = Space$(UBound(Columninfo.wszDescription))
           ' Columninfo.wszTitle = Space$(UBound(Columninfo.wszTitle))
           
           CopyMemory ColumnInfo.wszDescription(0), spaced(0), UBound(ColumnInfo.wszDescription)
           CopyMemory ColumnInfo.wszTitle(0), spaced(0), UBound(ColumnInfo.wszTitle)
            retArray(I).GetColumnInfo 0, ColumnInfo
            
            
            Stop
            
            k = k + 1
        'Loop
        Set addrefpunk = retArray(I)
        addrefpunk.Release
        
    Next I




End Sub
Public Function CreateObject(ByVal Class As String, Optional ByVal ServerName As String) As Object

    If StrComp(Class, "BASeSearch.CMP3Columns", vbTextCompare) = 0 Then
        Set CreateObject = New CMP3Columns
    ElseIf StrComp(Class, "BASeSearch.CAlternateStreamColumns", vbTextCompare) = 0 Then
        Set CreateObject = New CAlternateStreamColumns
    ElseIf StrComp(Class, "BASeSearch.CExtraColumns", vbTextCompare) = 0 Then
        Set CreateObject = New CExtraColumns
    ElseIf StrComp(Class, "BASeSearch.CVersionColumns", vbTextCompare) = 0 Then
        Set CreateObject = New CVersionColumns
    Else
        If ServerName <> "" Then
            Set CreateObject = VBA.CreateObject(Class, ServerName)
        Else
            Set CreateObject = VBA.CreateObject(Class)
    End If




End If



End Function
Public Sub ShowDonate()
'

End Sub
Private Function ArrayStr(ParamArray Strings()) As String()
    Dim I As Long
    Dim retValue() As String
    ReDim retValue(0 To UBound(Strings))
    For I = 0 To UBound(Strings)
        retValue(I) = Strings(I)
    Next I

   ArrayStr = retValue


End Function
Public Function ExpandEnvironment(ByVal Str As String) As String
    Dim Dest As String
    Static EnvNames() As String, EnvValues() As String
    Dest = Space$(32768)
    
    
    
    
    ExpandEnvironmentStrings Str, Dest, 32767

    ExpandEnvironment = Replace$(Trim$(Dest), vbNullChar, "")
    
End Function

Public Function GetFirstAvailableFile(ParamArray FileNames()) As String
Dim I As Long
For I = 0 To UBound(FileNames)
    If VarType(FileNames(I)) = vbString Then
        If bcfile.Exists(FileNames(I)) Then
            GetFirstAvailableFile = CStr(FileNames(I))
            Exit Function
        End If
    
    
    End If

Next I

GetFirstAvailableFile = ""

End Function
'def Translate(Str, dict):
'    words = string.split(string.lower(str))
'    keys = dict.keys();
'    for i in range(0,len(words)):
'        if words[i] in keys:
'            words[i] = dict[words[i]]
'    return string.join(words)


Public Function IsBinary64bit(ByVal Dllpath As String) As Boolean
    Dim freader As FileStream
    Dim peOffset As Long, pehead As Long
    Dim machinetype As Integer
    Set freader = bcfile.OpenStream(Dllpath)
    freader.SeekTo &H3C, STREAM_BEGIN
    peOffset = freader.readLong
    freader.SeekTo peOffset, STREAM_BEGIN
    pehead = freader.readLong
    CDebug.Post pehead
    machinetype = freader.ReadInteger
    CDebug.Post machinetype
    freader.CloseStream
    
    
    

End Function
Public Sub CheckBytes(onFile As String)
Dim StreamOfFile As bcfile.FileStream
Dim stringread As String, countmismatch As Long
Set StreamOfFile = bcfile.OpenStream(onFile)

Do Until StreamOfFile.AtEndOfStream
    stringread = StreamOfFile.readstring(32768, StrRead_ANSI, False)
    If StrComp(stringread, String$(32768, vbNullChar)) <> 0 Then
        CDebug.Post "compare mismatch"
        countmismatch = countmismatch + 1
    
    End If
Loop
StreamOfFile.CloseStream
CDebug.Post "number of 32K blocks containing characters that are not Null:" & countmismatch







End Sub

'public static MachineType GetDllMachineType(string dllPath)
'    {
'      //see http://download.microsoft.com/download/9/c/5/
'      //             9c5b2167-8017-4bae-9fde-d599bac8184a/pecoff_v8.doc
'      //offset to PE header is always at 0x3C
'      //PE header starts with "PE\0\0" =  0x50 0x45 0x00 0x00
'      //followed by 2-byte machine type field (see document above for enum)
'      FileStream fs = new FileStream(dllPath, FileMode.Open);
'      BinaryReader br = new BinaryReader(fs);
'      fs.Seek(0x3c, SeekOrigin.Begin);
'      Int32 peOffset = br.ReadInt32();
'      fs.Seek(peOffset, SeekOrigin.Begin);
'      UInt32 peHead = br.ReadUInt32();
'      if(peHead!=0x00004550) // "PE\0\0", little-endian
'        throw new Exception("Can't find PE header");
'      MachineType machineType = (MachineType) br.ReadUInt16();
'      br.Close();
'      fs.Close();
'      return machineType;
'    }


Public Sub TestExecutive()
    Dim X As CExecutive
    Set X = New CExecutive
    
    X.launch "C:\windows\system32\notepad.exe"


End Sub
Public Sub Prepender()

    Dim fsopen As FileStream
    Dim fsout As FileStream
    Dim codecb() As Byte
    ReDim codecb(1 To 1)
    
    Set fsopen = OpenStream("D:\document.doc")
    Set fsout = CreateStream("D:\docout.doc")
    fsout.writebytes codecb()
    fsout.Writestream fsopen
    fsout.CloseStream
    fsopen.CloseStream
End Sub

Public Function GetDirString(ByVal Path As String) As String
    Dim buildstring As cStringBuilder
    Set buildstring = New cStringBuilder
    
    Dim gotDir As Directory
    Dim GotDrive As CVolume
    Dim driveletter As String
    Dim LoopObject As Object
    Set gotDir = GetDirectory(Path)
    
    Set GotDrive = gotDir.Volume
    driveletter = Left$(GotDrive.RootFolder.Path, 1)
    buildstring.Append "Volume in drive " & driveletter & " is " & GotDrive.Label & vbCrLf
    buildstring.Append "Volume Serial Number is " & GotDrive.GetFormattedSerialNumber() & vbCrLf & vbCrLf
    buildstring.Append "Directory of " & gotDir.Path & vbCrLf & vbCrLf
    
    Dim gotwalker As CDirWalker
    Set gotwalker = gotDir.GetWalker(, 0, 0)
    Dim FileCount As Long
    Dim DirCount As Long, runsize As Double
    
    Do Until gotwalker.GetNext(LoopObject) Is Nothing
    buildstring.Append FormatDateTime(LoopObject.DateModified, vbShortDate) & vbTab & FormatDateTime(LoopObject.DateModified, vbLongTime)
'08/02/2009  01:54 PM    <DIR>          .scorched3d
'10/23/2009  07:14 AM             1,047 ccmanifest.zip
'07/31/2009  02:34 AM    <DIR>          Contacts


'Column 1: date.

    Dim sizestring As String
        If TypeOf LoopObject Is Directory Then
            buildstring.Append vbTab & "<DIR>"
            
            DirCount = DirCount + 1
            buildstring.Append LoopObject.Name
        ElseIf TypeOf LoopObject Is CFile Then
            buildstring.Append vbTab & Format$(LoopObject.size, "###,###") & vbTab
            
            FileCount = FileCount + 1
            runsize = runsize + LoopObject.size
            buildstring.Append LoopObject.Filename
        End If
        
        buildstring.Append vbCrLf
    
    Loop
    buildstring.Append FileCount & " file(s)" & vbTab & Format$(runsize, "###,###") & vbTab & "bytes." & vbCrLf
    buildstring.Append DirCount & " dir(s)" & vbTab & Format$(GotDrive.FreeSpace, "###,###") & vbTab & " bytes free." & vbCrLf
    
    GetDirString = buildstring.ToString



End Function
Public Sub TestAutoString()

    Dim fStreamUse As FileStream
    Dim stringread As String
    Set fStreamUse = OpenStream("C:\teststream4.txt")
    stringread = fStreamUse.ReadStringAuto(5)
    
    Stop

End Sub
Public Function CalcTimepassedOnEarth(ByVal TimeOnObject As Date, ByVal velocity As Variant)
Const light As Double = 300000 'km/h
'time passed for object * 1/sqrt(1-((v*v)/(c*c)))
CalcTimepassedOnEarth = CDate(TimeOnObject * (1 / Sqr(1 - ((velocity * velocity) / (light * light)))))

End Function
Public Function DateToTimePassed(ByVal dateuse As Date)
Dim doubleeqv As Double
doubleeqv = dateuse
Dim Numdays As Double
Dim numhours As Long

Numdays = Fix(doubleeqv)
doubleeqv = doubleeqv - Numdays
numhours = doubleeqv / (1 / 24)
doubleeqv = doubleeqv - (numhours * (1 / 24))

DateToTimePassed = Numdays & " days," & numhours & "hours"
End Function
Private Function Quadform(a, b, C)
Dim result()
ReDim result(1 To 2)
'(-b+-Sqr(b^2-4ac))/2a
result(1) = (-b + Sqr(b ^ 2 - (4 * a * C)) / (2 * a))
result(2) = (-b - Sqr(b ^ 2 - (4 * a * C)) / (2 * a))
Quadform = result
End Function
Public Sub testquadform(a, b, C)
    Dim result
    result = Quadform(a, b, C)
    Stop
End Sub
Public Sub test()
Dim I As Long, X As Double
X = 100
For I = 1 To 15
X = X - (X * 0.06)
Next I
MsgBox X
End Sub
Private Function IsInIDE() As Boolean

On Error GoTo InIDE
    Debug.Print 1 / 0
    Exit Function
InIDE:
    IsInIDE = True


End Function
Public Function LoadResData(ID, resType) As Byte()
Static resourceini As CINIData
On Error GoTo errorloadres
Dim respath As String, retbytes() As Byte
 If IsInIDE Then
 
    'check "res" folder, underneath this folder, for a ini file.
    Dim buildinipath As String, loadedstream As FileStream
    
       If Right$(App.Path, 1) = "\" Then buildinipath = App.Path Else buildinipath = App.Path & "\"
       respath = buildinipath & "res\"
       If resourceini Is Nothing Then
       buildinipath = respath & "ide.ini"
       
       If Dir$(buildinipath) <> "" Then
          Set resourceini = New CINIData
          resourceini.LoadINI buildinipath
       End If
    End If
    If Not resourceini Is Nothing Then
       Dim splitvalue() As String
    
       'values are in "the "[mappings]" section.
       Dim EnumValues() As String, vcount As Long, CurrValue As Long
       Call resourceini.EnumerateValues(Setting_User, "mappings", EnumValues(), vcount)
       
       For CurrValue = 1 To UBound(EnumValues)
           splitvalue = Split(EnumValues(CurrValue), ":")
           'filename:id:restype
           'ubound 2....
           Dim gotfname As String, gotid As String, gotrestype As String
           If UBound(splitvalue) >= 2 Then
               gotfname = splitvalue(0)
               gotid = splitvalue(1)
               gotrestype = splitvalue(2)
               If gotid = ID And gotrestype = resType Then
               
                Set loadedstream = OpenStream(respath & gotfname)
                retbytes = loadedstream.readbytes(loadedstream.size)
                loadedstream.CloseStream
                LoadResData = retbytes
                Exit Function
               End If
           
           
           End If
       
       Next CurrValue
    
    
    End If
 
 Else
    'not in the IDE...
  LoadResData = VB.LoadResData(ID, resType)
 
 End If
Exit Function
errorloadres:
Debug.Assert False
Resume
End Function

'loadresdata(id,type)

Public Function intbytes(intcheck As Long) As String

Dim Bytes(3) As Byte
CopyMemory Bytes(0), intcheck, 4

intbytes = Hex$(Bytes(0)) & " " & Hex$(Bytes(1)) & " " & Hex$(Bytes(2)) & " " & Hex$(Bytes(3))

End Function

Public Sub SetTabStop(VisibleFrames As Variant, InvisibleFrames As Variant)


Dim visFrames, InvisFrames
If IsArray(VisibleFrames) Then
    visFrames = VisibleFrames
Else
    ReDim visFrames(0)
    Set visFrames(0) = VisibleFrames
End If

If IsArray(InvisibleFrames) Then
    InvisFrames = InvisibleFrames
Else
    ReDim InvisFrames(0)
    Set InvisFrames(0) = InvisibleFrames
End If



Dim loopcontrol As Object
Dim loopframe
On Error Resume Next
For Each loopcontrol In visFrames.Container.Controls
    For Each loopframe In visFrames
        If loopcontrol.Container Is loopframe Then
            loopcontrol.TabStop = True
        End If
    Next
    For Each loopframe In InvisFrames
        If loopcontrol.Container Is loopframe Then
            loopcontrol.TabStop = False
        End If
    Next

Next





End Sub
Public Function Replace(ByVal sSrc As String, _
                        ByVal sTerm As String, _
                        ByVal sNewTerm As String, _
                        Optional lStart As Long = 1, _
                        Optional lHitCnt As Long, _
                        Optional ByVal lCompare As _
                        VbCompareMethod = vbBinaryCompare _
                        ) As String ' ©Rd

    Dim lSize As Long, lHit As Long, lHitPos As Long, lPos As Long
    Dim lLenOrig As Long, lOffset As Long, lOffStart As Long
    Dim lLenOld As Long, lLenNew As Long, lCnt As Long
    Dim s1 As String, s2 As String, al() As Long

    'On Error GoTo FreakOut

    lLenOrig = Len(sSrc)
    If (lLenOrig = 0) Then Exit Function ' No text

    lLenOld = Len(sTerm): lLenNew = Len(sNewTerm)
    Replace = sSrc
    If (lLenOld = 0) Then Exit Function

    If lCompare = vbBinaryCompare Then
        s1 = sTerm: s2 = sSrc
    Else
        s1 = LCase$(sTerm): s2 = LCase$(sSrc)
    End If

    lOffset = lLenNew - lLenOld
    lCnt = 0: lSize = 8000 ' lSize = Arr chunk size
    ReDim al(0 To lSize) As Long

    lHit = InStr(lStart, s2, s1)
    Do While (lHit <> 0) And (lHit <= lLenOrig)
        al(lCnt) = lHit: lCnt = lCnt + 1
        If (lCnt = lHitCnt) Then Exit Do
        If (lCnt = lSize) Then
            lSize = lSize + 8000
            ReDim Preserve al(0 To lSize) As Long
        End If
        lOffStart = lHit + lLenOld ' offset start pos
        lHit = InStr(lOffStart, s2, s1)
    Loop

    If (lCnt = 0) Then GoTo FreakOut ' No hits
    lHitCnt = lCnt
    If lCompare = vbBinaryCompare Then
        If StrComp(s1, sNewTerm) = 0 Then Exit Function
    End If

    lSize = (lLenOrig + lOffset * lCnt) ' lSize = result str size
    Replace = Space$(lSize)

    lCnt = lCnt - 1: lOffStart = 1: lPos = 1
    For lHit = 0 To lCnt
        lHitPos = al(lHit)
        Mid$(Replace, lOffStart) = Mid$(sSrc, lPos, lHitPos - lPos)
        lOffStart = lHitPos + (lOffset * lHit)
        If (lLenNew <> 0) Then
            Mid$(Replace, lOffStart) = sNewTerm
            lOffStart = lOffStart + lLenNew
        End If
        lPos = lHitPos + lLenOld ' No offset orig str
    Next

    If lOffStart <= lSize Then
        Mid$(Replace, lOffStart) = Mid$(sSrc, lPos)
    End If
    'lHitCnt = lHit
FreakOut:
End Function    ' Rd - cryptic but crazy :)

Public Function PerformUpdateCheck(Optional ByVal OwnerWnd As Long = 0) As Boolean
Dim UpdateProg As Object
Dim strData() As String
If App.StartMode = ApplicationStartConstants.vbSModeAutomation Then Exit Function
    Dim DlID As String, Version As String, fullURL As String, DlSummary As String, DocURL As String, FileSize As String, dateuse As String, DlName As String
If CurrApp.Settings.ReadProfileSetting("BCSearch", "UpdateCheck", "1") = "1" Then
'Public Function CheckForUpdate(ByVal BCAppID As Long, ByVal CurrMajor As Integer, ByVal CurrMinor As Integer, ByVal CurrRevision As Integer, Optional ByVal ShowDialog As Boolean) As Boolean
    On Error GoTo updateerror
    Set UpdateProg = CreateObject("BCUpdate.CAppUpdate")
    If UpdateProg.CheckForUpdate(1, App.Major, App.Minor, App.Revision, False) Then
        strData = UpdateProg.verinfo
        DlID = strData(0)
        Version = strData(1)
        fullURL = strData(2)
        DlSummary = strData(3)
        DocURL = strData(4)
        FileSize = strData(5)
        dateuse = strData(6)
        DlName = strData(7)
        If MsgBox("a new update for " & DlName & " is available." & vbCrLf & _
        "Current Version:" & Trim$(App.Major) & "." & Trim$(App.Minor) & "." & Trim$(App.Revision) & vbCrLf & _
        "New Version:" & Version & vbCrLf & vbCrLf & _
        "Would you like to download it?", vbYesNo, "Update available") = vbYes Then
        
        ShellExec OwnerWnd, fullURL
        PerformUpdateCheck = True
        End If
        
    End If
    
    
    
End If
Set UpdateProg = Nothing
Exit Function
updateerror:

End Function

Public Function PerformNameSubstitution(currItem As cListItem, ByVal StrSubstitutionMask As String) As String
    Dim tempflt As CActionFilter
    Set tempflt = New CActionFilter
   PerformNameSubstitution = tempflt.DoSubstitute(StrSubstitutionMask, currItem)
   Set tempflt = Nothing
End Function
