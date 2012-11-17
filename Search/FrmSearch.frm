VERSION 5.00
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.1#0"; "vbalcmdbar6.ocx"
Object = "{AFFDD50D-733B-4E1C-8F98-E88F1ED6980D}#1.0#0"; "vbaListView6BC.ocx"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Begin VB.Form FrmSearch 
   Caption         =   "Search..."
   ClientHeight    =   7140
   ClientLeft      =   780
   ClientTop       =   4275
   ClientWidth     =   9990
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   Begin VB.PictureBox PicMain 
      Height          =   6015
      Left            =   2520
      ScaleHeight     =   5955
      ScaleWidth      =   6825
      TabIndex        =   2
      Top             =   660
      Width           =   6885
      Begin VB.PictureBox PicClientArea 
         BorderStyle     =   0  'None
         Height          =   5730
         Left            =   360
         ScaleHeight     =   382
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   471
         TabIndex        =   3
         Top             =   60
         Width           =   7065
         Begin VB.PictureBox PicUpperPane 
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   60
            ScaleHeight     =   249
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   465
            TabIndex        =   7
            Top             =   0
            Width           =   6975
            Begin VB.PictureBox PicSimple 
               Height          =   2835
               Left            =   540
               ScaleHeight     =   2775
               ScaleWidth      =   5475
               TabIndex        =   21
               Top             =   2100
               Visible         =   0   'False
               Width           =   5535
               Begin VB.ComboBox cboSimpleFileMask 
                  Height          =   315
                  Left            =   840
                  TabIndex        =   23
                  Top             =   120
                  Width           =   4515
               End
               Begin VB.CheckBox chkSimpleregExp 
                  Caption         =   "Use pattern matching"
                  Height          =   315
                  Left            =   3420
                  TabIndex        =   22
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.Label lblMask 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "&File Mask:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   24
                  Top             =   180
                  Width           =   720
               End
            End
            Begin VB.PictureBox PicANI 
               BorderStyle     =   0  'None
               Height          =   1095
               Left            =   5670
               ScaleHeight     =   73
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   77
               TabIndex        =   20
               Top             =   1350
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.CheckBox chksubfolders 
               Caption         =   "Search Subfolders"
               Height          =   195
               Left            =   2640
               TabIndex        =   19
               Top             =   480
               Width           =   2175
            End
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "..."
               Height          =   315
               Left            =   3960
               TabIndex        =   18
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox PicFilters 
               BorderStyle     =   0  'None
               Height          =   2835
               Left            =   1200
               ScaleHeight     =   189
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   381
               TabIndex        =   13
               Top             =   1080
               Width           =   5715
               Begin VB.CommandButton cmdEditFilter 
                  Caption         =   "&Edit"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   16
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.CommandButton CmdremoveFilter 
                  Caption         =   "&Remove"
                  Height          =   375
                  Left            =   0
                  TabIndex        =   15
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.CommandButton cmdAddFilter 
                  Caption         =   "&Add..."
                  Height          =   375
                  Left            =   0
                  TabIndex        =   14
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin vbalCmdBar6.vbalCommandBar cmdbarmenu 
                  Height          =   1815
                  Index           =   3
                  Left            =   0
                  Top             =   0
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   3201
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Orientation     =   1
                  Style           =   3
               End
               Begin vbaBClListViewLib6.vbalListViewCtl lvwfilters 
                  Height          =   1950
                  Left            =   1140
                  TabIndex        =   17
                  Top             =   240
                  Width           =   4425
                  _ExtentX        =   7805
                  _ExtentY        =   3440
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  View            =   1
                  LabelEdit       =   0   'False
                  FullRowSelect   =   -1  'True
                  AutoArrange     =   0   'False
                  HeaderButtons   =   0   'False
                  HeaderTrackSelect=   0   'False
                  HideSelection   =   0   'False
                  InfoTips        =   0   'False
               End
            End
            Begin VB.ComboBox cboLookin 
               Height          =   315
               Left            =   780
               TabIndex        =   12
               Top             =   120
               Width           =   3135
            End
            Begin VB.CommandButton CmdFind 
               Caption         =   "Find Now"
               Height          =   375
               Left            =   5700
               TabIndex        =   11
               Top             =   60
               Width           =   1095
            End
            Begin VB.CommandButton CmdStop 
               Caption         =   "&Stop"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5700
               TabIndex        =   10
               Top             =   540
               Width           =   1095
            End
            Begin VB.CommandButton cmdNewSearch 
               Caption         =   "&New Search"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5700
               TabIndex        =   9
               Top             =   960
               Width           =   1095
            End
            Begin VB.PictureBox PicBasic 
               Height          =   1515
               Left            =   480
               ScaleHeight     =   1455
               ScaleWidth      =   4695
               TabIndex        =   8
               Top             =   1920
               Width           =   4755
            End
            Begin VB.Label lblFilters 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Filters:"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   840
               Width           =   450
            End
            Begin VB.Label lblLookIn 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "&Look In:"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   180
               Width           =   585
            End
         End
         Begin VB.PictureBox PicLowerPane 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   0
            ScaleHeight     =   113
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   466
            TabIndex        =   5
            Top             =   4005
            Width           =   6990
            Begin vbalCmdBar6.vbalCommandBar cmdbarmenu 
               Height          =   375
               Index           =   2
               Left            =   60
               Top             =   0
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   661
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   3
            End
            Begin vbaBClListViewLib6.vbalListViewCtl lvwfiles 
               Height          =   1275
               Left            =   135
               TabIndex        =   6
               Top             =   450
               Width           =   6360
               _ExtentX        =   11218
               _ExtentY        =   2249
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               View            =   1
               MultiSelect     =   -1  'True
               AutoArrange     =   0   'False
               HeaderButtons   =   0   'False
               HeaderTrackSelect=   0   'False
               HideSelection   =   0   'False
               InfoTips        =   0   'False
               LabelTips       =   -1  'True
               OLEDropMode     =   1
               ScaleMode       =   3
               DoubleBuffer    =   -1  'True
            End
         End
         Begin VB.PictureBox PicSplit 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   60
            MousePointer    =   7  'Size N S
            ScaleHeight     =   13
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   457
            TabIndex        =   4
            Top             =   3780
            Width           =   6855
         End
      End
   End
   Begin VB.Timer tmrFilterResults 
      Enabled         =   0   'False
      Left            =   7740
      Top             =   4740
   End
   Begin VB.Timer TmrmenuHighlight 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7740
      Top             =   2580
   End
   Begin VB.PictureBox PicStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   666
      TabIndex        =   0
      Top             =   6825
      Width           =   9990
   End
   Begin vbalCmdBar6.vbalCommandBar cmdbarmenu 
      Align           =   1  'Align Top
      Height          =   300
      Index           =   0
      Left            =   0
      Negotiate       =   -1  'True
      Tag             =   "noPersist"
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MainMenu        =   -1  'True
      Style           =   3
   End
   Begin vbalCmdBar6.vbalCommandBar cmdbarmenu 
      Align           =   1  'Align Top
      Height          =   420
      Index           =   1
      Left            =   0
      Negotiate       =   -1  'True
      Tag             =   "noPersist"
      Top             =   300
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl ExplorerBar 
      Height          =   6075
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   10716
      BackColorEnd    =   0
      BackColorStart  =   0
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Implements IContextCallback
Implements IFileSearchCallback
Implements IFilterChangeCallback
Implements IProgress
Implements iSuperClass
'Implements iSuperClass
Private Enum FilterViewMode
    FilterView_Simple
    FilterView_Advanced

End Enum
Private Const HDM_FIRST As Long = &H1200
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
Private Const HDM_HITTEST As Long = (HDM_FIRST + 6)

Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2


Private mSearching As Boolean
Private mInmenu As Boolean
'12/29/2009 12:43:07
'Discovered some bugs in the handling of right-click for columns. Turns out that the listview wwasn't passing the right button arguments.
'I tried a few hacky solutions that tried to take advantage of the curious fact that when right-clicking a listview the "button" argument was the same as the last button argument
'passed to the actual listview (the file list, that is). I ended up realizing something better: when ColumnClick was fired, there was no mousedown or mouseup, so I simply took it on faith that mouse clicks within the header height were in fact on the header.
'seems to work like a charm! :P

Private Currgrouping As String

Private mLvwclasser As cSuperClass
Const Magic_Number = 87495874
'for IContextCallback:
Private Type ColumnProviderData
    ColumnTitle As String
    ColumnDescription As String
    Defwidth As Long 'converted from characters...
    lvwformat As Long
    flags As SHCOLSTATE
    
End Type
'typedef struct {
'    POINT ptReserved;
'    POINT ptMaxSize;
'    POINT ptMaxPosition;
'    POINT ptMinTrackSize;
'    POINT ptMaxTrackSize;
'} MINMAXINFO;

    

'a private var; indicates when nonlocal mode (for viewing saved searches, for example) is active.
Private mNonLocal As Boolean
Private mclasser As cSuperClass
Private mColumnProviders() As IColumnProvider
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer


Private Declare Function AppendMenu Lib "user32.dll" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Const MF_STRING As Long = &H0&
Private Const MF_SEPARATOR As Long = &H800&


'Stored Menu handles from ContextMenu interface methods- we create the popup menus in the Before() method, and destroy them all in the After() method.
Private mScannerMenu As Long 'Popup menu for virus scanners.

Private mSearchPath As String

'Private Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetLogicalDriveStringsW Lib "kernel32.dll" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Enum EHitTestAreas
    HTERROR = (-2)
    HTTRANSPARENT = (-1)
    HTNOWHERE = 0
    HTCLIENT = 1
    HTCAPTION = 2
    HTSYSMENU = 3
    HTGROWBOX = 4
    HTMENU = 5
    HTHSCROLL = 6
    HTVSCROLL = 7
    HTMINBUTTON = 8
    HTMAXBUTTON = 9
    HTLEFT = 10
    HTRIGHT = 11
    HTTOP = 12
    HTTOPLEFT = 13
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTBORDER = 18
End Enum
Private Const WMSZ_BOTTOM As Long = 6
Private Const WMSZ_BOTTOMLEFT As Long = 7
Private Const WMSZ_BOTTOMRIGHT As Long = 8
Private Const WMSZ_LEFT As Long = 1
Private Const WMSZ_RIGHT As Long = 2
Private Const WMSZ_TOP As Long = 3
Private Const WMSZ_TOPLEFT As Long = 4
Private Const WMSZ_TOPRIGHT As Long = 5
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11

Private mSbar As cNoStatusBar
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type


'API routines for making frames transparent....
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
Private Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long



Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Private tt As ExToolTip
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Implements IChangeNotification
Private mTip As clsTooltip
Private mNotify As CFileChangeNotify
'Implements iSuperClass
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private WithEvents mFormSizer As CFormSizeControl
Attribute mFormSizer.VB_VarHelpID = -1
Private WithEvents mtxtSearchBar As TextBox
Attribute mtxtSearchBar.VB_VarHelpID = -1

'typedef struct _HD_HITTESTINFO {
'  POINT pt;
'  UINT flags;
'  int iItem;
'} HD_HITTESTINFO;
Private Const HHT_ONHEADER As Long = &H2

Private Type HD_HITTESTINFO
 pt As POINTAPI
 flags As Long
 iItem As Long
End Type


Private mLookInAutoComplete As CAutoCompleteCombo
Private mFilemaskAutoComplete As CAutoCompleteCombo
Private mmenuImages As cVBALImageList
Private mHeaderIML As cVBALImageList


Private mPaneSplitter As CSplitterBar
Private mtaskSplitter As CSplitterBar
Private mThemeUse As String, mThemeSection As String
Public mFileSearch As CFileSearchEx
Private WithEvents mCmdBarLoader As CXMLLoader
Attribute mCmdBarLoader.VB_VarHelpID = -1
Private mcancel As Boolean
Private SearchAnimation As cAnimControl
Private mDateStrs(0 To 5) As String
Private mmove As Boolean
Private mCurrSelection As Collection 'Collection of selected ListItems in the search results.

'Private mCachedDateUIValues(1 To 3) As mDateUIValues
Private mFormSaver As CFormDataSaver
Private mTotalRecurse As Long, mTotalFileCount As Long
Private testnotify As CDevChangeNotify
Dim mzipfilter As CSearchFilter
Private mFireResizeNextTmr As Boolean 'hack alert.
Private mFilterViewMode As FilterViewMode
Private mSimpleFilter As CSearchFilter
Private mAdvancedFilters As SearchFilters
Private mSearchAnimator As cPicAnimator
Private mSButton As Long, mSShift As Long
Private mCachedRemovedItems As Collection   'cached removed items. Note that it actually holds the string tags, not the listitem objects. this will be set to nothing if the results are not filtered.
Private mSearchColumnskeys As String
Private mMenuTempLookups() As String  'Array used to store the columns between BeforeShowMenu and AfterShowMenu.
Private mexpdragging As Boolean, medragstart As Double
Private Type ItemDataStruct
    ListItemObj As cListItem
    DataObj As CItemExtraData
End Type
Private mItemData() As ItemDataStruct
Private mItemDataCount As Long

Private Sub AddDirectoryExclude(ToFilters As SearchFilters)
   
    Dim firstfilter As CSearchFilter, zipperfilter As CSearchFilter
    Set firstfilter = New CSearchFilter
    Set zipperfilter = New CSearchFilter
    

    firstfilter.Attributes = FILE_ATTRIBUTE_DIRECTORY
    firstfilter.SearchOperation = Filter_Exclude
    firstfilter.Name = "Directory Excluder"
    zipperfilter.FileSpec = "*.zip"
    zipperfilter.SearchOperation = Filter_Or
    zipperfilter.Name = "Zip Viewer"
    ToFilters.Add firstfilter
    ToFilters.Add zipperfilter
  



End Sub
Private Sub UpdateSimpleFilter()
    'called in the Validate routine of each control used for manipulating the "Simple" filter in simple mode. these are found on the PicSimple Picturebox.
    If mSimpleFilter Is Nothing Then
        
        Set mSimpleFilter = New CSearchFilter
        If mFilterViewMode = FilterView_Simple Then
            'perform the same cache-ing of the value that is done in the changing routine...
            Set mAdvancedFilters = mFileSearch.Filters
            Set mFileSearch.Filters = New SearchFilters
            AddDirectoryExclude mFileSearch.Filters
            mFileSearch.Filters.Add mSimpleFilter
        
        End If
    End If
    
    With mSimpleFilter
        .FileSpec = cboSimpleFileMask.Text
        .FileSpecIsRegExp = chkSimpleregExp.Value = vbChecked
    
    End With
End Sub

Public Function ListItemFromSearchFilter(SFilter As CSearchFilter) As cListItem
    Dim I As Long
    
        For I = 1 To lvwfilters.ListItems.Count
            If lvwfilters.ListItems.Item(I).ItemData = Val(SFilter.Tag.Tag) Then
                Set ListItemFromSearchFilter = lvwfilters.ListItems.Item(I)
                Exit Function
            End If
        Next I
        Set ListItemFromSearchFilter = Nothing



End Function
Private Function GetFilterObjectFromlistItem(lvwitem As cListItem) As CSearchFilter

    
    Dim I As Long
    With mFileSearch.Filters
    For I = 1 To .Count
        If Val(.Item(I).Tag.Tag) = lvwitem.ItemData Then
            Set GetFilterObjectFromlistItem = .Item(I)
            Exit Function
        End If
    
    
    Next I
    End With

    Set GetFilterObjectFromlistItem = Nothing

End Function
Private Sub ShowViewPopup(Optional ByVal Attach As Boolean = False)
Dim entrybar As cCommandBar, loopcol As Long, currcol As cColumn, cursorX As Long, cursorY As Long, pointap As POINTAPI
  GetCursorPos pointap
   cursorX = pointap.X
   cursorY = pointap.Y
    On Error Resume Next
     Set entrybar = cmdbarmenu(0).CommandBars("FileColumnBar")
    If Err <> 0 Then
        Err.Clear
        Set entrybar = cmdbarmenu(0).CommandBars.Add("FileColumnBar", "Show/Hide Columns")
        
        
        
    End If
    Dim newbutton As cButton
    entrybar.Buttons.Clear
    For loopcol = 1 To lvwfiles.columns.Count
    Set currcol = lvwfiles.columns(loopcol)
        Err.Clear
        Set newbutton = cmdbarmenu(0).Buttons.Item(currcol.Key)
        If Err <> 0 Then
            Set newbutton = cmdbarmenu(0).Buttons.Add(currcol.Key, , currcol.Text, eCheck)
        End If
        newbutton.checked = Not (currcol.Width = 0)
        newbutton.Tag = "SHOWHIDE:" & currcol.Key
        entrybar.Buttons.Add newbutton
    Next
    On Error Resume Next
    cmdbarmenu(0).Buttons.Add "SEPFileLvwList", , "-", eSeparator
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("SEPFileLvwList")
    cmdbarmenu(0).Buttons.Add("LVWSHOW:ALL", , "Show All Columns:").Tag = "LVWSHOW:ALL"
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWSHOW:ALL")
    cmdbarmenu(0).Buttons.Add "SEPFileLvwList2", , "View", eSeparator
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("SEPFileLvwList2")
    'Dim listviewbar As cCommandBar
    'Set listviewbar = CmdBarMenu(0).CommandBars.Add("LISTVIEWMODEBAR", "List View Display Mode")
    cmdbarmenu(0).Buttons.Add("LVWColumnViewDetails", mmenuImages.ItemIndex("DETAILS"), "Details", eCheck).Tag = "SetView:Details"
    cmdbarmenu(0).Buttons.Add("LVWColumnViewIcon", mmenuImages.ItemIndex("ICON"), "Icon", eCheck).Tag = "SetView:Icon"
    cmdbarmenu(0).Buttons.Add("LVWColumnViewList", mmenuImages.ItemIndex("LIST"), "List", eCheck).Tag = "SetView:List"
    cmdbarmenu(0).Buttons.Add("LVWColumnViewSmallIcon", mmenuImages.ItemIndex("SMALLICONS"), "Small Icons", eCheck).Tag = "SetView:SmallIcon"
    cmdbarmenu(0).Buttons.Add("LVWColumnViewTile", , "Tile", eCheck).Tag = "SetView:Tile"
    
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWColumnViewDetails")
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWColumnViewIcon")
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWColumnViewList")
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWColumnViewSmallIcon")
    entrybar.Buttons.Add cmdbarmenu(0).Buttons("LVWColumnViewTile")
   'remove the check from each view item:
    cmdbarmenu(0).Buttons("LVWColumnViewDetails").checked = False
     cmdbarmenu(0).Buttons("LVWColumnViewIcon").checked = False
     cmdbarmenu(0).Buttons("LVWColumnViewList").checked = False
     cmdbarmenu(0).Buttons("LVWColumnViewSmallIcon").checked = False
     cmdbarmenu(0).Buttons("LVWColumnViewTile").checked = False
   
   
   'set the Checked item to the appropriate view style...
    Select Case lvwfiles.View
        Case eViewIcon
            cmdbarmenu(0).Buttons.Item("LVWColumnViewIcon").checked = True
        Case eViewSmallIcon
            cmdbarmenu(0).Buttons.Item("LVWColumnViewSmallIcon").checked = True
        Case eViewList
            cmdbarmenu(0).Buttons.Item("LVWColumnViewTile").checked = True
        Case eViewDetails
            cmdbarmenu(0).Buttons.Item("LVWColumnViewDetails").checked = True
        Case eViewTile
            cmdbarmenu(0).Buttons.Item("LVWColumnViewTile").checked = True
    End Select
    
    
    
    'Set CmdBarMenu(0).Buttons.Add("Lvwviewmode", , "View", eNormal).Bar = listviewbar
    'entrybar.Buttons.Add CmdBarMenu(0).Buttons("lvwviewmode")
    
    
    
    
    
    If Attach Then
    Set cmdbarmenu(0).Buttons("VIEW").bar = entrybar
    
    Else
    cmdbarmenu(0).ShowPopupMenu cursorX, cursorY, entrybar
    End If
   'cancelled = True
   Exit Sub
End Sub
Private Sub GenFileListing(fname As String, ProgressObject As IProgress)
Dim StreamOut As FileStream, loopfile As CFile, I As Long
Dim loopItem As cListItem
If fname = "" Then Exit Sub
If lvwfiles.ListItems.Count = 0 Then Exit Sub
Dim MaxFileNameSize As Long, maxSizeLen As Long, lencurr As Long
Dim MaxFilterNameLength As Long, MaxFilterFilterLength As Long
For I = 1 To lvwfiles.ListItems.Count
    Set loopItem = lvwfiles.ListItems(I)
    If Len(loopItem.Tag) > MaxFileNameSize Then MaxFileNameSize = Len(loopItem.Tag)
    If loopItem.Tag = "" Then
        loopItem.Tag = loopItem.SubItems(1).caption & loopItem.Text
    
    End If
    Set loopfile = GetFile(loopItem.Tag)
    lencurr = Len(FormatSize(loopfile.size))
    If lencurr > maxSizeLen Then maxSizeLen = lencurr
    
    
Next I
Dim LoopFilter As CSearchFilter, countfilters As Long
'streamout.writestring
Set StreamOut = bcfile.CreateStream(fname)
StreamOut.WriteString String$(80, "-") & vbCrLf
StreamOut.WriteString "Filters list:" & vbCrLf
Dim CurrFilter As Long
ProgressObject.UpdateUI 0, "Writing Active Filter info..."
For CurrFilter = 1 To mFileSearch.Filters.Count
    Set LoopFilter = mFileSearch.Filters.Item(CurrFilter)
    countfilters = countfilters + 1
    StreamOut.WriteString "Filter #" & Trim$(Str$(countfilters))
    StreamOut.WriteString "Filter name:    """ & LoopFilter.Name & """" & vbCrLf
    StreamOut.WriteString "Filter Spec:    """ & LoopFilter.FileSpec & """" & vbCrLf
    StreamOut.WriteString "Filter Type:    """ & LoopFilter.GetMatchTypeString & """" & vbCrLf
    StreamOut.WriteString String$(80, "-")


Next



     'Set StreamOut = CreateStream(fname)
     StreamOut.WriteString "BCSearch File Search Listing"
     StreamOut.WriteString vbCrLf
     StreamOut.WriteString "Created " & FormatDateTime(Now, vbLongDate) & " at " & FormatDateTime(Now, vbLongTime)
     StreamOut.WriteString vbCrLf
     StreamOut.WriteString vbCrLf
     StreamOut.WriteString "Filename" & Space$(MaxFileNameSize - Len("Filename") + 1)
     StreamOut.WriteString "Size" & Space$(maxSizeLen - Len("Size") + 1) & vbCrLf
        For I = 1 To lvwfiles.ListItems.Count
        'For Each LoopItem In lvwFiles.ListItems
        ProgressObject.UpdateUI CDbl(lvwfiles.ListItems.Count) / CDbl(I), "Writing... (" & I & " of " & lvwfiles.ListItems.Count & ") records processed."
        Set loopItem = lvwfiles.ListItems(I)
            On Error Resume Next
            Set loopfile = GetFile(loopItem.Tag)
            If Err <> 0 Then
                StreamOut.WriteString "*** ERROR: could not access """ & loopItem.Tag & """- error #" & Err.Number & "(" & Err.Description & ")" & vbCrLf
            Else
            StreamOut.WriteString loopfile.FullPath & Space$(MaxFileNameSize - Len(loopfile.FullPath) + 1)
            StreamOut.WriteString FormatSize(loopfile.size) & Space$(maxSizeLen - Len(FormatSize(loopfile.size)))
            StreamOut.WriteString vbCrLf
                'StreamOut.WriteString loopfile.FullPath & ", Attr:" & loopfile.GetAttributeString & ", Size:" & loopfile.Size & "," & " Created:" & loopfile.DateCreated & "," & " Modified:" & loopfile.DateModified & _
                '", Accessed:" & loopfile.DateLastAccessed & vbCrLf
                
            End If
            StreamOut.Flush
            
            
            
        
        
        Next
        StreamOut.CloseStream
End Sub
Public Property Get SaveControlStates()
    SaveControlStates = True
End Property

Private Function addfilter(ByVal filterSpec As String) As CSearchFilter

'
    Dim newfilter As CSearchFilter, added As cListItem
    Dim newform As FrmSearchSpec
    Static NewKeyIndex  As Long
    Set newfilter = New CSearchFilter
    Set newfilter.Callback = Me
    Set newfilter.Tag = New CExtraFilterData
    'Set newform = New FrmSearchSpec
    
    'newform.EditFilter newfilter, Me
    
    CDebug.Post newfilter.FileSpec
    
    mFileSearch.Filters.Add newfilter
    NewKeyIndex = NewKeyIndex + 1
    newfilter.FileSpec = filterSpec
    newfilter.Name = "Filter" & Trim$(NewKeyIndex)
    newfilter.Tag.Tag = Int(Rnd * 16777216#)
    'with lvwFilters.ListItems.Add ,,"NONAME"
    Set added = lvwfilters.ListItems.Add(, "TAG" & Trim$(Str$(NewKeyIndex)))
    added.ItemData = Val(newfilter.Tag.Tag)
    PopulateFilterItem added, newfilter
    Set addfilter = newfilter
End Function
Private Sub cboLookin_Validate(Cancel As Boolean)
    Dim IsDir As Boolean
    Dim splitdirs() As String, I As Long
    Dim gotfilter As String, rebuildPath() As Variant
    'if isDirectory(cbolookin.Text)
    If InStr(cboLookin.Text, ";") > 0 Then
    splitdirs() = Split(cboLookin.Text, ";")
    Else
        ReDim splitdirs(0)
        splitdirs(0) = cboLookin.Text
    
    End If
    
    ReDim rebuildPath(0 To UBound(splitdirs))
    
    
    
    For I = 0 To UBound(splitdirs)
    
        'if it isn't a directory, take the file part of the name and create a new filter from it.
        If Not IsDirectory(splitdirs(I)) Then
            gotfilter = GetFilenamePart(splitdirs(I))
            rebuildPath(I) = Mid$(splitdirs(I), 1, Len(splitdirs(I)) - Len(gotfilter))
    
            'add the filter...
            Dim addedfilter As CSearchFilter
            Set addedfilter = addfilter(gotfilter)
            addedfilter.ScriptLanguage = "VBScript"
            addedfilter.ScriptCode = "Function DoFilter(FileObj)" & vbCrLf & _
            "DoFilter=Instr(Fileobj.FullPath,""" & rebuildPath(I) & """) > 0" & vbCrLf & _
            "End Function"
        Else
            rebuildPath(I) = splitdirs(I)
        End If
    Next I
    cboLookin.Text = Join(rebuildPath, ";")
    
    
  
    If Not bcfile.Exists(cboLookin.Text) Then
'        MsgBox "The specified Directory Could not be found, or you do not have permissions to access it."
'        cboLookin.SetFocus
'        cboLookin.SelStart = 1
'        cboLookin.SelLength = Len(cboLookin.Text)
        
        'cboLookin.ForeColor = vbRed
        'Call mTip.ShowToolTip(cboLookin.hwnd, "Location Not found", "the location specified, """ & cboLookin.Text & """ could not be found on the system. Enter a valid Folder Name and try again.", 0, 80)
    Else
        cboLookin.BackColor = vbWindowBackground
        cboLookin.ForeColor = vbWindowText
    
    End If
End Sub





Private Sub cboSimpleFileMask_Validate(Cancel As Boolean)
UpdateSimpleFilter
End Sub

Private Sub chkSimpleregExp_Validate(Cancel As Boolean)
UpdateSimpleFilter
End Sub

Private Sub CmdBarMenu_ButtonDropDown(Index As Integer, btn As vbalCmdBar6.cButton, Cancel As Boolean)
    CDebug.Post "buttondropdown " & btn.Key & ";" & btn.Tag
    
    Dim bar As cCommandBar
    Set bar = btn.bar
    If bar.Key = "FILELVW::SORTASCENDING::SUBMENU" Or bar.Key = "FILELVW::SORTDESCENDING::SUBMENU" Or Left$(bar.Key, 6) = "SORT::" Then
    CmdBarMenu_BeforeShowMenu Index, bar
'    'Clear the buttons. because I wanna.
'    bar.Buttons.Clear
'     Dim flDescending As Boolean, columnI As Long
'     flDescending = (bar.Key = "FILELVW::SORTDESCENDING::SUBMENU")
'     Dim sortmode As String
'
'     'What do we do here? Populate the bar with buttons that represent each column. The name of the button will be either "SORTASCENDING::<ListviewName>::<COLUMNKEY>" or "SORTDESCENDING::<ListViewName>::<COLUMNKEY>"
'     Dim addkey As String
'     Dim AddedButton As cButton, loopcolumn As cColumn
'     sortmode = IIf(flDescending, "SORTDESCENDING", "SORTASCENDING")
'     For columnI = 1 To lvwfiles.Columns.Count
'        Set loopcolumn = lvwfiles.Columns(columnI)
'        addkey = sortmode & "::" & "lvwfiles::" & loopcolumn.Key
'        On Error Resume Next
'        Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, loopcolumn.IconIndex, "By " & loopcolumn.Text)
'        If Err.Number <> 0 Then
'         Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
'
'        End If
'        bar.Buttons.Add AddedButton
'     Next columnI
     
    
    End If
    
'        PopulateScanMenu bar
End Sub

Private Sub CmdBarMenu_ButtonHighlight(Index As Integer, btn As vbalCmdBar6.cButton)
'CDebug.Post  "highlight:" & btn.caption
Dim StatusBarText As String
StatusBarText = mCmdBarLoader.GetMenuItemXMLAttribute(btn.Key, "StatusText")
If StatusBarText = "" Then StatusBarText = btn.ToolTip
If StatusBarText <> "" Then
    
    mSbar.SimpleMode = True
    mSbar.SimpleText = StatusBarText
    TmrmenuHighlight.enabled = True
    'reset the delay...
    TmrmenuHighlight.Tag = 0
    'mInmenu = True

End If



End Sub

Private Sub CmdBarMenu_RightClick(Index As Integer, btn As vbalCmdBar6.cButton, ByVal X As Long, ByVal Y As Long)
'
End Sub

Private Sub cmdBrowse_Click()
    Dim browseObj As CDirBrowser, gotfolder As Directory
    Set browseObj = New CDirBrowser
    Set gotfolder = browseObj.BrowseForDirectory(Me.hWnd, "Select Folder", BIF_EDITBOX + BIF_NEWDIALOGSTYLE + BIF_RETURNONLYFSDIRS)
    If Not gotfolder Is Nothing Then
    cboLookin.Text = gotfolder.Path
    End If
End Sub

Private Sub cmdEditFilter_Click()
    Dim getitem As cListItem, newform As FrmSearchSpec, getfilter As CSearchFilter
    Set getitem = lvwfilters.SelectedItem
    If getitem Is Nothing Then Exit Sub
    Set getfilter = GetFilterObjectFromlistItem(getitem)
    Set newform = New FrmSearchSpec
    newform.EditFilter getfilter, Me, "Edit Filter", False
End Sub

Private Sub CmdremoveFilter_Click()
    Dim itemget As CSearchFilter
    'remove the selected filter.
    If Not lvwfilters.SelectedItem Is Nothing Then
        Set itemget = GetFilterObjectFromlistItem(lvwfilters.SelectedItem)
        mFileSearch.Filters.Remove mFileSearch.Filters.getindex(itemget)
        
        lvwfilters.ListItems.Remove lvwfilters.SelectedItem.Index
    
    End If
End Sub

Private Sub ExplorerBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    Debug.Print "itemclick:" & itm.Key
'SINGLEITEM:: Copy
'SINGLEITEM::Move
'SINGLEITEM::Print
'SINGLEITEM::Rename
'SINGLEITEM::Delete
Dim confirmdelete As Boolean
Dim dirbrowse As CDirBrowser, Targetdirectory As Directory
Dim grabfile As CFile
Set dirbrowse = New CDirBrowser


Select Case UCase$(itm.Key)
    
    Case "SINGLEITEM::COPY"
        Set Targetdirectory = dirbrowse.BrowseForDirectory(Me.hWnd, "Copy " & mCurrSelection.Item(1).Text, BIF_RETURNONLYFSDIRS + BIF_USENEWUI)
        Debug.Print "Copy file " & mCurrSelection.Item(1).Text & " to " & Targetdirectory.FullPath
        Call bcfile.PerformFileOperation(mCurrSelection.Item(1).Tag, Targetdirectory.FullPath, FO_COPY, 0, Me.hWnd, False, "Copying...")
    Case "SINGLEITEM::MOVE"
     Set Targetdirectory = dirbrowse.BrowseForDirectory(Me.hWnd, "Move " & mCurrSelection.Item(1).Text, BIF_RETURNONLYFSDIRS + BIF_USENEWUI)
        Debug.Print "Move file " & mCurrSelection.Item(1).Text & " to " & Targetdirectory.FullPath
        Call bcfile.PerformFileOperation(mCurrSelection.Item(1).Tag, Targetdirectory.FullPath, FO_MOVE, 0, Me.hWnd, False, "Moving...")
    Case "SINGLEITEM::PRINT"
        GetFile(mCurrSelection.Item(1).Text).Execute Me.hWnd, "Print"
    Case "SINGLEITEM::RENAME"
       ' lvwfiles.LabelEdit = True
       Dim newname As String
       Dim createpath As String
       'createpath = bcfile.GetPathPart(mCurrSelection.Item(1).Text)
       createpath = Mid$(mCurrSelection.Item(1).Tag, 1, InStrRev(mCurrSelection.Item(1).Tag, "\"))
       If Right$(createpath, 1) <> "\" Then createpath = createpath & "\"
       newname = InputBox$("Rename File " & mCurrSelection.Item(1).Text & " to:", "Rename", "")
       
        If newname <> "" Then
            newname = PerformNameSubstitution(mCurrSelection.Item(1), newname)
            
            
            createpath = createpath & newname
            If bcfile.PerformFileOperation(mCurrSelection.Item(1).Tag, createpath, FO_RENAME, 0, Me.hWnd, True, "Renaming") = 0 Then
                'successful... change the name of the item...
                
                RefreshItemData mCurrSelection.Item, bcfile.GetFile(createpath)
            
            End If
        End If
       
       
     
       
       
       
    Case "SINGLEITEM::DELETE"
        Set grabfile = GetFile(mCurrSelection.Item(1).Tag)
        If GetKeyState(VK_SHIFT) = 0 Then
            confirmdelete = MsgBox("Are you sure you want to delete " & grabfile.basename & "." & grabfile.Extension & "?", vbYesNo + vbQuestion, "Confirm File delete") = vbYes
            
        
        End If
    Case "MULTIITEM::COPY"
    Case "MULTIITEM::MOVE"
    Case "MULTIITEM::PRINT"
    Case "MULTIITEM::RENAME"
    Case "MULTIITEM::DELETE"
End Select
    
'MULTIITEM:: Copy
'MULTIITEM::Move
'MULTIITEM::Print
'MULTIITEM::Rename
'MULTIITEM::Delete
    
    
    
End Sub

Private Sub Form_Initialize()
    Set mFormSaver = New CFormDataSaver
        mFormSaver.Initialize Me, CurrApp.Settings
    

End Sub
Public Function MainMenu() As Object
    Set MainMenu = cmdbarmenu(0)

End Function

Private Function GetSelectedCount(OfView As Long)

    GetSelectedCount = SendMessage(OfView, LVM_GETSELECTEDCOUNT, 0&, 0&)


End Function


Private Sub PopulateFilterItem(Mitem As cListItem, FilterObj As CSearchFilter)
'    With lvwFilters
'        .Columns.Add , "TYPE", "Type"
'        .Columns.Add , "NAME", "Name"

'        .Columns.Add , "FILESPEC", "FileSpec"
'        .Columns.Add , "ATTRIBUTES", "Attributes"
'        .Columns.Add , "LARGERTHAN", "Larger Than:"
'        .Columns.Add , "SMALLERTHAN", "Smaller Than"
'        .Columns.Add , "STARTDATESPECS", "After Date"
'        .Columns.Add , "BEFOREDATESPECS", "Before Date"
'
'    End With





    With Mitem
            
            .Text = FilterObj.GetMatchTypeString
            .SubItems(1).caption = FilterObj.Name
            
            .SubItems(2).caption = FilterObj.FileSpec
            .SubItems(3).caption = GetAttributeString(FilterObj.Attributes, False)
            '.Tag = "TAG" & Trim$(mFileSearch.Filters.Count)
            .SubItems(4).caption = bcfile.FormatSize(FilterObj.SizeLargerThan, 0)
            .SubItems(5).caption = bcfile.FormatSize(FilterObj.SizeSmallerThan, 0)
            '.SubItems(6).Caption=
            
    End With
End Sub



Private Sub cmdAddFilter_Click()
    Dim newfilter As CSearchFilter, added As cListItem
    Dim newform As FrmSearchSpec
    Static NewKeyIndex  As Long
    Set newfilter = New CSearchFilter
    Set newfilter.Callback = Me
    Set newfilter.Tag = New CExtraFilterData
    Set newform = New FrmSearchSpec
    newform.EditFilter newfilter, Me, "Add Filter", True
    
    CDebug.Post newfilter.FileSpec
    
    mFileSearch.Filters.Add newfilter
    NewKeyIndex = NewKeyIndex + 1
    newfilter.Name = "Filter" & Trim$(NewKeyIndex)
    newfilter.Tag.Tag = Int(Rnd * 16777216#)
    'with lvwFilters.ListItems.Add ,,"NONAME"
    Set added = lvwfilters.ListItems.Add(, "TAG" & Trim$(Str$(NewKeyIndex)))
    added.ItemData = Val(newfilter.Tag.Tag)
    PopulateFilterItem added, newfilter
    
End Sub
Private Sub PopulateScanMenu(BarPopulate As vbalCmdBar6.cCommandBar)
    'populate this commandbar with items for known AV/malware programs.
    Dim gotmbampath As String
    Dim parsedPath() As String
    
'HKEY_CLASSES_ROOT\mbam.script\shell\open\command "default" value is "mbam install location" %1
 With CurrApp.registry
    .Classkey = hhkey_classes_root
    .SectionKey = "mbam.script\shell\open\command"
    .ValueKey = ""
    gotmbampath = .Value
  '  parsedPath = ParseCommandLine(gotmbampath)
 
 End With
    
    
    




End Sub
'script access functions...
'VBScript and the commandbar don't quite mix...
Public Sub AddButtonToBar(BarTo, Button)
    Dim castbar As cCommandBar
    Dim castbutton As cButton
    CDebug.Post "addButtonToBar"
    Set castbar = BarTo
    Set castbutton = Button
    castbar.Buttons.Add castbutton



End Sub

Private Sub CmdBarMenu_BeforeShowMenu(Index As Integer, bar As vbalCmdBar6.cCommandBar)

        Dim SendToFolder   As Directory, loopfile As CFile

        Dim iconkey        As String

        Dim newbutton      As cButton

        Dim Loopobj        As CFile, FilePathSend As String

        Dim loopcolumn     As cColumn, columnI As Long

        Dim splitstrings() As String

        Dim addkey         As String

        Dim AddedButton    As cButton, useiconindex As Long

        On Error GoTo reportandresume

        CDebug.Post "BEFORESHOWMENU:" & bar.Key
        Dim I As Long, currbutton As cButton
        Static SendToInitialized As Boolean
       
        
        
        If bar.Key = "TOOLBAR::FILTERSEARCHIN::SUBMENU" Then
            'populate with column names...
            bar.Buttons.Clear
            For I = lvwfiles.columns.Count To 1 Step -1
                Set loopcolumn = lvwfiles.columns(I)
                'only if it's visible, though.
                If loopcolumn.Width > 5 Then
                    
                    Set currbutton = CreateButton("FILTERSEARCH::COLUMN::" & Trim$(loopcolumn.Key), , loopcolumn.Text, eCheck)
                    If InStr(mSearchColumnskeys, "&+" & loopcolumn.Key & "&+") > 0 Then
                        currbutton.checked = True
                    End If
                    
                    
                    bar.Buttons.Add currbutton
                End If
            Next I
        ElseIf bar.Key = "VIEW::MODE::SUBMENU" Then
            bar.Buttons("VIEW::MODE::SIMPLE").checked = mFilterViewMode = FilterView_Simple
            bar.Buttons("VIEW::MODE::ADVANCED").checked = mFilterViewMode = FilterView_Advanced
            cmdbarmenu(0).Buttons.Item("VIEW::MODE::SIMPLE").checked = bar.Buttons("VIEW::MODE::SIMPLE").checked
            cmdbarmenu(0).Buttons.Item("VIEW::MODE::ADVANCED").checked = bar.Buttons("VIEW::MODE::ADVANCED").checked
            If cmdbarmenu(0).Buttons.Item("VIEW::MODE::SIMPLE").checked = cmdbarmenu(0).Buttons.Item("VIEW::MODE::ADVANCED").checked Then
               ' Debug.Assert False
            End If
        ElseIf bar.Key = "POPUP::COPYCLIP::SUBMENU" Then
            'add a item for each column
            
            'First, COPYCLIP::ALLVISIBLE  'copy all visible columns to clipboard (tab delimited)
            'and a separator.
            
            'COPYCLIP::Columnkey
            bar.Buttons.Clear
            'COPYCLIP::ALLVISIBLE "All Visible Columns"
            'COPYCLIP::ALL "All Visible"
            'COPYCLIP::SEPARATOR separator...
            'COPYCLIP::COLUMN::KEY  'key name

            Set currbutton = CreateButton("COPYCLIP::ALLVISIBLE", -1, "All Visible Columns", eNormal, "Copy all Visible Columns to the clipboard")
            bar.Buttons.Add currbutton
            Set currbutton = CreateButton("COPYCLIP::ALL", -1, "All Columns", eNormal, "Copy All Column values to the clipboard.")
            bar.Buttons.Add currbutton
            Set currbutton = CreateButton("COPYCLIP:SEPARATOR", -1, , eSeparator)
            bar.Buttons.Add currbutton
            
            'add one item for each visible column.
            
            
            
            
            
            
        
            For I = lvwfiles.columns.Count To 1 Step -1
            Set loopcolumn = lvwfiles.columns(I)
            'only if it's visible, though.
            If loopcolumn.Width > 5 Then
                Set currbutton = CreateButton("COPYCLIP::COLUMN::" & Trim$(loopcolumn.Key), , loopcolumn.Text, eNormal)
                
            
            
                bar.Buttons.Add currbutton
            End If
            Next I
            
            
            
        ElseIf bar.Key = "VIEW::GROUP::SUBMENU" Then
            'tasks:
            'if Currgrouping is "", then check the "none" item.
            'otherwise, highlight the "GROUP::<currgrouping>" item.
            
            
            'Dim I As Long
            For I = 1 To bar.Buttons.Count
                If Left$(bar.Buttons(I).Key, 7) = "GROUP::" Then
                    bar.Buttons(I).checked = False
                End If
                If bar.Buttons(I).Key = "GROUP::NONE" And Currgrouping = "" Then
                    bar.Buttons(I).checked = True
                ElseIf bar.Buttons(I).Key = "GROUP::" & UCase$(Trim$(Currgrouping)) Then
                    bar.Buttons(I).checked = True
                End If
                
            Next I
            
                
            
            
            
        ElseIf bar.Key = "TOOLBAR::EXPORT::SUBMENU" Then
        With cmdbarmenu(0).Buttons("EXPORT_WORD")
            If gWordInstalled Then .Visible = True Else .Visible = False
        End With
        
        
        
        ElseIf Left$(bar.Key, 9) = "SHOWDIR::" Then
            bar.Buttons.Clear
            'SHOWDIR::FOLDERPATH
            splitstrings = Split(bar.Key, "::")

            Dim FileWalker As CDirWalker, showfolder As Directory

            Dim CurrFile   As CFile
            Static INIFile As CINIData, INIFilename As String
            
            
            
            splitstrings(1) = ExpandEnvironment(splitstrings(1))
            Set showfolder = bcfile.GetDirectory(splitstrings(1))
            Set FileWalker = showfolder.GetWalker(splitstrings(2), 0, FILE_ATTRIBUTE_DIRECTORY)

            'INIFilename = showfolder.Files.
            Dim newINIfilename
            newINIfilename = showfolder.Path & showfolder.Name & ".ini"
            If StrComp(INIFilename, newINIfilename, vbTextCompare) <> 0 Or INIFile Is Nothing Then
             On Error Resume Next
                INIFilename = newINIfilename
                Set INIFile = New CINIData
                INIFile.LoadINI INIFilename
                If Err <> 0 Then
                    Set INIFile = Nothing
                    INIFilename = ""
                End If
            End If
            Do Until FileWalker.GetNext(CurrFile) Is Nothing
               If StrComp(CurrFile.basename, showfolder.Name, vbTextCompare) <> 0 Then
                addkey = "SHOWFILE::" & CurrFile.FullPath
                Dim gotfileicon As Long, gotfilecaption As String
                Dim fileiconstr As String, filecaptionstr As String
                On Error Resume Next
                If Not INIFile Is Nothing Then
                    If INIFile.SectionExists(CurrFile.Filename) Then
                    'icon and title items.
                    fileiconstr = INIFile.ReadProfileSetting(CurrFile.Filename, "icon", CurrFile.FullPath)
                    fileiconstr = ExpandEnvironment(fileiconstr)
                    filecaptionstr = INIFile.ReadProfileSetting(CurrFile.Filename, "title", CurrFile.Filename)
                    
                    Else
                        fileiconstr = CurrFile.FullPath
                        filecaptionstr = CurrFile.Filename
                    
                    
                    
                    End If
                
                
                End If



                iconkey = "FileImage:" & CurrFile.FullPath
                Call mmenuImages.AddFromHandle(bcfile.GetFile(fileiconstr).GetFileIcon(ICON_SMALL), IMAGE_ICON, iconkey)
        
                Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, mmenuImages.ItemIndex(iconkey), filecaptionstr)

                If Err <> 0 Then
                    Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
                    AddedButton.caption = CurrFile.Filename
                    AddedButton.IconIndex = mmenuImages.ItemIndex(iconkey)
                End If
     
                bar.Buttons.Add AddedButton
            End If
            Loop

        ElseIf Left$(bar.Key, Len("SORT::COLUMN::")) = "SORT::COLUMN::" Then
     
            splitstrings = Split(bar.Key, "::")

            Select Case splitstrings(UBound(splitstrings) - 1)
        
                Case "ASCENDING", "DESCENDING"
                    '0,SORT,1,COLUMN,2,<column key>,3,ASCENDING or DESCENDING.
        
                Case "SORTASCENDING", "SORTDESCENDING"
        
                Case Else
        
                    bar.Buttons.Clear

                    On Error Resume Next

                    Dim coltitle As String

                    coltitle = lvwfiles.columns.Item(splitstrings(2)).Text
                    addkey = "SORT::COLUMN::" & splitstrings(2) & "::" & "SORTASCENDING"
                    Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, , "Sort " & coltitle & " Ascending", eCheck)

                    If Err <> 0 Then
                        Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
                    End If

                    If lvwfiles.columns.Item(splitstrings(2)).IconIndex = mHeaderIML.ItemIndex("ASCENDING") Then
                        AddedButton.checked = True
                    Else
                        AddedButton.checked = False
                    End If

                    bar.Buttons.Add AddedButton
                    addkey = "SORT::COLUMN::" & splitstrings(2) & "::" & "SORTDESCENDING"
                    Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, , "Sort Descending", eCheck)

                    If Err <> 0 Then
                        Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
                    End If

                    If lvwfiles.columns.Item(splitstrings(2)).IconIndex = mHeaderIML.ItemIndex("DESCENDING") Then
                        AddedButton.checked = True
                    Else
                        AddedButton.checked = False
                    End If
        
                    bar.Buttons.Add AddedButton
        
            End Select
     
        ElseIf bar.Key = "SORT::FILELVW::SUBMENU" Then
     
            'the sort submenu
            'contains all the columns as buttons, named SORT::COLUMN::<NAME>
            'each one will be populated with two buttons, "SORT::COLUMN::<NAME>::ASCENDING, and DESCENDING respectively.
            bar.Buttons.Clear

            For columnI = 1 To lvwfiles.columns.Count
                Set loopcolumn = lvwfiles.columns(columnI)
                addkey = "SORT::COLUMN::" & loopcolumn.Key

                On Error Resume Next

                If loopcolumn.IconIndex > 0 Then
                    If loopcolumn.SortOrder = eSortOrderAscending Then useiconindex = mmenuImages.ItemIndex("SORTEDASCENDING") Else useiconindex = mmenuImages.ItemIndex("SORTEDDESCENDING")
                Else
                    useiconindex = -1
                End If

                Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, useiconindex, "Sort by """ & loopcolumn.Text & """")

                If Err.Number <> 0 Then
                    Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
                    AddedButton.IconIndex = useiconindex
                End If

                bar.Buttons.Add AddedButton

                Dim madebar As cCommandBar

                On Error Resume Next

                Set madebar = cmdbarmenu(0).CommandBars.Add(addkey & "::SUBMENU")

                If Err.Number <> 0 Then
                    Set madebar = cmdbarmenu(0).CommandBars.Item(addkey & "::SUBMENU")
                End If

                madebar.Buttons.Add cmdbarmenu(0).Buttons("GHOST")
                Set AddedButton.bar = madebar
            Next
     
        ElseIf bar.Key = "FILELVW::SORTASCENDING::SUBMENU" Or bar.Key = "FILELVW::SORTDESCENDING::SUBMENU" Then
            'Clear the buttons. because I wanna.
            bar.Buttons.Clear

            Dim flDescending As Boolean

            flDescending = (bar.Key = "FILELVW::SORTDESCENDING::SUBMENU")

            Dim sortmode As String
     
            'What do we do here? Populate the bar with buttons that represent each column. The name of the button will be either "SORTASCENDING::<ListviewName>::<COLUMNKEY>" or "SORTDESCENDING::<ListViewName>::<COLUMNKEY>"

            sortmode = IIf(flDescending, "SORTDESCENDING", "SORTASCENDING")

            For columnI = 1 To lvwfiles.columns.Count
                Set loopcolumn = lvwfiles.columns(columnI)
                addkey = sortmode & "::" & "lvwfiles::" & loopcolumn.Key

                On Error Resume Next
        
                If loopcolumn.IconIndex > 0 Then
                    If loopcolumn.SortOrder = eSortOrderAscending Then useiconindex = mmenuImages.ItemIndex("SORTEDASCENDING") Else useiconindex = mmenuImages.ItemIndex("SORTEDDESCENDING")
                Else
                    useiconindex = -1
                End If
        
                Set AddedButton = cmdbarmenu(0).Buttons.Add(addkey, useiconindex, "By " & loopcolumn.Text)

                If Err.Number <> 0 Then
                    Set AddedButton = cmdbarmenu(0).Buttons.Item(addkey)
        
                End If

                AddedButton.IconIndex = useiconindex
                bar.Buttons.Add AddedButton
            Next columnI
    
        ElseIf bar.Key = "POPUP::SCAN::SUBMENU" Then
            'Bar.Buttons.Add CmdBarMenu(0).Buttons.Add("TEST::SCANNER1", , "Norton", , "Don't scan with this, you fool.")
            'Bar.Buttons.Add CmdBarMenu(0).Buttons.Add("TEST::SCANNER2", , "Mcaffee", , "Don't scan with this, either.")
    
        ElseIf bar.Key = "VIEW::FILECONTEXTMENU::SUBMENU" Then
            'bar.Buttons.Clear
        ElseIf bar.Key = "VIEW::SUBMENU" Then
            cmdbarmenu(0).Buttons.Item("VIEW::FILTERSPANE").checked = PicUpperPane.Visible
            cmdbarmenu(0).Buttons.Item("VIEW::EXPLORERBAR").checked = ExplorerBar.Visible
            
        ElseIf bar.Key = "VIEW" Or bar.Key = "TOOLBAR::VIEW" Then
    
            'determine which item should recieve the checkmark.
            Dim buttonloop As cButton, CurrIndex As Long

            'Dim x As EViewStyleConstants
            '    <MENU NAME="VIEW::SETTILE" OPERATION="SetView:Tile" CAPTION="{}Tile"></MENU>
            '<MENU NAME="VIEW::SETICON" OPERATION="SetView:Icon" CAPTION="{}Icons"></MENU>
            '<MENU NAME="VIEW::.SETSMALLICON" OPERATION="SetView:SmallIcon" CAPTION="{}Small Icons"></MENU>
            '<MENU NAME="VIEW::SETLIST" OPERATION="SetView:List" CAPTION="{}List"></MENU>
            '<MENU NAME="VIEW::SETDETAILS" OPERATION="SetView:Details" CAPTION="{}Details"></MENU>
            For CurrIndex = 1 To bar.Buttons.Count
                bar.Buttons(CurrIndex).checked = False

            Next
            
            Select Case lvwfiles.View

                Case eViewIcon
                    cmdbarmenu(0).Buttons("VIEW::SETICON").checked = True
            
                Case eViewSmallIcon
                    cmdbarmenu(0).Buttons("VIEW::SETSMALLICON").checked = True

                Case eViewList
                    cmdbarmenu(0).Buttons("VIEW::SETLIST").checked = True

                Case eViewDetails
                    cmdbarmenu(0).Buttons("VIEW::SETDETAILS").checked = True

                Case eViewTile
                    cmdbarmenu(0).Buttons("VIEW::SETTILE").checked = True
            End Select
    
        ElseIf bar.Key = "VIEW::COLUMNS::SUBMENU" Then
   
            Dim currcol     As Long

            Dim lvwstylestr As String

            bar.Buttons.Clear
   
            If lvwfiles.columns.Count = 0 Then

                On Error Resume Next

                Set newbutton = cmdbarmenu(0).Buttons.Add("VIEW::COLUMNS::NOCOLUMNS", , "<No Columns>", eNormal)

                If Err <> 0 Then
                    Set newbutton = cmdbarmenu(0).Buttons.Item("VIEW::COLUMNS::NOCOLUMNS")
   
                End If

                newbutton.enabled = False
                bar.Buttons.Add newbutton
   
            End If

            For currcol = 1 To lvwfiles.columns.Count
        
                Set loopcolumn = lvwfiles.columns.Item(currcol)
                CDebug.Post "column: key=" & loopcolumn.Key & " tag = " & loopcolumn.Tag

                On Error Resume Next

                Set newbutton = cmdbarmenu(0).Buttons.Add("COLUMNVIEW::" & loopcolumn.Key, , loopcolumn.Text, eCheck, "Show/Hide Column """ & loopcolumn.Text & """")
        
                If Err <> 0 Then
                    Set newbutton = cmdbarmenu(0).Buttons.Item("COLUMNVIEW::" & loopcolumn.Key)
                Else
                    newbutton.Tag = "SHOWHIDE:" & loopcolumn.Key
                End If

                If loopcolumn.Width > 0 Then
                    CDebug.Post "checking off Key=" & newbutton.Key & " Tag=" & newbutton.Tag
                    newbutton.checked = True
                Else
                    CDebug.Post "checking on Key=" & newbutton.Key & " Tag=" & newbutton.Tag
                    newbutton.checked = False
                End If

                bar.Buttons.Add newbutton
        
            Next currcol
    
        ElseIf bar.Key = "FILE::SUBMENU" Then
            'emulate windows explorer popup.
            'and how, you ask, do we do that?
            CDebug.Post "File::SUBMENU"
        ElseIf bar.Key = "OWFILELIST" Then

            'CDebug.Post  "openfilelist"
            'clear existing buttons...
            Dim LooperFile As CFile, fpath As String

            Dim tempbutton As cButton, loopItem As cListItem

            bar.Buttons.Clear
            Set mCurrSelection = GetSelectedItems
            'add a button for each file; we will use the filename as a key- if the filename exists, we use that button. easy as pie.
            CDebug.Post "selection count=" & mCurrSelection.Count

            For Each loopItem In mCurrSelection

                Set loopfile = GetFile(loopItem.Tag)
                fpath = loopfile.FullPath

                On Error Resume Next

                Set tempbutton = cmdbarmenu(0).Buttons(fpath)

                If Err.Number <> 0 Then
                    Set tempbutton = cmdbarmenu(0).Buttons.Add(fpath, , loopfile.Filename, eNormal, "Select a program to open " & loopfile.Filename & " With.")
                End If

                iconkey = "EXT" & loopfile.Extension & loopfile.FileIndex
                'EXT" & fileobj.Extension & fileobj.FileIndex
                tempbutton.IconIndex = mmenuImages.ItemIndex(iconkey)
                'add to the bar...
                bar.Buttons.Add tempbutton
            
            Next loopItem
        
        ElseIf bar.Key = "POPUP::OPENWITH::SUBMENU" Then

            Dim OpenwithInfo()      As OpenWithListItem, owfile As cButton

            Dim selitems            As Collection, Extension As String, lcount As Long

            Dim OpenwithProg        As CFile, owbutton As cButton

            Static flOpenWithCached As Boolean

            'If Not flOpenWithCached Then
            flOpenWithCached = True

            Dim owbar As cCommandBar, owghost As cButton, owfilelist As cButton

            'OpenwithInfo = GetOpenWithList(getfile(
            'alright- now that we have the open with list, proceed to create the submenu.
            'we will "cheat" by using the "SENDTO:
            bar.Buttons.Clear
            bar.Title = "Open With"
            
            'add initial button if needed.
            On Error Resume Next

            Set selitems = GetSelectedItems
            Set owbutton = cmdbarmenu(0).Buttons.Add("OPENWITHDIALOG", 0, "Select File", eNormal)
            
            '"owbutton" gets a ghost item- because it will be used to list the files in selitems.
            'note we only add the item if there is more then one item selected.
            If selitems.Count > 1 Then

                'add the bar and the ghost.
                'add ghost.
                On Error Resume Next

                Set owbar = cmdbarmenu(0).CommandBars.Item("OWFILELIST")

                If Err.Number <> 0 Then
                    Set owbar = cmdbarmenu(0).CommandBars.Add("OWFILELIST", "Open With File List")
                End If

                'owbar should be the bar.
                On Error Resume Next

                Set owghost = cmdbarmenu(0).Buttons.Item("OWGHOST")

                If Err.Number <> 0 Then
                    Set owghost = cmdbarmenu(0).Buttons.Add("OWGHOST", , "GHOST")
                End If

                owbar.Buttons.Clear
                owbar.Buttons.Add owghost
                'right now, don't place the button on the bar at all.
                'Set owbutton.Bar = owbar
                'set owfilelist = cmdbarmenu(0).Buttons.Add("OWFILELIST::LIST",,"Open With
                'they might exist, so check for that.
                '
            Else
                'remove the bar and the ghost....
                'although we always use the shell menu for single-click...
            
            End If

            'add a button with that bar.
           
            On Error Resume Next

            cmdbarmenu(0).Buttons.Add "OWSEPSEP", , "-", eSeparator
            '
            '
            bar.Buttons.Add cmdbarmenu(0).Buttons("OPENWITHDIALOG")
            bar.Buttons.Add cmdbarmenu(0).Buttons("OWSEPSEP")
            
            'iterate through each openwithitem...
            'if more then one file is selected, show the * menu. otherwise show the menu for the appropriate extension.
            On Error GoTo reportandresume

            CDebug.Post "iterating."
            Extension = vbNullChar

            If selitems.Count > 1 Then
                
                For Each loopItem In selitems

                    'If BCFile.pat loopitem.tag
                    If "." & bcfile.ParseExtension(loopItem.Tag) <> Extension And Extension <> vbNullChar Then
                        'multiple file types selected.
                        Extension = "*"
                        CDebug.Post "extension set to *," & Extension & " is not " & bcfile.ParseExtension(loopItem.Tag)

                        Exit For

                    ElseIf Extension = vbNullChar Then
                        Extension = "." & bcfile.ParseExtension(loopItem.Tag)
                    
                    End If

                Next loopItem

                CDebug.Post "extension is: " & Extension
            
            Else
                Extension = GetFile(lvwfiles.SelectedItem.Tag).Extension
            
            End If

            'alrighty then.
            CDebug.Post "acquiring list for extension, " & Extension
            lcount = 0
            OpenwithInfo = bcfile.GetOpenWithList(Extension, lcount)
            CDebug.Post "acquired. list has " & lcount & " Items."

            For I = 1 To lcount

                If OpenwithInfo(I).strCommandLine <> "" Then

                    On Error GoTo RAISEANDEXIT

                    CDebug.Post "acquiring Open With EXE CFile..."
                    CDebug.Post "strcommand=" & OpenwithInfo(I).strCommandLine
                    Set OpenwithProg = GetFile(OpenwithInfo(I).strCommandLine)
                    CDebug.Post "creating iconkey..."
                    iconkey = OpenwithProg.Filename & OpenwithProg.FileIndex

                    On Error Resume Next

                    mmenuImages.AddFromHandle bcfile.GetObjIcon(OpenwithProg.FullPath, ICON_SMALL), IMAGE_ICON, iconkey

                    On Error Resume Next

                    CDebug.Post "adding button..."
                    Set newbutton = cmdbarmenu(0).Buttons.Add("OPENWITH:" & OpenwithProg.Filename, mmenuImages.ItemIndex(iconkey), OpenwithInfo(I).strName, , "Open file(s) with " & OpenwithInfo(I).strName)

                    If Err = 0 Then
                        CDebug.Post "error adding:" & Err.Description

                    Else
                        Set newbutton = Nothing
                    End If

                    If newbutton Is Nothing Then
                        Set newbutton = cmdbarmenu(0).Buttons.Item("OPENWITH:" & OpenwithProg.Filename)
                    Else
                        
                    End If

                    newbutton.Tag = OpenwithProg.FullPath
                    newbutton.Visible = True
                    bar.Buttons.Add newbutton
                End If

            Next I
            
            'End If
        
        ElseIf bar.Key = "POPUP::SENDTO::SUBMENU" Then
        
            'clear all buttons from this bar....
            'Add from sendto menu, found Here: GetSpecialFolder(0,CSIDL_SENDTO)

            On Error GoTo reportandresume
      
100         Set SendToFolder = GetSpecialFolder(CSIDL_SENDTO)
   
            If Not SendToFolder Is Nothing Then
            
            End If
        
101         If Not SendToInitialized Then
             
103             bar.Buttons.Clear

                On Error Resume Next

                Set newbutton = cmdbarmenu(0).Buttons.Add("SENDTO:CLIPBOARDASNAME", mmenuImages.ItemIndex("PASTE"), "Clipboard as Name", eNormal, "Send The file names to the clipboard.")

                If Err <> 0 Then
                    Set newbutton = cmdbarmenu(0).Buttons.Item("SENDTO:CLIPBOARDASNAME")

                End If

                newbutton.Tag = "SENDTO:CLIPBOARDASNAME"
                bar.Buttons.Add newbutton

                Set newbutton = cmdbarmenu(0).Buttons.Add("SENDTO:CLIPBOARDASCONTENTS", mmenuImages.ItemIndex("PASTE"), "Clipboard as Contents", eNormal, "Send The Contents of the selected file(s) to the clipboard.")

                If Err <> 0 Then
                    Set newbutton = cmdbarmenu(0).Buttons.Item("SENDTO:CLIPBOARDASCONTENTS")

                End If

                newbutton.Tag = "SENDTO:CLIPBOARDASCONTENTS"
                bar.Buttons.Add newbutton
                
                'Separator...
                
                Set newbutton = cmdbarmenu(0).Buttons.Add("SENDTO:CLIPBOARDSEPARATOR", , "-", eSeparator)

                If Err <> 0 Then
                    Set newbutton = cmdbarmenu(0).Buttons.Item("SENDTO:CLIPBOARDSEPARATOR")

                End If

                bar.Buttons.Add newbutton
                
                Set newbutton = Nothing

                With SendToFolder.Files.GetWalker

                    '104            For Each LoopObj In SendToFolder.Files
                    'Print #fNum, "looping via .GetNext..."
104                 Do Until .GetNext(Loopobj) Is Nothing
               
                        'Dim resolver As CshellLink
                        Dim filepathfull As String

105                     iconkey = Loopobj.Filename & Loopobj.FileIndex
                
                        CDebug.Post iconkey
106                     mmenuImages.AddFromHandle bcfile.GetObjIcon(Loopobj.FullPath, ICON_SMALL), IMAGE_ICON, iconkey
        
107                     Set newbutton = cmdbarmenu(0).Buttons.Add("SENDTO:" & Loopobj.Filename, mmenuImages.ItemIndex(iconkey), Loopobj.basename, , "Send files here.")

108                     If StrComp(Loopobj.Extension, "LNK", vbTextCompare) = 0 Then
                            'resolve shortcut...
109                         filepathfull = bcfile.ResolveShortcut(Loopobj.FullPath)
110                         newbutton.Tag = filepathfull

111                     Else
112                         newbutton.Tag = Loopobj.FullPath
            
113                     End If
        
114                     bar.Buttons.Add newbutton
        
                        '115            Next
                    Loop

                    .Reset
                End With

116             CDebug.Post "items added."
                'Add removable drives, too...
                'commented out for testing.
            
                Set newbutton = cmdbarmenu.Item(0).Buttons.Add("SEPSENDTO", , "Removable Drives", eSeparator)
                bar.Buttons.Add newbutton

                Dim drivestrings As String

                Dim Drvs()       As String

                drivestrings = Space$(255)
                GetLogicalDriveStrings 254, drivestrings

                drivestrings = Trim$(Replace$(drivestrings, vbNullChar, " "))

                CDebug.Post drivestrings
                Drvs = Split(drivestrings, " ")

                For I = 0 To UBound(Drvs)

                    If GetDriveType(Drvs(I)) = DRIVE_REMOVABLE Then
                        If ((Asc(Drvs(I)) <> 65 And Asc(Drvs(I)) <> 66)) Or GetKeyState(VK_SHIFT) > 0 Then
                            CDebug.Post "removable drive """ & Drvs(I) & """ found."
                            'add an icon...
                            mmenuImages.AddFromHandle GetObjIcon(Drvs(I), ICON_SMALL), IMAGE_ICON, Drvs(I)
                        
                            Set newbutton = cmdbarmenu.Item(0).Buttons.Add(Drvs(I), mmenuImages.ItemIndex(Drvs(I)), Drvs(I) & "[" & GetVolumelabel(Drvs(I)) & "]")
                            newbutton.Tag = "SENDTO:" & Drvs(I)
                            bar.Buttons.Add newbutton
                        End If
                    End If

                Next I
            
                SendToInitialized = True
            End If
            
        End If
        
        CDebug.Post "beforeShowMenu end " & bar.Key

        Exit Sub

RAISEANDEXIT:
reportandresume:
        CDebug.Post "line:" & Erl & " Desc:" & Err.Description
        CDebug.Post Err.Description

        Resume Next

End Sub
Private Function CreateButton(skey As String, Optional iIcon As Long = -1, Optional sCaption As String, Optional eStyle As EButtonStyle = eNormal, Optional sToolTip As String, Optional vShortcutKey As Integer, Optional eShortcutModifier As ShiftConstants = vbCtrlMask) As vbalCmdBar6.cButton



    Dim madebutton As cButton
    On Error Resume Next
    Set madebutton = cmdbarmenu(0).Buttons.Add(skey, iIcon, sCaption, eStyle, sToolTip, vShortcutKey, eShortcutModifier)
    
    
    If Err <> 0 Then
        Set madebutton = cmdbarmenu(0).Buttons.Item(skey)
        With madebutton
            .IconIndex = iIcon
            .caption = sCaption
            .ToolTip = sToolTip
            .ShortcutKey = vShortcutKey
            .ShortcutModifiers = eShortcutModifier
        
        End With
    
    
    End If
    
    Set CreateButton = madebutton
    


End Function
Private Function ListViewStyleFromString(ByVal LvwStr As String) As EViewStyleConstants

    If StrComp(LvwStr, "Icon", vbTextCompare) = 0 Then
        ListViewStyleFromString = eViewIcon
    ElseIf StrComp(LvwStr, "Small Icon", vbTextCompare) = 0 Then
        ListViewStyleFromString = eViewSmallIcon
    ElseIf StrComp(LvwStr, "List", vbTextCompare) = 0 Then
        ListViewStyleFromString = eViewList

    ElseIf StrComp(LvwStr, "Details", vbTextCompare) = 0 Then
        ListViewStyleFromString = eViewDetails
    ElseIf StrComp(LvwStr, "Tile", vbTextCompare) = 0 Then
        ListViewStyleFromString = eViewTile
    End If
End Function
Private Function ListViewStyleToString(ByVal LvwStyle As EViewStyleConstants) As String

Select Case LvwStyle

    Case eViewIcon
        ListViewStyleToString = "Icon"
    Case eViewSmallIcon
        ListViewStyleToString = "Small Icon"
    Case eViewList
        ListViewStyleToString = "List"
    Case eViewDetails
        ListViewStyleToString = "Details"
    Case eViewTile
        ListViewStyleToString = "Tile"
End Select



End Function
Private Function ImmIf(ByVal Expression, ByVal TrueResult, ByVal FalseResult)
 If Expression Then ImmIf = TrueResult Else ImmIf = FalseResult
End Function
Private Sub DoFileGroup(ByVal StrGroupType As String)

   Static mgrouper As CListViewGrouper
    If mgrouper Is Nothing Then
    Set mgrouper = New CListViewGrouper
    End If
    
    If StrGroupType = "ATTRIB" Then
    mgrouper.AssignGroups lvwfiles, GroupingTypeEnum.GroupbyAttributes
    ElseIf StrGroupType = "SIZE" Then
        mgrouper.AssignGroups lvwfiles, GroupingTypeEnum.GroupBySize
    
    ElseIf StrGroupType = "NONE" Then
        mgrouper.Ungroup lvwfiles
    End If

End Sub
Private Sub CmdBarMenu_ButtonClick(Index As Integer, btn As vbalCmdBar6.cButton)
    CDebug.Post "Clicked: caption=" & btn.caption & " key=" & btn.Key & " TAG=" & btn.Tag
    Dim CommandStr As String
    Dim loopItem As cListItem
    Dim gfile As CFile
    Dim CliRect As RECT, WndRect As RECT
    Dim gotkey As String
    Dim makep() As Variant
    CDebug.Post "ButtonClick"
    
    
    
    'Oh well- just force it to recognize it...
    ReDim makep(1 To 1)
    Set makep(1) = btn
    'mCmdBarLoader.HandleControlEvent CmdBarMenu(0), Nothing, "ButtonClick", makep
    Dim lefttemp As String
    lefttemp = Right$(btn.Key, (Len(btn.Key) - InStrRev(btn.Key, "::") - 1))
    'sortmode & "::" & "lvwfiles::" & loopcolumn.Key
    If (StrComp(lefttemp, "SORTASCENDING", vbTextCompare) = 0) Or (StrComp(lefttemp, "SORTDESCENDING", vbTextCompare) = 0) Then
    'Sort ascending or descending by the column name.
    Dim SplitItems() As String
    Dim sortmodeuse As String, ColumnSort As cColumn
    SplitItems = Split(btn.Key, "::")
    'Sortmode::lvwfiles::columnkey
    sortmodeuse = SplitItems(UBound(SplitItems))
    Set ColumnSort = lvwfiles.columns.Item(SplitItems(2))
    
    ColumnSort.SortOrder = IIf(sortmodeuse = "SORTASCENDING", eSortOrderAscending, eSortOrderDescending)
     Dim loopcol As cColumn, currcol As Long
   For currcol = 1 To lvwfiles.columns.Count
    Set loopcol = lvwfiles.columns.Item(currcol)
    loopcol.IconIndex = -1
   Next
   
   If ColumnSort.SortOrder = eSortOrderNone Then
    ColumnSort.IconIndex = -1
   ElseIf ColumnSort.SortOrder = eSortOrderAscending Then
    ColumnSort.IconIndex = mHeaderIML.ItemIndex("ASCENDING")
    ElseIf ColumnSort.SortOrder = eSortOrderDescending Then
    ColumnSort.IconIndex = mHeaderIML.ItemIndex("DESCENDING")
   End If
    
    
    lvwfiles.ListItems.SortItems
    ElseIf btn.Tag = "UTILITY::ACTION" Then
        FrmAction.ShowDialog Me
    
    ElseIf btn.Tag = "POPUP::ACTION" Then
        FrmAction.Show
    
    ElseIf btn.Key = "VIEW::REFRESH" Then
    
        Dim currItem As cListItem, I As Long
        For I = 1 To lvwfiles.ListItems.Count
            Set currItem = lvwfiles.ListItems.Item(I)
            RefreshItemData currItem, bcfile.GetFile(currItem.Tag)
        Next I
    
    
    ElseIf btn.Tag = "VIEW::SIZEFIT" Then
    
    Debug.Print "sizefit"
    SizeColumnsToFit lvwfiles
    ElseIf btn.Key = "VIEW::FILTERSPANE" Then
    'show/hide the "filter pane" which is the top level pane. note that this essentially puts the program into a mode where
    'it only shows the results, maximizing screen real estate.
        
        PicUpperPane.Visible = Not PicUpperPane.Visible
        PicUpperPane.enabled = PicUpperPane.Visible
        PicSplit.Visible = PicUpperPane.Visible
        'SendMessage Me.hWnd, CWM_SIZE
        'Me.Move Me.Left, Me.Top, Me.Width, Me.Height
        'Me.Refresh
        Form_Resize
        
        mFireResizeNextTmr = True 'hack alert.
        Me.Move Me.Left, Me.Top, Me.Width, Me.Height
        'we need to call move so that the messages get triggered in the subclasser
        ElseIf btn.Tag = "VIEW::MODE::SIMPLE" Then
        Set mzipfilter = New CSearchFilter
       
        
            mFilterViewMode = FilterView_Simple
            PicSimple.Visible = True
 '           PicSimple.enabled = True
            PicFilters.Visible = False
 '           PicFilters.enabled = False
            PicUpperPane_Resize
            'Cache the old reference...
            Set mAdvancedFilters = mFileSearch.Filters
            'recreate the filters...
            Set mFileSearch.Filters = New SearchFilters
            AddDirectoryExclude mFileSearch.Filters
            If mSimpleFilter Is Nothing Then
                'we haven't used simple mode- recreate the object.
                Set mSimpleFilter = New CSearchFilter
            End If
            'either way, add it to the collection.
            
            mFileSearch.Filters.Add mSimpleFilter
            SetTabStop PicSimple, PicFilters
        ElseIf btn.Tag = "VIEW::MODE::ADVANCED" Then
            mFilterViewMode = FilterView_Advanced
            PicSimple.Visible = False
'            PicSimple.enabled = False
            PicFilters.Visible = True
'            PicFilters.enabled = True
            Set mFileSearch.Filters = mAdvancedFilters
            SetTabStop PicFilters, PicSimple
            'mFileSearch.Filters.Clear
            'restore the cached value...
            PicUpperPane_Resize
            
        
    ElseIf btn.Tag = "FILE::NEW" Then
        cmdNewSearch.Value = 1
    
    ElseIf btn.Tag = "VIEW::FILECONTEXT" Then
        'show popup menu....
        
        'CmdBarMenu(0).get
        
        
        ShowFilepopup
    ElseIf btn.Key = "VIEW::EXPLORERBAR" Then
        ExplorerBar.Visible = Not ExplorerBar.Visible
        Form_Resize
    ElseIf StrComp(Left$(btn.Tag, 8), "SETVIEW:", vbTextCompare) = 0 Then
    Select Case UCase$(Mid$(btn.Tag, 9))

    Case "DETAILS"
        lvwfiles.View = eViewDetails
    Case "ICON"
        lvwfiles.View = eViewIcon
    Case "LIST"
        lvwfiles.View = eViewList
    Case "SMALLICON"
        lvwfiles.View = eViewSmallIcon
    Case "TILE"
        lvwfiles.View = eViewTile
    End Select
    
    ElseIf Left$(btn.Key, 22) = "FILTERSEARCH::COLUMN::" Then
       gotkey = Mid$(btn.Key, 23)
    'the remainder of the string is the index of the column.
    If InStr(1, mSearchColumnskeys, "&+" & gotkey & "&+", vbTextCompare) > 0 Then
        'it exists... uncheck it.
        mSearchColumnskeys = Replace$(mSearchColumnskeys, "&+" & gotkey & "&+", "")
    Else
        'doesn't check it.
        mSearchColumnskeys = mSearchColumnskeys + "&+" & gotkey & "&+"
    End If
    ElseIf btn.Tag = "GREATPOSSUM" Then
     On Error Resume Next
    Dim shellobj As Object
     Set shellobj = CreateObject("WScript.Shell")
     If Err <> 0 Then
        'MsgBox "error:" & Err
    Else
        shellobj.CurrentDirectory = CurrApp.GetDataFolder & "resource\other\ovum"
        shellobj.Run "psmpassovr.html"
'        If Err <> 0 Then
'            MsgBox Error & " " & Err.Number
'        End If
     End If
    
    
    ElseIf btn.Tag = "HELP::DONATE" Then
    
       
    On Error Resume Next
    'Dim shellobj As Object
     Set shellobj = CreateObject("WScript.Shell")
     If Err <> 0 Then
        'MsgBox "error:" & Err
    Else
        shellobj.CurrentDirectory = App.Path
        shellobj.Run "donate.html"
'        If Err <> 0 Then
'            MsgBox Error & " " & Err.Number
'        End If
     End If
   
    ElseIf btn.Key = "TOOLBAR::PRINT" Then
        'print...
        
    ElseIf Left$(UCase$(btn.Key), 7) = "GROUP::" Then
        Currgrouping = Mid$(btn.Key, 8)
        DoFileGroup Currgrouping
        btn.checked = True
    ElseIf btn.Key = "OPENWITHDIALOG" Then
            'MsgBox "show open with dialog here."



            Set gfile = GetFile(lvwfiles.SelectedItem.Tag)
            gfile.OpenWith Me.hWnd


    ElseIf Left$(btn.Tag, 9) = "SHOWHIDE:" Then
        Dim showhidecol As cColumn
        Set showhidecol = lvwfiles.columns(Mid$(btn.Tag, 10))
        If showhidecol.Width = 0 Then
            showhidecol.Width = showhidecol.ItemData
            If showhidecol.Width = 0 Then
                showhidecol.Width = 50
            End If
        Else
            showhidecol.ItemData = showhidecol.Width
            showhidecol.Width = 0
        End If
    ElseIf btn.Tag = "LVWSHOW:ALL" Then
        'display all hidden columns...
        'Dim I As Long
        For I = 1 To lvwfiles.columns.Count
            With lvwfiles.columns(I)
                If lvwfiles.columns(I).Width = 0 Then lvwfiles.columns(I).Width = lvwfiles.columns(I).ItemData
                    If lvwfiles.columns(I).Width = 0 Then
                        'if it's STILL zero set to something visible.
                        lvwfiles.columns(I).Width = 50
                    
                    End If
            End With
        Next I

    ElseIf btn.Key = "FILTERS::EDIT" Then
        'edit currently selected filter.
        'Set getitem = lvwFilters.ListItems("TAG" & Trim$(mFileSearch.Filters.Count))
        cmdEditFilter_Click
    ElseIf btn.Key = "FILTERS::ADD" Then
        cmdAddFilter_Click
    ElseIf btn.Key = "FILTERS::REMOVE" Then
        CmdremoveFilter_Click
    
    
     ElseIf btn.Tag = "MENU::ABOUT" Then
        FrmAbout.Show
        'show the about box. there. done.
    ElseIf btn.Tag = "UTILITY::MANAGECOL" Then
        FrmColumnPlugins.Show
        
        'lvwfilters_Columns
    ElseIf btn.Tag = "UTILITY::CLEARPOSITION" Then
        'clear position data...
        CurrApp.Settings.DeleteSection "lvwfilters_Columns"
        CurrApp.Settings.DeleteSection "lvwfiles_Columns"
        'set the tag so they are not saved back to the INI...
        lvwfiles.Tag = "nopersist"
        lvwfilters.Tag = "nopersist"
        MsgBox "Column Position Data cleared; restart BCSearch for changes to take effect."
    ElseIf btn.Tag = "UTILITY::CLEARHISTORY" Then

        CurrApp.Settings.DeleteSection "persistproperties.cbolookin"
        CurrApp.Settings.DeleteSection "persistproperties.cbofilter"
        cboLookin.Clear
        
    ElseIf btn.Tag = "SHOW::DIRSIZE" Then
        FrmDirSizeAnalyzer.Show
    ElseIf btn.Tag = "MENU::FILE" Then
    'NOTE: one of several possible options here.
    'if we have an item(s) selected in the Results list, then we fire off the same logic that we would in that case. (IE to determine wether to show an explorer or custom menu)
    'otherwise, we show a different menu, with various search options. In either case, the options as present on the standard File menu should exist on the Explorer style menu as well.
    'the trick NOW is to position the menu so it appears as if it is actually the menu of the button.
'    GetClientRect CmdBarMenu(0).hwnd, CliRect
'    GetWindowRect CmdBarMenu(0).hwnd, WndRect
'    'cmdbarmenu(0).ClientCoordinatesToScreen(clirect.Left,clirect.TOP,cmdbarmenu(0).hWnd)
'
'    ShowFilepopup WndRect.Left, WndRect.Bottom
    ElseIf btn.Tag = "FILE::EXIT" Then
        Unload Me
    
    ElseIf btn.Tag = "FILE::OPEN" Then
        Dim cdlg As CFileDialog, fileget As CFile
        Set cdlg = New CFileDialog
        cdlg.caption = "Open Saved Search..."
        cdlg.Filter = "Saved Searches(*.BCSEARCH)|*.BCSEARCH|All Files(*.*)|*.*"

        Set fileget = cdlg.SelectOpenFile(Me.hWnd)
        If Not fileget Is Nothing Then
            CDebug.Post "proceed to open " & fileget.FullPath
            OpenSavedSearch fileget

        End If
        'loading code here...
    ElseIf btn.Tag = "FILE::SAVE" Then
        Dim fname As String
        Set cdlg = New CFileDialog
        cdlg.caption = "Save Search..."
        cdlg.Filter = "Saved Searches(*.BCSEARCH)|*.BCSEARCH|CSV file(*.CSV)|*.CSV|All Files(*.*)|*.*"
        cdlg.DefExt = "BCSEARCH"

        fname = cdlg.SelectSaveFile(Me.hWnd)
        If fname <> "" Then
        SaveSearch fname
        End If
        CDebug.Post "proceed to save to " & fname

        'Saveing code here...
    ElseIf btn.Key = "EXPORT_MDB" Then
        
        Set cdlg = New CFileDialog
        cdlg.caption = "Export Results to Database"
        cdlg.Filter = "Jet Database (*.mdb)|*.mdb"
        cdlg.DefExt = "mdb"
        fname = cdlg.SelectSaveFile(Me.hWnd)
        ExportResultsToMDB fname
    
    
    ElseIf btn.Key = "EXPORT_TEXT" Then


        Set cdlg = New CFileDialog
        cdlg.caption = "Save File listing..."
        cdlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
        cdlg.DefExt = "txt"

        fname = cdlg.SelectSaveFile(Me.hWnd)
        GenFileListing fname, Me
           'other keys:
        'COPYCLIP::ALLVISIBLE 'copy all visible columns
        'COPYCLIP::ALL 'copy all columns
        'COPYCLIP::COLUMN::<KEY> 'copy that column to the clipboard
    ElseIf btn.Key = "EXPORT_HTML" Then
    'MsgBox "TODO:// export to HTML"
    ExportToHTML

    ElseIf btn.Key = "EXPORT_CSV" Then
    MsgBox "TODO:// export to CSV"
    'ExportToCSV
    ElseIf btn.Key = "EXPORT_WORD" Then
        'export to word document.
'          Set cdlg = New CFileDialog
'        cdlg.caption = "Save File listing..."
'        cdlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*"
'        cdlg.DefExt = "txt"

'        fname = cdlg.SelectSaveFile(Me.hWnd)
'         Export_Word
        
        
    ElseIf btn.Key = "COPYCLIP::ALLVISIBLE" Or Left$(btn.Key, 16) = "COPYCLIP::COLUMN" Then
        'copy all visible columns.
        
        Dim colloop As cColumn
        Dim getcolumn As String
        Dim strbuilder As cStringBuilder
        Set strbuilder = New cStringBuilder
        If Left$(btn.Key, 16) = "COPYCLIP::COLUMN" Then
            getcolumn = Mid$(btn.Key, 17)
            Debug.Print "copy column " & getcolumn
            
        End If
        Set mCurrSelection = GetSelectedItems
        For Each loopItem In mCurrSelection
            Dim CurrPosition As Long, columnappend As cColumn
            CurrPosition = 1
  
                If getcolumn = "" Then
                'if the variable is empty, copy all columns.
'                For Each litem In mCurrSelection
                
                
                
                
                    
                    For CurrPosition = 1 To loopItem.SubItems.Count
                        Set columnappend = GetListViewColumnByPosition(lvwfiles, CurrPosition)
                        If columnappend.Width > 5 Then
                            strbuilder.Append loopItem.SubItems(CurrPosition).caption & vbTab
                        End If
                    Next
              
                
                Else
                'otherwise,
                Set columnappend = lvwfiles.columns(getcolumn)
                strbuilder.Append loopItem.SubItems(columnappend.position).caption
                End If
      
            
            
'            Do Until CurrPosition > lvwfiles.Columns.Count
'
'                Set columnappend = GetListViewColumnByPosition(lvwfiles, CurrPosition)
'
'                strbuilder.Append LoopItem.SubItems(columnappend.Position).caption
'
'
'
'
'                CurrPosition = CurrPosition + 1
'                If CurrPosition < lvwfiles.Columns.Count Then
'                    strbuilder.Append vbTab
'
'                End If
'            Loop
            
        
        
            strbuilder.Append vbCrLf
        
        
        Next
        Clipboard.Clear
        Clipboard.SetText strbuilder.ToString
        Clipboard.SetText strbuilder.ToString, vbCFText
    
        'specific column
    
    
    ElseIf Left$(btn.Tag, 7) = "SENDTO:" Or Left$(btn.Key, 9) = "OPENWITH:" Then
        'Assemble the exec string... may as well use shell... wtf, right?

        'CommandStr = """" & btn.Tag & """"

        'first exec name. Now- loop through the selecteditems collection, and add the TAG of each one- that is the filename.
        If Not mCurrSelection Is Nothing Then
            For Each loopItem In mCurrSelection
                CommandStr = CommandStr & " """ & loopItem.Tag & """"
    
    
            Next loopItem
        Else
        Set mCurrSelection = New Collection
        End If

        'IF the btn.tag file is a directory, then we copy each file/folder to the specified destination.
        'Shortcut targets will have been resolves already.
        
        
     
'
'
'
'            Next i
        
        If StrComp(Mid$(btn.Key, 8), "CLIPBOARDASNAME", vbTextCompare) = 0 Then
        CommandStr = ""
        If Not mCurrSelection Is Nothing Then
            For Each loopItem In mCurrSelection
                CommandStr = CommandStr & loopItem.Tag & vbCrLf
    
    
            Next loopItem
            Clipboard.SetText CommandStr, vbCFText
        End If
        ElseIf StrComp(Mid$(btn.Key, 8), "CLIPBOARDASCONTENTS", vbTextCompare) = 0 Then
        Dim StrCopy As String
        Dim readFile As FileStream
        'this is the PITA...
        If Not mCurrSelection Is Nothing Then
            For Each loopItem In mCurrSelection
               ' CommandStr = CommandStr & LoopItem.Tag & vbCrLf
               Set readFile = OpenStream(loopItem.Tag)
               'Read the contents of the file...
               StrCopy = StrCopy & readFile.ReadAllStr
               
               readFile.CloseStream
    
    
            Next loopItem
            Clipboard.SetText StrCopy, vbCFText
        End If
        
        
        
        Else
        
            Dim resolved As String
            
            resolved = ResolveShortcut(btn.Tag)
    
            
            CDebug.Post "Executed:" & CommandStr
            ShellExecute Me.hWnd, "Open", btn.Tag, Trim$(CommandStr), "", vbNormalFocus
    
            'Shell CommandStr
    
            CDebug.Post "proceed to send " & mCurrSelection.Count & " files to send to item " & btn.Tag
        End If

    
    End If
End Sub
Private Sub ExportResultsToMDB(mdbfile As String, Optional ByVal TableName As String = "")
    '
    Dim constring As String
    Dim catuse As ADOX.Catalog
    Dim usecolumn As cColumn
    Dim testconnection As Connection
    Dim Newtable As Table, currcolumn As Long, newcolumn As Column
    If TableName = "" Then TableName = "exporttable"
    constring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mdbfile
    
    
    
    
    Set catuse = New ADOX.Catalog
    catuse.Create constring
    
    'add the table...
    Set Newtable = New Table
    Newtable.Name = TableName
    Dim madename As String
    'Alright- add a Field for every column...
    For currcolumn = 1 To lvwfiles.columns.Count
        Set usecolumn = lvwfiles.columns.Item(currcolumn)
            madename = Replace$(usecolumn.Text, ".", "")
            madename = Replace$(madename, ",", "")
            
            
            
            
           Newtable.columns.Append madename, adWChar, 80
           'Newtable.columns(madename).Properties = adPropOptional
           
           
        
    
    
    Next
    catuse.Tables.Append Newtable
    
    Set catuse = Nothing
    
    Dim recordsetuse As ADODB.Recordset
    Dim connuse As ADODB.Connection, I As Long
    Dim currfield As Long
    Set connuse = New Connection
    connuse.ConnectionString = constring
    connuse.Open
    Set recordsetuse = New Recordset
    Call recordsetuse.Open(TableName, connuse, adOpenDynamic, adLockOptimistic)
    'recordsetuse.MoveLast
    
    
    Dim lfileitem As cListItem
    For I = 1 To lvwfiles.ListItems.Count
        recordsetuse.AddNew
        Set lfileitem = lvwfiles.ListItems(I)
        For currfield = 1 To lvwfiles.columns.Count
            recordsetuse.Fields(currfield - 1).Value = lfileitem.SubItems(currfield - 1).caption
        
            If Trim$(recordsetuse.Fields(currfield - 1).Value) = "" Then Debug.Assert False
        Next currfield
    
    
        
    
    
    Next I
    
    
    recordsetuse.Close
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub
Private Sub SizeColumnsToFit(Onlvw As vbalListViewCtl, Optional includeheaderwidth As Boolean = True)

Dim currcol As Long
    If Not includeheaderwidth Then
   For currcol = 0 To Onlvw.columns.Count - 1
      Call SendMessage(Onlvw.hWndListView, _
                       LVM_SETCOLUMNWIDTH, _
                       currcol, _
                       ByVal LVSCW_AUTOSIZE)
    Next

Else
    For currcol = 0 To Onlvw.columns.Count - 1
      Call SendMessage(Onlvw.hWndListView, _
                       LVM_SETCOLUMNWIDTH, _
                       currcol, _
                       ByVal LVSCW_AUTOSIZE_USEHEADER)

    Next
End If


'below: old version.
''sizes all columns in onLvw to fit the contents.
'
''Step one: find optimal size for each column:
'
'Dim Optimal() As Long, longest() As String
'Dim I As Long, currli As cListItem, originalFont As Object
'Dim currcol As Long, subitemtext As String
'ReDim Optimal(1 To OnLvw.Columns.Count)
'ReDim longest(1 To OnLvw.Columns.Count)
'Set originalFont = Me.Font
'For I = 1 To OnLvw.ListItems.Count
'    Set currli = OnLvw.ListItems.Item(I)
'    Set Me.Font = currli.Font
'    'loop through subitems...
'    For currcol = 1 To currli.SubItems.Count + 1
'
'        If currcol = 1 Then
'            subitemtext = currli.Text
'        Else
'            subitemtext = currli.SubItems(currcol - 1).caption
'        End If
'
'        If OnLvw.Columns.Item(currcol).Width > 5 Then
'
'        Debug.Print subitemtext; " in "; OnLvw.Columns.Item(currcol).Text
'            If TextWidth(subitemtext) > Optimal(currcol) Then
'                Optimal(currcol) = TextWidth(subitemtext)
'                longest(currcol) = subitemtext
'            End If
'        Else
'            Optimal(currcol) = 0
'        End If
'    Next
'
'Next
'
'
'For currcol = 1 To OnLvw.Columns.Count
'Debug.Print "longest entry in column " & OnLvw.Columns(currcol).Text & " was " + longest(currcol) & " at " & Optimal(currcol) & " pixels."
'OnLvw.Columns(currcol).Width = Optimal(currcol)
'Next
'
'Set Me.Font = originalFont
'

End Sub
Private Sub ExportToHTML()
    'export search results to an HTML file.
    'First version :P.
    
    'maybe I'll add an "export" plugin architecture....
    Dim useBasicHTML As Boolean
    Dim SelFile As CFile, Selector As CFileDialog, selected As String
    Dim OutStream As FileStream
    Dim ScriptCode As String
    Dim CSSCode As String
    Set Selector = New CFileDialog
    selected = Selector.SelectSaveFile(Me.hWnd)
    
    If selected = "" Then Exit Sub
    'Set SelFile = GetFile(selected)
    On Error GoTo FileOpenError
    'Set OutStream = SelFile.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS)
    
    If Exists(selected) Then
        If MsgBox(selected & " already exists. Would you like to overwrite the existing file?", vbYesNo + vbQuestion, "Overwrite?") = vbNo Then
            Exit Sub
        End If
    End If
    If InStr(selected, "\") = 0 And InStr(selected, "/") = 0 Then selected = IIf(Right$(CurDir$, 1) = "\", CurDir$ & selected, CurDir$ & "\" & selected)
    
    Set OutStream = bcfile.CreateStream(selected)
    'useBasicHTML = CBool(CurrApp.Settings.ReadProfileSetting("export", "exportHTMLUseCSS", "0"))
    OutStream.Stringmode = StrRead_unicode
    OutStream.WriteString "<html><head>" & vbCrLf
    OutStream.WriteString "<title>BCSearch Search Results created " & FormatDateTime(Now, vbGeneralDate) & "</title>" & vbCrLf
    
    'write out the script...
    OutStream.WriteString "<script type=""text/javascript"">" & vbCrLf
    Dim scriptbytes() As Byte
    scriptbytes = LoadResData("SORTSCRIPT", "JSS")
    OutStream.WriteString StrConv(scriptbytes, vbUnicode)
    
    
    OutStream.WriteString vbCrLf
    OutStream.WriteString "</script>"
    
    OutStream.WriteString "<style>" & vbCrLf
    
    
    'write out the style code:
    'SORTCSS, type CSS:
    Dim cssbytes() As Byte
    cssbytes = LoadResData("SORTCSS", "CSS")
    OutStream.WriteString StrConv(cssbytes(), vbUnicode)
    OutStream.WriteString vbCrLf
    
    OutStream.WriteString "</style>" & vbCrLf
                        
    OutStream.WriteString "<body>" & vbCrLf
    'class="sortable" id="anyid"
    OutStream.WriteString "<table class=""sortable"" id=""sortid"">"
    Dim currcol As Long, poscol As Long
    OutStream.WriteString "<tr>"
    For currcol = 1 To lvwfiles.columns.Count
        OutStream.WriteString "<th>"
        OutStream.WriteString lvwfiles.columns(currcol).Text
        OutStream.WriteString "</th>"
    
    Next
    OutStream.WriteString "</tr>"
    Dim loopItem As cListItem, LoopPos As Long
        For LoopPos = 1 To lvwfiles.ListItems.Count
        Set loopItem = lvwfiles.ListItems(LoopPos)
        If LoopPos Mod 2 Then
            
        'OutStream.WriteString "<tr>"
        OutStream.WriteString "<tr backcolor=""#C0C0C0"">"
        Else
        OutStream.WriteString "<tr backcolor=""#E0E0E0"">"
        End If
    For currcol = 0 To loopItem.SubItems.Count
    
        OutStream.WriteString "<td>"
        If currcol = 0 Then
            OutStream.WriteString loopItem.Text
        Else
            OutStream.WriteString loopItem.SubItems(currcol).caption
        End If
        OutStream.WriteString "</td>"
    
    Next
    OutStream.WriteString "</tr>" & vbCrLf
        
        
        
        
        
        Next
    
    OutStream.WriteString "</table></html>"
    OutStream.CloseStream
    
    
    
    
    
    
    
    
Exit Sub
FileOpenError:
    MsgBox "Error Exporting HTML:" & Err.Description
    Debug.Assert False
    Resume
End Sub



Private Sub CmdBarMenu_RequestNewInstance(Index As Integer, ctl As Object)
On Error Resume Next
Dim NewIndex As Long
NewIndex = cmdbarmenu.UBound + 1
Load cmdbarmenu(NewIndex)
cmdbarmenu(NewIndex).Align = 0
cmdbarmenu(NewIndex).MenuImageList = mmenuImages.hImageList
Set ctl = cmdbarmenu(NewIndex)
With cmdbarmenu(0)
ctl.Style = .Style
ctl.BackgroundImage = .BackgroundImage
Set ctl.Font = .Font
Dim I As Long
For I = [_eccCustomColorFirst] To [_eccCustomColorLast]
    'duplicate...!
    ctl.CustomColor(I) = .CustomColor(I)
    ctl.UseStyleColor(I) = .UseStyleColor(I)
Next I
'.CustomColor(
CDebug.Post "setting menuimagelist to " & mmenuImages.hImageList
ctl.MenuImageList = mmenuImages.hImageList
End With
  CDebug.Post "New Control Instance Obtained:" & ctl.hWnd
  mCmdBarLoader.AddHandledControl ctl
End Sub

Private Sub CmdFind_Click()

Dim strDownArrow As String
Dim egg(0 To 1) As String
egg(0) = "evol e"
egg(1) = "zylana"
mItemDataCount = -1
Erase mItemData

mSbar.PanelText("FOUND") = ""
If StrComp(cboLookin.Text, StrReverse(egg(0) & egg(1)), vbTextCompare) = 0 Then
    MsgBox " A strange game." & vbCrLf & _
        "the only winning move is to not play."
    
    Exit Sub
End If

If GetKeyState(VK_CONTROL) > 0 Then
    If GetKeyState(vbKeyReturn) > 0 Then
    
    
    'Break out...


    Else
    
    CDebug.Post "control+Enter pressed."
    End If
    
    
    
End If

Me.caption = "BCSearch"
mNonLocal = False
strDownArrow = ChrW$(&H25BC)
mcancel = False
CmdFind.enabled = False
cmdNewSearch.enabled = False
CmdStop.enabled = True
lvwfiles.Visible = True
'fire resize...
Form_Resize  'Should fix resizing issue.

'erase any cached items...
Set mCachedRemovedItems = Nothing


lvwfiles.ListItems.Clear
lvwfiles.columns.Clear
lvwfiles.Imagelist(eLVSmallIcon) = CurrApp.SystemIML(Size_Small).himl
lvwfiles.Imagelist(eLVLargeIcon) = CurrApp.SystemIML(Size_Large).himl
ExplorerBar.BarTitleImageList = CurrApp.SystemIML(Size_Large).himl
If mHeaderIML Is Nothing Then
    Set mHeaderIML = mCmdBarLoader.LoadImageList("HEADERICONS")

    



End If
lvwfiles.Imagelist(eLVHeaderImages) = mHeaderIML.hImageList
'On Error GoTo HasColumns
lvwfiles.HeaderDragDrop = True
lvwfiles.HeaderButtons = True
'lvwFiles.HeaderTrackSelect = True
'mSmallImages.AddFromResourceID "DOWNARROW", App.hInstance, IMAGE_CURSOR, "DOWNARROW"
If Not lvwfiles.columns.Exists("FILENAME") Then

    With lvwfiles.columns.Add(, "FILENAME", "Filename", -1, 230)
        .Width = TextWidth("################.###")
        .SortType = eLVSortStringNoCase
    End With
'    With lvwFiles.Columns.Add(, "DROPDOWN:FILENAME", "")
'        '.Width = ScaleX(TextWidth(ChrW$(&H25BC)), vbTwips, vbPixels)
'        'show a down arrow, via the icon...
'        'we added this first...
'        '.IconIndex = "DOWNARROW"
'        .Width = ScaleX(TextWidth("##"), vbTwips, vbPixels)
'    End With
    With lvwfiles.columns.Add(, "DIRECTORY", "Dir", -1, 230)
        .Width = TextWidth("C:\windows\System32\Folder\Folder")
        .SortType = eLVSortStringNoCase
        
    End With
    With lvwfiles.columns.Add(, "EXTENSION", "Extension", -1, 230) '
        .Width = TextWidth("HTML")
        .SortType = eLVSortStringNoCase
    End With
    With lvwfiles.columns.Add(, "TYPE", "File Type", -1, 230)
        .Width = TextWidth("######### Document")
        .SortType = eLVSortStringNoCase
        End With
    With lvwfiles.columns.Add(, "SIZE", "Size", -1, 230)
        .Width = TextWidth("#1024 bytes#")
        .SortType = eLVSortItemData
    
        
    End With
    With lvwfiles.columns.Add(, "DISKSIZE", "Size On Disk", -1, 230)
        .Width = TextWidth("#1024 bytes#")
        .SortType = eLVSortItemData
        
    End With
    With lvwfiles.columns.Add(, "MODIFIED", "Modified", -1, 230)
        .Width = TextWidth("##/##/## at ##:##.##")
        .SortType = eLVSortDate
    End With
    With lvwfiles.columns.Add(, "ACCESSED", "Accessed", -1, 230)
        .Width = TextWidth("##/##/## at ##:##.##")
        .SortType = eLVSortDate
    End With
    With lvwfiles.columns.Add(, "CREATED", "Created", -1, 230)
        .Width = TextWidth("##/##/## at ##:##.##")
        .SortType = eLVSortDate
    End With

    'add the plugin ones, too. we will find these loaded by the "Application" object during bootstrap.
    Dim CurrPlugin As Long
    Dim grabplugs() As IColumnPlugin
    Dim currcolumn As Long
    Dim colinfo As ColumnInfo
    Dim grabplug As IColumnPlugin
    grabplugs = CurrApp.ColumnPlugins
    For CurrPlugin = 1 To CurrApp.ColumnCount
        Set grabplug = grabplugs(CurrPlugin)
            For currcolumn = 1 To grabplug.GetColumnCount
                colinfo = grabplug.GetColumnInfo(currcolumn)
                Dim newcolumn As cColumn
                
                'add the new Column... we will associate it with the original plugin and column index via the tag property, which will be
                'PLUGIN:<pluginindex currplugin>:<currcolumn> which can be parsed when required, via a simple split function call.
                Set newcolumn = lvwfiles.columns.Add(, "PLUGIN:" & Trim$(CurrPlugin) & ":" & Trim$(currcolumn), colinfo.ColumnTitle, , (colinfo.ColumnDefaultWidth))
                'ta-da!
            
            
            Next currcolumn
        
        
        
        
    
    Next CurrPlugin
'lastly, reset the width and ordering...
'On Error Resume Next
If CurrApp.Settings.SectionExists("lvwfiles_Columns") Then
mFormSaver.LoadListViewColumnConfig lvwfiles, "lvwfiles_Columns"
End If

End If
On Error Resume Next

''SearchAnimation.StartPlay
''mTotalFileCount = 0
''mTotalRecurse = 0
''PicANI.AutoRedraw = False
mSearchAnimator.StartAnimation
mSearching = True
On Error GoTo ReportError
mFileSearch.Search "*", cboLookin.Text, Me, True

'Stop
Exit Sub
ReportError:
    CDebug.Post "Unknown Error:" & Err.Description & "(" & Err.Number & ")" & " From " & Err.Source, Severity_Warning
'    Select Case MsgBox(Err.Description, vbAbortRetryIgnore, "Untrapped Error")
'        Case vbAbort, vbIgnore
'            IFileSearchCallback_Cancelled
'        Case vbRetry
'            Resume
'
'
'    End Select
End Sub

Private Sub Command1_Click()
'
'GroupBySize "Small", 10# * 1024#, "Medium", 20# * 1024#, "Large", 50# * 1024# * 5#, "Huge"
FrmDirSizeAnalyzer.Show

End Sub

Private Sub cmdLookinBrowse_Click()
    Dim dirget As Directory
    Dim browser As CDirBrowser
    Set browser = New CDirBrowser
    Set dirget = browser.BrowseForDirectory(Me.hWnd, "Browse for folder...", BIF_EDITBOX + BIF_NEWDIALOGSTYLE + BIF_RETURNONLYFSDIRS)
    cboLookin.Text = dirget.Path
End Sub

Private Sub cmdNewSearch_Click()
    lvwfiles.ListItems.Clear
    lvwfilters.ListItems.Clear
    'mFileSearch.Filters.Clear
    mNonLocal = False
    Dim loopItem As cListItem
    Dim litem As Long
'    For litem = 1 To lvwfilters.ListItems.Count
'    Set loopItem = lvwfilters.ListItems.Item(litem)
'    mFileSearch.Filters.Remove loopItem.ItemData
'
'    Next
    Set mFileSearch = New CFileSearchEx
    Set mFileSearch.Filters = New SearchFilters
    'recreate "filters" collection...
    AddDirectoryExclude mFileSearch.Filters
    cmdbarmenu(2).Visible = False
    PicLowerPane_Resize
End Sub

Private Sub CmdStop_Click()
    mcancel = True
    mFileSearch.Cancel
End Sub


Private Sub Form_Load()
    Dim I As Long, loopctl As Control
    Dim loadedIco As Long
    'TODO: fix crash with XP and earlier :(
    Static flinit As Boolean
    
    On Error Resume Next
'    If mSClass Is Nothing Then
'        Set mSClass = New cSuperClass
'        mSClass.AddMsg CWM_GETMINMAXINFO, False
'
'        mSClass.Subclass Me.hWnd, Me, False
'    End If

    If mFormSizer Is Nothing Then
        Set mFormSizer = New CFormSizeControl
        
        mFormSizer.InitClass Me.hWnd, , , 128, 405, , , , , False
        
    End If
    'Set mToolTip = New clsTooltip
    
    SetIcon Me.hWnd, "AAA", True
    If flinit Then Exit Sub
'    Set mNotify = New CFileChangeNotify
'    mNotify.Start "D:\vbproj\vb\", False, Me
    Set mSbar = New cNoStatusBar
    Set mFileSearch = New CFileSearchEx
    Set mFileSearch.Filters = New SearchFilters
    AddDirectoryExclude mFileSearch.Filters
    Set mCmdBarLoader = New CXMLLoader
   mCmdBarLoader.AddScriptObject Me, "SearchForm"



    Me.Show


    mSbar.Create PicStatusBar
    Set mSbar.Font = New StdFont
    mSbar.Font.Name = "MS Shell Dlg 2"
    mSbar.AddPanel estbrStandard, "Files Found", 0, 72 * 2.5, False, False, , "FOUND"
    mSbar.AddPanel estbrStandard, "Monitoring New Items", , , , , , "MONITOR"
    mSbar.SizeGrip = True



    
    
'    Set mLargeImages = New cVBALImageList
'    mLargeImages.IconSizeX = 32
'    mLargeImages.IconSizeY = 32
'    mLargeImages.ColourDepth = ILC_COLOR32
'    mLargeImages.Create
        On Error Resume Next
    'If GetSwitchArguments(CommandLine, "menuxml") <> "" Then
    Dim MenuXMLFile As String
    'MenuXMLFile = CurrApp.GetDataFolder & "MenuXML.XML"
    If cmdParser.Switches.Exists("MENUXML") Then
    MenuXMLFile = cmdParser.Switches.Item("MENUXML").Arguments.Item(1).ArgString
        
        If Dir$(MenuXMLFile) = "" Then
            'not found... try relative to appdata path...
            On Error Resume Next
            MenuXMLFile = CurrApp.GetDataFolder & cmdParser.Switches.Item("MENUXML").Arguments.Item(1).ArgString
            If Dir$(MenuXMLFile) = "" Then
                MenuXMLFile = CurrApp.GetDataFolder & "MenuXML.XML"
            End If
        End If
    ElseIf CurrApp.Settings.ReadProfileSetting("BCSearch", "MenuXML", "") <> "" Then
         MenuXMLFile = CurrApp.Settings.ReadProfileSetting("BCSearch", "MenuXML", "")
    Else
        MenuXMLFile = CurrApp.GetDataFolder & "MenuXML.XML"
    End If
    If Err <> 0 Then
        MenuXMLFile = CurrApp.GetDataFolder & "MenuXML.XML"
    End If
    'ChDrive left$(App.Path, 1)
    'ChDir App.Path
        mCmdBarLoader.LoadXMLFile MenuXMLFile
        
    If Err.Number <> 0 Then
        'Me.Hide
        MsgBox "An Error occured while loading the menu data from """ & MenuXMLFile & """." & vbCrLf & "Number:" & Err.Number & vbCrLf
        'MsgBox CurrApp.GetDataFolder & "MenuXML.XML"
        MenuXMLFile = CurrApp.GetDataFolder & "MenuXML.XML"
        mCmdBarLoader.LoadXMLFile MenuXMLFile
        
        
        'Unload Me
        'End
    End If
    mCmdBarLoader.LoadScripts
    'initialize the picturebox animation class...
    
    Set mSearchAnimator = New cPicAnimator
    Set mSearchAnimator.PaintObject = PicANI
    Set mSearchAnimator.Imagelist = mCmdBarLoader.LoadImageList("SearchANI")
   
 
    Set mmenuImages = mCmdBarLoader.LoadImageList("MAINMENU16")
   
    'CmdBarMenu(0).MenuImageList = mmenuImages.hImageList
    'CmdBarMenu(0).ToolbarImageList = mmenuImages.hImageList
    If mmenuImages Is Nothing Then
        Set mmenuImages = New cVBALImageList
        With mmenuImages
            .IconSizeX = 16
            .IconSizeY = 16
            .ColourDepth = ILC_COLOR32
            
            .Create
            
            

            '.AddFromHandle Me.Icon, IMAGE_ICON, "APP"
        End With
    End If
    'mmenuImages.AddFromFile "J:\mylogo.bmp", IMAGE_BITMAP, "SSL"
    
    


   
    
    
    With lvwfilters
        .columns.Add , "TYPE", "Type"
        .columns.Add , "NAME", "Name"

        .columns.Add , "FILESPEC", "FileSpec"
        .columns.Add , "ATTRIBUTES", "Attributes"
        .columns.Add , "LARGERTHAN", "Larger Than:"
        .columns.Add , "SMALLERTHAN", "Smaller Than"
        .columns.Add , "STARTDATESPECS", "After Date"
        .columns.Add , "BEFOREDATESPECS", "Before Date"
        
    End With

 
    



    
    
    mCmdBarLoader.LoadCommandBar cmdbarmenu(0)
    
    cmdbarmenu(0).Toolbar = cmdbarmenu(0).CommandBars.Item("MAINMENU")
    
    
    'CmdBarMenu(1).Toolbar = CmdBarMenu(1).CommandBars.Item("MAINMENU")
    
    cmdbarmenu(1).ToolbarImageList = mmenuImages.hImageList

    Set cmdbarmenu(1).Toolbar = cmdbarmenu(1).CommandBars.Item("MAINTOOLBAR")

    cmdbarmenu(1).MenuImageList = mmenuImages.hImageList

    cmdbarmenu(1).Visible = True

    mCmdBarLoader.AddHandledControl cmdbarmenu(1)
    
    CDebug.Post "setting toolbarimagelist(2)..."
    cmdbarmenu(2).ToolbarImageList = mmenuImages.hImageList
    Set cmdbarmenu(2).Toolbar = cmdbarmenu(2).CommandBars.Item("FILELVWOPTIONS")
    cmdbarmenu(2).MenuImageList = mmenuImages.hImageList
    cmdbarmenu(2).Visible = False
    mCmdBarLoader.AddHandledControl cmdbarmenu(2)
    
    'now, add in the "filters" toolbar, cmdbarmenu(3) with key "FILTERSTOOLBAR"
    'VIEW::DIRSIZE::SUBMENU
    cmdbarmenu(3).ToolbarImageList = mmenuImages.hImageList
    cmdbarmenu(3).Toolbar = cmdbarmenu(0).CommandBars.Item("FILTERS::MENU::SUBMENU")
    cmdbarmenu(3).MenuImageList = mmenuImages.hImageList
    cmdbarmenu(3).Visible = True
    mCmdBarLoader.AddHandledControl cmdbarmenu(3)
    
    mCmdBarLoader.LoadExplorerBar ExplorerBar, "taskspane"
    'Do the same for the "toolbar" (cmdbarmenu(1))
    'NOTE: without a "toolbar" set, it crashes!
    
    'CmdBarMenu(0).Buttons.Add "SSL", mmenuImages.ItemIndex("SSL"), "SSL Caption", eNormal
    'CmdBarMenu(0).Buttons.Add "TESTER2", mmenuImages.ItemIndex("CONFIG"), "CONFIG", eNormal
    
    
    ' Set created = FromCollection.Add(KeyUse, PicIndex, captionuse, styleadd, tooltipuse)
    
    'CmdBarMenu(0).Buttons("SSL").Visible = True
    'CmdBarMenu(0).CommandBars("FILE::SUBMENU").Buttons.Add CmdBarMenu(0).Buttons("SSL")
    'CmdBarMenu(0).CommandBars("FILE::SUBMENU").Buttons.Add CmdBarMenu(0).Buttons("TESTER2")
    
    'Me.Move Me.left, Me.top, 7170

    Set mLookInAutoComplete = New CAutoCompleteCombo
    mLookInAutoComplete.Init cboLookin, False
    Set mFilemaskAutoComplete = New CAutoCompleteCombo
    mFilemaskAutoComplete.Init cboSimpleFileMask, False
    flinit = True
    
    'load commandbar appearance data from INI...
 
    mThemeUse = CurrApp.Settings.ReadProfileSetting("BCSearch", "Theme", "Default")
    If mThemeUse = "" Then mThemeUse = "Default"
    mThemeSection = "Appearance." & mThemeUse
   
    cmdbarmenu(0).Style = CommandBarThemeFromStr(CurrApp.Settings.ReadProfileSetting(mThemeSection, "CommandBarStyle"))
    If IsEmpty(CurrApp.Settings.ReadProfileSetting(mThemeSection, "CommandBarStyle")) Then
    cmdbarmenu(0).Style = eOfficeXP 'default.
    End If
    Err.Clear
    
    cmdbarmenu(0).Font = StringToFont(CurrApp.Settings.ReadProfileSetting(mThemeSection, "CommandBarFont"))
 
    'colour data...
    Dim strINI As String
    Dim readthemesection As String
   ' MsgBox "load colour data"
    
    'Error occurs between here...
    For I = [_eccCustomColorFirst] To [_eccCustomColorLast]
    'custom colour 4 crashes....
        'MsgBox "custom color " & I
        strINI = CommandBarStyleConv(I, True)
        'MsgBox "strINI=" & strINI
        On Error Resume Next
        If strINI <> "" Then
            readthemesection = CurrApp.Settings.ReadProfileSetting(mThemeSection, strINI, -1)
            'MsgBox "readthemesection=" & readthemesection
            If readthemesection <> "" Then
                cmdbarmenu(0).CustomColor(I) = readthemesection
                cmdbarmenu(0).UseStyleColor(I) = cmdbarmenu(0).CustomColor(I) <> -1
            End If
        Else
            
        End If
    Next I
    
    'lvwfiles.View = CurrApp.Settings.ReadProfileSetting(mThemeSection, "FilesListViewMode", eViewDetails)
    'lvwfiles.GridLines = CBool(CurrApp.Settings.ReadProfileSetting(mThemeSection, "FilesListViewGridLines", True))
    
    lvwfiles.FullRowSelect = CBool(CurrApp.Settings.ReadProfileSetting(mThemeSection, "FilesListViewFullRowSelect", True))
    
    
    '(error occurs between...) and here
    
    For I = 1 To cmdbarmenu.UBound
        cmdbarmenu(I).Style = cmdbarmenu(0).Style
        cmdbarmenu(I).Font = cmdbarmenu(0).Font
        cmdbarmenu(I).UseStyleColor(I) = cmdbarmenu(0).UseStyleColor(I)
        cmdbarmenu(I).CustomColor(I) = cmdbarmenu(0).CustomColor(I)
        
        
    
    Next I
    If Err <> 0 Then
   ' CmdBarMenu(0).Font.Name = GetDefaultUIFont
    
    End If
   ' ExtendFrame Me
    Dim loopbutton As cButton
    Dim loopindex As Long
   'subclass the listview....
'   Stop
   'Set mLvwclasser = New cSuperClass
   
   
    mFilterViewMode = CurrApp.Settings.ReadProfileSetting("BCSearch", "FilterViewMode", "0")

     'the simple pane will be visible, and  the advanced pane hidden.
     PicSimple.Visible = mFilterViewMode = FilterView_Simple
     PicFilters.Visible = Not PicSimple.Visible
    
   mLvwclasser.AddMessages True, CWM_RBUTTONUP, CWM_RBUTTONDOWN, CWM_MOUSEMOVE
   
   mLvwclasser.Subclass lvwfiles.hWnd, Me, False
   
   
   Dim filterresultssplitscalevalue As Single
   On Error Resume Next
   
   filterresultssplitscalevalue = CurrApp.Settings.ReadProfileSetting("BCSearch", "FilterResultsSplitScale", "0.5")
   If Err <> 0 Then filterresultssplitscalevalue = 0.5
 Set mPaneSplitter = New CSplitterBar
 
mPaneSplitter.Init PicUpperPane, PicLowerPane, PicSplit, SO_VERTICAL, cmdNewSearch.Top + cmdNewSearch.Height + 5, -1 * PicStatusBar.Height, filterresultssplitscalevalue
 
'Set mtaskSplitter = New CSplitterBar
ExplorerBar.Visible = CurrApp.Settings.ReadProfileSetting("BCSearch", "TaskpaneVisible", True)

    
    'InitAutocomplete cboLookin
    'Set tt = New clsTooltip
    'tt.AddTool cboLookin, "Enter the Location to begin the Search from.", vbBlack, vbYellow, True, "Look In", ttiInfo
    
    'Set tt = New ExToolTip
End Sub

Private Sub Form_Resize()
'resize the upper and lower, and statusbar pictureboxes to be the appropriate size forthe form.
'Status bar remains the same height- the others will be proportional...

Dim baseLeft As Double

If Me.WindowState = vbMinimized Then Exit Sub
    baseLeft = Me.ScaleLeft

If ExplorerBar.Visible Then
   If ExplorerBar.Width < 180 Then ExplorerBar.Width = 180
   
   
   ExplorerBar.Height = PicStatusBar.Top - ExplorerBar.Top
   baseLeft = ExplorerBar.Width
    If ScaleX(Me.Width, vbTwips, vbPixels) < baseLeft Then Me.Width = baseLeft * 1.1
    PicMain.Move baseLeft, cmdbarmenu(1).Top + cmdbarmenu(1).Height, Me.ScaleWidth - Me.ScaleLeft - ExplorerBar.Width, PicStatusBar.Top - PicMain.Top
Else
'picexpsplitter.Visible = False
    PicMain.Move baseLeft, cmdbarmenu(1).Top + cmdbarmenu(1).Height, Me.ScaleWidth - Me.ScaleLeft, PicStatusBar.Top - PicMain.Top
End If


'PicStatusBar.ZOrder 0
refreshmonitorpanel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mclasser Is Nothing Then
        mclasser.UnSubclass
    End If
    '
    CurrApp.Settings.WriteProfileSetting "BCSearch", "FilterResultsSplitScale", mPaneSplitter.SizeScale, "percentage of vertical width the Filters pane consumes"
    CurrApp.Settings.WriteProfileSetting "BCSearch", "FilterViewMode", mFilterViewMode
    CurrApp.Settings.WriteProfileSetting "BCSearch", "TaskPaneVisible", ExplorerBar.Visible
End Sub





Private Sub IChangeNotification_Change(FromNotify As bcfile.CFileChangeNotify, Changed As bcfile.NotifyChangeStruct)
Dim I As Long
CDebug.Post Changed.NumChanges & " changes detected."

For I = 0 To Changed.NumChanges
    CDebug.Post "Type=" & Changed.Changes(I).Changetype & " Name=" & Changed.Changes(I).ChangedItem

Next I

End Sub
Private Sub OpenContaining()
    'Opens the containing folder(s) of the currently selected items.
    '
    ' Explorer /e,/root,C:\TestDir\TestApp.exe


    'if more then one file is selected in a folder, don't select any of them.
    


End Sub
Private Sub IContextCallback_AfterShowMenu(idChosen As Long)
 CDebug.Post "after show menu ID=" & idChosen
 Debug.Print idChosen
 If idChosen = 5677 Then
    'remove from list.
    'lvwFiles.ListItems.Remove lvwFiles.SelectedItem.Index
    'this was the "delete" item.
    CDebug.Post "remove it!"
    On Error Resume Next 'no errors while subclassing... please...
    lvwfiles.ListItems.Remove lvwfiles.SelectedItem.Index
ElseIf idChosen = 5678 Then

    '"open containing folder"
    If Not lvwfiles.SelectedItem Is Nothing Then
        Shell "Explorer /e,/SELECT," & Replace$(lvwfiles.SelectedItem.Tag, "/", "\"), vbNormalFocus
    
    End If
    
    'the menu added to the submenu of "Copy To Clipboard"- there is one for each column.
    'they start at ID# 9000; each one is one higher based on the column position.
    
 ElseIf idChosen > 9000 And idChosen < 10000 Then
    Dim selectedcolumn As cColumn
    On Error Resume Next
    'Set selectedcolumn = lvwfiles.Columns.Item(idChosen - 9000)
    If idChosen = 9801 Then
        Dim BuildIt As String
        Dim I As Long
        BuildIt = lvwfiles.SelectedItem.Text
        For I = 1 To lvwfiles.SelectedItem.SubItems.Count
            BuildIt = BuildIt & vbTab & lvwfiles.SelectedItem.SubItems(I).caption
        Next I
        Clipboard.Clear
        Clipboard.SetText BuildIt, vbCFText
        Clipboard.SetText BuildIt
    Else
        'Set selectedcolumn = GetListViewColumnByPosition(lvwfiles, idChosen - 9000)
       Set selectedcolumn = lvwfiles.columns.Item(mMenuTempLookups(idChosen - 9000))
       'CopyMemory selectedcolumn, idChosen - 9000, 4
       
       If Not selectedcolumn Is Nothing Then
           Dim colpos As Long
           Debug.Print "copy to clicked for column:" & selectedcolumn.Text
           BuildIt = ""
           colpos = selectedcolumn.position - 1
           
           
           If colpos = 0 Then
               BuildIt = lvwfiles.SelectedItem.Text
           Else
               BuildIt = lvwfiles.SelectedItem.SubItems(selectedcolumn.position - 1).caption
           End If
           
           Clipboard.Clear
           Clipboard.SetText BuildIt, vbCFText
           Clipboard.SetText BuildIt
           
       Else
           'CDebug.Post  "error acquiring column for "; idChosen - 9000 & ":" & Err.Description
       End If
    
    End If
 End If
'ElseIf idChosen = 5679 Then
    

 'End If
End Sub
Private Function CreateScanWithMenu() As Long
    Static flInitialized As Boolean, ScanHandle As Long
    Dim IniFileLoad As String, iniloader As CINIData
    If Not flInitialized Then
        flInitialized = True
        
        'TODO:// actually get AV info- for now load INI file :(
        IniFileLoad = GetSpecialFolder(CSIDL_PROFILE).Path
        
        If Right$(IniFileLoad, 1) <> "\" Then IniFileLoad = IniFileLoad & "\"
        IniFileLoad = IniFileLoad & "BCSearchScan.ini"
        Set iniloader = New CINIData
        iniloader.LoadINI IniFileLoad
        
        
        
    
    
    
    End If
        CreateScanWithMenu = ScanHandle
    




End Function
Private Sub InitColumnsArray(columns As cColumns)

    Dim I As Long
    Erase mMenuTempLookups
    ReDim mMenuTempLookups(1 To columns.Count)
    For I = 1 To columns.Count
        mMenuTempLookups(I) = columns.Item(I).Key
        
    
    Next I






End Sub
Private Function FindColumnIndex(ColumnObj As cColumn) As Long

    Dim I As Long
    For I = 1 To UBound(mMenuTempLookups)
        If mMenuTempLookups(I) = ColumnObj.Key Then
            FindColumnIndex = I
            Exit Function
        End If
    
    Next I


End Function

Private Sub IContextCallback_BeforeShowMenu(ByVal hMenu As Long, _
                                            Optional CancelShow As Boolean)

    Static Captions() As String, NewHandle As Long

    Static IDs() As Long, flInitialized

    Static CopySubmenu As Long

    Const MIIM_STRING As Long = &H40

    Const MIIM_ID As Long = &H2

    Const MIIM_SUBMENU As Long = &H4

    'initialize data values. *Ideally* we would get this from an external source, btw...
    Const MENUID_REMOVEFROMLIST = 5677

    Const MENUID_SCAN = 5700 '(+1 for items within this menu- (the scanners)

    'step one: create the "columns" array...
    InitColumnsArray lvwfiles.columns


    'newmenuCaptions(0) = "Remove From List"
    CDebug.Post "before show menu " & hMenu
    
    'Steps:
    'we want to customize this menu, to add "Remove from list"- and a few others later maybe.
    CDebug.Post "Count of items is " & GetMenuItemCount(hMenu)

    'add separator...
    'Call AppendMenu(hMenu, MF_SEPARATOR, 0, ByVal &H0&)
    'NewHandle = AppendMenu(hMenu, MF_STRING, 5677, ByVal "Remove from List")
    Dim MenuItem As MENUITEMINFO
    
    Dim Ret As Long

    ' ret = InsertMenuItem(hMenu, 1, -1, MenuItem)
    
    '  NewHandle = AppendMenu(hMenu, MF_STRING, 5677, ByVal "Remove from List")
    
    'Dim ret As Long
    
    Dim PrependItems() As MENUITEMINFO

    ReDim PrependItems(1 To 4)

    With PrependItems(1)
       
        .fType = MF_STRING
        .dwTypeData = "Remove from List" & vbNullChar
        .cch = Len(.dwTypeData)
        .wID = 5677
        .cbSize = Len(PrependItems(1))
        .fMask = MIIM_STRING Or MIIM_ID
    End With
'TODO:// add some options for previewing of alternate streams, or extracting alternate streams.
    With PrependItems(2)
       
        .fType = MF_STRING
        .dwTypeData = "Open Containing Folder" & vbNullChar
        .cch = Len(.dwTypeData)
        .wID = 5678
        .cbSize = Len(PrependItems(2))
        .fMask = MIIM_STRING Or MIIM_ID

        With PrependItems(3)
            .fType = MF_STRING
            .dwTypeData = "Copy To Clipboard" & vbNullChar
            .cch = Len(.dwTypeData)
            .wID = 8999
            .cbSize = Len(PrependItems(3))
            .fMask = MIIM_STRING Or MIIM_ID Or MIIM_SUBMENU
        End With

        'Create the Copy To Submenu
        'note, we keep the menu handle as a static; if we already have one, we destroy it and recreate it.
        
        If CopySubmenu <> 0 Then
            DestroyMenu CopySubmenu
        End If

        CopySubmenu = CreatePopupMenu()

        Dim addthis As MENUITEMINFO

        Dim loopcolumn As cColumn, LoopI As Long

     
        'lastly, we add a few more menu items; "all visible columns" which copies each visible item in a tab-delimited fashion.
        'add a separator first...
               
        
                addthis.fType = MF_STRING
                addthis.dwTypeData = "All Visible Columns"
                addthis.cch = Len(addthis.dwTypeData)
                addthis.wID = 9801
                addthis.cbSize = Len(addthis)
                addthis.fMask = MIIM_STRING Or MIIM_ID
                InsertMenuItem CopySubmenu, 1, -1, addthis
                
           
        
        
           'For Each loopcolumn In lvwfiles.Columns
        For LoopI = lvwfiles.columns.Count To 1 Step -1
            Set loopcolumn = lvwfiles.columns(LoopI)
            'only add this column if it's visible (width>1)
            If loopcolumn.Width > 1 Then
                addthis.fType = MF_STRING
                addthis.dwTypeData = loopcolumn.Text & vbNullChar
                addthis.cch = Len(addthis.dwTypeData)
                addthis.wID = 9000 + FindColumnIndex(loopcolumn)
                addthis.cbSize = Len(addthis)
                addthis.fMask = MIIM_STRING Or MIIM_ID
                InsertMenuItem CopySubmenu, 1, -1, addthis
            End If
        Next
        
        addthis.fType = MF_SEPARATOR
         addthis.cch = 0
         addthis.dwItemData = 0
         addthis.dwTypeData = ""
         addthis.fMask = 0
         addthis.fState = 0
         addthis.hSubMenu = 0
         addthis.wID = 0
         'addthis.dwTypeData = "-"
         'addthis.cch = Len(addthis.dwTypeData)
        addthis.cbSize = Len(addthis)
        
         InsertMenuItem CopySubmenu, 9001, 0, addthis
        
        PrependItems(3).hSubMenu = CopySubmenu
        
    End With

    With PrependItems(4)
        .fType = MF_SEPARATOR
        .cbSize = Len(PrependItems(4))
    End With
        
    Dim I As Long

    For I = UBound(PrependItems) To 1 Step -1
        Ret = InsertMenuItem(hMenu, 1, -1, PrependItems(I))
    Next I

    ' ret = InsertMenuItem(hMenu, 1, -1, MenuItem)
    
    'Add items to a new submenu, the "Scan with" submenu.
    CDebug.Post "ret=" & Ret
End Sub
Private Function refreshmonitorpanel()
Dim panelrectuse As RECT, usewidth As Long
If mSearchPath = "" Then Exit Function
 With panelrectuse
 
    mSbar.GetPanelRect mSbar.PanelIndex("MONITOR"), .Left, .Top, .Right, .Bottom
    usewidth = .Right - .Left
    End With
mSbar.PanelText("MONITOR") = "Searching:" & bcfile.ShortenPath(mSearchPath, PicStatusBar.hdc, usewidth)
End Function
Private Function IFileSearchCallback_AllowRecurse(InDir As String) As Boolean
'
'TODO:// add MORE options to literally allow for a way to mask which dirs get recursed into.

'Don't recurse into Reparse points...
Dim getDir As Directory, panelrectuse As RECT, usewidth As Long
On Error GoTo NoRecurse
Set getDir = GetDirectory(InDir)
IFileSearchCallback_AllowRecurse = (chksubfolders.Value = vbChecked) And Not (getDir Is Nothing)
If chksubfolders.Value = vbUnchecked Then
    Exit Function
End If


If Not getDir Is Nothing Then
    If (getDir.Attributes And FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT Then
    CDebug.Post InDir & " is a reparse point... not recursing..."
        IFileSearchCallback_AllowRecurse = False
        Exit Function
        
    End If
   
    mSearchPath = InDir
    'mSbar.PanelText("MONITOR") = "Searching:" & BCFile.ShortenPathToCharLength(InDir, ScaleX(usewidth, vbPixels, vbCharacters))
    refreshmonitorpanel
End If

mTotalRecurse = mTotalRecurse + 1 * (Abs(getDir Is Nothing))
'IFileSearchCallback_AllowRecurse = ChkSubfolder.Value = vbChecked
Exit Function
NoRecurse:
mSbar.PanelText("MONITOR") = "Not recursing into " & InDir & " (Error occured during access)"

End Function

Private Function IFileSearchCallback_Cancelled() As Boolean
    IFileSearchCallback_Cancelled = mcancel
    mSearchAnimator.StopAnimation
    mSearching = False
    CmdStop.enabled = False
    CmdFind.enabled = True
    cmdNewSearch.enabled = True
End Function

Private Sub IFileSearchCallback_ExecuteComplete(Sender As Object)
'
If Sender Is mFileSearch Then
    're-enable search controls...
    CmdStop.enabled = False
    CmdFind.enabled = True
    cmdNewSearch.enabled = CBool(lvwfiles.ListItems.Count)
   ' SearchAnimation.StopPlay
   mSbar.PanelText("MONITOR") = "Search completed. Traversed " & mTotalRecurse & " directories, examined " & mTotalFileCount & " files."
End If
ClickHack
mSearchAnimator.StopAnimation
mSearching = False
mcancel = False
If Not mtxtSearchBar Is Nothing Then
    mtxtSearchBar.Text = mtxtSearchBar.Text
End If
End Sub

Private Function GetColumnIndex(ByVal StrKey As String) As Long

    Dim currcol As cColumn
    Dim I As Long
    For I = 1 To lvwfiles.columns.Count
        Set currcol = lvwfiles.columns.Item(I)
        If StrComp(currcol.Key, StrKey, vbTextCompare) = 0 Then
            GetColumnIndex = I
            Exit Function
        End If
    Next I


End Function

Public Sub RefreshItemData(newitem As cListItem, fileobj As CFile)
With newitem
    Set .Font = New StdFont
    .Font.Name = "MS Shell Dlg 2"
    
    .IconIndex = CurrApp.SystemIML(Size_Small).ItemIndex(fileobj.FullPath)
   
    
    
    If (fileobj.FileAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED Then
        'TODO// (Use system "compressed" colour)
        .ForeColor = vbBlue
    ElseIf (fileobj.FileAttributes And FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED Then
        .ForeColor = vbGreen
    End If
    
    If (fileobj.FileAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
        .Cut = True
    End If
    'filename,directory,type,size,modified,accessed,created
    '.SubItems("FILENAME").Caption = FileObj.Name
    
    'FILENAME,"DIRECTORY","EXTENSION","TYPE","SIZE",DISKSIZE,MODIFIED,ACCESSED,CREATED
    
    '.SubItems(GetColumnIndex("DIRECTORY")).caption = fileobj.Directory.Path
    
    '.SubItems(1).IconIndex = mSmallImages.ItemIndex(diriconkey)
    If .Text = "" Then .Text = fileobj.basename & "." & fileobj.Extension
    .SubItems(1).caption = fileobj.Directory.Path
   .SubItems(2).caption = fileobj.Extension
    .SubItems(3).caption = fileobj.FileType
     .SubItems(4).caption = bcfile.FormatSize(fileobj.size)
     .SubItems(4).ShowInTile = True
    '.ItemData = fileobj.Size
    .SubItems(5).caption = bcfile.FormatSize(fileobj.compressedsize)
    
    .SubItems(6).caption = fileobj.DateModified
    .SubItems(7).caption = fileobj.DateLastAccessed
    .SubItems(8).caption = fileobj.DateCreated
    .SubItems(7).ShowInTile = True
    .SubItems(8).ShowInTile = True
    
    '.ItemData = fileobj.size
    .Tag = fileobj.FullPath
    If IsEmpty(.Tag) Then
        Debug.Assert False
    End If
    '.indent = Len(fileobj.Fullpath) - Len(Replace$(fileobj.Fullpath, "\", ""))
    
    
    
    
    
End With
Dim loopcolumn As cColumn, currcol As Long
Dim strsplit() As String, coldata As ColumnData, gotplugs() As IColumnPlugin
gotplugs = CurrApp.ColumnPlugins
On Error Resume Next
'Final step: iterate through each ColumnHeader, and stop at those whose key starts with "PLUGIN:"...
For currcol = 1 To lvwfiles.columns.Count
    Set loopcolumn = lvwfiles.columns(currcol)
    If Left$(loopcolumn.Key, 7) = "PLUGIN:" Then
        strsplit = Split(loopcolumn.Key, ":")
        '0 is PLUGIN, 1 is plugin index, 2 is column index to pass to that plugin.
        coldata.ColumnData = ""
        coldata.ColumnIcon = 0
        coldata = gotplugs(Val(strsplit(1))).GetColumnData(newitem, Val(strsplit(2)))
        'newitem.SubItems(loopcolumn.Position - 1).caption = coldata.ColumnData
        newitem.SubItems(loopcolumn.position - 1).caption = coldata.ColumnData
    End If
    loopcolumn.ImageOnRight = True
Next currcol





End Sub

Private Sub IFileSearchCallback_Found(Sender As Object, found As String, Optional Cancel As Boolean, Optional FiltersFound As Variant)
'
Dim fileobj As CFile, diricon As Long, diriconkey As String
Dim tempicon As Long, iconkey As String, CurrIndex As Long
Dim tempcast As CSearchFilter
On Error Resume Next
mTotalFileCount = mTotalFileCount + 1
If LCase$(Right$(found, 3)) = "zip" Then
    Debug.Print "found zip file:" & found

End If
'If UCase$(GetExtension(found)) = "ZIP" Then Stop

Dim zipfound As Boolean
zipfound = True
If Not FiltersFound Is Nothing Then
    'For currindex = LBound(FiltersFound) To UBound(FiltersFound)
    For Each tempcast In FiltersFound
        If tempcast.Name <> "Zip Viewer" Then
            zipfound = False
            Exit For
        End If
    
        'Debug.Print "Filter Name:" & tempcast.Name & " REcontainmatchcolcount=" & tempcast.recontainmatchcol.Count
    Next
    
    'If zipfound Then
        
    
    'End If
    If UCase$(GetExtension(found)) = "ZIP" Then
        Debug.Print "zip file was found-(" & found & ")" & "investigate here"
        
        
        
        'if the only filter that found the zip was "Zip viewer" break out, otherwise, we'll continue so we add the file to the list.
        If zipfound Then Exit Sub
    
    
    End If
    
End If
Err.Clear

Set fileobj = GetFile(found)

'CDebug.Post  Found
If Err <> 0 Then
    CDebug.Post "error accessing " & found & Err.Description
    Exit Sub
End If
'lvwFiles.Columns.Add , "FILENAME", "Filename"
'lvwFiles.Columns.Add , "MODIFIED", "Modified"
'lvwFiles.Columns.Add , "ACCESSED", "Accessed"
'lvwFiles.Columns.Add , "CREATED", "Created"
'
'



'add icon to imagelists...

'stupid me! Don't try adding the image if we already have it !!!
'on the other hand- files like ".ICO" could have different icons... damn it. This complicates matters.
'diriconkey = "DIR" & fileobj.Directory.Path
''iconkey = "EXT" & fileobj.Extension & fileobj.FileIndex
''On Error Resume Next
''mLargeImages.AddFromHandle fileobj.GetFileIcon(icon_large), IMAGE_ICON, iconkey
''mSmallImages.AddFromHandle fileobj.GetFileIcon(ICON_SMALL), IMAGE_ICON, iconkey

'mLargeImages.AddFromHandle fileobj.Directory.GetFileIcon(icon_large), IMAGE_ICON, diriconkey
'mSmallImages.AddFromHandle fileobj.Directory.GetFileIcon(icon_shell), IMAGE_ICON, diriconkey
'mlargeimages.AddFromHandle fileobj.Directory.getfile
'If Err.Number <> 0 Then Stop
On Error Resume Next
Dim newitem As cListItem
Set newitem = lvwfiles.ListItems.Add(, , fileobj.DisplayName, CurrApp.SystemIML(Size_Large).ItemIndex(fileobj.FullPath), CurrApp.SystemIML(Size_Small).ItemIndex(fileobj.FullPath))
If Not cmdbarmenu(2).Visible Then
    cmdbarmenu(2).Visible = True
    mtxtSearchBar.Visible = True
    PicLowerPane_Resize
End If
'With newitem
'    Set .Font = New StdFont
'    .Font.Name = "MS Shell Dlg 2"
'
'    If TypeOf Sender Is CFileSearchEx Then
'        Dim getcol As Collection, filtermatched As CSearchFilter
'        Dim appearancedata As CExtraFilterData
'        Set getcol = Sender.matched
'        If getcol.Count > 0 Then
'            Set filtermatched = getcol.Item(getcol.Count)
'            Set appearancedata = filtermatched.Tag
'            .BackColor = appearancedata.BackColor
'            Set .Font = appearancedata.Font
'            .ForeColor = appearancedata.ForeColor
'
'
'        End If
'
'
'    End If
'
'
'    If (fileobj.Fileattributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED Then
'        'TODO// (Use system "compressed" colour)
'        .ForeColor = vbBlue
'    ElseIf (fileobj.Fileattributes And FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED Then
'        .ForeColor = vbGreen
'    End If
'
'    If (fileobj.Fileattributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN Then
'        .Cut = True
'    End If
'    'filename,directory,type,size,modified,accessed,created
'    '.SubItems("FILENAME").Caption = FileObj.Name
'
'    'FILENAME,"DIRECTORY","EXTENSION","TYPE","SIZE",DISKSIZE,MODIFIED,ACCESSED,CREATED
'
'    '.SubItems(GetColumnIndex("DIRECTORY")).caption = fileobj.Directory.Path
'
'    '.SubItems(1).IconIndex = mSmallImages.ItemIndex(diriconkey)
'    .SubItems(1).caption = fileobj.Directory.Path
'   .SubItems(2).caption = fileobj.Extension
'    .SubItems(3).caption = fileobj.FileType
'     .SubItems(4).caption = BCFile.FormatSize(fileobj.Size, True)
'     .SubItems(4).ShowInTile = True
'    '.ItemData = fileobj.Size
'    .SubItems(5).caption = BCFile.FormatSize(fileobj.CompressedSize, True)
'
'    .SubItems(6).caption = fileobj.DateModified
'    .SubItems(7).caption = fileobj.DateLastAccessed
'    .SubItems(8).caption = fileobj.DateCreated
'    .SubItems(9).ShowInTile = True
'    .SubItems(10).ShowInTile = True
'
'    .ItemData = fileobj.Size
'    .Tag = fileobj.Fullpath
'    '.indent = Len(fileobj.Fullpath) - Len(Replace$(fileobj.Fullpath, "\", ""))
'
'
'
'
'
'End With
'Dim loopcolumn As cColumn, currcol As Long
'Dim strsplit() As String, coldata As ColumnData, gotplugs() As IColumnPlugin
'gotplugs = CurrApp.ColumnPlugins
''Final step: iterate through each ColumnHeader, and stop at those whose key starts with "PLUGIN:"...
'For currcol = 1 To lvwfiles.Columns.Count
'    Set loopcolumn = lvwfiles.Columns(currcol)
'    If Left$(loopcolumn.key, 7) = "PLUGIN:" Then
'        strsplit = Split(loopcolumn.key, ":")
'        '0 is PLUGIN, 1 is plugin index, 2 is column index to pass to that plugin.
'        coldata = gotplugs(Val(strsplit(1))).GetColumnData(newitem, Val(strsplit(2)))
'        newitem.SubItems(loopcolumn.Position - 1).caption = coldata.ColumnData
'    End If
'
'Next currcol
CurrApp.IncrementStatistic "FilesFound"
RefreshItemData newitem, fileobj
'newitem.Tag
'SetItemExtraData newitem, matched
 If TypeOf Sender Is CFileSearchEx Then
        Dim getcol As Collection, filtermatched As CSearchFilter
        Dim appearancedata As CExtraFilterData
        Set getcol = Sender.matched
        If getcol.Count > 0 Then
            Set filtermatched = getcol.Item(getcol.Count)
            Set appearancedata = filtermatched.Tag
            newitem.BackColor = appearancedata.BackColor
            Set newitem.Font = appearancedata.Font
            newitem.ForeColor = appearancedata.ForeColor
            
            
        End If
    
    
    End If
If newitem.Tag = "" Then Debug.Assert False

mSbar.PanelText("FOUND") = lvwfiles.ListItems.Count & " Files found."

Set fileobj = Nothing
Exit Sub
ErrorOccur:
CDebug.Post Error$ & " (file:""" & fileobj.FullPath & """)"
Set fileobj = Nothing
'Resume
'Stop


End Sub
Private Function GetItemExtraData(forItem As cListItem) As CItemExtraData
    Dim I As Long
    For I = 0 To mItemDataCount
        If mItemData(I).ListItemObj Is forItem Then
            Set GetItemExtraData = mItemData(I).DataObj
        End If
    
    Next I



End Function
Private Sub SetItemExtraData(Item As cListItem, FiltersFound As Variant)

Dim CreateMatchArray() As Object
Dim CreateREArray() As String
Dim Count As Long, tempcast As CSearchFilter
Count = -1
 For Each tempcast In FiltersFound
    If tempcast.ContainsIsRegExp Then
        Count = Count + 1
        ReDim Preserve CreateMatchArray(Count)
        ReDim Preserve CreateREArray(Count)
        Set CreateMatchArray(Count) = tempcast.recontainmatchcol
        CreateREArray(Count) = tempcast.ContainsStr
        Debug.Print "Filter Name:" & tempcast.Name & " REcontainmatchcolcount=" & tempcast.recontainmatchcol.Count
        
        
        
        
        
        
    End If
    Next
        Dim newci As CItemExtraData
        Set newci = New CItemExtraData
   newci.SetREMatchCol CreateMatchArray()
   newci.setREmatchExpr CreateREArray()
    
    'newci.SetREMatchCol
    mItemDataCount = mItemDataCount + 1
    ReDim Preserve mItemData(mItemDataCount)
    Set mItemData(mItemDataCount).ListItemObj = Item
    Set mItemData(mItemDataCount).DataObj = newci
    
End Sub
Private Sub IFileSearchCallback_ProgressMessage(ByVal strMessage As String)
'
'mSBar.PanelText
End Sub

Private Sub IFileSearchCallback_SearchError(ErrCode As Long, ErrDesc As String, Cancel As Boolean)
'

End Sub

Private Sub IFilterChangeCallback_Change(changedObj As bcfile.CSearchFilter)
Dim getitem As cListItem
CDebug.Post "Change"

'grab filter listview item...
'Set getitem = lvwfilters.ListItems("TAG" & Trim$(mFileSearch.Filters.Count))
Set getitem = ListItemFromSearchFilter(changedObj)
If getitem Is Nothing Then
    Exit Sub
End If

PopulateFilterItem getitem, changedObj

'getitem.Text = changedObj.Name
'getitem.SubItems(1).caption = changedObj.FileSpec
'

End Sub

Private Sub IProgress_UpdateUI(ByVal PercentComplete As Double, ByVal StatusMessage As String)
'
If StatusMessage <> "" Then
    mSbar.SimpleMode = True
    mSbar.SimpleText = StatusMessage

End If

End Sub

'Private Sub iSuperClass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
''
''
'Dim Rct As RECT
'Select Case uMsg
'    Case CWM_SIZING
'    'wparam is the state constant-
'    'lparam points to rect.
'    CopyMemory Rct, ByVal lParam, Len(Rct)
'
'
'    '
'     'restrict width to no less then 412 pixels.
'     If Rct.right - Rct.left < 412 Then Rct.right = Rct.left + 412
'
'
'    'copy it back.
'    CopyMemory ByVal lParam, Rct, Len(Rct)
'
'
'End Select
'
'
'End Sub
Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
Debug.Print "After,", lReturn, hWnd, uMsg, wParam, lParam
End Sub



Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
Debug.Print "before ", lHandled, lReturn, hWnd, uMsg, wParam, lParam
End Sub


Private Sub lvwfiles_AfterLabelEdit(Cancel As Boolean, NewString As String, Item As vbaBClListViewLib6.cListItem)
'rename the file to the newstring...
'TAG is the filename, the edited value is the new name.

Dim gotfile As CFile
Set gotfile = GetFile(Item.Tag)
gotfile.Rename NewString, Me.hWnd, True
'Item.Text = gotfile.DisplayName
Item.Tag = gotfile.FullPath
'Item.IconIndex = CurrApp.SystemIML(Size_Small).ItemIndex(Item.Tag)

RefreshItemData Item, gotfile

End Sub

Private Sub lvwFiles_BeforeLabelEdit(Cancel As Boolean, Item As vbaBClListViewLib6.cListItem)
'
End Sub

Private Sub lvwFiles_ColumnClick(Column As cColumn)
   ' Sort according to the column type:
   ''FILENAME,"TYPE","SIZE","MODIFIED","ACCESSED","CREATED","DIR"
   
   'new addition (for emulation of Vista columns)
   'Drop-down:
   Debug.Print "Columnclick"
   If lvwfiles.View <> eViewDetails Then Exit Sub
   If Left$(Column.Tag, 9) = "DROPDOWN:" Then
    'if it is a drop-down column; show a drop-down...
    Dim dropfor As String
    dropfor = Mid$(Column.Tag, 10)
    CDebug.Post "display drop down for column:" & dropfor
    
   End If
  
'   If GetAsyncKeyState(VK_CONTROL) <> 0 Then
'
'          ' End If
'        Dim cancelled As Boolean
'        cancelled = False
'        ShowViewPopup
'        Exit Sub
'    End If
'If cancelled Then Exit Sub
'   Exit Sub
  
  
If Left$(Column.Key, 7) = "PLUGIN:" Then
    'get the appropriate plugin to handle it.
    Dim splitkey() As String
    Dim plugs() As IColumnPlugin
    Dim grabplugin As IColumnPlugin
    plugs = CurrApp.ColumnPlugins
    splitkey = Split(Column.Key, ":")
    '1 is plugin...
        Set grabplugin = plugs(Val(splitkey(1)))
        grabplugin.ColumnClick lvwfiles, Column, Val(splitkey(2))
            
    'Exit Sub
Else
  
  
   Select Case Column.Key
   
   Case "FILENAME", "DIR"
      Column.SortType = eLVSortStringNoCase
      
   Case "TYPE"
      Column.SortType = eLVSortStringNoCase
      'Column.SortOrder = NewSortOrder(Column.SortOrder)
    
   Case "SIZE"
     ' Column.SortType = eLVSortItemData
     Column.SortType = eLVSortCustom
      
      'Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case "MODIFIED", "ACCESSED", "CREATED"
        Column.SortType = eLVSortDate
       ' Column.SortOrder = NewSortOrder(Column.SortOrder)
    Case Else
        'Column.SortType = eLVSortStringNoCase
        
        'Column.SortOrder = NewSortOrder(Column.SortOrder)
   End Select
End If
   Column.SortOrder = NewSortOrder(Column.SortOrder)
   
   Dim loopcol As cColumn, currcol As Long
   For currcol = 1 To lvwfiles.columns.Count
    Set loopcol = lvwfiles.columns.Item(currcol)
    loopcol.IconIndex = -1
   Next
   
   If Column.SortOrder = eSortOrderNone Then
    Column.IconIndex = -1
   ElseIf Column.SortOrder = eSortOrderAscending Then
    Column.IconIndex = mHeaderIML.ItemIndex("ASCENDING")
    ElseIf Column.SortOrder = eSortOrderDescending Then
    Column.IconIndex = mHeaderIML.ItemIndex("DESCENDING")
   End If
   
   
   
   lvwfiles.ListItems.SortItems
End Sub

Private Function NewSortOrder(ByVal SortOrder As ESortOrderConstants) As ESortTypeConstants
   Select Case SortOrder
   Case eSortOrderNone, eSortOrderDescending
      NewSortOrder = eSortOrderAscending
   Case eSortOrderAscending
      NewSortOrder = eSortOrderDescending
   End Select
End Function
Private Function GetListItemsBetween(OnView As vbalListViewCtl, ItemA As cListItem, ItemB As cListItem) As Collection
    Dim currcol As Collection
    



End Function

Private Sub lvwfiles_CustomSortCompare(Item1 As vbaBClListViewLib6.cListItem, Item2 As vbaBClListViewLib6.cListItem, SortColumn As vbaBClListViewLib6.cColumn, ReturnValue As Long)
'had to cheat, and manually add this event myself. This changes the dependencies a bit, but for the most part the VBAccelerator control likely won't be present on most systems anyway.
Dim size(1 To 2) As Double
If SortColumn.Key = "SIZE" Then
    Debug.Print "Sorting size..."
    'ReturnValue = GetFile(Item1.Tag).Size > GetFile(Item2.Tag).Size
    If Item1.Tag = "" Then
        Item1.Tag = Item1.SubItems(1).caption & Item1.Text
    End If
    If Item2.Tag = "" Then
        Item2.Tag = Item1.SubItems(1).caption & Item2.Text
    End If
    size(1) = GetFile(Item1.Tag).size
    size(2) = GetFile(Item2.Tag).size
    If size(1) > size(2) Then ReturnValue = 1
    If size(1) < size(2) Then ReturnValue = -1
    If size(1) = size(2) Then ReturnValue = 0
    'ReturnValue = Int(Rnd * 3) - 1
    'If SortColumn.SortOrder = eSortOrderAscending Then ReturnValue = ReturnValue * -1
ElseIf Left$(SortColumn.Key, 7) = "PLUGIN:" Then
    'get the appropriate plugin to handle it.
    Dim splitkey() As String
    Dim plugs() As IColumnPlugin
    Dim grabplugin As IColumnPlugin
    plugs = CurrApp.ColumnPlugins
    splitkey = Split(SortColumn.Key, ":")
    '1 is plugin...
        Set grabplugin = plugs(Val(splitkey(1)))
        'grabplugin.ColumnClick lvwfiles, Column, Val(splitkey(2))
        ReturnValue = grabplugin.PluginColumnCompare(Item1, Item2)
End If



End Sub

Private Sub lvwfiles_ItemClick(Item As vbaBClListViewLib6.cListItem)
'    'an item was clicked.
'    Static lastClicked As cListItem    'for when we detect that "SHIFT" was held down....
'    'if the item is selected, add it to the collection. Of not, remove it.
'    'If it already does or doesn't exist, ignore.
'
'    If mCurrSelection Is Nothing Then Set mCurrSelection = New Collection
'    On Error GoTo reportandexit
'    If Item.Selected Then
'    'add
'        mCurrSelection.Add Item, "ITEM" & Item.index
'        'If GetAsyncKeyState(VK_SHIFT) <> 0 And GetAsyncKeyState(VK_CONTROL) = 0 Then
'            'if shift is down and ctrl is not...
'            'iterate from lastClicked to Item, and Select items (if "item" is not selected) or unselect (if item is selected)
'            'lastClicked.index
'            'use LVM_GETNEXTITEM logic...
'
'        'End If
'    Else
'    'remove
'        mCurrSelection.Remove "ITEM" & Item.index
'
'    End If
'    CDebug.Post  "Count of selected items = " & mCurrSelection.Count
'    Set lastClicked = Item
'    Exit Sub
'
'reportandexit:
'    CDebug.Post  Error$
End Sub
Public Function GetSelectedItems() As Collection
    Static mCol As Collection
    Dim loopItem As cListItem, I As Long
    Set mCol = New Collection
    For I = 1 To lvwfiles.ListItems.Count
    Set loopItem = lvwfiles.ListItems.Item(I)
        If loopItem.selected Then mCol.Add loopItem
    
    Next I
    Set GetSelectedItems = mCol
End Function

Private Sub lvwFiles_ItemDblClick(Item As vbaBClListViewLib6.cListItem)
    Dim Filegrab As CFile
    On Error Resume Next
    Set Filegrab = GetFile(Item.Tag)
    
    If Err <> 0 Then
        VBA.Beep
        
    
    End If
    
    
    Filegrab.Execute Me.hWnd
End Sub

Private Sub lvwfiles_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print "Keydown:"; KeyCode

Const VK_APPCOMMAND = 93
'?VK_APPS (0x5D)
If KeyCode = VK_APPCOMMAND Then

    'application key.
    'simulate a mouse click on the listview.
    If Not lvwfiles.SelectedItem Is Nothing Then
    With lvwfiles.SelectedItem
      lvwfiles_MouseUp vbRightButton, 0, .Left, .Top
    End With
    End If
Else

End If
End Sub

Private Sub lvwfiles_KeyUp(KeyCode As Integer, Shift As Integer)
    SelectionChange GetSelectedItems
End Sub

Private Sub lvwfiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Debug.Print "mousedown, button:"; Button, "shift:", Shift, X, Y


End Sub

Private Sub lvwFiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Static currhighlight As cListItem, oldColour As Long
'    Dim lvwtest As cListItem
'    Set lvwtest = lvwFiles.HitTest(X, Y)
'    CDebug.Post  lvwtest Is Nothing
'    If Not currhighlight Is lvwtest Then
'        If Not currhighlight Is Nothing Then
'            currhighlight.ForeColor = oldColour
'        End If
'            oldColour = lvwtest.ForeColor
'            lvwtest.ForeColor = &H909090
'            Set currhighlight = lvwtest
'
'    End If



'Dim FirstX As Long, FirstY, flInit As Boolean
'Static gotmetrics As New systemmetrics, dragging As Boolean
'Dim dx As Long, dy As Long
'dx = gotmetrics.GetMetric(CSM_CXDRAG)
'dy = gotmetrics.GetMetric(CSM_CYDRAG)
'
'If dragging Then
'    Debug.Print "dragging...."
'Else
'    If Not flInit And (Button And vbLeftButton) = vbLeftButton Then
'        flInit = True
'        FirstX = x
'        FirstY = y
'    ElseIf flInit And (Button And vbLeftButton) = vbLeftButton Then
'        'check to see if we've moved enough.
'        If Abs(FirstX - x) > dx And Abs(firsy - y) > dy Then
'
'
'
'        End If
'    ElseIf flInit Then
'        finit = False
'        FirstX = 0
'        FirstY = 0
'
'    End If
'End If

CDebug.Post "mousemove button=" & Button & ",shift=" & Shift & ",x=" & X & ",y=" & Y
End Sub
Private Function IsPressed(VirtualKey As Long) As Boolean
 IsPressed = GetKeyState(VirtualKey) < 0
End Function
Private Function ColumnHitTest(ByVal X As Long, ByVal Y As Long) As cColumn

Dim Ret As cColumn, headerhandle As Long


'retrieve header handle:

headerhandle = SendMessage(lvwfiles.hWndListView, LVM_GETHEADER, 0, ByVal 0&)

'retrieve using HDM hittest...
Dim ht As HD_HITTESTINFO
ht.pt.X = X
ht.pt.Y = Y
'ht.flags=
Call SendMessage(headerhandle, HDM_HITTEST, 0, VarPtr(ht))
If ht.flags = HHT_ONHEADER Then
    Set Ret = lvwfiles.columns(ht.iItem)
End If

Set ColumnHitTest = Ret

End Function
Private Function GetHeaderHandle(ByVal ofLvw As Long) As Long
    GetHeaderHandle = SendMessage(ofLvw, LVM_GETHEADER, 0, ByVal 0&)
End Function
Private Sub lvwfiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fileobjgrab As CFile, Ret As Long, selitems As Collection
    Dim cursor As POINTAPI, lvheaderheight As Long, buttonuse As Long
    'CDebug.Print "mouseup "; Button, Shift, x, y
    If X < 0 Or Y < 0 Then Exit Sub
    CDebug.Post "mouseup button=" & Button & ",shift=" & Shift & ",x=" & X & ",y=" & Y
    Const VK_RBUTTON As Long = &H2
    Const VK_LBUTTON As Long = &H1
    Const VK_MBUTTON As Long = &H4
    Const VK_ALT = &H12&
    'rbuttonpress = GetAsyncKeyState(VK_RBUTTON)
    If lvwfiles.SelectedItem Is Nothing Then
        If lvwfiles.ListItems.Count > 0 Then
            lvwfiles.ListItems(1).selected = True
        End If
    End If
    'BUG: seems the listview doesn't pass button arguments when we have no selected item.
    'use GetAsyncKeyState to acquire the buttons pressed, and use those instead to populate the button argument.
    'If Button = 0 Then
'        If IsPressed(VK_LBUTTON) Then
'            Debug.Print "left button pressed."
'            Button = vbLeftButton
'        End If
'        If IsPressed(VK_RBUTTON) Then
'            Button = Button + vbRightButton
'        End If
'        If IsPressed(VK_MBUTTON) Then
'            Button = Button + vbMiddleButton
'        End If
'        'shift...
'
'        If IsPressed(VK_SHIFT) Then
'            Shift = vbShiftMask
'        End If
'        'control
'        If IsPressed(VK_CONTROL) Then
'            Shift = Shift + vbCtrlMask
'        End If
'        'alt
'        If IsPressed(VK_ALT) Then
'            Shift = Shift + vbAltMask
'        End If
'
'        Debug.Print "button=" & Button, "shift=" & Shift
'
    'End If
    
    
    
    buttonuse = Button
    lvheaderheight = GetListViewHeaderHeight(lvwfiles.hWndListView)
    'CDebug.Post  "mouseup:", x, y, "Headerheight:" & lvheaderheight
    
    
    
        
    
    
    
    
    
    If Not lvwfiles.HitTest(X, Y) Is Nothing Then
    
        CDebug.Post lvwfiles.HitTest(X, Y).Tag
    
    End If
    Set mCurrSelection = GetSelectedItems
'    If GetSelectedCount(lvwFiles.hwnd) > 1 Then
'    'multiple selections...
'    Dim Cursor As POINTAPI
'    GetCursorPos Cursor
'
'    'CmdBarMenu.Item(0).ShowPopupMenu Cursor.x, Cursor.y, CmdBarMenu(0).CommandBars("POPUP")
'    'items within the popup will act on those files/items in the collection...
'
'
    'Else
    
    'wed march 25th
    'moved logic from this procedure to separate "showfilepopup" routine as I needed it for the File menu "button" at the top.
    
    If Y < lvheaderheight Then
        Dim selColumn As cColumn
        
        
        'use columnhittest...
        'Set selColumn = ColumnHitTest(x, y)
        'GetCursorPos cursor
        cursor.X = X
        cursor.Y = lvheaderheight / 2
         ClientToScreen lvwfiles.hWnd, cursor
        
        cmdbarmenu(0).ShowPopupMenu cursor.X, cursor.Y, cmdbarmenu(0).CommandBars("VIEW::SUBMENU")
        'ShowViewPopup
        
    
    
        Exit Sub
    End If
    If (buttonuse And vbRightButton) = vbRightButton Then
        If mCurrSelection.Count > 0 Then
            If Not mNonLocal Then
    '          ClientToScreen lvwfiles.hWnd, cursor
    '          cursor.x = cursor.x + mCurrSelection.Item(1).Left
    '          cursor.y = cursor.y + mCurrSelection.Item(mCurrSelection.Count).Top
              Dim LvwRect As RECT
              Call GetWindowRect(lvwfiles.hWnd, LvwRect)
              With LvwRect
              Debug.Print "Left:"; .Left; " top:"; .Top; " Right:"; .Right; " Bottom:"; .Bottom
              End With
              'ShowFilepopup x + lvwfiles.Left + lvwfiles.Parent.Left, y + lvwfiles.Top + lvwfiles.Parent.Left
              ShowFilepopup LvwRect.Left + X, LvwRect.Top + Y
              'ClickHack
            Else
                MsgBox "The currently displayed search results offer a non-local view; context menus are disabled in non-local views."
            
            End If
              'GetCursorPos cursor
              'CmdBarMenu(0).ShowPopupMenu cursor.x, cursor.y, CmdBarMenu(0).CommandBars("VIEW::SUBMENU")
        End If
    Else
        SelectionChange GetSelectedItems
    
    
    End If
'    End If 'multiple selections...
End Sub
Private Sub SelectionChange(SelectedItems As Collection)
    Debug.Print "Selection has changed Count:" & SelectedItems.Count
    'there are <at least> three "bars" on the explorer bar.
    'the "default" bar, loaded via XML (with the basic new search, start search, etc commands as links)
    Dim expbar As vbalExplorerBarLib6.cExplorerBar, gotcfile As CFile
    Dim exifinfo As CExifData, propnames() As String, propValues() As Variant, pnCount As Long, pvCount As Long, I As Long
    Set exifinfo = New CExifData
    On Error Resume Next
    ExplorerBar.Bars.Remove "SINGLEITEM"
    ExplorerBar.Bars.Remove "SINGLEITEMPROP"
    ExplorerBar.Bars.Remove "MULTIITEM"
    ExplorerBar.Bars.Remove "MULTIITEMPROP"
    Err.Clear
    If SelectedItems.Count = 1 Then
        'show some statistics about the file in the explorer bar...
        Dim selitem As Object
        Set selitem = SelectedItems(1)
        Set expbar = ExplorerBar.Bars.Add(, "SINGLEITEM", "Item Tasks")
        expbar.items.Add , "SINGLEITEM::COPY", "Copy  " & selitem.Text, mmenuImages.ItemIndex("COPY"), eItemLink
        expbar.items.Add , "SINGLEITEM::MOVE", "Move " & selitem.Text, mmenuImages.ItemIndex("MOVETO"), eItemLink
        expbar.items.Add , "SINGLEITEM::RENAME", "REname " & selitem.Text, , eItemLink
        expbar.items.Add , "SINGLEITEM::OPEN", "Open " & selitem.Text, mmenuImages.ItemIndex("RUN"), vbalExplorerBarLib6.EExplorerBarItemTypes.eItemLink
        expbar.items.Add , "SINGLEITEM::PRINT", "Print " & selitem.Text, mmenuImages.ItemIndex("PRINTERS"), eItemLink
        expbar.items.Add , "SINGLEITEM::DELETE", "Delete " & selitem.Text, mmenuImages.ItemIndex("CUT"), eItemLink
        'populates the "tasks" pane...
        Set gotcfile = GetFile(SelectedItems.Item(1).Tag)
        If Not gotcfile Is Nothing Then
            'create the "properties" pane, which will show selection properties...
            Set expbar = ExplorerBar.Bars.Add(, "SINGLEITEMPROP", "Properties")
            expbar.IconIndex = CurrApp.SystemIML(Size_Large).ItemIndex(gotcfile.FullPath)
            'expbar.IconIndex = m
            With expbar
                .items.Add , "SIPROP::LOCATION", "Location:" & vbTab & gotcfile.Directory.Path, , eItemText
                .items.Add , "SIPROP::SIZE", "Size:" & vbTab & bcfile.FormatSize(gotcfile.size), , eItemText
                .items.Add , "SIPROP::TYPE", "Type:" & vbTab & gotcfile.FileType, , eItemText
                .items.Add , "SIPROP::ATTRIBUTES", "Attributes:" & vbTab & gotcfile.GetAttributeString(True), , eItemText
                Dim gdip As GDIPImage, gdithumb As GDIPImage
                Set gdip = New GDIPImage
                
                On Error Resume Next
                 gdip.FromFile gotcfile.FullPath
                If Err = 0 Then
                .items.Add , "SIPROP::SIZE", "Size:" & gdip.Width & " x " & gdip.Height & " pixels.", , eItemText
                End If
                
                
                     On Error Resume Next
                exifinfo.LoadEXIF gotcfile.FullPath
                If Err = 0 Then
                    propnames = exifinfo.GetPropertyNames(pnCount)
                    propValues = exifinfo.GetPropertyValues(pvCount)
                    For I = LBound(propnames) To UBound(propnames)
                        .items.Add , "SIPROP::" & propnames(I), propnames(I) & ":" & vbTab & propValues(I), , eItemText
                    
                    Next I
                
                End If
            
            
            
            
            End With
        Else
        
        
        End If
        
    ElseIf SelectedItems.Count > 1 Then
        Set expbar = ExplorerBar.Bars.Add(, "MULTIITEM", "Item Tasks")
        
        
        expbar.items.Add , "MULTIITEM::COPY", "Copy the selected files", mmenuImages.ItemIndex("COPY"), eItemLink
        expbar.items.Add , "MULTIITEM::MOVE", "Move the selected files", mmenuImages.ItemIndex("MOVETO"), eItemLink
        expbar.items.Add , "MULTIITEM::RENAME", "Rename the selected files", eItemLink
        expbar.items.Add , "MULTIITEM::DELETE", "Delete the selected files", mmenuImages.ItemIndex("CUT"), eItemLink
        Dim Loopobj As Object, Getcf As CFile, accumsize As Double
        Set expbar = ExplorerBar.Bars.Add(, "MULTIITEMPROP", "Details")
        expbar.items.Add(, "MULTIITEM::Count", SelectedItems.Count & " Items selected.", , vbalExplorerBarLib6.EExplorerBarItemTypes.eItemText).SpacingAfter = 16
        
        For Each Loopobj In SelectedItems
            Set Getcf = GetFile(Loopobj.Tag)
            
            accumsize = Getcf.size + accumsize
            
        
        Next
        
        
        expbar.items.Add , "MIPROP::SIZE", "Total file size:" & FormatSize(accumsize), , eItemText
        
    End If
    
    
End Sub
Private Sub ClickHack()
'SendMessage lvwfiles.hWndListView, CWM_RBUTTONDOWN, &H2, ByVal MakeLong(-5, -5)
End Sub
Private Sub ShowFilepopup(Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1)
100    Dim fileobjgrab As CFile, selitems As Collection, Ret As Long, cursor As POINTAPI
101    Dim SameFolder As Boolean, loopItem As cListItem
102    Dim FFile As CFile, fdir As Directory
103    Dim LastDir As String, FileArray() As String, FileCount As Long
    
104    On Error GoTo PopError
105    GetCursorPos cursor
106    FileCount = -1
    'change cursor position if arguments were passed.
107    If X > -1 Then cursor.X = X
108    If Y > -1 Then cursor.Y = Y
109  Set selitems = GetSelectedItems
  'TODO:// change to use Shell menu if all selected items are in the same folder.

110        If selitems.Count > 1 Then
111         SameFolder = True
112            For Each loopItem In selitems
                'grab the CFile object from the tag...
113                Set FFile = GetFile(loopItem.Tag)
114                Set fdir = FFile.Directory
115                If LastDir <> "" And StrComp(LastDir, fdir.Path, vbTextCompare) = 0 Then
                    'they are different.
116                    CDebug.Post LastDir & " differs from " & fdir.Path & "."
117                    SameFolder = False
118                    Exit For
119                End If
                'Add it to the string array that we might use.
120                FileCount = FileCount + 1
121                ReDim Preserve FileArray(FileCount)
122               FileArray(FileCount) = loopItem.Tag
123                CDebug.Post "adding " & loopItem.Tag & " to array"
124                LastDir = fdir.Path
125            Next
        
        
        
126            If Not SameFolder Then
127                GetCursorPos cursor
128                Set mCurrSelection = selitems
                    On Error GoTo PopError
                    Dim popupitem As cCommandBar
                    Set popupitem = cmdbarmenu(0).CommandBars("POPUP")
                    'Err.Clear
                    If Not popupitem Is Nothing Then
                    
129                    cmdbarmenu(0).ShowPopupMenu cursor.X, cursor.Y, popupitem
                        CDebug.Post "after showpopupmenu."
                    Else
                        CDebug.Post "popupitem was nothing."
                    End If
                    If Err.Number <> 0 Then
                        CDebug.Post Err.Number, Err.Description, Erl
                        Err.Clear
                    End If
130            Else
                'Target:
                'use exposed functions.
131                CDebug.Post "showing multi menu"
                    
132                ShowExplorerMenuMulti Me.hWnd, LastDir, FileArray(), cursor.X, cursor.Y, Me
                
                
133            End If
        
134        Else
135            If Not lvwfiles.SelectedItem Is Nothing Then
                
                'If Button And vbRightButton Then
                
136                    Set fileobjgrab = GetFile(lvwfiles.SelectedItem.Tag)
                    'display the explorer menu for this file.
                    
137                    Ret = fileobjgrab.ShowExplorerMenu(Me.hWnd, cursor.X, cursor.Y, Me)
                
                
                    
                'End If
138            Else
                'generic menu...
139                cmdbarmenu.Item(0).ShowPopupMenu cursor.X, cursor.Y, cmdbarmenu(0).CommandBars("DEFFILE")
            
            
140            End If
141        End If
142    Exit Sub
PopError:
If Err.Number <> 0 Then
    If Erl = 136 Then
        'error occured accessing file.
        lvwfiles.SelectedItem.ForeColor = vbBlack
        lvwfiles.SelectedItem.BackColor = vbRed
        MsgBox "Failed to access file, " & lvwfiles.SelectedItem.Tag & "(" & Err.Description & ")"
    Else


        'CDebug.Post Err.Number, Err.Description, "Line Number:" & Erl
        Resume Next
    End If
End If
End Sub

Private Sub lvwfiles_OLECompleteDrag(Effect As Long)
'
End Sub

Private Sub lvwFiles_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'
End Sub

Private Sub lvwFiles_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'
End Sub

Private Sub lvwFiles_OLESetData(Data As DataObject, DataFormat As Integer)
Dim loopItem As cListItem

If mCurrSelection.Count > 0 Then
    For Each loopItem In mCurrSelection
        Data.Files.Add loopItem.Tag
        

    Next loopItem
End If
End Sub

Private Sub lvwfiles_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Dim loopItem As cListItem
Debug.Print "OLEStartDrag"
Set mCurrSelection = GetSelectedItems()
If mCurrSelection Is Nothing Then Exit Sub
If mCurrSelection.Count > 0 Then
    Data.SetData , vbCFFiles
    
    For Each loopItem In mCurrSelection
        Data.Files.Add loopItem.Tag
        

    Next loopItem
    Debug.Print "OLEStartDrag of " & Data.Files.Count & " files..."
    AllowedEffects = 1
End If
End Sub

Private Sub lvwFilters_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cursor As POINTAPI

GetCursorPos cursor
If Not lvwfilters.SelectedItem Is Nothing Then
    If Button = vbRightButton Then
    cmdbarmenu.Item(0).ShowPopupMenu cursor.X, cursor.Y, cmdbarmenu(0).CommandBars("FILTER")
    
    End If

End If
End Sub

Private Sub mCmdBarLoader_ControlAdded(ControlObj As Object, BelongsTo As Object)
'
If StrComp(ControlObj.Name, "txtFilterResultsBar", vbTextCompare) = 0 Then
    Debug.Print "txtFilterResultsBar Added"
    
    Set mtxtSearchBar = ControlObj
    mtxtSearchBar.Visible = True
End If

End Sub

Private Sub mCmdBarLoader_GetImageList(ByVal Keyof As String, retList As cVBALImageList)
'
Debug.Assert False
End Sub

Private Sub mCmdBarLoader_ResolveFilename(ByVal FilenameRes As String, forNode As Object, TargetObject As Object, handled As Boolean)
'special handling for BCRS files (BASeCamp Resource). They basically will be a CSimpleFilePackage.

Dim castList As cVBALImageList
Dim castnode As XMLNode
Dim resourcegrabber As cResource

Dim FilePath As String
Dim resourcetype As String, ResourceID As String
Dim retArray() As Byte
Dim splitreverse() As String

'replace "res://" and "res:\\" and so forth with uppercase equivalents.
Mid$(FilenameRes, 1, 6) = UCase$(Left$(FilenameRes, 6))
If Left$(FilenameRes, 6) = "RES:\\" Or Left$(FilenameRes, 6) = "RES://" Then
    'Syntax: <listimage src="RES://<filename>,restype,resid
    Set resourcegrabber = New cResource
    FilenameRes = Mid$(FilenameRes, 6)

    'deal with filenameres and depending on what "forNode" happens to be, either add items to the imagelist or whatever the targetobject is.
    
    'tad kludgy...
    splitreverse = Split(StrReverse(FilenameRes), ",", 2)
    
    ResourceID = StrReverse(splitreverse(0))
    resourcetype = StrReverse(splitreverse(1))
    FilenameRes = StrReverse(splitreverse(2))
    
    
    
    Set resourcegrabber = New cResource
    With resourcegrabber
        .LibraryFilePath = FilenameRes
        Call .RES_LoadData(Val(ResourceID), retArray(), , , resourcetype)
        
    
    
    End With
    Set castnode = forNode
    Set castList = TargetObject
    
    
    
    
    
    
    handled = True

End If
End Sub

Private Sub mCmdBarLoader_ResolveImageKey(ByVal KeyString As String, ByRef PicIndex As Long)
PicIndex = mmenuImages.ItemIndex(KeyString)
End Sub

Private Sub mFormSizer_GetMinmaxSizeInfo(MinWidth As Long, MinHeight As Long, MaxXPos As Long, MaxYPos As Long, maxWidth As Long, MaxHeight As Long)
If ExplorerBar.Visible Then
    MinWidth = MinWidth + ExplorerBar.Width
End If
    
End Sub

Private Sub mtxtSearchBar_Change()
'
If mSearching Then Exit Sub
If tmrFilterResults.enabled = True Then
    tmrFilterResults.enabled = False
End If

tmrFilterResults.Interval = 1000
tmrFilterResults.enabled = True

End Sub

Private Sub PicANI_Click()
'If Not testnotify Is Nothing Then testnotify.Stop_
'Set testnotify = New CDevChangeNotify
'testnotify.Start
'Load frmRename
'frmRename.Init lvwfiles
'frmRename.Show

FrmAction.ShowDialog Me


End Sub

Private Sub PicClientArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'PaneSplitterHorz.MouseDown Button, Shift, X, Y
    
End Sub

Private Sub PicClientArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'PaneSplitterHorz.MouseMove Button, Shift, X, Y
End Sub

Private Sub PicClientArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'PaneSplitterHorz.MouseUp Button, Shift, X, Y
End Sub

Private Sub PicClientArea_Resize()
'
If mPaneSplitter Is Nothing Then Exit Sub 'still initializing the form.
If PicUpperPane.Visible Then
mPaneSplitter.resizetocontainer
Else
    PicLowerPane.ZOrder 0
    PicLowerPane.Move 0, 0, ScaleWidth, ScaleHeight - PicSplit.Height



End If
End Sub




Private Sub PicFilters_Resize()
On Error Resume Next

 'cmdAddFilter.Move 0, 0
 '   lvwfilters.Move cmdAddFilter.Width + cmdAddFilter.Left, 15, PicFilters.ScaleWidth - cmdAddFilter.Width - 15, PicFilters.Height - lvwfilters.Top - 30
 cmdbarmenu(3).Move 0, 0
 lvwfilters.Move cmdbarmenu(3).Left + cmdbarmenu(3).Width, 0, PicFilters.ScaleWidth - cmdbarmenu(3).Width - 0, PicFilters.Height - 10
End Sub

Private Sub PicLowerPane_Resize()
 '11/4/2009-06:05:10
 'Changed resize code here to accomodate "cmdbarmenu(2)" which is the menu that will now be shown above the file list when there are files in the list.
 
    Dim lvwleft As Long, lvwtop As Double, lvwwidth As Double, lvwheight As Double

    If cmdbarmenu(2).Visible = False Then

        With PicLowerPane

            lvwleft = .ScaleLeft
            lvwtop = .ScaleTop
            lvwwidth = .ScaleWidth
            lvwheight = .ScaleHeight
            lvwfiles.Move lvwleft, lvwtop, lvwwidth, lvwheight
        End With

    Else

        With cmdbarmenu(2)
            If PicUpperPane.Visible Then
                cmdbarmenu(2).Orientation = eTop
                
                On Error Resume Next
                    .Move PicLowerPane.ScaleLeft, PicLowerPane.ScaleTop, PicLowerPane.ScaleWidth, .Height
                    If Err <> 0 Then
                    .Move PicLowerPane.ScaleLeft, PicLowerPane.ScaleTop, PicLowerPane.ScaleWidth, 20
                    End If
                    lvwleft = PicLowerPane.ScaleLeft
                    lvwtop = Abs(.Height * (.Align = vbAlignNone))
                    lvwheight = PicLowerPane.ScaleHeight - lvwtop
                    lvwwidth = PicLowerPane.ScaleWidth
                    
                
                
        
                
                    lvwfiles.Move lvwleft, lvwtop, lvwwidth, lvwheight '- PicStatusBar.ScaleHeight
                    mtxtSearchBar.Visible = True
                    mtxtSearchBar.Move 1440, 2, 2500, cmdbarmenu(2).Height - 15
                    'lvwFiles.Visible = True
                    'If Not mtxtSearchBar Is Nothing Then
                        'mtxtSearchBar.Height = cmdbarmenu(2).Height - 15
                        
            Else
                'if the upper pane is not visible, we will align cmdbarmenu to the left side.
              
                    .Orientation = eLeft
                    
                    .Move PicLowerPane.ScaleLeft, PicLowerPane.ScaleTop, 20, PicLowerPane.ScaleHeight
                    lvwfiles.Move .Width, PicLowerPane.ScaleTop, PicLowerPane.ScaleWidth - .Width, PicLowerPane.ScaleHeight
                    
              
            
            End If
        End With
        End If
    
        
    
    
    End Sub

Private Sub PicMain_Resize()
    PicClientArea.Move PicMain.ScaleLeft, PicMain.ScaleTop, PicMain.ScaleWidth, PicMain.ScaleHeight
End Sub

Private Sub PicSimple_Resize()
On Error Resume Next
    cboSimpleFileMask.Move cboSimpleFileMask.Left, cboSimpleFileMask.Top, PicSimple.ScaleWidth - cboSimpleFileMask.Left - (PicSimple.ScaleWidth / 10)
    chkSimpleregExp.Move cboSimpleFileMask.Left + cboSimpleFileMask.Width - (PicSimple.ScaleWidth / 11) - chkSimpleregExp.Width, chkSimpleregExp.Top, chkSimpleregExp.Width, chkSimpleregExp.Height
    
End Sub

Private Sub PicStatusBar_Paint()
    mSbar.Draw
End Sub



Private Sub PicTabDateSpec_Click()
'
End Sub



Private Sub PicTopLev_Click()

End Sub

Private Sub PicTopLev_Resize()

End Sub

Private Sub PicUpperPane_Resize()
Dim newleft As Double
If Me.WindowState = vbMinimized Then Exit Sub
newleft = PicUpperPane.ScaleWidth - (CmdFind.Width + PicUpperPane.ScaleX(2, vbMillimeters, PicUpperPane.ScaleMode))
CmdFind.Move newleft
cmdNewSearch.Move newleft
CmdStop.Move newleft

'move cboLookin to be wider, and move cmdbrowse as well.
On Error Resume Next

cmdBrowse.Move newleft - cmdBrowse.Width - (cmdBrowse.Width / 5)

'and resize the combobox, too...
cboLookin.Width = (cmdBrowse.Left - cboLookin.Left) - 5


Dim I As Long
On Error Resume Next
PicFilters.Move lblFilters.Left, lblFilters.Top + lblFilters.Height, PicUpperPane.ScaleWidth - (PicUpperPane.ScaleWidth - CmdFind.Left) - lblFilters.Left, PicUpperPane.Height - (lblFilters.Top + lblFilters.Height)
PicSimple.Move lblFilters.Left, lblFilters.Top + lblFilters.Height, PicUpperPane.ScaleWidth - (PicUpperPane.ScaleWidth - CmdFind.Left) - lblFilters.Left, PicUpperPane.Height - (lblFilters.Top + lblFilters.Height)
PicBasic.Move PicFilters.Left, PicFilters.Top, PicFilters.Width, PicFilters.Height
PicANI.Move PicFilters.Left + PicFilters.Width, cmdNewSearch.Top + cmdNewSearch.Height
'otherresizing...


End Sub

Private Sub LoadXMLMenus()

    'load the XML menus.
    'currently loads from file:
    'app.Path & menuxml.xml.
    Set mCmdBarLoader = New CXMLLoader
    'mCmdBarLoader.LoadXMLFile App.Path & "\menuxml.xml"
    'mCmdBarLoader.LoadCommandBar CmdBarMenu(0)
    


End Sub
Private Sub OpenSavedSearch(Fget As CFile)
    'open a saved search...
    Dim PropBag As PropertyBag, gotbytes() As Byte
    Dim Instream As FileStream
    Dim FileCount As Long
    Dim readLong As Long
    Dim DateValue As Date, username As String, CompName As String
    Dim LenCompname As Long, LenUserName As Long
    Dim Filename As String, fnamelen As Long
    
    Dim LenFilterData As Long
    Dim FilterData() As Byte
    
    
    On Error GoTo FileError
    Set Instream = Fget.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING)
    readLong = Instream.readLong
    
    If readLong <> Magic_Number Then
        Err.Raise 9, "FrmSearch::OpenSavedSearch", "The file, """ & Fget.FullPath & """ is not a Saved Search file."
    
    End If
    'go on...
    
'File format:
'Magic_Number is the first 2 bytes. (Long)

DateValue = CDbl(Instream.ReadDouble())

'then a "double" value, casted from the Now() Date value (cast back for the date the file was written originally)
'length of username

LenUserName = Instream.readLong
username = Instream.readstring(LenUserName)
LenCompname = Instream.readLong
CompName = Instream.readstring(LenCompname)

mNonLocal = True
Me.caption = "BCSearch - " & Fget.Filename & " (" & username & " On \\" & CompName & "\)"
readLong = Instream.readLong
'read the bytes for the property bag...
FilterData = Instream.readbytes(readLong)
Set PropBag = New PropertyBag

'Dim bytecopy() As Byte
'ReDim bytecopy(0 To UBound(FilterData) - 1)
'CopyMemory bytecopy(0), FilterData(1), readLong

PropBag.Contents = FilterData
Set mFileSearch.Filters = PropBag.ReadProperty("FILTERS")

Dim LoopFilter As CSearchFilter
Dim newitem As cListItem
lvwfilters.ListItems.Clear
Dim CurrFilter As Long

For CurrFilter = 1 To mFileSearch.Filters.Count - 1
    Set LoopFilter = mFileSearch.Filters.Item(CurrFilter)
    
    Set newitem = lvwfilters.ListItems.Add
    PopulateFilterItem newitem, LoopFilter

Next




'length of "filters" Saved byte propertybag
'"filters" Saved byte property bag.
'number of Files in the search results
'clear the listview... also the columns...
lvwfiles.ListItems.Clear
lvwfiles.columns.Clear
FileCount = Instream.readLong
Dim CurrFile As Long


Dim lenBuildStr As Long, buildstr As String

'now, loop...
For CurrFile = 1 To FileCount

    'repeated for each file (the above value number of times)
    
    'length of the description string
    'description string
    
    fnamelen = Instream.readLong
    Filename = Instream.readstring(fnamelen)
    
    lenBuildStr = Instream.readLong
    buildstr = Instream.readstring(lenBuildStr)
    
    
    'split buildstr at null characters...
    
    
    CDebug.Post buildstr
    
    'the description string is a string that is null delimited; the format is "column name\0column value\0column name\0column value
Next CurrFile
    
    Exit Sub
    
FileError:
    
    Stop
    Resume
    
    
    
End Sub
Private Sub SaveSearch(ByVal StrFile As String)
'save to the given file.
Dim OutStream As FileStream
Dim PropBag As PropertyBag, gotbytes() As Byte
Dim currItem As Long, currlvw As cListItem

Set PropBag = New PropertyBag



'File format:
'Magic_Number is the first 2 bytes. (Long)
'then a "double" value, casted from the Now() Date value (cast back for the date the file was written originally)
'length of username
'username
'length of computername
'computername

'length of "filters" Saved byte propertybag
'"filters" Saved byte property bag.
'number of Files in the search results


'repeated for each file (the above value number of times)

'length of the description string
'description string

'the description string is a string that is null delimited; the format is "column name\0column value\0column name\0column value


PropBag.WriteProperty "FILTERS", mFileSearch.Filters


gotbytes = PropBag.Contents
Set OutStream = CreateStream(StrFile)
'first, write out a but of header info to the file

OutStream.WriteLong Magic_Number
OutStream.WriteDouble Now
OutStream.WriteLong Len(GetUserName)
OutStream.WriteString GetUserName
OutStream.WriteLong Len(GetComputerName)
OutStream.WriteString GetComputerName



CDebug.Post "saved length:" & UBound(gotbytes) + 1
OutStream.WriteLong UBound(gotbytes) + 1
OutStream.writebytes gotbytes
Dim gotfile As CFile, LoopStream As CAlternateStream
OutStream.WriteLong lvwfiles.ListItems.Count
For currItem = 1 To lvwfiles.ListItems.Count
    Set currlvw = lvwfiles.ListItems(currItem)
    OutStream.WriteLong Len(currlvw.Tag)
    OutStream.WriteString currlvw.Tag, StrRead_ANSI
    
    
    'write out some info about the file.
    
    
    Set gotfile = GetFile(currlvw.Tag)
    
    Dim buildstr As String
    buildstr = "Size" & vbNullChar & gotfile.size & vbNullChar
    buildstr = buildstr & "Compressed Size" & vbNullChar & gotfile.compressedsize & vbNullChar
    buildstr = buildstr & "type" & vbNullChar & gotfile.FileType & vbNullChar
    buildstr = buildstr & "Created" & vbNullChar & Format$(gotfile.DateCreated, "MM/DD/YYYY hh:mm:ss") & vbNullChar
    buildstr = buildstr & "Modified" & vbNullChar & Format$(gotfile.DateModified, "MM/DD/YYYY hh:mm:ss") & vbNullChar
    buildstr = buildstr & "Accessed" & vbNullChar & Format$(gotfile.DateLastAccessed, "MM/DD/YYYY hh:mm:ss") & vbNullChar
    buildstr = buildstr & "attributes" & vbNullChar & gotfile.FileAttributes & vbNullChar
    buildstr = buildstr & "AlternateStreamCount" & vbNullChar & gotfile.AlternateStreams.Count
    'store the current length of this string to disk...
    'it will be used to load it.
    OutStream.WriteLong (Len(buildstr))
    
    'store the string itself...
    OutStream.WriteString buildstr
    'Now: we create a block of data containing the alternate streams of the file; these are generally small, thankfully...
'    Dim AltStrBlock As String
'    Dim BinStream As FileStream, hasaltstreams As Boolean
'    Dim readdata() As Byte, count As Long
'    On Error GoTo AltStreamError
'    count = GotFile.AlternateStreams.count
'    Outstream.WriteLong count
'    For Each LoopStream In GotFile.AlternateStreams
'        hasaltstreams = True
'        Set BinStream = LoopStream.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_DELETE + FILE_SHARE_READ, OPEN_EXISTING)
'        readdata = BinStream.ReadBytes(LoopStream.Size)
'
'
'
'    Next
'AltStreamError:
'    hasaltstreams = False
'
    
    
Next


OutStream.Flush
OutStream.CloseStream




End Sub
Private Function GetColumnProviderHeaders() As ColumnProviderData()
    'returns an array of strings. each one is a columnheader.
    'note that the order corresponds to the output from GetColumnProvidersData()...
    Dim currprovider As Long
    Dim Ret As Long, mprov As IColumnProvider
    Dim ColumnInfo As SHCOLUMNINFO
    Dim retdata() As ColumnProviderData
    ReDim retdata(1 To UBound(mColumnProviders))
    For currprovider = 1 To UBound(mColumnProviders)
        Set mprov = mColumnProviders(currprovider)
        mprov.GetColumnInfo 1, ColumnInfo
        'columninfo.
        With retdata(currprovider)
            .ColumnDescription = StrConv(ColumnInfo.wszDescription, vbFromUnicode)
            .ColumnTitle = ColumnInfo.wszTitle
            .Defwidth = ScaleX(ColumnInfo.cChars, vbCharacters, vbPixels)
            .flags = ColumnInfo.csFlags
            .lvwformat = ColumnInfo.fmt
        
        End With
    Next

    GetColumnProviderHeaders = retdata
End Function

Private Sub InitColumnProviders(FolderInit As String)
    'retrieve the ColumnProviders and initialize them.
    Dim pcount As Long, currprovider As Long
    Dim shinit As SHCOLUMNINIT, Ret As Long, addrefpunk As olelib.IUnknown
    Dim sFolder(0 To 519) As Byte, strfold As String
    
    strfold = StrConv(FolderInit, vbUnicode)
    CopyMemory sFolder(0), strfold, LenB(strfold)
    Erase mColumnProviders
    mColumnProviders = GetColumnProviders(pcount)
    'SHinit.wszFolder = sfolder(0)
    CopyMemory shinit.wszFolder(0), sFolder(0), LenB(strfold)
    shinit.dwFlags = 0
    shinit.dwReserved = 0
    For currprovider = 1 To pcount
    'initialize each columnprovider...
   
        Call mColumnProviders(currprovider).Initialize(shinit)
    
    
    Next



End Sub
Public Sub TestColumns()
    Dim ColumnData() As ColumnProviderData
    InitColumnProviders "C:\"
    ColumnData = GetColumnProviderHeaders()
    Stop


End Sub
Private Sub GroupBySize(ParamArray GroupSizeSpecs())

'groups all the listitems in lvwfiles by their size.

'groups:

'GroupSizeSpecs: passed in should be something like:

'GroupBySize("Small Files",50*1024*1024,"Large Files",50*1024*1024*4,"Huge Files")


'That is- it goes from the first item set to the last- the last one is used for all items larger then the second to last value.


Dim GroupKeys() As String, Sizes() As Double, Titles() As String
Dim groups() As vbaBClListViewLib6.cItemGroup
Dim I As Long

Dim CurrCount As Long
Dim currItem As cListItem
Dim currgrp As Long
lvwfiles.ItemGroups.Clear

'step one: populate our array using the paramarray.
For I = 0 To UBound(GroupSizeSpecs) - 1 Step 2
    ReDim Preserve GroupKeys(CurrCount)
    ReDim Preserve Sizes(CurrCount)
    ReDim Preserve Titles(CurrCount)
    ReDim Preserve groups(CurrCount)
    Titles(CurrCount) = GroupSizeSpecs(I)
    If I = UBound(GroupSizeSpecs) Then
        Sizes(CurrCount) = -1
    End If
    GroupKeys(CurrCount) = "GROUPNUM" & Trim$(Str$(I))
    
    Set groups(CurrCount) = lvwfiles.ItemGroups.Add(, GroupKeys(CurrCount), Titles(CurrCount))
    Sizes(CurrCount) = GroupSizeSpecs(I + 1)
    CurrCount = CurrCount + 1
Next I

'we've added the groups. Iterate through each item in the list...
For I = 1 To lvwfiles.ListItems.Count
    Set currItem = lvwfiles.ListItems(I)
    For currgrp = 0 To CurrCount - 1
        If currItem.ItemData > Sizes(currgrp) Or currgrp = CurrCount - 1 Then
            Set currItem.Group = groups(currgrp)
            'Stop
            Exit For
        
        End If
    
    Next currgrp
    
lvwfiles.ItemGroups.enabled = True


Next I




'SetIcon Me.hWnd, "AAA", True


End Sub
Private Function CreateSendToItemForFile(FileAdd As CFile, bar As cCommandBar, Optional ByVal StrCaption As String = "") As cButton
        'Dim resolver As CshellLink
                Dim filepathfull As String, iconkey As String
                Dim newbutton As cButton
                
505                iconkey = FileAdd.Filename & FileAdd.FileIndex
                
                CDebug.Post iconkey
506                mmenuImages.AddFromHandle bcfile.GetObjIcon(FileAdd.FullPath, ICON_SMALL), IMAGE_ICON, iconkey
        
                On Error Resume Next
                    Set newbutton = cmdbarmenu(0).Buttons.Item("SENDTO:" & FileAdd.Filename)
                    If Err <> 0 Then
                        Err.Clear
                        
507                    Set newbutton = cmdbarmenu(0).Buttons.Add("SENDTO:" & FileAdd.Filename, mmenuImages.ItemIndex(iconkey), FileAdd.DisplayName, , "Send files here.")
                    
                    
                    
                    End If
                    newbutton.Tag = FileAdd.FullPath
                   If StrComp(FileAdd.Extension, "EXE", vbTextCompare) = 0 Then
                        newbutton.caption = GetEXEFriendlyName(FileAdd.FullPath)
508                ElseIf StrComp(FileAdd.Extension, "LNK", vbTextCompare) = 0 Then
                    'resolve shortcut...
509                    filepathfull = bcfile.ResolveShortcut(FileAdd.FullPath)
510                    newbutton.Tag = filepathfull


511                Else
'512                    newbutton.Tag = FileAdd.Fullpath
            
513                End If
        
514                bar.Buttons.Add newbutton
                    Set CreateSendToItemForFile = newbutton
        
        
        
        
'115            Next



End Function
Public Function CreateSendToItemFromPath(Tobar, PathCreate)
    Dim castbar As cCommandBar
    Set castbar = Tobar
    Set CreateSendToItemFromPath = CreateSendToItemForFile(GetFile(PathCreate), castbar)

End Function
Public Function GetApplicationFolder()
    GetApplicationFolder = App.Path
End Function

Private Sub tmrFilterResults_Timer()
'as the text in the textbox changes, filter the results we have in the listview.
'step one: check if we are currently filtered...
Dim LoopListItem As cListItem
Dim Progress As IProgress, SearchString As String
'first, disable the timer...
tmrFilterResults.enabled = False
'second, stop window redraw...
'SendMessage lvwfiles.hWnd, CWM_SETREDRAW, 0, ByVal 0&
SearchString = mtxtSearchBar.Text
Set Progress = Me

If Not mCachedRemovedItems Is Nothing Then
  'it would appear that we already may have cached entries.
    'add them all back to the listview.
    Dim currItem As Variant, readded As cListItem, gotfref As CFile
    For Each currItem In mCachedRemovedItems
        Set gotfref = GetFile(currItem, False)
        If Not gotfref Is Nothing Then
            Set readded = lvwfiles.ListItems.Add
            readded.Tag = currItem
            RefreshItemData readded, gotfref
        End If
    Next

End If
    'loop through all our listitems and search for the text in this textbox...
    'if it doesn't match, remove it from the listview but also plop the tag into the cachedremoveditems collection.
    If SearchString <> "" Then
        Dim doremove As Boolean, loopingindex As Long, totaldone As Long, originalcount As Long
        originalcount = lvwfiles.ListItems.Count
        Set mCachedRemovedItems = New Collection
        Screen.MousePointer = vbHourglass
        For loopingindex = 1 To lvwfiles.ListItems.Count - 1
            On Error Resume Next
            Set LoopListItem = lvwfiles.ListItems.Item(loopingindex)
            If Err <> 0 Then Exit For
        'For Each LoopListItem In lvwfiles.ListItems
            'first, check the text...
            'TODO:look in each column....
            
            If InStr(1, LoopListItem.Text, SearchString, vbTextCompare) > 0 Then
                'leave it be...
                doremove = False
            Else
                'otherwise, remove it.
                doremove = True
            End If
            
            If doremove Then
                mCachedRemovedItems.Add LoopListItem.Tag
                lvwfiles.ListItems.Remove LoopListItem.Index
                loopingindex = loopingindex - 1
                
            End If
            
            totaldone = totaldone + 1
            Progress.UpdateUI CDbl(totaldone) / CDbl(originalcount), "Filtering..."
            
            
        Next loopingindex
        Screen.MousePointer = vbDefault
    End If
'reenable redraw...
'SendMessage lvwfiles.hWnd, CWM_SETREDRAW, 1, ByVal 0&
'and, refresh.
lvwfiles.Refresh
'frmsearch.

Dim ctlloop As Object
On Error Resume Next
For Each ctlloop In Me.Controls
    
    ctlloop.Refresh
Next
DoEvents
End Sub

Private Sub TmrmenuHighlight_Timer()
'Debug.Print "tmrmenuhighlight"

If mFireResizeNextTmr Then
    Me.Move Me.Left, Me.Top, Me.Width, Me.Height
    mFireResizeNextTmr = False
End If

If mInmenu Then 'if we are in a menu, perform the check for changing the status text back from "simple" mode.
    TmrmenuHighlight.Tag = Val(TmrmenuHighlight.Tag) + 1
    
    If Val(TmrmenuHighlight.Tag) > 60 Then
        mSbar.SimpleMode = False
        TmrmenuHighlight.enabled = False
        'mInmenu = False
        TmrmenuHighlight.Tag = 0
    End If
End If
'If mSearching Then mSearchAnimator.NextFrame

End Sub
Private Sub ParseCmdLine(ByVal StrCommand As String)
    Debug.Print "command line sent:" & StrCommand
End Sub
