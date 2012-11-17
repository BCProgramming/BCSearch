Attribute VB_Name = "mdlTooltip"
'*********************************************************************************************
'
' Multiline Tooltips
'
' Declarations and Hook module
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 01/20/1999
'
'*********************************************************************************************
Option Explicit

' All intances of clsTooltip
' uses the same tooltip window
' and hook. Only the first one
' receives events.
Public m_TTWnd As Long
Public m_hHook As Long

' Object count. Stores the
' number of clsTooltip
' objects created so
' the last one destroy
' the hook and the tooltip
' window.
Public m_ObjectCount As Long

Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    RECT As RECT
    hinst As Long
    lpszText As Long
End Type

Type NMHDR
    hWndFrom As Long
    idFrom As Long
    code As Long         ' NM_ code
End Type
   
Type NMCUSTOMDRAW
   hdr As NMHDR
   dwDrawStage As Long
   hdc As Long
   rc As RECT
   dwItemSpec As Long
   uItemState As Long
   lParam As Long
End Type

Type NMTTCUSTOMDRAW
   NMCD As NMCUSTOMDRAW
   uDrawFlags As Long
End Type

Type TOOLTIPTEXT
    hdr As NMHDR
    lpszText As Long
    szText As String * 80
    hinst As Long
    uFlags As Long
End Type

Declare Function SetWindowsHookEx Lib "User32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hhk As Long) As Boolean
Declare Function CallNextHookEx Lib "User32" (ByVal hhk As Long, ByVal nCode As Long, wParam, ByVal lParam As Long) As Long

Declare Sub AnimateWindow Lib "User32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long)

Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_NOTIFY = &H4E

Declare Sub InitCommonControls Lib "comctl32" ()

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Length As Long)

Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2
Public Const TTS_NOANIMATE = &H10
Public Const TTS_NOFADE = &H20
Public Const TTS_BALLOON = &H40

Public Const TTF_IDISHWND = &H1
Public Const TTF_CENTERTIP = &H2
Public Const TTF_RTLREADING = &H4
Public Const TTF_SUBCLASS = &H10
Public Const TTF_TRACK = &H20
Public Const TTF_ABSOLUTE = &H80
Public Const TTF_TRANSPARENT = &H100
Public Const TTF_DI_SETITEM = &H8000

Public Const TTDT_AUTOMATIC = 0
Public Const TTDT_RESHOW = 1
Public Const TTDT_AUTOPOP = 2
Public Const TTDT_INITIAL = 3

Public Const TTM_ACTIVATE = (&H400 + 1)
Public Const TTM_SETDELAYTIME = (&H400 + 3)
Public Const TTM_ADDTOOL = (&H400 + 4)
Public Const TTM_ADDTOOLW = (&H400 + 50)
Public Const TTM_DELTOOL = (&H400 + 5)
Public Const TTM_DELTOOLW = (&H400 + 51)
Public Const TTM_NEWTOOLRECT = (&H400 + 6)
Public Const TTM_NEWTOOLRECTW = (&H400 + 52)
Public Const TTM_RELAYEVENT = (&H400 + 7)
Public Const TTM_GETTOOLINFO = (&H400 + 8)
Public Const TTM_GETTOOLINFOW = (&H400 + 53)
Public Const TTM_SETTOOLINFO = (&H400 + 9)
Public Const TTM_SETTOOLINFOW = (&H400 + 54)
Public Const TTM_HITTEST = (&H400 + 10)
Public Const TTM_HITTESTW = (&H400 + 55)
Public Const TTM_GETTEXT = (&H400 + 11)
Public Const TTM_GETTEXTW = (&H400 + 56)
Public Const TTM_UPDATETIPTEXT = (&H400 + 12)
Public Const TTM_UPDATETIPTEXTW = (&H400 + 57)
Public Const TTM_GETTOOLCOUNT = (&H400 + 13)
Public Const TTM_ENUMTOOLS = (&H400 + 14)
Public Const TTM_ENUMTOOLSW = (&H400 + 58)
Public Const TTM_GETCURRENTTOOL = (&H400 + 15)
Public Const TTM_GETCURRENTTOOLW = (&H400 + 59)
Public Const TTM_WINDOWFROMPOINT = (&H400 + 16)
Public Const TTM_TRACKACTIVATE = (&H400 + 17)
Public Const TTM_TRACKPOSITION = (&H400 + 18)
Public Const TTM_SETTIPBKCOLOR = (&H400 + 19)
Public Const TTM_SETTIPTEXTCOLOR = (&H400 + 20)
Public Const TTM_GETDELAYTIME = (&H400 + 21)
Public Const TTM_GETTIPBKCOLOR = (&H400 + 22)
Public Const TTM_GETTIPTEXTCOLOR = (&H400 + 23)
Public Const TTM_SETMAXTIPWIDTH = (&H400 + 24)
Public Const TTM_GETMAXTIPWIDTH = (&H400 + 25)
Public Const TTM_SETMARGIN = (&H400 + 26)
Public Const TTM_GETMARGIN = (&H400 + 27)
Public Const TTM_POP = (&H400 + 28)
Public Const TTM_UPDATE = (&H400 + 29)
Public Const TTM_GETBUBBLESIZE = (&H400 + 30)
Public Const TTM_ADJUSTRECT = (&H400 + 31)
Public Const TTM_SETTITLE = (&H400 + 32)
Public Const TTM_SETTITLEW = (&H400 + 33)

Public Const NM_CUSTOMDRAW = -12

Public Const TTN_FIRST = -520
Public Const TTN_GETDISPINFO = (TTN_FIRST - 0)
Public Const TTN_GETDISPINFOW = (TTN_FIRST - 10)
Public Const TTN_SHOW = (TTN_FIRST - 1)
Public Const TTN_POP = (TTN_FIRST - 2)

Public Const LPSTR_TEXTCALLBACK = -1

Public Const TOOLTIPS_CLASS = "tooltips_class32"

Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_BORDER = &H800000
Public Const WS_POPUP = &H80000000

Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Sub OleTranslateColor Lib "olepro32" (ByVal OLECLR As Long, ByVal hPal As Long, ColorRef As Long)

Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "User32" (ByVal hwnd As Long) As Boolean

Public Const HWND_TOPMOST = (-1)

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20       ' The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200     ' Don't do owner Z ordering
Public Const SWP_NOSENDCHANGING = &H400    ' Don't send WM_WINDOWPOSCHANGING

Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Boolean

Public Const WH_CALLWNDPROC = 4

Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_NOCLIP = &H100
Public Const DT_CALCRECT = &H400

Declare Function DrawEdge Lib "User32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

'*********************************************************************************************
'
' Hook procedure.
'
'*********************************************************************************************
Public Function CallWndHookProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim CWPS As CWPSTRUCT, NM As NMHDR

   ' Get the CWPS struct from
   ' the pointer passed in lParam
   CopyMemory CWPS, ByVal lParam, Len(CWPS)
   
   ' Process only WM_NOTIFY messages
   If CWPS.message = WM_NOTIFY Then
      
      ' Get the NMHDR struct from
      ' the CWPS.lParam pointer
      CopyMemory NM, ByVal CWPS.lParam, Len(NM)
      
      ' Process only if the message
      ' was sent by the tooltip
      If NM.hWndFrom = m_TTWnd Then
         
         Dim clsTT As clsTooltip
         
         ' Get a reference to the
         ' object so we can raise events.
         
         CopyMemory clsTT, GetWindowLong(m_TTWnd, GWL_USERDATA), 4
         
         ' Call the Friend callback sub.
         clsTT.HookCallback CWPS.lParam
         
         ' Do not use Set clsTT = Nothing,
         ' since the object was obtained
         ' without incrementing the
         ' reference count.
         CopyMemory clsTT, 0&, 4
         
      End If
   
   End If

   CallWndHookProc = CallNextHookEx(m_hHook, code, wParam, lParam)
  
End Function


Public Function TranslateColor(ByVal OLEColor As Long) As Long

   OleTranslateColor OLEColor, 0, TranslateColor
   
End Function


