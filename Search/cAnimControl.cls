VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnimControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

' Common Controls:
Private Type InitCommonControlsEx
   animSize As Long
   animICC As Long
End Type
Private Const ICC_ANIMATE_CLASS = &H80
Private Const ANIMATE_CLASSA = "SysAnimate32"
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As InitCommonControlsEx) As Boolean
'// begin_r_commctrl
Private Const ACS_CENTER = &H1
Private Const ACS_TRANSPARENT = &H2
Private Const ACS_AUTOPLAY = &H4
Private Const ACS_TIMER = &H8        '// don't use threads... use timers
'// end_r_commctrl
'// Standard messages
Private Const WM_USER = &H400
Private Const ACM_OPEN = (WM_USER + 100)
Private Const ACM_PLAY = (WM_USER + 101)
Private Const ACM_STOP = (WM_USER + 102)
'// Notification messages... if you want them
Private Const ACN_START = 1
Private Const ACN_STOP = 2

' Windows
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Window Styles
Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_TABSTOP = &H10000
Private Const WS_VISIBLE = &H10000000
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
' Show window
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

' Libraries
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' General
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)

Private m_bAutoPlay As Boolean
Private m_sFileName As String
Private m_sResFileName As String
Private m_lResID As Long
Private m_bLoaded As Boolean
Private m_hMod As Long
Private m_hWnd As Long
Private m_bComCtlLoaded As Boolean
Private m_oOwner As Object
Private m_bAnimTransparent As Boolean
Private m_oBackColor As OLE_COLOR
Private m_bCentre As Boolean
Private m_tAnimR As RECT

Public Sub StartPlay(Optional ByVal From As Integer, Optional ByVal To_ As Integer = -1, Optional ByVal Repeat As Long = -1)
Dim lFromTo As Long
   If (m_bLoaded) And Not (m_hWnd = 0) Then
      SendMessageLong m_hWnd, ACM_PLAY, Repeat, MakeLong(From, To_)
   End If
End Sub

Public Sub StopPlay()
   If (m_bLoaded) And Not (m_hWnd = 0) Then
      SendMessageLong m_hWnd, ACM_STOP, 0, 0
   End If
End Sub

Public Property Get Filename() As String
   Filename = m_sFileName
End Property

Public Property Let Filename(ByVal sFile As String)
   pUnloadAnimation
   m_lResID = 0
   m_sResFileName = ""
   m_sFileName = sFile
   pLoadAnimation
End Property

Public Property Get ResourceId() As Long
   ResourceId = m_lResID
End Property
Public Property Let ResourceId(ByVal lId As Long)
   pUnloadAnimation
   m_sFileName = ""
   m_lResID = lId
   pLoadAnimation
End Property
Public Property Get ResourceFileName() As String
   ResourceFileName = m_sResFileName
End Property
Public Property Let ResourceFileName(ByVal sFileName As String)
   pUnloadAnimation
   m_sFileName = ""
   m_sResFileName = sFileName
   pLoadAnimation
End Property

Private Sub pLoadAnimation()
   If Not (m_hWnd = 0) Then
      Dim hMod As Long
      If Not (m_lResID = 0) Then
         If Len(m_sResFileName) = 0 Then
            hMod = App.hInstance
         Else
            hMod = LoadLibraryEx(m_sResFileName, 0, 0)
            If Not (hMod = 0) Then
               m_hMod = hMod
            Else
               Err.Raise 45003, App.EXEName & ".cAnimControl", "Failed to load the resource library '" & m_sResFileName & "'"
               Exit Sub
            End If
         End If
         If SendMessageLong(m_hWnd, ACM_OPEN, hMod, m_lResID) = 0 Then
            Err.Raise 45002, App.EXEName & ".cAnimControl", "Failed to open the AVI from the resource " & m_lResID
         Else
            ShowWindow m_hWnd, SW_SHOW
            GetWindowRect m_hWnd, m_tAnimR
            pSizeWindow
            m_bLoaded = True
         End If
      ElseIf Len(m_sFileName) > 0 Then
         If SendMessage(m_hWnd, ACM_OPEN, 0, ByVal m_sFileName) = 0 Then
            Err.Raise 45002, App.EXEName & ".cAnimControl", "Failed to open the file '" & m_sFileName & "'"
         Else
            ShowWindow m_hWnd, SW_SHOW
            GetWindowRect m_hWnd, m_tAnimR
            pSizeWindow
            m_bLoaded = True
         End If
      End If
   End If
End Sub
Private Sub pUnloadAnimation()
   If Not (m_hWnd = 0) Then
      StopPlay
      SendMessage m_hWnd, ACM_OPEN, 0, ByVal 0&
      ShowWindow m_hWnd, SW_HIDE
   End If
   If Not (m_hMod = 0) Then
      FreeLibrary m_hMod
      m_hMod = 0
   End If
   m_bLoaded = False
End Sub

Public Property Get Owner() As Object
   Set Owner = m_oOwner
End Property
Public Property Let Owner(ByVal oFormOrControl As Object)
   pSetOwner oFormOrControl
End Property
Public Property Set Owner(ByVal oFormOrControl As Object)
   pSetOwner oFormOrControl
End Property
Public Property Get Transparent() As Boolean
   Transparent = m_bAnimTransparent
End Property
Public Property Let Transparent(ByVal bState As Boolean)
   If Not (m_bAnimTransparent = bState) Then
      m_bAnimTransparent = bState
      If Not (m_hWnd = 0) Then
         pbCreate
      End If
   End If
End Property
Public Property Get Centre() As Boolean
   Centre = m_bCentre
End Property
Public Property Let Centre(ByVal bState As Boolean)
   If Not (m_bCentre = bState) Then
      m_bCentre = bState
      If Not (m_hWnd = 0) Then
         pbCreate
      End If
   End If
End Property
Public Property Get AutoPlay() As Boolean
   AutoPlay = m_bAutoPlay
End Property
Public Property Let AutoPlay(ByVal bState As Boolean)
   If Not (m_bAutoPlay = bState) Then
      m_bAutoPlay = bState
      If Not (m_hWnd = 0) Then
         pbCreate
      End If
   End If
End Property
Public Property Get Width() As Long
Dim tR As RECT
   GetWindowRect m_hWnd, tR
   Width = (m_tAnimR.right - m_tAnimR.left)
End Property
Public Property Get Height() As Long
   Height = (m_tAnimR.bottom - m_tAnimR.top)
End Property

Private Sub pSetOwner(oFormOrControl As Object)
   If Not (oFormOrControl.hwnd = 0) Then
      Set m_oOwner = oFormOrControl
      If Not pbCreate() Then
         Err.Raise 45001, App.EXEName & ".cAnimControl", "Could not create Animation Control."
      Else
         pLoadAnimation
      End If
   End If
End Sub

Public Property Get hwnd() As Long
   hwnd = m_hWnd
End Property

Private Function MakeLong(ByVal HIWORD As Integer, ByVal LOWORD As Integer) As Long
   CopyMemory MakeLong, HIWORD, 2
   CopyMemory ByVal VarPtr(MakeLong) + 2, LOWORD, 2
End Function

Private Function pIIf(ByVal bCondition As Boolean, ByVal lTrue As Long, Optional ByVal lFalse As Long = 0) As Long
   If bCondition Then
      pIIf = lTrue
   Else
      pIIf = lFalse
   End If
End Function

Private Function pbCreate() As Boolean
Dim lStyle As Long
   
   ' Destroy the previous Anim control if any
   If Not (m_hWnd = 0) Then
      pDestroy
   End If
   
   If pbLoadComCtl() Then
   
      ' Set styles
      lStyle = WS_CHILD Or WS_VISIBLE
      lStyle = lStyle Or pIIf(m_bAutoPlay, ACS_AUTOPLAY, &H0)
      lStyle = lStyle Or pIIf(m_bAnimTransparent, ACS_TRANSPARENT, &H0)
      lStyle = lStyle Or pIIf(m_bCentre, ACS_CENTER, &H0)
      
      ' Create the Animation Control
      m_hWnd = CreateWindowEx( _
                  0, ANIMATE_CLASSA, vbNullString, lStyle, _
                  0, 0, 0, 0, _
                  m_oOwner.hwnd, 0&, _
                  App.hInstance, ByVal 0&)
      If Not (m_hWnd = 0) Then
         SetParent m_hWnd, m_oOwner.hwnd
         pSizeWindow
         pbCreate = True
      End If
      
   End If
   
End Function
Private Sub pDestroy()
   
   ' Destroys the animation control
   
   If Not (m_hWnd = 0) Then
      pUnloadAnimation
      ShowWindow m_hWnd, SW_HIDE
      SetParent m_hWnd, 0
      DestroyWindow m_hWnd
      m_hWnd = 0
   End If
   
End Sub

Private Sub pSizeWindow()
Dim tR As RECT

   GetClientRect m_oOwner.hwnd, tR
   MoveWindow m_hWnd, tR.left, tR.top, tR.right - tR.left, tR.bottom - tR.top, 1

End Sub


Private Function pbLoadComCtl() As Boolean
   If Not m_bComCtlLoaded Then
      ' Load common control 32 bit library
      Dim hMod As Long
      hMod = LoadLibrary("comctl32.dll")
      ' If the handle is valid, try to get the function address.
      If Not (hMod = 0) Then
         Dim lPtr As Long
         lPtr = GetProcAddress(hMod, "InitCommonControlsEx")
         FreeLibrary hMod
      End If
      
      If Not (lPtr = 0) Then
         Dim tICCEx As InitCommonControlsEx
         tICCEx.animICC = LenB(tICCEx)
         tICCEx.animICC = ICC_ANIMATE_CLASS
         InitCommonControlsEx tICCEx
         
         m_bComCtlLoaded = True
      End If
   End If
   pbLoadComCtl = m_bComCtlLoaded
End Function

Private Sub Class_Terminate()
   pDestroy
End Sub
