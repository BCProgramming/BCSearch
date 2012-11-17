VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Multiline Tooltips
'
' Tooltip class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 01/20/1999
' Updated: 04/27/2000 * Fixed a bug in the UpdateTool method that
'                       updates the tooltip text only for controls
'                       where Container = Parent.
'
'*********************************************************************************************
Option Explicit

' Tooltip icons
Enum TooltipIcon
   ttiNone = 0
   ttiInfo = 1
   ttiWarning = 2
   ttiError = 3
End Enum

' Tooltip animations
Enum TooltipShowMode
   ttsmDefault
   ttsmLeftToRight = &H1
   ttsmRightToLeft = &H2
   ttsmTopToBottom = &H4
   ttsmBottomToTop = &H8
   ttsmCenter = &H10
End Enum

Enum TooltipAlignment
   ttaLeft = 0
   ttaCentered = 1
   ttaRight = 2
   
End Enum

' Private struct where
' properties are stored

Private Type TTTool
   ID As Long
   ObjPtr As Long
   BackColor As Long
   ForeColor As Long
   Baloon As Boolean
   Title As String * 80
   Icon As TooltipIcon
   ShowMode As TooltipShowMode
   Margin As Long
   Alignment As TooltipAlignment
   Width As Long
End Type

Dim m_ToolIndex As Long       ' Last tool index
Dim m_Tools() As TTTool       ' Tools array

Event NeedText(ByVal Tool As Object, TOOLTIPTEXT As String)
Event Pop()
Event Show()

'*********************************************************************************************
'
' AddTool
' -------
'
' Registers a control with the tooltip.
'
' Parameters:
' -----------
'
' ToolObj:       The control for which the tooltip is set.
' TooltipText:   Optional. The tooltip text. If no text is specified
'                the class will raise the NeedText event.
' ForeColor:     Optional. The text color. Default is vbInfoText
' BackColor:     Optional. The tooltip background color. Default
'                is vbInfoBackground.
' Baloon:        Optional. Specifies whether the tooltip have baloon style.
' Title:         Optional. A title that is displayed in bold font.
'                The title is shown only if the tooltip has the
'                baloon style.
' Icon:          Optional. An icon that is displayed next to the title.
' ShowMode:      Optional. Specifies the animation used to show the
'                tooltip. Only applies to Win98/2000
' Margin:        Optional. Margin from text to the tooltip border.
' Centered:      Optional. Specifies whether the tooltip is shown
'                centered bellow the control.
' Transparent:   Optional. Specifies whether the tooltip will ignore
'                mouse events.
'
'*********************************************************************************************
Public Function AddToolObject( _
               ByVal ToolObj As Object, _
               Optional ByVal TOOLTIPTEXT, _
               Optional ByVal ForeColor, _
               Optional ByVal BackColor, _
               Optional ByVal Baloon As Boolean, _
               Optional ByVal Title As String, _
               Optional ByVal Icon As TooltipIcon = ttiNone, _
               Optional ByVal ShowMode As TooltipShowMode = ttsmDefault, _
               Optional ByVal Margin As Long, _
               Optional ByVal CenteredTooltip As Boolean, _
               Optional ByVal TRANSPARENT As Boolean, _
               Optional ByVal Alignment As TooltipAlignment, _
               Optional ByVal Width As Long) As Boolean
               
               End Function
               
Public Function AddTool( _
               ByVal ToolObj As Object, _
               Optional ByVal TOOLTIPTEXT, _
               Optional ByVal ForeColor, _
               Optional ByVal BackColor, _
               Optional ByVal Baloon As Boolean, _
               Optional ByVal Title As String, _
               Optional ByVal Icon As TooltipIcon = ttiNone, _
               Optional ByVal ShowMode As TooltipShowMode = ttsmDefault, _
               Optional ByVal Margin As Long, _
               Optional ByVal CenteredTooltip As Boolean, _
               Optional ByVal TRANSPARENT As Boolean, _
               Optional ByVal Alignment As TooltipAlignment, _
               Optional ByVal Width As Long) As Boolean
               
Dim TTTI As TOOLINFO, Text As String
   
   On Error Resume Next
   
   ' Fill the TOOLINFO struct
   With TTTI
      .cbSize = Len(TTTI)
      
      .hWnd = ToolObj.Container.hWnd
      
      .uFlags = TTF_SUBCLASS Or _
               (-TTF_CENTERTIP * CenteredTooltip) Or _
               (-TTF_TRANSPARENT * TRANSPARENT)
            
      ' Try using the object
      ' hWnd property
      .uId = ToolObj.hWnd
      
      If Err.Number = 0 Then
         .uFlags = .uFlags Or TTF_IDISHWND
      Else
         
         ' The object does not
         ' have an hWnd property,
         ' so use the RECTangle.
         
         .uId = m_ToolIndex
         
         With .RECT
            .Left = ToolObj.Parent.ScaleX(ToolObj.Left, ToolObj.Parent.ScaleMode, vbPixels)
            .Top = ToolObj.Parent.ScaleY(ToolObj.Top, ToolObj.Parent.ScaleMode, vbPixels)
            .Right = .Left + ToolObj.Parent.ScaleX(ToolObj.Width, ToolObj.Parent.ScaleMode, vbPixels)
            .Bottom = .Top + ToolObj.Parent.ScaleY(ToolObj.Height, ToolObj.Parent.ScaleMode, vbPixels)
         End With
         
      End If
      
      ' Add the control to the tools array
      m_ToolIndex = m_ToolIndex + 1
      ReDim Preserve m_Tools(1 To m_ToolIndex)
      
      ' Fill the tool properties
      With m_Tools(m_ToolIndex)
         .ID = TTTI.uId
         .ObjPtr = ObjPtr(ToolObj)
         
         If IsMissing(BackColor) Then
            .BackColor = TranslateColor(vbInfoBackground)
         Else
            .BackColor = TranslateColor(BackColor)
         End If
         
         If IsMissing(ForeColor) Then
            .ForeColor = TranslateColor(vbInfoText)
         Else
            .ForeColor = TranslateColor(ForeColor)
         End If
      
         .Baloon = Baloon
         .Title = Title
         .Icon = Icon
         .ShowMode = ShowMode
         .Margin = Margin
         .Alignment = Alignment
         .Width = Width
      End With
            
      If IsMissing(TOOLTIPTEXT) Then
         .lpszText = LPSTR_TEXTCALLBACK
      Else
         Text = CStr(TOOLTIPTEXT)
         .lpszText = StrPtr(Text)
      End If
      
   End With
   
   ' Add the control to the tooltip
   AddTool = SendMessage(m_TTWnd, TTM_ADDTOOLW, 0, TTTI)
      
End Function

'*********************************************************************************************
'
' HookCallback
' ------------
'
' Processes tooltip notifications.
'
' Parameters:
' -----------
'
' lParam:   Pointer to a NMHDR struct.
'
'*********************************************************************************************
Friend Sub HookCallback(ByVal lParam As Long)
Dim NM As NMHDR, Idx As Long

   On Error Resume Next
   
   ' Copy the NMHDR from pointer
   CopyMemory NM, ByVal lParam, Len(NM)
   
   Select Case NM.code
   
      Case NM_CUSTOMDRAW
   
         Dim r As mdlTooltip.RECT, NMCD As mdlTooltip.NMTTCUSTOMDRAW
         
         CopyMemory NMCD, ByVal lParam, Len(NMCD)

         ' Find the control in the array
         For Idx = 1 To UBound(m_Tools)
            If m_Tools(Idx).ID = NM.idFrom Then
               Exit For
            End If
         Next
         
         ' Update properties
         With m_Tools(Idx)
            
            ' Change colors
            SendMessage m_TTWnd, TTM_SETTIPBKCOLOR, .BackColor, 0&
            SendMessage m_TTWnd, TTM_SETTIPTEXTCOLOR, .ForeColor, 0&
            
            ' Change baloon style
            pvMakeBaloon .Baloon
            
            ' Set/Remove title and icon
            If Trim$(.Title) <> "" Then
               SendMessage m_TTWnd, TTM_SETTITLE, .Icon, ByVal Trim$(.Title)
            Else
               SendMessage m_TTWnd, TTM_SETTITLE, 0&, ByVal 0&
            End If

            ' Change margins
            r.Left = .Margin
            r.Right = .Margin
            r.Bottom = .Margin
            r.Top = .Margin
   
            If .Width > 0 Then
               If (NMCD.uDrawFlags And DT_CALCRECT) Then
                  NMCD.NMCD.rc.Right = NMCD.NMCD.rc.Left + .Width
                  NMCD.uDrawFlags = (NMCD.uDrawFlags Or DT_NOCLIP Or DT_WORDBREAK) And Not DT_SINGLELINE
               End If
            End If

            SendMessage m_TTWnd, TTM_SETMARGIN, 0, r
                        
            NMCD.uDrawFlags = NMCD.uDrawFlags Or .Alignment
            
            CopyMemory ByVal lParam, NMCD, Len(NMCD)
            
         End With
         
         
      Case TTN_POP
      
         ' The tooltip was hidden.
         RaiseEvent Pop
      
      Case TTN_SHOW
         
         ' The tooltip will be show
         
         ' Find the control in the array
         For Idx = 1 To UBound(m_Tools)
            If m_Tools(Idx).ID = NM.idFrom Then
               Exit For
            End If
         Next
         
         ' Use AnimateWindow API to
         ' animate the tooltip.
         ' AnimateWindow API is available
         ' only in win98/2000. In other OS
         ' the call will raise the 453
         ' error (Can't find DLL entry point),
         ' that will be ignored because we
         ' are using On Error.
         If m_Tools(Idx).ShowMode <> 0 And m_Tools(Idx).Baloon = False Then
            AnimateWindow m_TTWnd, 200, m_Tools(Idx).ShowMode
         End If
         
         RaiseEvent Show
      
      Case TTN_GETDISPINFO
      
         ' The tooltip needs the
         ' text.
         
         Dim TTGDI As TOOLTIPTEXT
         Dim ToolObject As Object, TmpObj As Object
         Static TTText As String
         
         ' Get the TOOLTIPTEXT struct
         CopyMemory TTGDI, ByVal lParam, Len(TTGDI)
                  
         ' Find the control in the array,
         ' and create a couted reference to it.
         For Idx = 1 To UBound(m_Tools)
            If m_Tools(Idx).ID = NM.idFrom Then
               CopyMemory TmpObj, m_Tools(Idx).ObjPtr, 4
               Set ToolObject = TmpObj
               CopyMemory TmpObj, 0&, 4
               Exit For
            End If
         Next
         
         ' Clear the TTText variable
         TTText = vbNullString
         
         ' Raise NeedText event
         RaiseEvent NeedText(ToolObject, TTText)
         
         ' Convert the string to ANSI
         TTText = StrConv(TTText, vbFromUnicode)
         
         ' Update the TOOLTIPTEXT struct
         ' with the text.
         With TTGDI
            .lpszText = StrPtr(TTText)
         End With
         
         ' Copy back the modified TOOLTIPTEXT
         CopyMemory ByVal lParam, TTGDI, Len(TTGDI)
         
   End Select
   
End Sub

'*********************************************************************************************
'
' SetDelayTimes
' -------------
'
' Sets the delay times used to show, hide and reshow the tooltip
'
'*********************************************************************************************
Public Sub SetDelayTimes(Optional ByVal Show, Optional ByVal AutoHide, Optional ByVal ReShow)

   If IsMissing(Show) And IsMissing(AutoHide) And IsMissing(ReShow) Then
      SendMessage m_TTWnd, TTM_SETDELAYTIME, TTDT_AUTOMATIC, ByVal 0&
   Else
      If Not IsMissing(Show) Then SendMessage m_TTWnd, TTM_SETDELAYTIME, TTDT_INITIAL, ByVal CLng(Show)
      If Not IsMissing(AutoHide) Then SendMessage m_TTWnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, ByVal CLng(AutoHide)
      If Not IsMissing(ReShow) Then SendMessage m_TTWnd, TTM_SETDELAYTIME, TTDT_RESHOW, ByVal CLng(ReShow)
   End If
   
End Sub

'*********************************************************************************************
'
' UpdateTool
' ----------
'
' Updates the properties of a control tooltip.
'
'*********************************************************************************************
Public Sub UpdateTool( _
               ByVal ToolObj As Object, _
               Optional ByVal TOOLTIPTEXT, _
               Optional ByVal ForeColor, _
               Optional ByVal BackColor, _
               Optional ByVal Baloon, _
               Optional ByVal Title, _
               Optional ByVal Icon, _
               Optional ByVal ShowMode, _
               Optional ByVal Margin, _
               Optional ByVal CenteredTooltip, _
               Optional ByVal Alignment, _
               Optional ByVal Width)
               
Dim TTTI As TOOLINFO, TText As String, Idx As Long
   
   On Error Resume Next
   
   ' Fill the TOOLINFO struct
   With TTTI
   
      .cbSize = LenB(TTTI)
      .hWnd = ToolObj.Container.hWnd
      
      ' Find then ID in the
      ' tools array
      For Idx = 1 To UBound(m_Tools)
         If ObjPtr(ToolObj) = m_Tools(Idx).ObjPtr Then
            Exit For
         End If
      Next
           
      ' If Idx > UBound(m_Tools) then
      ' control was not found
      ' in the array, and will be
      ' ignored.
      If Idx <= UBound(m_Tools) Then
      
         ' Instead of using m_Tools(Idx).ID
         ' we try again with the hWnd to
         ' add the TTF_IDISHWND if the
         ' ID is a hWnd.
         .uId = ToolObj.hWnd
         
         If Err.Number = 0 Then
         
            .uFlags = .uFlags Or TTF_IDISHWND
         
            ' Get the tool info
            SendMessage m_TTWnd, TTM_GETTOOLINFOW, 0, TTTI
            
         Else
            
            .uId = m_Tools(Idx).ID
            
            ' Get the tool info
            SendMessage m_TTWnd, TTM_GETTOOLINFOW, 0, TTTI

            ' Get the control RECTangle
            With .RECT
               .Left = ToolObj.Parent.ScaleX(ToolObj.Left, ToolObj.Parent.ScaleMode, vbPixels)
               .Top = ToolObj.Parent.ScaleY(ToolObj.Top, ToolObj.Parent.ScaleMode, vbPixels)
               .Right = .Left + ToolObj.Parent.ScaleX(ToolObj.Width, ToolObj.Parent.ScaleMode, vbPixels)
               .Bottom = .Top + ToolObj.Parent.ScaleY(ToolObj.Height, ToolObj.Parent.ScaleMode, vbPixels)
            End With
            
         End If
         
         ' Update text
         If Not IsMissing(TOOLTIPTEXT) Then
            
            If TOOLTIPTEXT = "" Then
               .lpszText = -1

            Else
               TText = CStr(TOOLTIPTEXT)
               .lpszText = StrPtr(TText)
            End If
            
         End If
          
         If Not IsMissing(CenteredTooltip) Then
            If CenteredTooltip Then
               .uFlags = .uFlags Or TTF_CENTERTIP
            Else
               .uFlags = .uFlags And Not TTF_CENTERTIP
            End If
         End If
         
         ' Update tool properties
         With m_Tools(Idx)
            If Not IsMissing(BackColor) Then .BackColor = TranslateColor(BackColor)
            If Not IsMissing(ForeColor) Then .ForeColor = TranslateColor(ForeColor)
            If Not IsMissing(Baloon) Then .Baloon = CBool(Baloon)
            If Not IsMissing(Title) Then .Title = CStr(Title)
            If Not IsMissing(Icon) Then .Icon = CLng(Icon)
            If Not IsMissing(ShowMode) Then .ShowMode = CLng(ShowMode)
            If Not IsMissing(Margin) Then .Margin = CLng(Margin)
            If Not IsMissing(Alignment) Then .Alignment = CLng(Alignment)
            If Not IsMissing(Width) Then .Width = CLng(Width)
         End With
                
         ' Update tool
         SendMessage m_TTWnd, TTM_SETTOOLINFOW, 0, TTTI
         
         ' Update tooltip
         SendMessage m_TTWnd, TTM_UPDATE, 0, 0&
         
      End If
   
   End With
   
End Sub



'*********************************************************************************************
' Removes a control from the tooltip.
'*********************************************************************************************
Public Sub RemoveTool(ByVal ToolObj As Object)
Dim TTTI As TOOLINFO, Idx As Long

   On Error Resume Next
   
   ' Fill the TOOLINFO struct
   With TTTI
   
      .cbSize = Len(TTTI)
      
      .hWnd = ToolObj.Parent.hWnd
      
      ' Try using the hWnd
      .uId = ToolObj.hWnd
      
      If Err.Number <> 0 Then
         
         ' The hWnd property does not
         ' exists. Find the control ID
         ' in the tools array
         
         For Idx = 1 To UBound(m_Tools)
            If ObjPtr(ToolObj) = m_Tools(Idx).ObjPtr Then
               .uId = m_Tools(Idx).ID
               Exit For
            End If
         Next
         
      End If
      
   End With
   
   ' Remove the control
   SendMessage m_TTWnd, TTM_DELTOOL, 0&, TTTI
   
End Sub

'*********************************************************************************************
'
' pvMakeBaloon
' ------------
'
' Changes the baloon style of the tooltip
'
'*********************************************************************************************
Private Sub pvMakeBaloon(ByVal Baloon As Boolean)
Dim lStyle As Long

   lStyle = GetWindowLong(m_TTWnd, GWL_STYLE)
   
   If Baloon Then
      lStyle = (lStyle And Not WS_BORDER) Or TTS_BALLOON
   Else
      lStyle = (lStyle Or WS_BORDER) And Not TTS_BALLOON
   End If
   
   SetWindowLong m_TTWnd, GWL_STYLE, lStyle
   
End Sub



Private Sub Class_Initialize()

   If m_TTWnd = 0 Then
   
      ' Initialize common controls
      InitCommonControls
   
      ' Create the tooltip window
      m_TTWnd = CreateWindowEx( _
         0, _
         TOOLTIPS_CLASS, _
         vbNullString, _
         WS_POPUP, _
         0, 0, _
         0, 0, _
         0&, _
         0&, _
         App.hInstance, _
         ByVal 0&)

      SetWindowLong m_TTWnd, GWL_USERDATA, ObjPtr(Me)
      
      SetWindowPos m_TTWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
            
      SendMessage m_TTWnd, TTM_SETMAXTIPWIDTH, 0&, ByVal &H7FFF&
      
      m_hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf CallWndHookProc, 0&, App.ThreadID)
      
   End If
   
   m_ObjectCount = m_ObjectCount + 1
   
End Sub

Private Sub Class_Terminate()
   
   m_ObjectCount = m_ObjectCount - 1
   
   If m_ObjectCount < 1 Then
      
      m_ObjectCount = 0

      ' Destroy the tooltip
      DestroyWindow m_TTWnd
   
      UnhookWindowsHookEx m_hHook
      
   End If
      
End Sub



