VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSuperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Name.......... cSuperClass
'File.......... cSuperClass.cls
'Dependencies.. Requires iSuperClass as the model implementation interface.
'Description... A novel window subclassing class that echews the use of a module by dynamically
'               assembling machine code.
'Author........ Paul_Caton@hotmail.com
'Date.......... June, 13th 2002
'Copyright..... None.
'
'v1.00 20020613 First cut......................................................................
'
'v1.01 20020621 Decided to split the single interface iSuperClass_Message into two,
'               iSuperClass_After and iSuperClass_Before. This is slightly more efficient
'               in that the more common *AFTER* the previous WndProc subclassing mode
'               was receiving a redundant parameter (lHandled) also, it reminds the
'               user in which of the two modes the message was added (AddMsg)..................
'
'               Optimized the assembler opcodes a bit.
'               Now using EIP relative calls.
'               WNDPROC_FILTERED is now 10 bytes shorter and slightly faster
'               WNDPROC_ALL is now 20 bytes shorter and slightly faster........................
'
'v1.02 20020627 Spotted that you could UnSubclass and still receive 1 more callback which
'               could stop an unload or worse. Scenario: you AddMsg WMES_NCLBUTTONDOWN and
'               click on the close button, the message goes to default processing first which
'               tells the form to unload wherein you call UnSubclass; at this point default
'               processing ends and execution returns to our WndProc who now wants to call
'               iSuperClass_After. The solution is to patch the WndProc code in UnSubclass
'               so that a return is patched between def processing and the call to
'               iSubClass_After................................................................
'
'v1.03 20020627 Added the AllMsgs mode of operation
'               I'm now reasonably confident that cSuperClass is immune to the IDE End button,
'               I think this is because the WndProc remains executable after the End button....
'
'v1.04 20020701 Added a couple of assembler optimizations to WndProc.asm
'               Zeroed lReturn before calling iSuperClass_Before
'               Fixed a few comments...........................................................
'
'v1.05 20020702 Cleaned up patching in SubClass
'               Cleaned up patching in Unsubclass
'               Re-inserted the commented out code to crash the app............................
'

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC       As Long = (-4)      'Get/Set the WndProc address with GetWindowLong/SetWindowLong
Private Const BUF_TOP           As Long = 511       'Max offset in opcode buffer. Requires 136 + (# Messages * 11)
Private Const OFFSET_BEFORE     As Long = 3         'Offset into the opcode bytes for the BEFORE default processing code
Private Const OFFSET_AFTER      As Long = 65        'Offset into the opcode bytes for the AFTER default processing code
Private Const CODE_RETURN       As Long = &H10C2C9  'Leave-return opcode sequence
Private Const OPCODE_CMP_EAX    As String = "3D"    'Opcode for cmp eax,????????
Private Const OPCODE_JE         As String = "0F84"  'Opcode for je with a 4 byte relative offset.
Private Const WNDPROC_ALL       As String = "558BEC83C4FCFF7514FF7510FF750CFF7508E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C21000"
Private Const WNDPROC_FILTERED  As String = "558BEC83C4F8EB6AC745FC000000008D45FC50C745F8000000008D45F850B8ptrOwner8BC88B0950FF5120837DF800753AFF7514FF7510FF750CFF7508E8wnd_procC9C21000E8wnd_proc8945FCFF7514FF7510FF750CFF75088D45FC50B8ptrOwner8BC88B0950FF511C8B45FCC9C210008B450CFF7514FF751050FF7508"
Private Const MSG_UNHANDLED     As String = "E8wnd_procC9C21000"
Private Const WMES_USER As Long = &H400
'Names have been changed to protect existing code.
Private Const WM_USER As Long = &H400
Public Enum WindowsMessages
    CWM_NULL = &H0
    CWM_CREATE = &H1
    CWM_DESTROY = &H2
    CWM_MOVE = &H3
    CWM_SIZE = &H5
    CWM_ACTIVATE = &H6
    CWM_SETFOCUS = &H7
    CWM_KILLFOCUS = &H8
    CWM_ENABLE = &HA
    CWM_SETREDRAW = &HB
    CWM_SETTEXT = &HC
    CWM_GETTEXT = &HD
    CWM_GETTEXTLENGTH = &HE
    CWM_PAINT = &HF
    CWM_CLOSE = &H10
    CWM_QUERYENDSESSION = &H11
    CWM_QUIT = &H12
    CWM_QUERYOPEN = &H13
    CWM_ERASEBKGND = &H14
    CWM_SYSColorCHANGE = &H15
    CWM_ENDSESSION = &H16
    CWM_SHOWWINDOW = &H18
    CWM_WININICHANGE = &H1A
    CWM_DEVMODECHANGE = &H1B
    CWM_ACTIVATEAPP = &H1C
    CWM_FONTCHANGE = &H1D
    CWM_TIMECHANGE = &H1E
    CWM_CANCELMODE = &H1F
    CWM_SETCURSOR = &H20
    CWM_MOUSEACTIVATE = &H21
    CWM_CHILDACTIVATE = &H22
    CWM_QUEUESYNC = &H23
    CWM_PAINTICON = &H26
    CWM_ICONERASEBKGND = &H27
    CWM_NEXTDLGCTL = &H28
    CWM_SPOOLERSTATUS = &H2A
    CWM_DRAWITEM = &H2B
    CWM_MEASUREITEM = &H2C
    CWM_DELETEITEM = &H2D
    CWM_VKEYTOITEM = &H2E
    CWM_CHARTOITEM = &H2F
    CWM_SETFONT = &H30
    CWM_GETFONT = &H31
    CWM_SETHOTKEY = &H32
    CWM_GETHOTKEY = &H33
    CWM_QUERYDRAGICON = &H37
    CWM_COMPAREITEM = &H39
    CWM_COMPACTING = &H41
    CWM_GETMINMAXINFO = &H24
    CWM_KEYFIRST = &H100
    CWM_KEYDOWN = &H100
    CWM_KEYUP = &H101
    CWM_CHAR = &H102
    CWM_DEADCHAR = &H103
    CWM_SYSKEYDOWN = &H104
    CWM_SYSKEYUP = &H105
    CWM_SYSCHAR = &H106
    CWM_SYSDEADCHAR = &H107
    CWM_KEYLAST = &H108
    CWM_INITDIALOG = &H110
    CWM_COMMAND = &H111
    CWM_SIZING = &H214
    CWM_ENTERSIZEMOVE = &H231
    CWM_EXITSIZEMOVE = &H232
    CWM_MOUSEHOVER = &H2A1&
    CWM_MOUSELEAVE = &H2A3&
    CWM_TIMER = &H113
    CWM_HSCROLL = &H114
    CWM_VSCROLL = &H115
    CWM_INITMENU = &H116
    CWM_INITMENUPOPUP = &H117
    CWM_MENUSELECT = &H11F
    CWM_MENUCHAR = &H120
    CWM_ENTERIDLE = &H121
    CWM_CTLColorMSGBOX = &H132
    CWM_CTLColorEDIT = &H133
    CWM_CTLColorLISTBOX = &H134
    CWM_CTLColorBTN = &H135
    CWM_CTLColorDLG = &H136
    CWM_CTLColorSCROLLBAR = &H137
    CWM_CTLColorSTATIC = &H138
    CWM_MOUSEFIRST = &H200
    





    


    
    CWM_MOUSELAST = &H209
    CWM_PARENTNOTIFY = &H210
    CWM_ENTERMENULOOP = &H211
    CWM_EXITMENULOOP = &H212
    CWM_MDICREATE = &H220
    CWM_MDIDESTROY = &H221
    CWM_MDIACTIVATE = &H222
    CWM_MDIRESTORE = &H223
    CWM_MDINEXT = &H224
    CWM_MDIMAXIMIZE = &H225
    CWM_MDITILE = &H226
    CWM_MDICASCADE = &H227
    CWM_MDIICONARRANGE = &H228
    CWM_MDIGETACTIVE = &H229
    CWM_MDISETMENU = &H230
    CWM_DROPFILES = &H233
    CWM_MDIREFRESHMENU = &H234
    CWM_CUT = &H300
    CWM_COPY = &H301
    CWM_PASTE = &H302
    CWM_CLEAR = &H303
    CWM_UNDO = &H304
    CWM_RENDERFORMAT = &H305
    CWM_RENDERALLFORMATS = &H306
    CWM_DESTROYCLIPBOARD = &H307
    CWM_DRAWCLIPBOARD = &H308
    CWM_PAINTCLIPBOARD = &H309
    CWM_VSCROLLCLIPBOARD = &H30A
    CWM_SIZECLIPBOARD = &H30B
    CWM_ASKCBFORMATNAME = &H30C
    CWM_CHANGECBCHAIN = &H30D
    CWM_HSCROLLCLIPBOARD = &H30E
    CWM_QUERYNEWPALETTE = &H30F
    CWM_PALETTEISCHANGING = &H310
    CWM_PALETTECHANGED = &H311
    CWM_HOTKEY = &H312
    
    CWM_NCMOUSEMOVE = &HA0
    CWM_NCLBUTTONDOWN = &HA1
    CWM_NCLBUTTONUP = &HA2
    CWM_NCLBUTTONDBLCLK = &HA3
    CWM_NCRBUTTONDOWN = &HA4
    CWM_NCMBUTTONDOWN = &HA7
    CWM_NCMBUTTONUP = &HA8
    CWM_NCMBUTTONDBLCLK = &HA9
    CWM_NCHITTEST = &H84
    CWM_SYSCOMMAND = &H112
    CWM_MOUSEMOVE = &H200
    CWM_LBUTTONDOWN = &H201
    CWM_LBUTTONUP = &H202
    CWM_LBUTTONDBLCLK = &H203
    CWM_RBUTTONDOWN = &H204
    CWM_RBUTTONUP = &H205
    CWM_RBUTTONDBLCLK = &H206
    CWM_MBUTTONDOWN = &H207
    CWM_MBUTTONUP = &H208
    CWM_MBUTTONDBLCLK = &H209
    CWM_MOUSEWHEEL = &H20A
    
    
          ' Edit Control Messages
    CEM_GETSEL = &HB0
    CEM_SETSEL = &HB1
    CEM_GETRECT = &HB2
    CEM_SETRECT = &HB3
    CEM_SETRECTNP = &HB4
    CEM_SCROLL = &HB5
    CEM_LINESCROLL = &HB6
    CEM_SCROLLCARET = &HB7
    CEM_GETMODIFY = &HB8
    CEM_SETMODIFY = &HB9
    CEM_GETLINECOUNT = &HBA
    CEM_LINEINDEX = &HBB
    CEM_SETHANDLE = &HBC
    CEM_GETHANDLE = &HBD
    CEM_GETTHUMB = &HBE
    CEM_LINELENGTH = &HC1
    CEM_REPLACESEL = &HC2
    CEM_GETLINE = &HC4
    CEM_LIMITTEXT = &HC5
    CEM_CANUNDO = &HC6
    CEM_UNDO = &HC7
    CEM_FMTLINES = &HC8
    CEM_LINEFROMCHAR = &HC9
    CEM_SETTABSTOPS = &HCB
    CEM_SETPASSWORDCHAR = &HCC
    CEM_EMPTYUNDOBUFFER = &HCD
    CEM_GETFIRSTVISIBLELINE = &HCE
    CEM_SETREADONLY = &HCF
    CEM_SETWORDBREAKPROC = &HD0
    CEM_GETWORDBREAKPROC = &HD1
    CEM_GETPASSWORDCHAR = &HD2
    
          ' Listbox messages
    CLB_ADDSTRING = &H180
    CLB_INSERTSTRING = &H181
    CLB_DELETESTRING = &H182
    CLB_SELITEMRANGEEX = &H183
    CLB_RESETCONTENT = &H184
    CLB_SETSEL = &H185
    CLB_SETCURSEL = &H186
    CLB_GETSEL = &H187
    CLB_GETCURSEL = &H188
    CLB_GETTEXT = &H189
    CLB_GETTEXTLEN = &H18A
    CLB_GETCOUNT = &H18B
    CLB_SELECTSTRING = &H18C
    CLB_DIR = &H18D
    CLB_GETTOPINDEX = &H18E
    CLB_FINDSTRING = &H18F
    CLB_GETSELCOUNT = &H190
    CLB_GETSELITEMS = &H191
    CLB_SETTABSTOPS = &H192
    CLB_GETHORIZONTALEXTENT = &H193
    CLB_SETHORIZONTALEXTENT = &H194
    CLB_SETCOLUMNWIDTH = &H195
    CLB_ADDFILE = &H196
    CLB_SETTOPINDEX = &H197
    CLB_GETITEMRECT = &H198
    CLB_GETITEMDATA = &H199
    CLB_SETITEMDATA = &H19A
    CLB_SELITEMRANGE = &H19B
    CLB_SETANCHORINDEX = &H19C
    CLB_GETANCHORINDEX = &H19D
    CLB_SETCARETINDEX = &H19E
    CLB_GETCARETINDEX = &H19F
    CLB_SETITEMHEIGHT = &H1A0
    CLB_GETITEMHEIGHT = &H1A1
    CLB_FINDSTRINGEXACT = &H1A2
    CLB_SETLOCALE = &H1A5
    CLB_GETLOCALE = &H1A6
    CLB_SETCOUNT = &H1A7
    CLB_MSGMAX = &H1A8
    CCB_FINDSTRING = &H14C
    CCB_SHOWDROPDOWN = &H14F
    CBM_GETCHECK = &HF0
    CBM_GETSTATE = &HF2
    CBM_SETCHECK = &HF1
    CBM_SETSTATE = &HF3
    CCDM_FIRST& = (WM_USER + 100)
    CCDN_FIRST = (-601)
    CCDM_GETFILEPATH& = (CCDM_FIRST + &H1)
    CCDM_GETFOLDERIDLIST& = (CCDM_FIRST + &H3)
    CCDM_GETFOLDERPATH& = (CCDM_FIRST + &H2)
    CCDM_GETSPEC& = (CCDM_FIRST + &H0)
    CCDM_HIDECONTROL& = (CCDM_FIRST + &H5)
    CCDM_LAST& = (WM_USER + 200)
    CCDM_SETCONTROLTEXT& = (CCDM_FIRST + &H4)
    CCDM_SETDEFEXT& = (CCDM_FIRST + &H6)
    CCDN_FILEOK& = (CCDN_FIRST - &H5)
    CCDN_FOLDERCHANGE& = (CCDN_FIRST - &H2)
    CCDN_HELP& = (CCDN_FIRST - &H4)
    CCDN_INCLUDEITEM& = (CCDN_FIRST - &H7)
    CCDN_INITDONE& = (CCDN_FIRST - &H0)
    CCDN_LAST& = (-699)
    CCDN_SELCHANGE& = (CCDN_FIRST - &H1)
    CCDN_SHAREVIOLATION& = (CCDN_FIRST - &H3)
    CCDN_TYPECHANGE& = (CCDN_FIRST - &H6)
    CD_FINDSTRINGEXACT = &H158
    LVW_SETCOLUMNWIDTH = &H1000 + 30
    LVW_USEHEADER = -2
End Enum
Private Type tCode
  Buf(0 To BUF_TOP) As Byte       'Opcode buffer
End Type

Private Type tCodeBuf
  code              As tCode      'WndProc opcodes
  nBuf              As Long       'Opcode buffer index
End Type

Private All         As Boolean    'All messages?
Private Running     As Boolean    'Is the subclasser running?
Private hWnd        As Long       'Window being subclassed
Private WndProcPrev As Long       'The address of the existing WndProc
Private pCode       As Long       'Pointer to the WndProc opcode buffer
Private CodeBuf     As tCodeBuf   'Opcode buffer

'Add a message to those that will call back either before or after the existing WndProc.
Public Sub AddMsg(MsgNum As WindowsMessages, Optional Before As Boolean = False)
  Debug.Assert (Running = False)                        'You don't add messages whilst the subclasser is running
 
  With CodeBuf
    If .nBuf = 0 Then                                   'If the buffer is empty (first msg to be added)
    
      Call AddCode(WNDPROC_FILTERED)                    'Add the filtered mode WndProc opcodes
    End If
    
    Call AddCode(OPCODE_CMP_EAX & Hex8(htonl(MsgNum)))  'Add the opcodes to compare the MsgNum
    
    'Add the opcodes to jump if matched
    Call AddCode(OPCODE_JE & Hex8(htonl(Not (.nBuf - IIf(Before, OFFSET_BEFORE, OFFSET_AFTER)))))
  End With
End Sub
Public Sub AddMessages(ByVal FlBefore As Boolean, ParamArray MessagesAdd())
'add all the messages in the MessagesAdd Array.
Dim I As Long
For I = 0 To UBound(MessagesAdd)
    AddMsg CLng(MessagesAdd(I)), FlBefore
Next I




End Sub

'Subclass the passed window handle.
Public Sub Subclass(hWndSub As Long, Owner As iSuperClass, Optional AllMsgs As Boolean = False)
  Dim pOwner  As Long                                   'Object address of the owner
  Dim nPos    As Long                                   'Buf pos temporary
  
  All = AllMsgs
  
  With CodeBuf
    Debug.Assert (Running = False)                      'Subclasser already running
    Debug.Assert (IsWindow(hWndSub))                    'Invalid hWnd
    Debug.Assert (Not All And .nBuf > 0) Or _
                 (All And .nBuf = 0)                    'Either filtered mode but no messages added OR All message mode but messages added.
    hWnd = hWndSub                                      'Save the window handle
    WndProcPrev = GetWindowLong(hWnd, GWL_WNDPROC)      'Save the address of the current WndProc
    pOwner = ObjPtr(Owner)                              'Get the address of the owner
    pCode = VarPtr(.code.Buf(0))                        'Get the address of our WndProc code
    
    If AllMsgs Then
    
      Call AddCode(WNDPROC_ALL)                         'Add the All messages WndProc opcodes
      Call PatchOffset(19)                              'Patch the WndProcPrev call
      Call PatchValue(43, pOwner)                       'Patch the owner
    Else
    
      Call PatchValue(31, pOwner)                       'Patch the owner
      Call PatchOffset(62)                              'Patch the BEFORE WndProcPrev call
      Call PatchOffset(71)                              'Patch the AFTER WndProcPrev call
      Call PatchValue(95, pOwner)                       'Patch the owner
      
      nPos = .nBuf + 1                                  'Save the buf pos
      Call AddCode(MSG_UNHANDLED)                       'Add the trailing unhandled WndProcPrev call
      Call PatchOffset(nPos)                            'Patch the WndProcPrev call
    End If
  End With
  
  'Debug support: uncomment the line below to crash the application which will (assuming VS is setup correctly)
  'allow you into the VS debugger where you can examine the generated opcodes and trace execution.
  'Don't call the Crash routine inside the IDE :)
  '
  'Call Crash
  
  Call SetWindowLong(hWnd, GWL_WNDPROC, pCode)          'Set our WndProc in place of the original
  Running = True
End Sub

'Unsubclass the window
Public Sub UnSubclass()
  If Running Then
    If All Then
      
      Call PatchValue(23, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
    Else
    
      CodeBuf.code.Buf(7) = &H29                        'Patch the WndProc entrance to jump to default processing JIC
      Call PatchValue(75, CODE_RETURN)                  'Patch a Leave-Return after default processing and before iSuperClass_After
    End If
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, WndProcPrev)  'Restore the previous WndProc
    CodeBuf.nBuf = 0                                    'Reset the opcode buffer
    Running = False                                     'Not running
  End If
End Sub

Private Sub Class_Terminate()
  If Running Then UnSubclass                            'Unsubclass if the Subclasser is running
End Sub

'Translate the passed hex string character pairs to bytes and stuff into the opcode buffer.
Private Sub AddCode(sOps As String)
  Dim I As Long
  Dim j As Long
  
  With CodeBuf
    j = Len(sOps)                                       'Get length of opcode string
    Debug.Assert (.nBuf + (j \ 2) <= BUF_TOP)           'Opcode buffer overflow, increase value of BUF_TOP
    
    For I = 1 To j Step 2                               'For each pair of hex chars
      
      .code.Buf(.nBuf) = Val("&H" & Mid$(sOps, I, 2))   'Convert from hex to byte, add to buffer at index
      .nBuf = .nBuf + 1                                 'Bump the opcode buffer index
    Next I
  End With
End Sub

'Return an 8 character hex representation of the passed 32 bit value
Private Function Hex8(lValue As Long) As String
  Dim s As String
  
  s = Hex$(lValue)
  Hex8 = String$(8 - Len(s), "0") & s
End Function

'Patch the passed code buffer offset with the passed value
Private Sub PatchValue(nOffset As Long, nValue As Long)
  Call CopyMemory(ByVal (pCode + nOffset), nValue, 4)
End Sub

'Patch the passed code buffer offset with the relative offset to the previous WndProc
Private Sub PatchOffset(nOffset As Long)
  Call CopyMemory(ByVal (pCode + nOffset), WndProcPrev - pCode - nOffset - 4, 4)
End Sub

'Debug Support:
'
'Crash the app allowing us into the debugger to examine opcodes
'Private Sub Crash()
'  Dim bCrash As Boolean
'
'  bCrash = True
'  If bCrash Then Call CopyMemory(ByVal 0, &HFFFFFFFF, 1)
'End Sub
