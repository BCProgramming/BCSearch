VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4620
   ClientLeft      =   6525
   ClientTop       =   1890
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   308
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtAdvanced 
      Height          =   4575
      Left            =   5160
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Timer TmrExpandCollapse 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3120
      Top             =   4020
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "More   >>"
      Height          =   435
      Left            =   3900
      TabIndex        =   4
      Top             =   3660
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   3900
      TabIndex        =   0
      Top             =   4140
      Width           =   1215
   End
   Begin VB.Label lblinfo 
      BackStyle       =   0  'Transparent
      Caption         =   "BASeSearch"
      Height          =   2595
      Left            =   2580
      TabIndex        =   3
      Top             =   0
      Width           =   2475
   End
   Begin VB.Label lblhyperlink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://bc-programming.com"
      DragIcon        =   "FrmAbout.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   0
      MouseIcon       =   "FrmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2340
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Uses Components from:"
      Height          =   195
      Left            =   3300
      TabIndex        =   1
      Top             =   3000
      Width           =   1680
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3240
      Picture         =   "FrmAbout.frx":0614
      Top             =   3195
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   2040
      Left            =   -45
      Picture         =   "FrmAbout.frx":2B0E
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   3240
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private mTargetWidth As Long 'Target width for timer.
Private mLastSgn As Integer
Private Property Get ExpandedWidth() As Long
'
ExpandedWidth = ScaleX(TxtAdvanced.Left + TxtAdvanced.Width + 2, vbPixels, vbTwips)
End Property
Private Property Get CollapsedWidth() As Long
Static mSysMetric As SystemMetrics
If mSysMetric Is Nothing Then Set mSysMetric = New SystemMetrics

CollapsedWidth = ScaleX(TxtAdvanced.Left + mSysMetric.GetMetric(CSM_CXDLGFRAME), vbPixels, vbTwips)



End Property
Private Sub cmdOK_Click()
Unload Me
End Sub
Private Sub PopulateAdvanced()
    '
    



End Sub
Private Sub cmdSysInfo_Click()

Static mcontCount As Long
If TmrExpandCollapse.enabled Then mcontCount = mcontCount + 1 Else mcontCount = 1

    If cmdSysInfo.caption = "More   >>" Then
        mTargetWidth = ExpandedWidth
        mLastSgn = (mTargetWidth - Me.Width)
        TmrExpandCollapse.enabled = True
        cmdSysInfo.caption = "Less  <<"
        
        
        If TxtAdvanced.Text = "" Then
        
        PopulateAdvanced
        
        
        
        End If
        
        
        
    ElseIf cmdSysInfo.caption = "Less  <<" Then
        mTargetWidth = CollapsedWidth
        mLastSgn = (mTargetWidth - Me.Width)
        TmrExpandCollapse.enabled = True
        cmdSysInfo.caption = "More   >>"
    
    End If
    
    
    
If mcontCount = 15 Then
    'easter egg of some sort. in the text box, obviously.

End If
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
Static strabout As String
Dim BCFMajor As Long, BCFMinor As Long, BCFrevision As Long, BCFDebug As Boolean
Me.Width = CollapsedWidth
'Set vbalCommandBar1.Toolbar = vbalCommandBar1.CommandBars("MAINMENU")
If strabout = "" Then
    bcfile.GetBCFileVersion BCFMajor, BCFMinor, BCFrevision, BCFDebug
    strabout = "BASeCamp BCSearch " & vbCrLf & _
          "Version " & App.Major & "." & Trim$(App.Minor) & " Revision " & App.Revision & vbCrLf & vbCrLf & _
            "Copyright 2008-2009 BASeCamp Corporation." & vbCrLf & _
            "BCFile Version: " & vbCrLf & Trim$(BCFMajor) & "." & Trim$(BCFMinor) & " Revision " & BCFrevision & " " & IIf(BCFDebug, " <Debug>", " <Release>")

End If
lblinfo.caption = strabout
'ExtendFrame Me
'BlurForm Me
mLastSgn = 1
End Sub

Private Sub Picture1_Click()
    End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub lblhyperlink_Click(Index As Integer)
HyperJump lblhyperlink(Index).caption
End Sub

Private Sub lblhyperlink_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
lblhyperlink_Click Index
End Sub

Private Sub lblhyperlink_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
 With lblhyperlink(Index)
    If State = vbLeave Then
            'now outside the control.
            
            .ForeColor = vbBlue
            .FontUnderline = False
            .FontBold = False
            .Drag vbEndDrag
        End If
        
    End With
End Sub



Private Sub HyperJump(ByVal URL As String)
   ShellExecute 0&, StrPtr(vbNullString), StrPtr(URL), StrPtr(vbNullString), StrPtr(vbNullString), vbNormalFocus
End Sub

Private Sub lblhyperlink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With lblhyperlink(Index)
    
        
    
   
        .ForeColor = RGB(25, 25, 255)
        .FontBold = False
        .FontUnderline = True
       ' Captured = True
    .Drag vbBeginDrag
    End With
End Sub

Private Sub TmrExpandCollapse_Timer()
'move towards mtargetwidth.
'the farther away we are from it, the faster we will go.... well, ideally.
Dim signuse As Long

CDebug.Post "Me.width=" & Me.Width
CDebug.Post "targetwidth=" & mTargetWidth
signuse = Sgn(mTargetWidth - Me.Width)
If signuse <> mLastSgn And Abs(mLastSgn) <= 1 Then
    TmrExpandCollapse.enabled = False
    Me.Width = mTargetWidth
    Exit Sub
Else
    Me.Width = Me.Width + signuse * 100
End If
mLastSgn = signuse

End Sub
