VERSION 5.00
Begin VB.Form FFontPreview 
   Caption         =   "TrueType Font Preview"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   Icon            =   "FFontPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   300
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6315
      Left            =   2700
      ScaleHeight     =   6255
      ScaleWidth      =   6195
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2355
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   2355
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   180
      TabIndex        =   2
      Top             =   3240
      Width           =   2355
   End
End
Attribute VB_Name = "FFontPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ****************************************************************
'  Copyright ©2001 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb/
' ****************************************************************
'  Author grants royalty-free rights to use this code within
'  compiled applications. Selling or otherwise distributing
'  this source code is not allowed without author's express
'  permission.
' ****************************************************************
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private m_Preview As CFontPreview
Private m_Debug As Boolean

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
   Call RefreshPreview
End Sub

Private Sub Form_Load()
   App.Title = Me.Caption
   Picture1.AutoRedraw = True
   If Compiled Then
      Drive1.Drive = CurDir
      Dir1.Path = CurDir
   Else
      Drive1.Drive = App.Path
      Dir1.Path = App.Path
   End If
   File1.Pattern = "*.ttf"
   If Len(Command$()) > 0 Or Compiled() = False Then m_Debug = True
End Sub

Private Sub Form_Resize()
   Dim Margin As Long
   Dim ListHeight As Long
   
   On Error Resume Next
   ' Resize controls to match new size.
   With Drive1
      Margin = .Top
      ListHeight = (Me.ScaleHeight - .Height - (4 * Margin)) \ 2
      Dir1.Move .Left, .Height + (2 * Margin), .Width, ListHeight
      File1.Move .Left, Me.ScaleHeight - ListHeight - Margin, .Width, ListHeight
      Picture1.Move .Left + .Width + Margin, .Top, Me.ScaleWidth - .Width - .Left - (3 * Margin), Me.ScaleHeight - (2 * Margin)
   End With
   ' Enable timer to repaint.
   Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
   Const HighBit As Integer = &H8000
   ' Repaint, if the user has taken finger off mouse.
   If ((GetAsyncKeyState(vbKeyLButton) And HighBit) = 0) Then
      Timer1.Enabled = False
      Call RefreshPreview
   End If
End Sub

' ****************************************************************
' Private Methods
' ****************************************************************
Private Function Compiled() As Boolean
   On Error Resume Next
   Debug.Print 1 / 0
   Compiled = (Err.Number = 0)
End Function

Private Function IsFontInstalled(ByVal FaceName As String) As Boolean
   Dim fnt As New StdFont
   ' Try assigning facename, and then
   ' compare to see if assignment took.
   fnt.Name = FaceName
   IsFontInstalled = (fnt.Name = FaceName)
End Function

Private Sub RefreshPreview()
   Const Pangram As String = "How quickly daft jumping zebras vex."
   Dim Sizes As Variant
   Dim fnt As StdFont
   Dim str As CStringBuilder
   Dim i As Long
   
   Picture1.Cls
   If File1.ListIndex >= 0 Then
      Set m_Preview = New CFontPreview
      m_Preview.FontFile = File1.Path & "\" & File1.FileName
      Picture1.CurrentX = 0
      Picture1.CurrentY = 0
      If Len(m_Preview.FaceName) Then
         Me.Caption = App.Title & ": " & m_Preview.FaceName
         Set fnt = New StdFont
         Set Picture1.Font = fnt
         fnt.Name = m_Preview.FaceName
         fnt.Bold = m_Preview.Bold
         fnt.Italic = m_Preview.Italic
         Sizes = Array(60, 48, 36, 24, 18, 14, 12, 10, 8)
         Picture1.Cls
         For i = LBound(Sizes) To UBound(Sizes)
            fnt.Size = Sizes(i)
            Picture1.Print Pangram; " ("; Sizes(i); ")"
         Next i
         fnt.Size = 24
         Picture1.Print "0123456789!@#$%^&*()~-_+=:;""',<.>/?"
      Else
         Me.Caption = App.Title
      End If
         
      Set Picture1.Font = Me.Font
      Set str = New CStringBuilder
      str.Append "FaceName: " & m_Preview.FaceName & vbCrLf
      str.Append "Family Name: " & m_Preview.FamilyName & vbCrLf
      str.Append "Subfamily Name: " & m_Preview.SubFamilyName & vbCrLf
      str.Append "Full Name: " & m_Preview.FullName & vbCrLf
      str.Append "Unique Identifier: " & m_Preview.UniqueIdentifier & vbCrLf
      str.Append "Postscript Name: " & m_Preview.PostscriptName & vbCrLf
      str.Append "Copyright: " & m_Preview.Copyright & vbCrLf
      str.Append "Trademark: " & m_Preview.Trademark & vbCrLf
      str.Append "Version: " & m_Preview.VersionString & vbCrLf
      str.Append "Installed: " & m_Preview.Installed & vbCrLf
      Picture1.Print str.ToString
      Set m_Preview = Nothing
      
      If m_Debug Then
         Clipboard.Clear
         Clipboard.SetText str.ToString
      End If
   End If
End Sub

