VERSION 5.00
Begin VB.Form frmtester 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   435
      Left            =   6960
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   675
      Left            =   1020
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1635
      Left            =   540
      ScaleHeight     =   1575
      ScaleWidth      =   8355
      TabIndex        =   5
      Top             =   4920
      Width           =   8415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3060
      Width           =   1395
   End
   Begin VB.ListBox List2 
      Height          =   3570
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   420
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   600
      TabIndex        =   0
      Top             =   420
      Width           =   1575
   End
End
Attribute VB_Name = "frmtester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFileSearchCallback
Implements IProgressCallback
Private Sub Command1_Click()
Dim X As Directory, Walker As CDirWalker, dirdir As Directory
Dim currfile As Object
Dim loopfile As CFile
Set X = GetDirectory("C:\windows\system32")
List1.Clear
'For Each dirdir In X.Directories
''Do
''    Set CurrFile = Walker.GetNext()
''    If CurrFile Is Nothing Then Exit Do
'Debug.Print dirdir.DateCreated
'   List1.AddItem dirdir.Name
'   If StrComp(Right$(dirdir.Name, InStrRev(dirdir.Name, "\") + 1), "Microsoft", vbTextCompare) = 0 Then
'   Debug.Print "Found ""Microsoft"" exiting for."
'    Exit For
'
'   End If
''
''
''Loop
'Next
List2.Clear
Set dirdir = GetDirectory("C:\")
For Each loopfile In dirdir.Files
    List2.AddItem loopfile.fullpath

Next loopfile
End Sub

Private Sub Command2_Click()
Dim tempA As CFile, TempB As CFile
Dim Tstr As FileStream, TStrB As FileStream

Set tempA = opentempfile("BC")
Set TempB = opentempfile("BC")
Set Tstr = tempA.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, CREATE_NEW, 0)
Set TStrB = TempB.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ, CREATE_NEW, 0)
Tstr.WriteString "this was the first temporary file.", StrRead_Unicode
TStrB.WriteString "this was the second temporary file.", StrRead_Unicode
Tstr.Flush
TStrB.Flush
Debug.Print "temporary files written. " & tempA.fullpath & "," & TempB.Name


Tstr.CloseStream
TStrB.CloseStream
Set tempA = getfile("C:\windows\system32\shell32.dll")
tempA.Copy "D:\", Me.hWnd
End Sub

Private Sub Command3_Click()
    Dim Msearcher As FileSearch
    Set Msearcher = New FileSearch
    Msearcher.Search "*.*", "D:\vbproj\vb\", Me
    
End Sub

Private Sub Command4_Click()
Dim VBAlist As cVBALImageList
Dim DirFile As String, currfile As CFile
Set VBAlist = New cVBALImageList
VBAlist.ColourDepth = ILC_COLOR32
VBAlist.IconSizeX = 32
VBAlist.IconSizeY = 32

VBAlist.Create

DirFile = Dir$("C:\")
Do Until DirFile = ""
    Set currfile = getfile("C:\" & DirFile)
    On Error Resume Next
    If VBAlist.KeyExists(LCase$(Right$(DirFile, 3))) Then
        
    Else
        VBAlist.AddFromHandle currfile.GetFileIcon(icon_shell), IMAGE_ICON, LCase$(Right$(DirFile, 3))
    End If
    'Picture1.Picture = VBAlist.ImagePictureStrip(1, VBAlist.ImageCount)
    'VBAlist.DrawImage "txt", Picture1.hDC, 32, 32
    Picture1.PaintPicture VBAlist.IconToPicture(VBAlist.ItemCopyOfIcon(LCase$(Right$(DirFile, 3)))), 60, 1
    Picture1.Refresh
    DirFile = Dir$
Loop

End Sub

Private Sub Command5_Click()
Display



End Sub

Private Sub Form_Click()
   ' Call ShowShellMenu(Me.hWnd, "C:\test\test.txt")
   
   
   Dim shell32 As CFile
   Set shell32 = getfile("C:\windows\system32\shell32.dll")
   shell32.CopyEx "C:\shelltest.dll", Me.hWnd, Me
End Sub

Private Function IFileSearchCallback_AllowRecurse(InDir As String) As Boolean
    IFileSearchCallback_AllowRecurse = True
    Debug.Print "recursing into subdir " & InDir
End Function

Private Sub IFileSearchCallback_ExecuteComplete(Sender As Object)
    Debug.Print "executeComplete"
End Sub

Private Sub IFileSearchCallback_Found(Sender As Object, Found As String, Optional Cancel As Boolean)
    Debug.Print Found
End Sub

Private Sub IFileSearchCallback_ProgressMessage(ByVal StrMessage As String)
'
End Sub

'Private Sub Iprogresscallback_UpdateProgress(Source As Object, Destination As Object, Optional FileSize As Double = -1#, Optional FileProgress As Double = -1#, Optional StreamSize As Double = -1#, Optional Streamprogress As Double = -1#)
'
'End Sub
Private Function IProgressCallback_UpdateProgress(Source As Object, Destination As Object, Optional FileSize As Double = -1#, Optional FileProgress As Double = -1#, Optional StreamSize As Double = -1#, Optional Streamprogress As Double = -1#) As Boolean
'

Me.Caption = "Copying: size=" & FileSize & " transferred=" & FileProgress

End Function
