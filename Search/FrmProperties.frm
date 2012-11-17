VERSION 5.00
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{AFFDD50D-733B-4E1C-8F98-E88F1ED6980D}#1.0#0"; "vbaListView6BC.ocx"
Begin VB.Form FrmProperties 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Properties"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4050
      TabIndex        =   0
      Top             =   5445
      Width           =   1095
   End
   Begin vbalTabStrip6.TabControl TabProperties 
      Height          =   5145
      Left            =   45
      TabIndex        =   1
      Top             =   180
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   9075
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox PicTabs 
         Height          =   4785
         Index           =   0
         Left            =   4095
         ScaleHeight     =   4725
         ScaleWidth      =   4815
         TabIndex        =   3
         Top             =   1215
         Width           =   4875
         Begin VB.TextBox TxtFilename 
            Height          =   285
            Left            =   990
            TabIndex        =   5
            Top             =   315
            Width           =   3750
         End
         Begin VB.PictureBox picIcon 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   135
            ScaleHeight     =   40
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   43
            TabIndex        =   4
            Top             =   90
            Width           =   645
         End
      End
      Begin VB.PictureBox PicTabs 
         Height          =   4785
         Index           =   1
         Left            =   360
         ScaleHeight     =   4725
         ScaleWidth      =   4770
         TabIndex        =   2
         Top             =   495
         Width           =   4830
         Begin vbaBClListViewLib6.vbalListViewCtl lvwstreams 
            Height          =   4380
            Left            =   0
            TabIndex        =   6
            Top             =   45
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   7726
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
            LabelEdit       =   0   'False
            AutoArrange     =   0   'False
            HeaderButtons   =   0   'False
            HeaderTrackSelect=   0   'False
            HideSelection   =   0   'False
            InfoTips        =   0   'False
            ScaleMode       =   3
         End
      End
   End
End
Attribute VB_Name = "FrmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msmallIml As cVBALImageList
Private mLargeIml As cVBALImageList
Private mFile As CFile
Public Sub ShowPropertiesForFile(ByVal Filename As String)

With TabProperties
        'add tabs...
        'General...

        .AddTab "General", , , "GENERAL"
        'Streams...
        .AddTab "Streams", , , "STREAMS"
        'ID3 tags, if supported...
'        .AddTab "ID3 Tags", , , "ID3"
'        .
'        If modMP3.MP3read_HasTag_ID3v2(Filename) Then
'
'        End If



    End With




    If Not msmallIml Is Nothing Then
        msmallIml.Destroy
        mLargeIml.Destroy
    End If
    Set msmallIml = New cVBALImageList
    msmallIml.IconSizeX = 16
    msmallIml.IconSizeY = 16
    msmallIml.ColourDepth = ILC_COLOR32
    msmallIml.Create
    
    Set mLargeIml = New cVBALImageList
    mLargeIml.IconSizeX = 32
    mLargeIml.IconSizeY = 32
    mLargeIml.ColourDepth = ILC_COLOR32
    mLargeIml.Create
    Set mFile = GetFile(Filename)
    'mFile.GetFileIcon (ICON_SMALL)
    msmallIml.AddFromHandle mFile.GetFileIcon(ICON_SMALL), IMAGE_ICON, "FILE"
    mLargeIml.AddFromHandle mFile.GetFileIcon(icon_large), IMAGE_ICON, "FILE"
    'Me.Icon = msmallIml.IconToPicture(mFile.GetFileIcon(ICON_SMALL))
    Const WM_SETICON As Long = &H80

    
    SendMessage Me.hWnd, WM_SETICON, ICON_SMALL, ByVal mFile.GetFileIcon(ICON_SMALL)
    mLargeIml.DrawImage "FILE", picIcon.hDC, 1, 1
    picIcon.Refresh
    TxtFilename.Text = mFile.FullPath
    
    
    'finally...
    
    
    'TabProperties_TabClick 1
    RefreshStreamData
    
    Me.Show
End Sub
Private Sub Text1_Change()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub RefreshStreamData()
Dim LoopStream As CAlternateStream
Dim newitem As cListItem

lvwstreams.ListItems.Clear
lvwstreams.Columns.Clear
On Error Resume Next
lvwstreams.Columns.Add , "NAME", "Name"
lvwstreams.Columns.Add , "SIZE", "Size"

For Each LoopStream In mFile.AlternateStreams

    Set newitem = lvwstreams.ListItems.Add(, , LoopStream.name)
    newitem.SubItems(0) = LoopStream.Size


Next
End Sub



Private Sub TabProperties_TabClick(ByVal lTab As Long)
Dim I As Long
    For I = PicTabs.LBound To PicTabs.UBound
    If I <> (lTab - 1) Then
        PicTabs(I).Visible = False
    End If
    Next I
'change visibility to avoid messing up taborder...
    PicTabs(lTab - 1).Visible = True
    PicTabs(lTab - 1).ZOrder 0
    PicTabs(lTab - 1).BorderStyle = 0
    PicTabs(lTab - 1).Move TabProperties.ClientLeft, TabProperties.ClientTop, TabProperties.ClientWidth - TabProperties.ClientLeft, TabProperties.ClientHeight - TabProperties.ClientTop
    
    CDebug.Post "tabclick " & lTab
'PicTabs(lTab).ZOrder vbBringToFront
'If lTab = 1 Then
'    RefreshStreamData
'
'
'
'End If
    
End Sub

