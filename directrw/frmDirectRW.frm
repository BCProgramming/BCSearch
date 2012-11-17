VERSION 5.00
Begin VB.Form frmDirectRW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   150
   ClientTop       =   1005
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   4515
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   0
      Width           =   7755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive Info"
      Height          =   3495
      Left            =   60
      TabIndex        =   21
      Top             =   360
      Width           =   2355
      Begin VB.Label Label2 
         Caption         =   "Label3"
         Height          =   3255
         Left            =   120
         TabIndex        =   23
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   2355
   End
   Begin VB.Frame Frame2 
      Caption         =   "Free clusters"
      Height          =   1095
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   4620
      Width           =   3675
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   17
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   15
         Top             =   660
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   4260
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3900
      Width           =   1515
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sector in view"
      Height          =   1095
      Left            =   7500
      TabIndex        =   2
      Top             =   4620
      Width           =   2715
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   315
         Left            =   2340
         TabIndex        =   5
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bad clusters"
      Height          =   1095
      Index           =   1
      Left            =   3780
      TabIndex        =   10
      Top             =   4620
      Width           =   3675
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   7
         Left            =   2520
         TabIndex        =   8
         Top             =   660
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   6
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   5
         Left            =   1320
         TabIndex        =   13
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Label Label3 
      Caption         =   "View area"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   735
   End
   Begin VB.Menu mnuUndelete 
      Caption         =   "VB Undelete"
   End
End
Attribute VB_Name = "frmDirectRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minSector As Long, maxSector As Long
Dim bFromCode As Boolean

Private Sub Combo1_Click()
   Dim bEnable As Boolean
   bEnable = (Combo1.ListIndex = 4)
   Command1(2).Enabled = bEnable
   Command1(3).Enabled = bEnable
   Command1(6).Enabled = bEnable
   Command1(7).Enabled = bEnable
   Select Case Combo1.ListIndex
       Case 0
            minSector = 0: maxSector = 0
       Case 1
            minSector = ReservedSectors: maxSector = ReservedSectors + SectorsPerFAT - 1
       Case 2
            minSector = ReservedSectors + SectorsPerFAT: maxSector = ReservedSectors + SectorsPerFAT * 2 - 1
       Case 3
            minSector = RootDirectoryStart: maxSector = RootDirectoryStart + RootDirectoryLength
       Case 4
            minSector = DataAreaStart: maxSector = TotalSectors
   End Select
   Text2 = minSector
   Command2_Click
End Sub

Private Sub Command1_Click(Index As Integer)
   Dim nCluster As Long, nFoundCluster As Long
   Dim ret As Boolean
   Dim sCaption As String
   nCluster = Int((Val(Text2) - DataAreaStart) / SectorsPerCluster)
   If nCluster < 0 Then nCluster = 0
   sCaption = Caption
   Select Case Index
       Case 0
            Caption = sCaption & " - Searching for first free cluster..."
            nFoundCluster = FindNextCluster(Left(Drive1.Drive, 2))
            ret = nFoundCluster
       Case 1
            Caption = sCaption & " - Searching for next free cluster..."
            nFoundCluster = FindNextCluster(Left(Drive1.Drive, 2), nCluster)
            ret = nFoundCluster
       Case 2
            Caption = sCaption & " - marking cluster " & nCluster & " as bad..."
            ret = MarkCluster(Left(Drive1.Drive, 2), nCluster, FAT32_BAD_CLUSTER)
       Case 3, 7
            Caption = sCaption & " - Writting data..."
            ret = DirectWriteDrive(Left(Drive1.Drive, 2), Val(Text2), 0, Text3((Index - 3) / 4).Text, DATA_AREA)
       Case 4
            Caption = sCaption & " - Searching for first bad cluster..."
            nFoundCluster = FindNextCluster(Left(Drive1.Drive, 2), , CLUSTER_BAD)
            ret = nFoundCluster
       Case 5
            Caption = sCaption & " - Searching for next bad cluster..."
            nFoundCluster = FindNextCluster(Left(Drive1.Drive, 2), nCluster, CLUSTER_BAD)
            ret = nFoundCluster
       Case 6
            Caption = sCaption & " - marking cluster " & nCluster & " as free..."
            ret = MarkCluster(Left(Drive1.Drive, 2), nCluster, FAT32_FREE_CLUSTER)
   End Select
   If ret = False Then
      MsgBox "Nothing found or operation failed!", vbCritical, "Direct RW demo"
      Caption = sCaption
      Exit Sub
   End If
   Caption = sCaption
   If nFoundCluster = 0 Then nFoundCluster = nCluster
   bFromCode = True
   Combo1.ListIndex = 4
   bFromCode = False
   Text2 = nFoundCluster * SectorsPerCluster + DataAreaStart
   Command2_Click
End Sub

Private Sub Command2_Click()
  If bFromCode Then Exit Sub
  Dim nLogClust As Long
  Text1 = ""
  If Val(Text2) < minSector Then Text2 = minSector
  If Val(Text2) > maxSector Then Text2 = maxSector
  ShowSector Text1, Left(Drive1.Drive, 2), Val(Text2)
  nLogClust = Int((Val(Text2) - DataAreaStart) / SectorsPerCluster) + RootDirStartCluster
  If nLogClust > 0 Then
     Label1(1) = "Logical cluster: " & nLogClust
  Else
     Label1(1) = "Logical cluster: N/A"
  End If
End Sub

Private Sub Command3_Click()
  If Val(Text2) > minSector Then
     Text2 = Val(Text2) - 1
     Command2_Click
  End If
End Sub

Private Sub Command4_Click()
  If Val(Text2) < maxSector Then
     Text2 = Val(Text2) + 1
     Command2_Click
  End If
End Sub

Private Sub Drive1_Change()
  Dim sInfo As String
  MousePointer = vbHourglass
  InitDriveInfo Left(Drive1.Drive, 2)
  sInfo = "Drive type: " & DriveType & vbCrLf
  sInfo = sInfo & "Volume Label: " & VolumeLabel & vbCrLf
  sInfo = sInfo & "Serial Number: " & "&H" & Hex(VolumeSerial) & vbCrLf
  sInfo = sInfo & "File System: " & FSName & vbCrLf
  sInfo = sInfo & "Bytes per sector: " & BytesPerSector & vbCrLf
  sInfo = sInfo & "Sectors per cluster: " & SectorsPerCluster & vbCrLf
  sInfo = sInfo & "Sectors on drive: " & TotalSectors & vbCrLf & vbCrLf
  sInfo = sInfo & "Hidden sectors: " & HiddenSectors & vbCrLf
  sInfo = sInfo & "Reserved sectors: " & ReservedSectors & vbCrLf
  sInfo = sInfo & "Number of FAT copies: " & NumberOfFATCopies & vbCrLf
  sInfo = sInfo & "Sectors per FAT: " & SectorsPerFAT & vbCrLf
  For i = 1 To NumberOfFATCopies
      sInfo = sInfo & "FAT copy(" & CStr(i) & ") start at: " & ReservedSectors + SectorsPerFAT * (i - 1) & vbCrLf
  Next i
  sInfo = sInfo & "Rootirectory starts at: " & RootDirectoryStart & vbCrLf
  sInfo = sInfo & "RootDirectoryLength: " & RootDirectoryLength
  Label2 = sInfo
  Text2 = "0"
  Combo1.ListIndex = 0
  MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Combo1.AddItem "MBR"
   Combo1.AddItem "FAT[1st copy]"
   Combo1.AddItem "FAT[2nd copy]"
   Combo1.AddItem "Root Directory"
   Combo1.AddItem "Data Entries"
   SizeForm Me, Text1
   Label1(0) = "Physical:"
   Label1(1) = "Logical cluster: N/A"
   Dim i As Integer
   For i = 0 To 1
      Command1(i * 4).Caption = "Find first"
      Command1(i * 4 + 1).Caption = "Find next"
      Command1(i * 4 + 3).Caption = "Write data"
      Text3(i) = "ARK_SECRET_KEY"
   Next i
   Command1(2).Caption = "Mark bad"
   Command1(6).Caption = "Mark free"
   Command2.Caption = "&Refresh"
   Command3.Font = "Marlett"
   Command4.Font = "Marlett"
   Command3.Font.Size = 12
   Command4.Font.Size = 12
   Command3.Caption = "3"
   Command4.Caption = "4"
   Frame1.Caption = "Drive Info"
   Label2.UseMnemonic = False
   Caption = "VB Direct Read/Write Demo"
   Drive1.Drive = "c:\"
End Sub

Private Sub mnuUndelete_Click()
   frmUndelete.Show vbModal
End Sub
