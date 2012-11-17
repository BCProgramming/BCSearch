VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saving data"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   540
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Save only first"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Save in original file size"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save data"
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "bytes"
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_StartCluster As Long
Public m_DataSize As Long
Public m_Drive As String

Private Sub Command1_Click()
   Dim nSize As Long, nFile As Long, nStartSector As Long
   Dim abTemp() As Byte
   Dim sTemp As String
   sTemp = ShowSave(hWnd, "00000001.tmp", , , OFN_CREATEPROMPT + OFN_OVERWRITEPROMPT, "All files|*.*")
   If sTemp = "" Then Exit Sub
   If Option1(0).Value Then
      nSize = m_DataSize
   Else
      nSize = Val(Text1)
   End If
   If nSize > m_DataSize Then nSize = m_DataSize
   If m_StartCluster = 0 Then
      nStartSector = RootDirectoryStart
   Else
      nStartSector = (m_StartCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart
   End If
   abTemp = DirectReadDrive(m_Drive, nStartSector, 0, nSize)
   nFile = FreeFile
   Open sTemp For Binary As #nFile
        Put #nFile, , abTemp
   Close #nFile
   Me.Hide
End Sub

Private Sub Command2_Click()
   Me.Hide
End Sub

Private Sub Form_Activate()
   If m_DataSize = 0 Then m_DataSize = BytesPerSector * SectorsPerCluster
   Text1 = m_DataSize
   Label1 = m_DataSize & " bytes."
   If Option1(0).Value = False And Option1(1).Value = False Then Option1(0).Value = True
End Sub

Private Sub Text1_GotFocus()
   Option1(1).Value = True
End Sub
