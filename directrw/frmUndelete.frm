VERSION 5.00
Begin VB.Form frmUndelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Undelete"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Save data"
      Height          =   315
      Left            =   9720
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command4"
      Height          =   315
      Left            =   5100
      TabIndex        =   9
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command3"
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4380
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   480
      Width           =   7875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   1935
      Left            =   60
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   1635
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2715
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Preview sector"
      Height          =   255
      Left            =   2940
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUndelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DIR_ENTRY_INFO
    dwDirEntryNum As Long
    sDosName As String
    sUnicodeName As String
    bLFNEntries As Byte
    dtDateCreated As Date
    dtLastAccesed As Date
    dtLastModified As Date
    dwFirstCluster As Long
    dwFileSize As Long
    dwRecoverSize As Long
    dwFirstFATEntry As Long
    bAttribute As VbFileAttribute
End Type
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long

Dim dei() As DIR_ENTRY_INFO
Dim sPath As String
Dim DirFirstCluster As Long
Dim MaxSectors As Long

Private Sub Command1_Click()
   If Val(Text2) > 1 Then
      Text2 = Val(Text2) - 1
      If dei(List1.ItemData(List1.ListIndex)).dwFirstCluster = 0 Then
         ShowSector Text1, Left(Drive1.Drive, 2), RootDirectoryStart + 1 - Val(Text2)
      Else
         ShowSector Text1, Left(Drive1.Drive, 2), (dei(List1.ItemData(List1.ListIndex)).dwFirstCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart + 1 - Val(Text2)
      End If
   End If
End Sub

Private Sub Command2_Click()
   If Val(Text2) < MaxSectors Then
      Text2 = Val(Text2) + 1
      If dei(List1.ItemData(List1.ListIndex)).dwFirstCluster = 0 Then
         ShowSector Text1, Left(Drive1.Drive, 2), RootDirectoryStart + 1 - Val(Text2)
      Else
         ShowSector Text1, Left(Drive1.Drive, 2), (dei(List1.ItemData(List1.ListIndex)).dwFirstCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart + 1 - Val(Text2)
      End If
   End If
End Sub

Private Sub Command3_Click()
   Dim nIndex As Long
   nIndex = List1.ItemData(List1.ListIndex)
   frmSave.m_StartCluster = dei(nIndex).dwFirstCluster
   frmSave.m_DataSize = dei(nIndex).dwFileSize
   frmSave.m_Drive = Left(Drive1.Drive, 2)
   frmSave.Show vbModal, Me
End Sub

Private Sub Drive1_Change()
  Erase dei
  InitDriveInfo Left(Drive1.Drive, 2)
  sPath = Left(Drive1.Drive, 2)
  FillUndeleteList List1
End Sub

Private Sub Form_Load()
   Drive1.Drive = "c:\"
   Label1.Font = "Terminal"
   Command1.Font = "Marlett"
   Command2.Font = "Marlett"
   Command1.Font.Size = 12
   Command2.Font.Size = 12
   Command1.Caption = "3"
   Command2.Caption = "4"
   SizeForm Me, Text1
   Text2.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload frmSave
End Sub

Private Sub List1_Click()
   Dim sName As String
   sName = dei(List1.ItemData(List1.ListIndex)).sDosName
   If sName = ".." Then
      Label2 = "of " & ParentFolder(sPath)
   Else
      Label2 = "of " & sPath & "\" & sName
   End If
   Text2 = "1"
   MaxSectors = Int((dei(List1.ItemData(List1.ListIndex)).dwFileSize / BytesPerSector / SectorsPerCluster) + 1) * 8
   ShowInfo List1.ItemData(List1.ListIndex)
   If dei(List1.ItemData(List1.ListIndex)).dwFirstCluster = 0 Then
      ShowSector Text1, Left(Drive1.Drive, 2), RootDirectoryStart
   Else
      ShowSector Text1, Left(Drive1.Drive, 2), (dei(List1.ItemData(List1.ListIndex)).dwFirstCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart
   End If
End Sub

Private Sub List1_DblClick()
   Dim sName As String, sRet As String
   Dim nIndex As Long, ret As Long
   nIndex = List1.ItemData(List1.ListIndex)
   sName = dei(nIndex).sDosName
   If sName = "" Then Exit Sub
   If Left(sName, 1) = "?" Then 'Undelete File/Folder
      If dei(nIndex).dwRecoverSize = 0 Then
         sRet = "Recovering impossible."
         If (dei(nIndex).bAttribute And vbDirectory) = False Then
            sRet = sRet & vbCrLf & "But you may save data into another file."
         End If
         MsgBox sRet, vbExclamation, "Undelete error"
         Exit Sub
      End If
      If dei(nIndex).dwRecoverSize < dei(nIndex).dwFileSize Then
         sRet = "Only first " & dei(nIndex).dwRecoverSize
         sRet = sRet & " bytes of this file can be undeleted."
         sRet = sRet & vbCrLf & "In most cases file will be corrupted."
         sRet = sRet & vbCrLf & "It's better to save data into another file."
         sRet = sRet & vbCrLf & "Proceed anywhere?"
         ret = MsgBox(sRet, vbExclamation + vbYesNo + vbDefaultButton2, "Undlete warning")
         If ret = vbNo Then Exit Sub
      End If
      sRet = ""
      sRet = InputBox("Enter first letter of deleted item which will replace " & Chr(34) & "?" & Chr(34) & " char in " & sName & ".", VBUndelete)
      If sRet = "" Then Exit Sub
      UndeleteFile List1.ItemData(List1.ListIndex), sRet
   Else 'Open subfolder
      If sName = ".." Then
         sPath = ParentFolder(sPath)
      Else
         sPath = sPath & "\" & sName
      End If
      If dei(nIndex).dwFirstCluster = 0 Then
         GetRootDir Left(Drive1.Drive, 2)
         DirFirstCluster = RootDirStartCluster
      Else
         SearchDirEntries Left(Drive1.Drive, 2), dei(nIndex).dwFirstCluster
         DirFirstCluster = dei(nIndex).dwFirstCluster
      End If
      FillUndeleteList List1
   End If
End Sub

Private Sub FillUndeleteList(lb As ListBox)
   Dim i As Long, LFNEntry As Long
   Dim sDirEntry As String
   Dim sShortName As String, sExt As String, sTemp As String
   Dim sUnicodeName As String
   Dim nListEntry As Long
   Dim iDosTime As Integer, iDosDate As Integer
   lb.Clear
   For i = 0 To UBound(aDirEntries)
       sDirEntry = StrConv(aDirEntries(i).abDirEntry, vbUnicode)
       If sDirEntry = String(32, 0) Then GoTo NextStep
       If aDirEntries(i).abDirEntry(2) = 0 Then 'LFN entry
          LFNEntry = LFNEntry + 1
          sUnicodeName = TrimNULL(StrConv(Mid(sDirEntry, 2, 10), vbFromUnicode) & StrConv(Mid(sDirEntry, 15, 12), vbFromUnicode) & StrConv(Right(sDirEntry, 4), vbFromUnicode)) & sUnicodeName
       Else 'Normal entry
          If aDirEntries(i).abDirEntry(11) And vbVolume Then GoTo NextStep
          ReDim Preserve dei(nListEntry)
          If sUnicodeName <> "" Then dei(nListEntry).sUnicodeName = sUnicodeName
          dei(nListEntry).bLFNEntries = LFNEntry
          sUnicodeName = ""
          LFNEntry = 0
          dei(nListEntry).dwDirEntryNum = i
          dei(nListEntry).bAttribute = aDirEntries(i).abDirEntry(11)
          sShortName = Trim(Left(sDirEntry, 8))
          If Asc(sShortName) = &HE5 Then Mid(sShortName, 1, 1) = "?"
          sTemp = String(Len(sShortName), 0)
          OemToChar sShortName, sTemp
          If sTemp <> "" Then sShortName = sTemp
          dei(nListEntry).sDosName = sShortName
          sExt = Trim(Mid(sDirEntry, 9, 3))
          If sExt <> "" Then dei(nListEntry).sDosName = dei(nListEntry).sDosName & "." & sExt
          CopyMemory iDosTime, aDirEntries(i).abDirEntry(14), 2
          CopyMemory iDosDate, aDirEntries(i).abDirEntry(16), 2
          dei(nListEntry).dtDateCreated = VBDateFromDosDate(iDosDate, iDosTime) + aDirEntries(i).abDirEntry(13) / 8640000
          CopyMemory iDosDate, aDirEntries(i).abDirEntry(18), 2
          dei(nListEntry).dtLastAccesed = VBDateFromDosDate(iDosDate, 0)
          CopyMemory iDosTime, aDirEntries(i).abDirEntry(22), 2
          CopyMemory iDosDate, aDirEntries(i).abDirEntry(24), 2
          dei(nListEntry).dtLastModified = VBDateFromDosDate(iDosDate, iDosTime)
          CopyMemory iDosTime, aDirEntries(i).abDirEntry(20), 2
          CopyMemory iDosDate, aDirEntries(i).abDirEntry(26), 2
          dei(nListEntry).dwFirstCluster = MakeDWord(iDosDate, iDosTime)
          CopyMemory dei(nListEntry).dwFileSize, aDirEntries(i).abDirEntry(28), 4
          If dei(nListEntry).bAttribute And vbDirectory Then
             dei(nListEntry).sDosName = UCase(dei(nListEntry).sDosName)
             If dei(nListEntry).sDosName <> "." Then
                lb.AddItem "(" & dei(nListEntry).sDosName & ")"
                lb.ItemData(lb.NewIndex) = nListEntry
             End If
          Else
             If Left(dei(nListEntry).sDosName, 1) = "?" Then
                dei(nListEntry).sDosName = LCase(dei(nListEntry).sDosName)
                lb.AddItem dei(nListEntry).sDosName
                lb.ItemData(lb.NewIndex) = nListEntry
             End If
          End If
          nListEntry = nListEntry + 1
       End If
NextStep:
   Next i
   If lb.ListCount Then lb.ListIndex = 0
   If DirFirstCluster = 0 Then DirFirstCluster = RootDirStartCluster
End Sub

Private Sub ShowInfo(ByVal nIndex As Long)
  Dim sInfo As String, sTemp As String
  Dim lTemp As Long
  Dim abTemp() As Byte
  sInfo = "Attributes:    " & FileAttributes(dei(nIndex).bAttribute)
  sTemp = dei(nIndex).sUnicodeName
  If sTemp = "" Then sTemp = "Not specified"
  sInfo = sInfo & vbCrLf & "Unicode name:  " & sTemp
  sInfo = sInfo & vbCrLf & "Date created:  " & Format(dei(nIndex).dtDateCreated, "Short Date") & Format(dei(nIndex).dtDateCreated, " hh:mm:ss")
  sInfo = sInfo & vbCrLf & "Last modified: " & Format(dei(nIndex).dtLastModified, "Short Date") & Format(dei(nIndex).dtDateCreated, " hh:mm:ss")
  sInfo = sInfo & vbCrLf & "Last access:   " & Format(dei(nIndex).dtLastAccesed, "Short Date")
  sInfo = sInfo & vbCrLf & "First cluster: " & dei(nIndex).dwFirstCluster
  If (dei(nIndex).bAttribute And vbDirectory) Then
     If Left(dei(nIndex).sDosName, 1) = "?" Then
        Call RecoverSize(nIndex)
       ' After checking FAT, check directory entry for validity -
       ' each directory, except of Root should start from "dot"
       ' and "dot-dot" entries
        abTemp = DirectReadDrive(Left(Drive1.Drive, 2), (dei(nIndex).dwFirstCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart, 0, 34)
        If abTemp(0) <> 46 Or abTemp(1) <> 32 Or abTemp(32) <> 46 Or abTemp(33) <> 46 Then
           dei(nIndex).dwRecoverSize = 0
        End If
        sInfo = sInfo & vbCrLf & "Recovering:    "
        If dei(nIndex).dwRecoverSize Then sTemp = "Possible" Else sTemp = "Impossible"
        sInfo = sInfo & sTemp
     End If
  Else
     Call RecoverSize(nIndex)
     sInfo = sInfo & vbCrLf & "File size:     " & dei(nIndex).dwFileSize & " bytes"
     sInfo = sInfo & vbCrLf & "Recover size:  " & dei(nIndex).dwRecoverSize & " bytes"
  End If
  Label1 = sInfo
End Sub

Private Function FileAttributes(ByVal bAttr As Byte) As String
   Dim sAttr As String
   If bAttr And vbVolume Then sAttr = "vbVolume,"
   If bAttr And vbDirectory Then sAttr = sAttr & "Directory,"
   If bAttr And vbHidden Then sAttr = sAttr & "Hidden,"
   If bAttr And vbSystem Then sAttr = sAttr & "System,"
   If bAttr And vbReadOnly Then sAttr = sAttr & "ReadOnly,"
   If bAttr And vbArchive Then sAttr = sAttr & "Archive,"
   If sAttr = "" Then sAttr = "Normal" Else sAttr = Left(sAttr, Len(sAttr) - 1)
   FileAttributes = sAttr
End Function

Private Function VBDateFromDosDate(ByVal iDosDate As Integer, iDosTime As Integer) As Date
   VBDateFromDosDate = DateSerial((iDosDate And &HFE00&) / &H200& + 1980, (iDosDate And &H1E0&) / &H20&, iDosDate And &HF1&) _
                     + TimeSerial((iDosTime And &HF800&) / &H800&, (iDosTime And &H7E0&) / &H20, (iDosTime And &H1F&) * 2)
End Function

Private Function RecoverSize(ByVal nIndex As Long) As Boolean
   Dim nFirstCluster As Long, nFileSize As Long
   Dim lTemp As Long, i As Long, BytesPerCluster As Long
   Dim lEOC As Long
   nFirstCluster = dei(nIndex).dwFirstCluster
   nFileSize = dei(nIndex).dwFileSize
   BytesPerCluster = CLng(BytesPerSector) * CLng(SectorsPerCluster)
   ' Check first cluster for recovery ability
   Select Case FSName
      Case "FAT12"
           lTemp = aFAT_12(nFirstCluster)
           lEOC = FAT12_END_OF_CHAIN_FIRST
      Case "FAT16"
           lTemp = aFAT_16(nFirstCluster + i)
           lEOC = FAT16_END_OF_CHAIN_FIRST
      Case "FAT32"
           lTemp = aFAT_32(nFirstCluster + i)
           lEOC = FAT32_END_OF_CHAIN_FIRST
   End Select
   'If FAT entry > 0 (not free), the only possibility to recover
   'is if FileSize < 1 cluster and FAT entry is END OF CHAIN mark
   If lTemp > 0 Then
      If lTemp >= lEOC And nFileSize <= BytesPerCluster Then
         dei(nIndex).dwRecoverSize = nFileSize
         RecoverSize = True
      End If
      Exit Function
   End If
   'If first cluster is 0 (free), we can recover all following
   'clusters with 0 (free) marks
   For i = 0 To nFileSize / BytesPerCluster
       Select Case FSName
          Case "FAT12"
               If aFAT_12(nFirstCluster + i) > 0 Then Exit For
          Case "FAT16"
               If aFAT_16(nFirstCluster + i) > 0 Then Exit For
          Case "FAT32"
               If aFAT_32(nFirstCluster + i) > 0 Then Exit For
       End Select
   Next i
   lTemp = BytesPerCluster * i
   If (lTemp > nFileSize) And (nFileSize > 0) Then lTemp = nFileSize
   If ((dei(nIndex).bAttribute And vbDirectory) = False) And (nFileSize = 0) Then lTemp = 0
   If lTemp Then
      dei(nIndex).dwRecoverSize = lTemp
      RecoverSize = True
   End If
End Function

Private Function ParentFolder(ByVal sPath As String) As String
   Dim i As Integer
   Dim sTemp As String
   For i = Len(sPath) To 1 Step -1
       sTemp = Left(sPath, i)
       If Right(sTemp, 1) = "\" Then Exit For
   Next i
   If sTemp <> "" Then
      ParentFolder = Left(sTemp, Len(sTemp) - 1)
   Else
      ParentFolder = Left(Drive1.Drive, 2)
   End If
End Function

Private Function UndeleteFile(ByVal nIndex As Long, ByVal sFirstLetter As String) As Boolean
'  First, repair DirEntries (change "&HE5" char in Dos and unicode names)
   Dim DirBaseAddress As Long, lArea As FAT_WRITE_AREA_CODE
   Dim lOffset As Long, i As Long, lFirstCluster As Long
   Dim iTemp_0 As Integer, iTemp_1 As Integer
   Dim lTemp As Long, nClusters As Long, lBytesWritten As Long
   Dim FATAddrBase(1 To 2) As Long
   Dim abFATEntry() As Byte
   Dim sFATEntry As String, sTemp As String
   sFirstLetter = Left(sFirstLetter, 1)
   sTemp = String(1, 0)
   CharToOem sFirstLetter, sTemp
   If sTemp <> "" Then sFirstLetter = sTemp
   lOffset = dei(nIndex).dwDirEntryNum * DIR_ENTRY_LENGTH
   If DirFirstCluster = RootDirStartCluster Then
      DirBaseAddress = (DirFirstCluster - RootDirStartCluster) * SectorsPerCluster + RootDirectoryStart
      lArea = ROOT_DIR_AREA
   Else
      DirBaseAddress = (DirFirstCluster - RootDirStartCluster) * SectorsPerCluster + DataAreaStart
      lArea = DATA_AREA
   End If
   'Change dos name
   DirectWriteDrive Left(Drive1.Drive, 2), DirBaseAddress, lOffset, UCase(sFirstLetter), lArea
   'Change LFN entries
   For i = 1 To dei(nIndex).bLFNEntries
       lOffset = lOffset - DIR_ENTRY_LENGTH
       If i < dei(nIndex).bLFNEntries Then
          DirectWriteDrive Left(Drive1.Drive, 2), DirBaseAddress, lOffset, Chr$(i), lArea
       Else
          DirectWriteDrive Left(Drive1.Drive, 2), DirBaseAddress, lOffset, Chr$(i Or &H40), lArea
       End If
   Next i
   'Now, restore FAT.
   'We already checked ability to restore, so if first FAT
   'entry marked as End Of Chain - we have nothing to do.
   'Otherwise, we have to restore FAT manually.
   If dei(nIndex).dwFirstFATEntry > 0 Then Exit Function
   nClusters = CInt(0.5 + dei(nIndex).dwRecoverSize / SectorsPerCluster / BytesPerSector)
   lFirstCluster = dei(nIndex).dwFirstCluster
   Select Case FSName
      Case "FAT32"
           ReDim abFATEntry(nClusters * 4 - 1)
           For i = 0 To nClusters - 2
               aFAT_32(lFirstCluster + i) = lFirstCluster + i + 1
               CopyMemory abFATEntry(i * 4), aFAT_32(lFirstCluster + i), 4
           Next i
           aFAT_32(lFirstCluster + nClusters - 1) = FAT32_END_OF_CHAIN_LAST
           CopyMemory abFATEntry((nClusters - 1) * 4), FAT32_END_OF_CHAIN_LAST, 4
           lOffset = lFirstCluster * 4
      Case "FAT16"
           ReDim abFATEntry(nClusters * 2 - 1)
           For i = 0 To nClusters - 2
               aFAT_16(lFirstCluster + i) = lFirstCluster + i + 1
               CopyMemory abFATEntry(i * 2), aFAT_16(lFirstCluster + i), 2
           Next i
           aFAT_16(lFirstCluster + nClusters - 1) = FAT16_END_OF_CHAIN_LAST
           CopyMemory abFATEntry((nClusters - 1) * 2), FAT16_END_OF_CHAIN_LAST, 2
           lOffset = lFirstCluster * 2
      Case "FAT12"
           Dim bFirstClusterOdd As Boolean, bLastClusterOdd As Boolean
           bFirstClusterOdd = (lFirstCluster And 1)
           bLastClusterOdd = ((lFirstCluster + nClusters - 1) And 1)
           ReDim abFATEntry((nClusters - bFirstClusterOdd + 1 + bLastClusterOdd) * 1.5 - 1)
           lOffset = (lFirstCluster + bFirstClusterOdd) * 1.5
           For i = 0 To nClusters - 2 Step 2
               aFAT_12(lFirstCluster + i) = lFirstCluster + i + 1
               aFAT_12(lFirstCluster + i + 1) = lFirstCluster + i + 2
               iTemp_0 = aFAT_12(lFirstCluster + i + bFirstClusterOdd)
               iTemp_1 = aFAT_12(lFirstCluster + i + 1 + bFirstClusterOdd)
               lTemp = MakeFAT12(iTemp_0, iTemp_1)
               CopyMemory abFATEntry(i * 3 / 2), lTemp, 3
               lBytesWritten = lBytesWritten + 3
           Next i
           aFAT_12(lFirstCluster + nClusters - 1) = FAT12_END_OF_CHAIN_LAST
           iTemp_0 = aFAT_12(lFirstCluster + nClusters - 1 + bLastClusterOdd)
           iTemp_1 = aFAT_12(lFirstCluster + nClusters + bLastClusterOdd)
           lTemp = MakeFAT12(iTemp_0, iTemp_1)
           CopyMemory abFATEntry((nClusters - bFirstClusterOdd + 1 + bLastClusterOdd) * 1.5 - 3), lTemp, 3
   End Select
   sFATEntry = StrConv(abFATEntry, vbUnicode)
   For i = 1 To NumberOfFATCopies
   'Calculate base FAT addresses (in sectors units) for each Fat
       FATAddrBase(i) = ReservedSectors + (i - 1) * SectorsPerFAT
       UndeleteFile = DirectWriteDrive(Left(Drive1.Drive, 2), FATAddrBase(i), lOffset, sFATEntry, FAT_AREA)
   Next i
   'Refresh list
   If DirFirstCluster = RootDirStartCluster Then
      GetRootDir Left(Drive1.Drive, 2)
   Else
      SearchDirEntries Left(Drive1.Drive, 2), DirFirstCluster
   End If
   FillUndeleteList List1
End Function
