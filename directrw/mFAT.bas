Attribute VB_Name = "mFAT"
'*****************************************************************
' Module for working with FAT-system drives
'
' Written by Arkadiy Olovyannikov (ark@fesma.ru
' Copyright 2001 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'*****************************************************************
'Cluster marks:
Public Enum FAT_CLUSTER_TYPE
       CLUSTER_FREE
       CLUSTER_RESERVED
       CLUSTER_BAD
       CLUSTER_END_OF_CHAIN
End Enum

Public Enum FAT12_CLUSTER_MARKS
       FAT12_FREE_CLUSTER = 0
       FAT12_RESERVED_CLUSTER_FIRST = &HFF0
       FAT12_RESERVED_CLUSTER_LAST = &HFF6
       FAT12_BAD_CLUSTER = &HFF7
       FAT12_END_OF_CHAIN_FIRST = &HFF8
       FAT12_END_OF_CHAIN_LAST = &HFFF
End Enum

Public Enum FAT16_CLUSTER_MARKS
       FAT16_FREE_CLUSTER = 0
       FAT16_RESERVED_CLUSTER_FIRST = &HFFF0
       FAT16_RESERVED_CLUSTER_LAST = &HFFF6
       FAT16_BAD_CLUSTER = &HFFF7
       FAT16_END_OF_CHAIN_FIRST = &HFFF8
       FAT16_END_OF_CHAIN_LAST = &HFFFF
End Enum

Public Enum FAT32_CLUSTER_MARKS
       FAT32_FREE_CLUSTER = 0
       FAT32_RESERVED_CLUSTER_FIRST = &HFFFFFF0
       FAT32_RESERVED_CLUSTER_LAST = &HFFFFFF6
       FAT32_BAD_CLUSTER = &HFFFFFF7
       FAT32_END_OF_CHAIN_FIRST = &HFFFFFF8
       FAT32_END_OF_CHAIN_LAST = &HFFFFFFF
End Enum

'Root directories entries

Public Const DIR_ENTRY_LENGTH = 32

Public Type FAT_DIR_ENTRY '32 bytes
   abDirEntry(31) As Byte
End Type

'***********************************************************
' Following structures are for reference only. Both they are
' byte aligned, while VB support only DWORD alignment, so we
' have to use CopyMemory for evry member of these structures
Public Type NORMAL_DIR_ENTRY '32 bytes
   abFileName(7) As Byte 'ANSI File Name (8 chars)
   abFileExt(2) As Byte  'ANSI File Ext  (3 chars)
   bFileAttribyte As Byte
   bReserved As Byte
   bTenthOfSeconds As Byte
   wTimeCreated As Integer
   wDateCreated As Integer
   wDateAccesed As Integer
   wFirstClusterHigh As Integer
   wTimeLastModified As Integer
   wDateLastModified As Integer
   wFirstClusterLow As Integer
   dwFileSize As Long
End Type

Public Type LFN_DIR_ENTRY '32 byte
   bEntryNum As Byte
   sFirst5Letters As String * 5 'Unicode string - 10 bytes
   bRSHVAttribute As Byte '&H0F
   bReserved As Byte
   bCheckSum As Byte
   sNext6Letters As String * 6  'Unicode string - 12 bytes
   wReserved As Integer '0
   sLast2Letters As String * 2  'Unicode string - 4 bytes
End Type
'***********************************************************

Public NumberOfFATCopies As Byte
Public HiddenSectors As Long, ReservedSectors As Long
Public SectorsPerFAT As Long, FATEntryLength As Single
Public RootDirectoryStart As Long, RootDirectoryLength As Long
Public DataAreaStart As Long
Public abFAT1() As Byte, abFAT2() As Byte
Public sFAT1 As String, sFAT2 As String
Public aFAT_32() As Long, aFAT_16() As Integer, aFAT_12() As Integer
Public aRootDirEntries() As FAT_DIR_ENTRY
Public aDirEntries() As FAT_DIR_ENTRY
Public RootDirStartCluster As Long

Public Function InitFAT(ByVal sDrive As String)
   Dim FileSystemType(7) As Byte
   Dim sTemp As String
   Dim intRootLength As Integer
   Dim abTemp() As Byte
   Dim lTemp As Long
   'Read first (MBR) sector:
   BytesPerSector = 512
   abMBR = DirectReadDrive(sDrive, 0, 0, BytesPerSector)
   BytesPerSector = 0
   CopyMemory BytesPerSector, abMBR(11), 2
   CopyMemory SectorsPerCluster, abMBR(13), 1
   If BytesPerSector = 0 Then BytesPerSector = 512
   'Collect extra FAT info
   CopyMemory NumberOfFATCopies, abMBR(16), 1
   CopyMemory ReservedSectors, abMBR(14), 2
   CopyMemory HiddenSectors, abMBR(28), 4
   CopyMemory TotalSectors, abMBR(19), 2
   If TotalSectors = 0 Then CopyMemory TotalSectors, abMBR(32), 4
   'Sometymes windows return "FAT" name both for 12/16 bit FAts
   'So, determine exact name:
   If FSName = "FAT" Then
      CopyMemory FileSystemType(0), abMBR(54), 8
      sTemp = StrConv(FileSystemType, vbUnicode)
      If UCase(Left(sTemp, 3)) = "FAT" Then
         FSName = Trim(sTemp)
      Else
         FSName = "FAT12"
      End If
   End If
   If FSName = "FAT32" Then
      FATEntryLength = 4
      CopyMemory SectorsPerFAT, abMBR(36), 4
      RootDirectoryStart = ReservedSectors + SectorsPerFAT * NumberOfFATCopies
      DataAreaStart = RootDirectoryStart
      CopyMemory RootDirStartCluster, abMBR(44), 4 'Logical cluster number!
   Else
      If FSName = "FAT16" Then FATEntryLength = 2 Else FATEntryLength = 1.5
      CopyMemory SectorsPerFAT, abMBR(22), 2
      CopyMemory intRootLength, abMBR(17), 2
      RootDirectoryStart = ReservedSectors + SectorsPerFAT * NumberOfFATCopies
      RootDirectoryLength = intRootLength * DIR_ENTRY_LENGTH / BytesPerSector - 1
      DataAreaStart = RootDirectoryStart + RootDirectoryLength + 1
      RootDirStartCluster = 2
   End If
   abFAT1 = DirectReadDrive(sDrive, ReservedSectors, 0, SectorsPerFAT * BytesPerSector)
   abFAT2 = DirectReadDrive(sDrive, ReservedSectors + SectorsPerFAT, 0, SectorsPerFAT * BytesPerSector)
   sFAT1 = StrConv(abFAT1, vbUnicode)
   sFAT2 = StrConv(abFAT2, vbUnicode)
   Select Case FSName
          Case "FAT32"
                ReDim aFAT_32((UBound(abFAT1) + 1) / 4 - 1)
                CopyMemory aFAT_32(0), abFAT1(0), UBound(abFAT1) + 1
          Case "FAT16"
                ReDim aFAT_16((UBound(abFAT1) + 1) / 2 - 1)
                CopyMemory aFAT_16(0), abFAT1(0), UBound(abFAT1) + 1
          Case "FAT12"
                ReDim aFAT_12((UBound(abFAT1) + 1) / 1.5 - 1)
                FillFat12
          Case Else
   End Select
   GetRootDir sDrive
End Function

Private Sub FillFat12()
   Dim i As Integer
   Dim lTemp As Long
   For i = 0 To UBound(abFAT1) / 3 - 3
       CopyMemory lTemp, abFAT1(i * 3), 3
       aFAT_12(i * 2) = lTemp And &HFFF&
       aFAT_12(i * 2 + 1) = (lTemp And &HFFF000) / &H1000
   Next i
End Sub

Public Sub GetRootDir(ByVal sDrive As String)
   Dim abTemp() As Byte
   Dim lTemp As Long, i As Long
   Dim sTemp As String
   Select Case FSName
          Case "FAT32"
                RootDirectoryLength = SearchDirEntries(sDrive, RootDirStartCluster)
                ReDim aRootDirEntries(UBound(aDirEntries))
                CopyMemory aRootDirEntries(0), aDirEntries(0), (UBound(aDirEntries) + 1) * DIR_ENTRY_LENGTH
          Case "FAT12", "FAT16"
                lTemp = (RootDirectoryLength + 1) * BytesPerSector / DIR_ENTRY_LENGTH
                ReDim aRootDirEntries(lTemp - 1)
                abTemp = DirectReadDrive(sDrive, RootDirectoryStart, 0, lTemp * DIR_ENTRY_LENGTH)
                CopyMemory aRootDirEntries(0), abTemp(0), lTemp * DIR_ENTRY_LENGTH
                For i = 0 To UBound(aRootDirEntries)
                    sTemp = StrConv(aRootDirEntries(i).abDirEntry, vbUnicode)
                    If sTemp = String(32, Chr(0)) Then Exit For
                Next i
                If i = 0 Then i = 1
                ReDim Preserve aRootDirEntries(i - 1)
                ReDim aDirEntries(i - 1)
                CopyMemory aDirEntries(0), aRootDirEntries(0), i * DIR_ENTRY_LENGTH
          Case Else
   End Select
End Sub

Public Function SearchDirEntries(ByVal sDrive As String, ByVal lStartCluster As Long) As Long
   Dim nDirEntries As Long, nEntriesInCluster As Long
   Dim curCluster As Long, EOC_MARK As Long
   Dim abTemp() As Byte
   If lStartCluster < RootDirStartCluster Then lStartCluster = RootDirStartCluster
   curCluster = lStartCluster
   Select Case FSName
          Case "FAT12": EOC_MARK = FAT12_END_OF_CHAIN_FIRST
          Case "FAT16": EOC_MARK = vbIntToUShort(FAT16_END_OF_CHAIN_FIRST)
          Case Else:    EOC_MARK = FAT32_END_OF_CHAIN_FIRST
   End Select
   nEntriesInCluster = CLng(SectorsPerCluster) * CLng(BytesPerSector) / DIR_ENTRY_LENGTH
   'Collect all valid entries
   ReDim aDirEntries(0)
   Do While curCluster < EOC_MARK
     DoEvents
     nDirEntries = nDirEntries + nEntriesInCluster
     ReDim Preserve aDirEntries(nDirEntries - 1)
     abTemp = DirectReadDrive(sDrive, DataAreaStart + (curCluster - RootDirStartCluster) * CLng(SectorsPerCluster), 0, CLng(SectorsPerCluster) * CLng(BytesPerSector))
     CopyMemory aDirEntries(nDirEntries - nEntriesInCluster), abTemp(0), UBound(abTemp) + 1
     Select Case FSName
          Case "FAT12": curCluster = aFAT_12(curCluster)
          Case "FAT16": curCluster = vbIntToUShort(aFAT_16(curCluster))
          Case Else:    curCluster = aFAT_32(curCluster)
     End Select
   Loop
   SearchDirEntries = nDirEntries * DIR_ENTRY_LENGTH / BytesPerSector - 1
End Function

Public Function FindNextCluster(ByVal sDrive As String, Optional nStartCluster As Long = 0, Optional iClusterType As FAT_CLUSTER_TYPE = CLUSTER_FREE) As Long
    nStartCluster = nStartCluster + 1
    If nStartCluster < 2 Then nStartCluster = nStartCluster + 2
    Select Case FSName
           Case "FAT32"
                FindNextCluster = FindNextCluster32(sDrive, nStartCluster, iClusterType)
           Case "FAT16"
                FindNextCluster = FindNextCluster16(sDrive, nStartCluster, iClusterType)
           Case "FAT12"
                FindNextCluster = FindNextCluster12(sDrive, nStartCluster, iClusterType)
    End Select
End Function

Private Function FindNextCluster32(ByVal sDrive As String, ByVal nStartCluster As Long, Optional iClusterType As FAT_CLUSTER_TYPE = CLUSTER_FREE)
   Dim i As Long
   For i = nStartCluster To UBound(aFAT_32)
       Select Case aFAT_32(i)
              Case FAT32_FREE_CLUSTER
                   If iClusterType = CLUSTER_FREE Then Exit For
              Case FAT32_RESERVED_CLUSTER_FIRST To FAT32_RESERVED_CLUSTER_LAST
                   If iClusterType = CLUSTER_RESERVED Then Exit For
              Case FAT32_BAD_CLUSTER
                   If iClusterType = CLUSTER_BAD Then Exit For
              Case FAT32_END_OF_CHAIN_FIRST To FAT32_END_OF_CHAIN_LAST
                   If iClusterType = CLUSTER_END_OF_CHAIN Then Exit For
              Case Else
       End Select
   Next i
   If i <= UBound(aFAT_32) Then FindNextCluster32 = i
End Function

Private Function FindNextCluster16(ByVal sDrive As String, ByVal nStartCluster As Long, Optional iClusterType As FAT_CLUSTER_TYPE = CLUSTER_FREE)
   Dim i As Long
   For i = nStartCluster To UBound(aFAT_16)
       Select Case aFAT_16(i)
              Case FAT16_FREE_CLUSTER
                   If iClusterType = CLUSTER_FREE Then Exit For
              Case FAT16_RESERVED_CLUSTER_FIRST To FAT16_RESERVED_CLUSTER_LAST
                   If iClusterType = CLUSTER_RESERVED Then Exit For
              Case FAT16_BAD_CLUSTER
                   If iClusterType = CLUSTER_BAD Then Exit For
              Case FAT16_END_OF_CHAIN_FIRST To FAT16_END_OF_CHAIN_LAST
                   If iClusterType = CLUSTER_END_OF_CHAIN Then Exit For
              Case Else
       End Select
   Next i
   If i < UBound(aFAT_16) Then FindNextCluster16 = i
End Function

Private Function FindNextCluster12(ByVal sDrive As String, ByVal nStartCluster As Long, Optional iClusterType As FAT_CLUSTER_TYPE = CLUSTER_FREE)
   Dim i As Long
   For i = nStartCluster To UBound(aFAT_12)
       Select Case aFAT_12(i)
              Case FAT12_FREE_CLUSTER
                   If iClusterType = CLUSTER_FREE Then Exit For
              Case FAT12_RESERVED_CLUSTER_FIRST To FAT12_RESERVED_CLUSTER_LAST
                   If iClusterType = CLUSTER_RESERVED Then Exit For
              Case FAT12_BAD_CLUSTER
                   If iClusterType = CLUSTER_BAD Then Exit For
              Case FAT12_END_OF_CHAIN_FIRST To FAT12_END_OF_CHAIN_LAST
                   If iClusterType = CLUSTER_END_OF_CHAIN Then Exit For
              Case Else
       End Select
   Next i
   If i <= UBound(aFAT_12) Then FindNextCluster12 = i
End Function

Public Function MarkCluster(ByVal sDrive As String, ByVal nCluster As Long, ByVal lClusterMark As FAT32_CLUSTER_MARKS) As Boolean
   Dim sFATEntry As String
   Dim iTemp_0 As Integer, iTemp_1 As Integer
   Dim lTemp As Long, lOffset As Long
   Dim FATAddrBase(1 To 2) As Long
   Dim abFATEntry() As Byte
   Select Case FSName
          Case "FAT12"
               ReDim abFATEntry(2)
               If (nCluster And 1) Then
                  iTemp_0 = aFAT_12(nCluster - 1)
                  CopyMemory iTemp_1, lClusterMark, 2
                  lOffset = (nCluster - 1) * 1.5
               Else
                  CopyMemory iTemp_0, lClusterMark, 2
                  iTemp_1 = aFAT_12(nCluster + 1)
                  lOffset = nCluster * 1.5
               End If
               aFAT_12(nCluster) = LongToUShort(lClusterMark And &HFFF&)
               lTemp = MakeFAT12(iTemp_0, iTemp_1)
               CopyMemory abFATEntry(0), lTemp, 3
               sFATEntry = StrConv(abFATEntry, vbUnicode)
          Case "FAT16"
               ReDim abFATEntry(1)
               CopyMemory abFATEntry(0), lClusterMark, 2
               aFAT_16(nCluster) = lClusterMark And &HFFFF&
               sFATEntry = StrConv(abFATEntry, vbUnicode)
               lOffset = nCluster * 2
          Case "FAT32"
               ReDim abFATEntry(3)
               CopyMemory abFATEntry(0), lClusterMark, 4
               aFAT_32(nCluster) = lClusterMark
               sFATEntry = StrConv(abFATEntry, vbUnicode)
               lOffset = nCluster * 4
   End Select
   For i = 1 To NumberOfFATCopies
   'Calculate base FAT addresses (in sectors units) for each Fat
       FATAddrBase(i) = ReservedSectors + (i - 1) * SectorsPerFAT
       MarkCluster = DirectWriteDrive(sDrive, FATAddrBase(i), lOffset, sFATEntry, FAT_AREA)
   Next i
End Function

