Attribute VB_Name = "mDriveInfo"
'*****************************************************************
' Module for collecting general drive info
' (type, capcity, file system etc.) as well as some other info.
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
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

'Public variables to store general drive info:
Public SectorsPerCluster As Byte
Public BytesPerSector As Integer
Public FSName As String, DriveType As String, VolumeLabel As String
Public VolumeSerial As Long, TotalSectors As Long
Public abMBR() As Byte

Public Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

Public Sub InitDriveInfo(ByVal sDrive As String)
    Dim asDriveType As Variant
    Dim intLength As Integer
    Dim sTemp As String
    Dim lTemp As Long
    ZeroData
    sDrive = Left(sDrive, 1) & ":\"
    VolumeLabel = String$(255, Chr$(0))
    FSName = String$(255, Chr$(0))
    GetVolumeInformation sDrive, VolumeLabel, 255, VolumeSerial, 0, 0, FSName, 255
    VolumeLabel = Left$(VolumeLabel, InStr(1, VolumeLabel, Chr$(0)) - 1)
    FSName = Left$(FSName, InStr(1, FSName, Chr$(0)) - 1)
    FSName = Trim(FSName)
    asDriveType = Array("Missing", "Unknown", "Removable", "Fixed", "Remote", "CDRom", "RAMDisk")
    DriveType = asDriveType(GetDriveType(sDrive))
    If UCase(Left(FSName, 3)) = "FAT" Then
       InitFAT sDrive
    Else
    End If
End Sub

Private Sub ZeroData()
  SectorsPerCluster = 0:  BytesPerSector = 0
  FSName = "": DriveType = "": VolumeLabel = ""
  VolumeSerial = 0: TotalSectors = 0
  NumberOfFATCopies = 0: HiddenSectors = 0
  RootDirectoryStart = 0: RootDirectoryLength = 0
  ReservedSectors = 0: SectorsPerFAT = 0
  sFAT1 = "": sFAT2 = ""
  Erase abFAT1: Erase abFAT2
  Erase aFAT_32: Erase aFAT_16: Erase aFAT_12
  Erase aRootDirEntries: Erase aDirEntries
  RootDirStartCluster = 0: DataAreaStart = 0
End Sub

Public Sub ShowSector(tx As TextBox, ByVal sDrive As String, ByVal nSector As Long)
   Dim i As Long, j As Long
   Dim s As String
   Dim dsply As String
   Dim buf() As Byte
   buf = DirectReadDrive(sDrive, nSector, 0, BytesPerSector)
   tx.FontSize = 8
   tx.FontName = "Courier New"
   For i = 0 To BytesPerSector - 1 Step 16
       s = Hex(i)
       If Len(s) = 1 Then dsply = dsply & "00"
       If Len(s) = 2 Then dsply = dsply & "0"
       dsply = dsply & s & " |"
       For j = 0 To 15
           s = Hex(buf(i + j))
           If Len(s) = 1 Then dsply = dsply & "0"
           dsply = dsply & s & " "
       Next j
       dsply = dsply & "|"
       For j = 0 To 15
           s = Chr(buf(i + j))
           If Asc(s) < 32 Then s = "."
           dsply = dsply & s
       Next j
       dsply = dsply & vbNewLine
   Next i
   tx.Text = dsply
End Sub

Public Sub SizeForm(frm As Form, txt As TextBox)
   Dim OldFont
   Dim lWidth As Long, lHeight As Long, dx As Long
   lHeight = frm.Height
   Set OldFont = frm.Font
   frm.Font = "Courier New"
   frm.FontSize = 8
   lWidth = frm.TextWidth(String(75, "0"))
   Set frm.Font = OldFont
   dx = lWidth - txt.Width
   txt.Width = lWidth
   frm.Width = frm.Width + dx
   frm.Height = lHeight
End Sub

