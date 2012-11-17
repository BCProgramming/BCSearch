Attribute VB_Name = "mDirectRW"
'*****************************************************************
' Module for performing Direct Read/Write access to disk sectors
'
' Written by Arkadiy Olovyannikov (ark@fesma.ru)
' Copyright 2001 by Arkadiy Olovyannikov
'
' This software is FREEWARE. You may use it as you see fit for
' your own projects but you may not re-sell the original or the
' source code.
'
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.
'*****************************************************************
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*************Win9x direct Read/Write Staff**********
Public Enum FAT_WRITE_AREA_CODE
    FAT_AREA = &H2001
    ROOT_DIR_AREA = &H4001
    DATA_AREA = &H6001
End Enum

Private Type DISK_IO
  dwStartSector As Long
  wSectors As Integer
  dwBuffer As Long
End Type
   
Private Type DIOC_REGISTER
  reg_EBX As Long
  reg_EDX As Long
  reg_ECX As Long
  reg_EAX As Long
  reg_EDI As Long
  reg_ESI As Long
  reg_Flags As Long
End Type

Private Const VWIN32_DIOC_DOS_IOCTL = 1&     'Int13 - 440X functions
Private Const VWIN32_DIOC_DOS_INT25 = 2&     'Int25 - Direct Read Command
Private Const VWIN32_DIOC_DOS_INT26 = 3&     'Int26 - Direct Write Command
Private Const VWIN32_DIOC_DOS_DRIVEINFO = 6& 'Extended Int 21h function 7305h

Private Const FILE_DEVICE_FILE_SYSTEM = &H9&
Private Const FILE_ANY_ACCESS = 0
Private Const FILE_READ_ACCESS = &H1
Private Const FILE_WRITE_ACCESS = &H2

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const INVALID_HANDLE_VALUE = -1&
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
'

'****************** NT direct Read/Write staff**************************************************
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

Private Const FILE_BEGIN = 0

Public Function DirectReadDrive(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal cBytes As Long) As Variant
   If IsWindowsNT Then
      DirectReadDrive = DirectReadDriveNT(sDrive, iStartSec, iOffset, cBytes)
   Else
      If FSName = "FAT12" Or FSName = "FAT16" Then
         DirectReadDrive = DirectReadFloppy9x(sDrive, iStartSec, iOffset, cBytes)
      Else
         DirectReadDrive = DirectReadDrive9x(sDrive, iStartSec, iOffset, cBytes)
      End If
   End If
End Function

Public Function DirectWriteDrive(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal sWrite As String, Optional AreaCode As FAT_WRITE_AREA_CODE = DATA_AREA) As Boolean
   If IsWindowsNT Then
      DirectWriteDrive = DirectWriteDriveNT(sDrive, iStartSec, iOffset, sWrite)
   Else
      If FSName = "FAT12" Or FSName = "FAT16" Then
         DirectWriteDrive = DirectWriteFloppy9x(sDrive, iStartSec, iOffset, sWrite)
      Else
         DirectWriteDrive = DirectWriteDrive9x(sDrive, iStartSec, iOffset, sWrite, AreaCode)
      End If
   End If
End Function

'===Direct Read/Write floppy using Int25/26===
'Works only for FAT12/16 systems, but much more quicker
'Then Int21 7305 function

Private Function DirectReadFloppy9x(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal cBytes As Long) As Variant
    Dim hDevice As Long
    Dim reg As DIOC_REGISTER
    Dim nSectors As Long
    Dim aOutBuff() As Byte
    Dim abResult() As Byte
    Dim nRead As Long
    nSectors = Int((iOffset + cBytes - 1) / BytesPerSector) + 1
    ReDim aOutBuff(nSectors * BytesPerSector)
    ReDim abResult(cBytes - 1) As Byte
    With reg
       .reg_EAX = Asc(UCase(sDrive)) - Asc("A")
       .reg_ESI = &H6000
       .reg_ECX = nSectors
       .reg_EBX = VarPtr(aOutBuff(0))
       .reg_EDX = iStartSec
    End With
    hDevice = CreateFile("\\.\VWIN32", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    Call DeviceIoControl(hDevice, VWIN32_DIOC_DOS_INT25, reg, Len(reg), reg, Len(reg), nRead, 0&)
    CloseHandle hDevice
    CopyMemory abResult(0), aOutBuff(iOffset), cBytes
    DirectReadFloppy9x = abResult
End Function

Private Function DirectWriteFloppy9x(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal sWrite As String) As Boolean
    Dim hDevice As Long
    Dim reg As DIOC_REGISTER
    Dim nSectors As Long
    Dim abBuff() As Byte
    Dim ab() As Byte
    Dim nRead As Long
    nSectors = Int((iOffset + Len(sWrite) - 1) / BytesPerSector) + 1
    abBuff = DirectReadFloppy9x(sDrive, iStartSec, 0, nSectors * BytesPerSector)
    ab = StrConv(sWrite, vbFromUnicode)
    CopyMemory abBuff(iOffset), ab(0), Len(sWrite)
    With reg
       .reg_EAX = Asc(UCase(sDrive)) - Asc("A")
       .reg_ESI = &H6000
       .reg_ECX = nSectors
       .reg_EBX = VarPtr(abBuff(0))
       .reg_EDX = iStartSec
    End With
    hDevice = CreateFile("\\.\VWIN32", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    DirectWriteFloppy9x = DeviceIoControl(hDevice, VWIN32_DIOC_DOS_INT26, reg, Len(reg), reg, Len(reg), nRead, 0&) And Not (reg.reg_Flags And 1)
    CloseHandle hDevice
End Function

'====Direct Read/Write drive using Int21 function 7305h====
'works with FAT12/16/32

Private Function DirectReadDrive9x(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal cBytes As Long) As Variant
    Dim hDevice As Long
    Dim reg As DIOC_REGISTER
    Dim dio As DISK_IO
    Dim abDioBuff() As Byte
    Dim nSectors As Long
    Dim aOutBuff() As Byte
    Dim abResult() As Byte
    Dim nRead As Long
    nSectors = Int((iOffset + cBytes - 1) / BytesPerSector) + 1
    ReDim abResult(cBytes - 1) As Byte
    ReDim aOutBuff(nSectors * BytesPerSector - 1)
    With dio
        .dwStartSector = iStartSec
        .wSectors = CInt(nSectors)
        .dwBuffer = VarPtr(aOutBuff(0))
    End With
    ReDim abDioBuff(LenB(dio) - 1)
    CopyMemory abDioBuff(0), dio, LenB(dio)
    CopyMemory abDioBuff(6), abDioBuff(8), 4&
    With reg
       .reg_EAX = &H7305 'function number
       .reg_ECX = -1&
       .reg_EBX = VarPtr(abDioBuff(0))
       .reg_EDX = Asc(UCase(sDrive)) - Asc("A") + 1
    End With
    hDevice = CreateFile("\\.\VWIN32", GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    Call DeviceIoControl(hDevice, VWIN32_DIOC_DOS_DRIVEINFO, reg, Len(reg), reg, Len(reg), nRead, 0&)
    CloseHandle hDevice
    CopyMemory abResult(0), aOutBuff(iOffset), cBytes
    DirectReadDrive9x = abResult
End Function

Private Function DirectWriteDrive9x(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal sWrite As String, ByVal AreaCode As FAT_WRITE_AREA_CODE) As Boolean
    Dim hDevice As Long, nSectors As Long
    Dim nRead As Long
    Dim reg As DIOC_REGISTER
    Dim dio As DISK_IO
    Dim abDioBuff() As Byte
    Dim abBuff() As Byte
    Dim ab() As Byte
    Dim bLocked As Boolean
    nSectors = Int((iOffset + Len(sWrite) - 1) / BytesPerSector) + 1
    abBuff = DirectReadDrive9x(sDrive, iStartSec, 0, nSectors * BytesPerSector)
    ab = StrConv(sWrite, vbFromUnicode)
    CopyMemory abBuff(iOffset), ab(0), Len(sWrite)
    With dio
        .dwStartSector = iStartSec
        .wSectors = CInt(nSectors)
        .dwBuffer = VarPtr(abBuff(0))
    End With
    ReDim abDioBuff(LenB(dio) - 1)
    CopyMemory abDioBuff(0), dio, LenB(dio)
    CopyMemory abDioBuff(6), abDioBuff(8), 4&
    With reg
       .reg_EAX = &H7305 'function number
       .reg_ECX = -1&
       .reg_EBX = VarPtr(abDioBuff(0))
       .reg_EDX = Asc(UCase(sDrive)) - Asc("A") + 1
       .reg_ESI = AreaCode
    End With
    hDevice = CreateFile("\\.\VWIN32", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    Dim i As Integer
    For i = 0 To 3
        If LockLogicalVolume(hDevice, Asc(UCase(sDrive)) - Asc("A") + 1, CByte(i), 0) Then
           bLocked = True
           Exit For
        End If
    Next i
    If Not bLocked Then GoTo WriteError
    DirectWriteDrive9x = DeviceIoControl(hDevice, VWIN32_DIOC_DOS_DRIVEINFO, reg, Len(reg), reg, Len(reg), nRead, 0&) And Not (reg.reg_Flags And 1)
    Call UnlockLogicalVolume(hDevice, Asc(UCase(sDrive)) - Asc("A") + 1)
WriteError:
    CloseHandle hDevice
End Function

Private Function LockLogicalVolume(hVWin32 As Long, bDriveNum As Byte, bLockLevel As Byte, wPermissions As Integer) As Boolean
    Dim fResult As Boolean
    Dim reg As DIOC_REGISTER
    Dim bDeviceCat As Byte  '  can be either 0x48 or 0x08
    Dim cb As Long
'   Try first with device category 0x48 for FAT32 volumes. If it
'   doesn 't work, try again with device category 0x08. If that
'   doesn 't work, then the lock failed.
    bDeviceCat = CByte(&H48)
ATTEMPT_AGAIN:
    reg.reg_EAX = &H440D&
    reg.reg_EBX = MAKEWORD(bDriveNum, bLockLevel)
    reg.reg_ECX = MAKEWORD(CByte(&H4A), bDeviceCat)
    reg.reg_EDX = wPermissions
    fResult = DeviceIoControl(hVWin32, VWIN32_DIOC_DOS_IOCTL, reg, LenB(reg), reg, LenB(reg), cb, ByVal 0&) And Not (reg.reg_Flags And 1)
    If (fResult = False) And (bDeviceCat <> CByte(&H8)) Then
        bDeviceCat = CByte(&H8)
        GoTo ATTEMPT_AGAIN
    End If
    LockLogicalVolume = fResult
End Function

Private Function UnlockLogicalVolume(hVWin32 As Long, bDriveNum As Byte) As Boolean
    Dim fResult As Boolean
    Dim reg As DIOC_REGISTER
    Dim bDeviceCat  As Byte '  // can be either 0x48 or 0x08
    Dim cb As Long
'   Try first with device category 0x48 for FAT32 volumes. If it
'   doesn 't work, try again with device category 0x08. If that
'   doesn 't work, then the unlock failed.
    bDeviceCat = CByte(&H48)
ATTEMPT_AGAIN:
    reg.reg_EAX = &H440D&
    reg.reg_EBX = bDriveNum
    reg.reg_ECX = MAKEWORD(CByte(&H6A), bDeviceCat)
    fResult = DeviceIoControl(hVWin32, VWIN32_DIOC_DOS_IOCTL, reg, LenB(reg), reg, LenB(reg), cb, ByVal 0&) And Not (reg.reg_Flags And 1)
    If (fResult = False) And (bDeviceCat <> CByte(&H8)) Then
        bDeviceCat = CByte(&H8)
        GoTo ATTEMPT_AGAIN
    End If
    UnlockLogicalVolume = fResult
End Function

'=============NT staff=============
'Read/Wrire drive with any file system

Private Function DirectReadDriveNT(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal cBytes As Long) As Variant
    Dim hDevice As Long
    Dim abBuff() As Byte
    Dim abResult() As Byte
    Dim nSectors As Long
    Dim nRead As Long
    nSectors = Int((iOffset + cBytes - 1) / BytesPerSector) + 1
    hDevice = CreateFile("\\.\" & UCase(Left(sDrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    Call SetFilePointer(hDevice, iStartSec * BytesPerSector, 0, FILE_BEGIN)
    ReDim abResult(cBytes - 1)
    ReDim abBuff(nSectors * BytesPerSector - 1)
    Call ReadFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
    CloseHandle hDevice
    CopyMemory abResult(0), abBuff(iOffset), cBytes
    DirectReadDriveNT = abResult
End Function

Private Function DirectWriteDriveNT(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByVal sWrite As String) As Boolean
    Dim hDevice As Long
    Dim abBuff() As Byte
    Dim ab() As Byte
    Dim nRead As Long
    Dim nSectors As Long
    nSectors = Int((iOffset + Len(sWrite) - 1) / BytesPerSector) + 1
    hDevice = CreateFile("\\.\" & UCase(Left(sDrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    abBuff = DirectReadDriveNT(sDrive, iStartSec, 0, nSectors * BytesPerSector)
    ab = StrConv(sWrite, vbFromUnicode)
    CopyMemory abBuff(iOffset), ab(0), Len(sWrite)
    Call SetFilePointer(hDevice, iStartSec * BytesPerSector, 0, FILE_BEGIN)
    Call LockFile(hDevice, LoWord(iStartSec * BytesPerSector), HiWord(iStartSec * BytesPerSector), LoWord(nSectors * BytesPerSector), HiWord(nSectors * BytesPerSector))
    DirectWriteDriveNT = WriteFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, LoWord(iStartSec * BytesPerSector), HiWord(iStartSec * BytesPerSector), LoWord(nSectors * BytesPerSector), HiWord(nSectors * BytesPerSector))
    CloseHandle hDevice
End Function

Private Function CTL_CODE(lngDeviceType, lngFunction, lngMethod, lngAccess) As Long
    CTL_CODE = ((lngDeviceType * (2 ^ 16)) Or (lngAccess * (2 ^ 14)) Or (lngFunction * (2 ^ 2)) Or lngMethod)
End Function

Private Function LockNTVolume(ByVal hDevice As Long) As Boolean
    Const LOCK_TIMEOUT = 5000   ' 5 Seconds
    Const LOCK_RETRIES = 20     ' try 20 times
    Const METHOD_BUFFERED = 0&
    Dim FSCTL_LOCK_VOLUME  As Long
    Dim lngSleepAmount As Long, ret As Long, lngTryCount As Long
    Dim bLocked As Boolean
    FSCTL_LOCK_VOLUME = CTL_CODE(FILE_DEVICE_FILE_SYSTEM, 6, METHOD_BUFFERED, FILE_ANY_ACCESS)
    lngSleepAmount = LOCK_TIMEOUT / LOCK_RETRIES
    For lngTryCount = 0 To LOCK_RETRIES
        If DeviceIoControl(hDevice, FSCTL_LOCK_VOLUME, ByVal 0, 0, ByVal 0, 0, ret, ByVal 0) Then
            bLocked = True
            Exit For
        End If
        Call Sleep(lngSleepAmount)
    Next lngTryCount
    LockNTVolume = bLocked
End Function
