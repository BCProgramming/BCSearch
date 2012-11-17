VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileVersionInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CFileVersionInformation:

'Uses Version information APIs to return binary version information.
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long ' = 0x3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type




Private Declare Function GetFileVersionInfoA Lib "version.dll" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function GetFileVersionInfoW Lib "version.dll" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long

Private Declare Function GetFileVersionInfoSizeA Lib "version.dll" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfoSizeW Lib "version.dll" (ByVal lptstrFilename As Long, ByRef lpdwHandle As Long) As Long

Private Function GetFileVersionInfoSize(ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
    If MakeWideCalls Then
        GetFileVersionInfoSize = GetFileVersionInfoSizeW(StrPtr(lptstrFilename), lpdwHandle)
    Else
        GetFileVersionInfoSize = GetFileVersionInfoSizeA(lptstrFilename, lpdwHandle)
        
End Function



Private Function GetFileVersionInfo(ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Long) As Long
    If MakeWideCalls Then
        GetFilever
End Function

