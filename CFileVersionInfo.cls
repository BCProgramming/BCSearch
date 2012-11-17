VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileVersionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mTInfo As FILEINFO
Private mVerInfo As FILE_VERSION_INFO

Private mFilePath As String




'Public Type VS_FIXEDFILEINFO
'    dwSignature As Long
'    dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
'    dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
'    dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
'    dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
'    dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
'    dwFileFlagsMask As Long ' = 0x3F for version "0.42"
'    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
'    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
'    dwFileType As Long ' e.g. VFT_DRIVER
'    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
'    dwFileDateMS As Long ' e.g. 0
'    dwFileDateLS As Long ' e.g. 0
'End Type
'
'
'Public Type FILE_VERSION_INFO
'    FileMajor As Long
'    FileMinor As Long
'    FileRevision As Long
'    ProductMajor As Long
'    productMinor As Long
'    ProductRevision As Long
'    Fileflags As Long
'    FileOS As Long
'    FileType As Long
'    FileSubtype As Long
'
'End Type
'Public Type FILEINFO
'CompanyName As String
'Filedescription As String
'FileVersion As String
'InternalName As String
'LegalCopyright As String
'originalFilename As String
'ProductName As String
'ProductVersion As String
'End Type

'Special enumerations:

'* ----- VS_VERSION.dwFileFlags ----- */

Public Enum FileVersionFlagsEnum
    VS_FF_DEBUG = &H1&
    VS_FF_PRERELEASE = &H2&
    VS_FF_PATCHED = &H4&
    VS_FF_PRIVATEBUILD = &H8
    VS_FF_INFOINFERRED = &H10
    VS_FF_SPECIALBUILD = &H20

    
End Enum


'#define VS_FFI_SIGNATURE        0xFEEF04BDL
'#define VS_FFI_STRUCVERSION     0x00010000L
'#define VS_FFI_FILEFLAGSMASK    0x0000003FL
'


'/* ----- VS_VERSION.dwFileFlags ----- */
'#define VS_FF_DEBUG             0x00000001L
'#define VS_FF_PRERELEASE        0x00000002L
'#define VS_FF_PATCHED           0x00000004L
'#define VS_FF_PRIVATEBUILD      0x00000008L
'#define VS_FF_INFOINFERRED      0x00000010L
'#define VS_FF_SPECIALBUILD      0x00000020L
'

Public Enum FileVersionOSFlagsEnum
    VOS_UNKNOWN = &H0&
    VOS_DOS = &H10000
    VOS_OS216 = &H20000
    VOS_OS232 = &H30000
    VOS_NT = &H40000
    VOS_BASE = &H0
    VOS_WINDOWS16 = &H1
    VIS_PM16 = &H2
    VOS_PM32 = &H3&
    VOS_WINDOWS32 = &H4
    VOS_DOS_WINDOWS16 = &H10001
    VOS_DOS_WINDOWS32 = &H10004
    VOS_OS216_PM16 = &H20002
    VOS_OS232_PM32 = &H30003
    VOS_NT_WINDOWS32 = &H40004
End Enum

'/* ----- VS_VERSION.dwFileOS ----- */
'#define VOS_UNKNOWN             0x00000000L
'#define VOS_DOS                 0x00010000L
'#define VOS_OS216               0x00020000L
'#define VOS_OS232               0x00030000L
'#define VOS_NT                  0x00040000L
'
'#define VOS__BASE               0x00000000L
'#define VOS__WINDOWS16          0x00000001L
'#define VOS__PM16               0x00000002L
'#define VOS__PM32               0x00000003L
'#define VOS__WINDOWS32          0x00000004L
'
'#define VOS_DOS_WINDOWS16       0x00010001L
'#define VOS_DOS_WINDOWS32       0x00010004L
'#define VOS_OS216_PM16          0x00020002L
'#define VOS_OS232_PM32          0x00030003L
'#define VOS_NT_WINDOWS32        0x00040004L
'


Public Enum FileVersionTypeFlagsEnum
    VFT_UNKNOWN = 0
    VFT_APP
    VFT_DLL
    VDT_DRV
    VFT_FONT
    VFT_VXD
    VFD_STATIC_LIB = &H7
End Enum

'/* ----- VS_VERSION.dwFileType ----- */
'#define VFT_UNKNOWN             0x00000000L
'#define VFT_APP                 0x00000001L
'#define VFT_DLL                 0x00000002L
'#define VFT_DRV                 0x00000003L
'#define VFT_FONT                0x00000004L
'#define VFT_VXD                 0x00000005L
'#define VFT_STATIC_LIB          0x00000007L
'

Public Enum FileVersionSubtypeFlagsEnum
 VFT2_UNKNOWN = 0
 VFT2_DRV_PRINTER
 VFT2_DRV_KEYBOARD
 VFT2_DRV_LANGUAGE
 VFT2_DRV_DISPLAY
 VFT2_DRV_MOUSE
 VFT2_DRV_NETWORK
 VFT2_DRV_SYSTEM
 VFT2_DRV_INSTALLABLE
 VFT2_DRV_SOUND
 VFT2_DRV_COMM
 VFT2_DRV_INPUTMETHOD
 
 VFT2_FONT_RASTER = &H1
 VFT2_FONT_VECTOR = &H2
 VFT2_FONT_TRUETYPE = &H3
End Enum
'/* ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV ----- */
'#define VFT2_UNKNOWN            0x00000000L
'#define VFT2_DRV_PRINTER        0x00000001L
'#define VFT2_DRV_KEYBOARD       0x00000002L
'#define VFT2_DRV_LANGUAGE       0x00000003L
'#define VFT2_DRV_DISPLAY        0x00000004L
'#define VFT2_DRV_MOUSE          0x00000005L
'#define VFT2_DRV_NETWORK        0x00000006L
'#define VFT2_DRV_SYSTEM         0x00000007L
'#define VFT2_DRV_INSTALLABLE    0x00000008L
'#define VFT2_DRV_SOUND          0x00000009L
'#define VFT2_DRV_COMM           0x0000000AL
'#define VFT2_DRV_INPUTMETHOD    0x0000000BL
'
'/* ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT ----- */
'#define VFT2_FONT_RASTER        0x00000001L
'#define VFT2_FONT_VECTOR        0x00000002L
'#define VFT2_FONT_TRUETYPE      0x00000003L
'    FileMajor As Long
'    FileMinor As Long
'    FileRevision As Long
'    ProductMajor As Long
'    productMinor As Long
'    ProductRevision As Long
'    Fileflags As Long
'    FileOS As Long
'    FileType As Long
'    FileSubtype As Long
Public Property Get Flags() As FileVersionFlagsEnum
    Flags = mVerInfo.Fileflags
End Property

Public Property Get Major() As Long
    Major = mVerInfo.FileMajor
    
End Property
Public Property Get Minor() As Long
    Minor = mVerInfo.FileMinor
End Property
Public Property Get Revision() As Long
    Revision = mVerInfo.FileRevision
End Property
Public Property Get ProductMajor() As Long
    ProductMajor = mVerInfo.ProductMajor
End Property
Public Property Get ProductMinor() As Long
    ProductMinor = mVerInfo.ProductMinor
End Property
Public Property Get ProductRevision() As Long
    ProductRevision = mVerInfo.ProductRevision
End Property
Public Property Get FileType() As FileVersionTypeFlagsEnum
    FileType = mVerInfo.FileType
End Property
Public Property Get FileSubType() As FileVersionSubtypeFlagsEnum
    FileSubType = mVerInfo.FileSubType
End Property
Public Property Get TargetOS() As FileVersionOSFlagsEnum
    TargetOS = mVerInfo.FileOS
    
End Property

'CompanyName As String
'Filedescription As String
'FileVersion As String
'InternalName As String
'LegalCopyright As String
'originalFilename As String
'ProductName As String
'ProductVersion As String
Public Property Get CompanyName() As String
    CompanyName = mTInfo.CompanyName
End Property
Public Property Get Description() As String
    Description = mTInfo.Filedescription
End Property
Public Property Get InternalName() As String
    InternalName = mTInfo.InternalName
    
End Property
Public Property Get LegalCopyright() As String
    LegalCopyright = mTInfo.LegalCopyright
    
End Property
Public Property Get OriginalFilename() As String
    OriginalFilename = mTInfo.OriginalFilename
End Property
Public Property Get ProductName() As String
    ProductName = mTInfo.ProductName
End Property
Public Property Get ProductVersion() As String
    ProductVersion = mTInfo.ProductVersion
End Property
Friend Function Init(ByVal WithFile As CFile) As Boolean
    

    If modFileVersion.GetFileVersionInformation(WithFile.Fullpath, mTInfo) = eOK Then
    
        mVerInfo = modFileVersion.ReadFileVersion(WithFile.Fullpath)
    
        Init = True
    Else
        
        Init = False
    End If



End Function
