VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'cVolume: represents a disk volume.
Public Enum DriveTypeConstants
 DRIVE_CDROM = 5
 DRIVE_FIXED = 3
 DRIVE_NO_ROOT_DIR = 1
 DRIVE_RAMDISK = 6
 DRIVE_REMOVABLE = 2
 DRIVE_REMOTE = 4
 DRIVE_UNKNOWN = 0
End Enum
Public Enum VolumeFSFlags
    FILE_CASE_PRESERVED_NAMES = &H2
    FILE_CASE_SENSITIVE_SEARCH = &H1
    FILE_FILE_COMPRESSION = &H10
    FILE_NAMED_STREAMS = &H40000
    FILE_PERSISTENT_ACLS = &H8
    FILE_SUPPORTS_ENCRYPTION = &H20000
    FILE_SUPPORTS_OBJECT_IDS = &H10000
    FILE_SUPPORTS_REPARSE_POINTS = &H80
    FILE_SUPPORTS_SPARSE_FILES = &H40
    FILE_VOLUME_IS_COMPRESSED = &H8000
    FILE_VOLUME_QUOTAS = &H20
End Enum
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, ByRef lpSectorsPerCluster As Long, ByRef lpBytesPerSector As Long, ByRef lpNumberOfFreeClusters As Long, ByRef lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToCaller As LARGE_INTEGER, ByRef lpTotalNumberOfBytes As LARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Private mvarRootPathName As String, VolumeName As String, VolumeSerialNumber As Long, MaximumComponentLength As Long, FilesystemFlags As Long, FileSystemName As String
Dim mfilesystemflags As Long
Private mBytesAvailableToCaller As LARGE_INTEGER
Private mTotalBytes As LARGE_INTEGER
Private mTotalFree As LARGE_INTEGER
Private MvarVolume As String
Public Property Get flags() As VolumeFSFlags
    flags = mfilesystemflags
End Property

Public Property Get Label() As String
    RefreshVolumeinfo
    Label = Trim$(VolumeName)
End Property
Public Property Get FileSystem() As String
    RefreshVolumeinfo
    FileSystem = FileSystemName


End Property

Public Sub RefreshVolumeinfo()
    FileSystemName = Space$(255)
    VolumeName = Space(255)
    GetVolumeInformation MvarVolume, VolumeName, Len(VolumeName), VolumeSerialNumber, MaximumComponentLength, mfilesystemflags, FileSystemName, Len(FileSystemName)
    VolumeName = TrimNull(VolumeName)
    FileSystemName = TrimNull(FileSystemName)

End Sub
Private Sub updateinfo()
    Dim ret As Long
    GetDiskFreeSpaceEx MvarVolume, mBytesAvailableToCaller, mTotalBytes, mTotalFree
    RefreshVolumeinfo
End Sub
Public Property Get FreeSpace() As Double
    updateinfo
    FreeSpace = LI2DBL(mTotalFree.lowpart, mTotalFree.highpart)

    


End Property
Public Property Get AvailableToCaller() As Double
    updateinfo
    AvailableToCaller = LI2DBL(mBytesAvailableToCaller.lowpart, mBytesAvailableToCaller.highpart)



End Property
Public Property Get TotalSize() As Double
    updateinfo
    TotalSize = LI2DBL(mTotalBytes.lowpart, mTotalBytes.highpart)
End Property
Public Property Get DriveType() As DriveTypeConstants
'
    DriveType = GetDriveType(MvarVolume)
End Property
Public Sub Init(ToDrive As String)
    'MvarVolume = Left$(ToDrive, 1)
    MvarVolume = ToDrive

End Sub

Public Function IsReady() As Boolean

'
Dim hfind As Long, findit As WIN32_FIND_DATA
On Error GoTo notready
hfind = FindFirstFile(Left$(MvarVolume, 1) & ":\*", findit)
If hfind <= 0 Then
    RaiseAPIError Err.LastDllError, "CVolume::IsReady"
Else
    IsReady = True
    FindClose hfind
End If
Exit Function
notready:
IsReady = False

End Function
Public Property Get RootFolder() As Directory

Dim getdir As Directory
If Not IsReady Then
    Set RootFolder = Nothing
End If
Set getdir = MdlFileSystem.FileSystem.GetDirectory(MvarVolume)
Set RootFolder = getdir



End Property

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount + -1
End Sub
