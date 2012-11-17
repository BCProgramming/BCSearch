VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cDrive: represents a disk drive.

Private Declare Function DeviceIoControl Lib "kernel32.dll" (ByVal hdevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As Any) As Long

Private mCylinders As Double
Private mMediatype As Long
Private mTracksPerCylinder As Long
Private mBytesPerSector As Long
Private mDiskSize As Double
Private mSectorsPerTrack As Long

Private Type DISK_GEOMETRY
    Cylinders As Currency
    MediaType As Long
    TracksPerCylinder As Long
    SectorsPerTrack As Long
    BytesPerSector As Long

End Type
Private Type DISK_GEOMETRY_EX
    Geometry As DISK_GEOMETRY
    diskSize As Currency
    data(0 To 1) As Byte
    


End Type

'typedef struct _PARTITION_INFORMATION_MBR {
'  BYTE    PartitionType;
'  BOOLEAN BootIndicator;
'  BOOLEAN RecognizedPartition;
'  DWORD   HiddenSectors;
'} PARTITION_INFORMATION_MBR, *PPARTITION_INFORMATION_MBR;
Private Type PARTITION_INFORMATION_MBR
    partitionType As Byte
    BootIndicator As Byte
    RecognizedPartition As Byte
    HiddenSectors As Long
End Type
Private Type PARTITION_INFORMATION_GPT
    partitionType As olelib.UUID  '16 bytes
    PartitionID As olelib.UUID    '16 bytes
    Attributes As Currency '8 bytes
    sName(0 To 36 * 2) As Byte  '74 bytes
End Type
'typedef struct {
'  PARTITION_STYLE PartitionStyle;
'  LARGE_INTEGER   StartingOffset;
'  LARGE_INTEGER   PartitionLength;
'  DWORD           PartitionNumber;
'  BOOLEAN         RewritePartition;
'  union {
'    PARTITION_INFORMATION_MBR Mbr;
'    PARTITION_INFORMATION_GPT Gpt;
'  } ;
'} PARTITION_INFORMATION_EX;


Private Type PARTITION_INFORMATION_EX
    PartitionStyle As Long
    StartingOffset As Currency
    Partitionlength As Currency
    Partitionnumber As Long
    RewritePartition As Integer
    UnionBuffer(1 To 114) As Byte
End Type

'typedef struct _PARTITION_INFORMATION_GPT {
'  GUID    PartitionType;
'  GUID    PartitionId;
'  DWORD64 Attributes;
'  WCHAR   Name[36];
'} PARTITION_INFORMATION_GPT, *PPARTITION_INFORMATION_GPT;

'typedef struct _PARTITION_INFORMATION {
'  LARGE_INTEGER StartingOffset;
'  LARGE_INTEGER PartitionLength;
'  DWORD         HiddenSectors;
'  DWORD         PartitionNumber;
'  BYTE          PartitionType;
'  BOOLEAN       BootIndicator;
'  BOOLEAN       RecognizedPartition;
'  BOOLEAN       RewritePartition;
'} PARTITION_INFORMATION, *PPARTITION_INFORMATION;


Private Type PARTITION_INFORMATION
    StartingOffset As Currency
    Partitionlength As Currency
    HiddenSectors As Long
    Partitionnumber As Long
    partitionType As Integer
    BootIndicator As Integer
    RecognizedPartition As Integer
    RewritePartition As Integer
End Type
'typedef struct _DISK_GEOMETRY_EX {
'  DISK_GEOMETRY Geometry;
'  LARGE_INTEGER DiskSize;
'  BYTE          Data[1];
'} DISK_GEOMETRY_EX, *PDISK_GEOMETRY_EX;
'typedef struct _DISK_GEOMETRY {
'  LARGE_INTEGER Cylinders;
'  MEDIA_TYPE    MediaType;
'  DWORD         TracksPerCylinder;
'  DWORD         SectorsPerTrack;
'  DWORD         BytesPerSector;
'} DISK_GEOMETRY;
'typedef enum _MEDIA_TYPE {
'  Unknown,
'  F5_1Pt2_512,
'  F3_1Pt44_512,
'  F3_2Pt88_512,
'  F3_20Pt8_512,
'  F3_720_512,
'  F5_360_512,
'  F5_320_512,
'  F5_320_1024,
'  F5_180_512,
'  F5_160_512,
'  RemovableMedia,
'  FixedMedia,
'  F3_120M_512,
'  F3_640_512,
'  F5_640_512,
'  F5_720_512,
'  F3_1Pt2_512,
'  F3_1Pt23_1024,
'  F5_1Pt23_1024,
'  F3_128Mb_512,
'  F3_230Mb_512,
'  F8_256_128,
'  F3_200Mb_512,
'  F3_240M_512,
'  F3_32M_512
'} MEDIA_TYPE;



Private Const IOCTL_DISK_GET_PARTITION_INFO = 475140
Private Const IOCTL_DISK_GET_PARTITION_INFO_EX = 458824
Private Const IOCTL_DISK_GET_DRIVE_GEOMETRY = 458752
Private Const IOCTL_DISK_GET_DRIVE_GEOMETRY_EX = 458912
Private Const IOCTL_DISK_CONTROLLER_NUMBER = 458820
Private Const IOCTL_DISK_EJECT_MEDIA = 477192
Private Const IOCTL_DISK_GET_DRIVE_LAYOUT = 475148
Private Const IOCTL_DISK_GET_DRIVE_LAYOUT_EX = 458832
Private Const IOCTL_DISK_GET_LENGTH_INFO = 475228
Private Const IOCTL_DISK_GET_MEDIA_TYPES = 461824

Private Const FSCTL_CREATE_OR_GET_OBJECT_ID = 590016
Private Const FSCTL_CREATE_USN_JOURNAL = 590055
Private Const FSCTL_FILESYSTEM_GET_STATISTICS = 589920
Private Const FSCTL_GET_COMPRESSION = 589884
Private Const FSCTL_SET_COMPRESSION = 639040
Private Const FSCTL_IS_VOLUME_DIRTY = 589944
Private Const FSCTL_READ_USN_JOURNAL = 590011
Private Const FSCTL_QUERY_USN_JOURNAL = 590068

Public Sub PopulateProperties(ByVal StrDrive As String)

    Dim hdevice As Long, retsize As Long, retval As Long
    Dim bufferin As DISK_GEOMETRY_EX
    Dim PartInfo(1 To 8) As PARTITION_INFORMATION
    hdevice = CreateFile("\\.\" & StrDrive, GENERIC_READ, FILE_SHARE_READ + FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)
    If hdevice = 0 Then Debug.Assert False
    
    'Windows 2000 (verson 5.0 has IOCTL_DISK_GET_GEOMETRY, but it doesn't have the EX version.
    If VerWinMajor = 5 And VerWinMinor = 0 Then
        'windows 2000....
        retval = DeviceIoControl(hdevice, IOCTL_DISK_GET_DRIVE_GEOMETRY, ByVal 0&, 0, bufferin.Geometry, Len(bufferin.Geometry), retsize, ByVal 0&)
        bufferin.diskSize = -1
    ElseIf IsXPOrLater Then
        retval = DeviceIoControl(hdevice, IOCTL_DISK_GET_DRIVE_GEOMETRY_EX, ByVal 0&, 0, bufferin, Len(bufferin), retsize, ByVal 0&)
    End If
    
 
    
    mDiskSize = bufferin.diskSize
    With bufferin.Geometry
        mBytesPerSector = .BytesPerSector
        mCylinders = .Cylinders
        mMediatype = .MediaType
        mSectorsPerTrack = .SectorsPerTrack
        mTracksPerCylinder = .TracksPerCylinder
    
    End With
    
    
    
    
    CloseHandle retval
    Stop



End Sub
