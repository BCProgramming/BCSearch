VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FileMappingPageProtectionConstants
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    SEC_COMMIT = &H8000000
    SEC_IMAGE = &H1000000
    SEC_NOCACHE = &H10000000
    SEC_RESERVE = &H4000000
End Enum
Private Const SECTION_QUERY As Long = &H1
Private Const SECTION_MAP_READ As Long = &H4
Private Const SECTION_MAP_WRITE As Long = &H2
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SECTION_EXTEND_SIZE As Long = &H10
Private Const SECTION_MAP_EXECUTE As Long = &H8

Private Const SECTION_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or _
SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

Public Enum FileMapOptionConstants
 FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS
 FILE_MAP_COPY = SECTION_QUERY
 FILE_MAP_READ = SECTION_MAP_READ
 FILE_MAP_WRITE = SECTION_MAP_WRITE
End Enum
Private mHFileMapping As Long

Public Function MapFileView(ByVal MapAccess As FileMapOptionConstants, ByVal FileOffset As Double, ByVal BytestoMap As Long) As CMappedView

    Dim CreateView As CMappedView
    Dim hmapaddr As Long
    Dim fileoffsetlow As Long, fileoffsethigh As Long
    
    Set CreateView = New CMappedView
    DBL2LI FileOffset, fileoffsetlow, fileoffsethigh
    hmapaddr = MapViewOfFile(mHFileMapping, MapAccess, fileoffsethigh, fileoffsetlow, BytestoMap)
    CreateView.InitToAddress hmapaddr, Me

End Function
Friend Sub InitializeToFile(ToFile As String, Optional ByVal MappingName As String = "", Optional desiredaccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal filesharemode As FileShareModeConstants = FILE_SHARE_READ + FILE_SHARE_WRITE, Optional ByVal CreateFlags As CreationDispositionConstants = OPEN_EXISTING, Optional PageProtection As FileMappingPageProtectionConstants = PAGE_EXECUTE_READWRITE, Optional ByVal MaximumSize As Double = -1)
    Dim hFile As Long, lowpart As Long, hipart As Long
    hFile = MdlFileSystem.CreateFile(ToFile, desiredaccess, filesharemode, 0, CreateFlags, FILE_FLAG_RANDOM_ACCESS, 0)
    
    
    If hFile <> INVALID_HANDLE_VALUE Then
        'OK... now create the mapping.
        'mhfilemapping = CreateFileMapping(hfile,byval 0&,
        If MaximumSize > 0 Then
            Call DBL2LI(MaximumSize, lowpart, hipart)
        Else
            lowpart = 0: hipart = 0
        
        End If
        mHFileMapping = CreateFileMapping(hFile, ByVal 0&, PageProtection, hipart, lowpart, MappingName)
    Else
        RaiseAPIError Err.LastDllError, "CFileMapping::InitializeToFile"
    
    End If

End Sub
Friend Sub InitializeSharedMapping(ByVal SharedName As String, Optional PageProtection As FileMappingPageProtectionConstants = PAGE_EXECUTE_READWRITE, Optional ByVal MaximumSize As Double = -1)
    
    
    Dim hipart As Long, lowpart As Long
    
    
    mHFileMapping = CreateFileMapping(-1, ByVal 0&, PageProtection, hipart, lowpart, SharedName)
    



End Sub

Private Sub Class_Terminate()
CloseHandle mHFileMapping
End Sub
