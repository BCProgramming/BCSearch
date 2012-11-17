Attribute VB_Name = "Common"
Option Explicit

Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub RtlZeroMemory Lib "kernel32" (ptr As Any, ByVal Length As Long)
Public Declare Function VirtualProtect Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32s As Long = 0


Private OSVersion As OSVERSIONINFO, M_WinNT As Boolean
Private flInit As Boolean
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Declare Sub GetSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

Private mSysInfo As SYSTEM_INFO
Public Property Get SystemInfo() As SYSTEM_INFO
    Init
    SystemInfo = mSysInfo
End Property

Private Sub Init()
    Dim lngreturn As Long
    If flInit Then Exit Sub
    flInit = True
   ' OSVersion.dwOSVersionInfoSize = Len(OSVersion)
   ' lngreturn = GetVersionEx(OSVersion)
    M_WinNT = IsWinNt()
    GetSystemInfo mSysInfo
End Sub
Public Function VerWinMajor() As Long
    Init
    VerWinMajor = OSVersion.dwMajorVersion
    
End Function
Public Function VerWinMinor() As Long
    Init
    VerWinMinor = OSVersion.dwMinorVersion
    
End Function
Public Function VerWinBuildNumber() As Long
    Init
    VerWinBuildNumber = OSVersion.dwBuildNumber
End Function

Public Function IsWinNt() As Boolean
'===========================================================================
'   IsWinNT - Returns true if we're running Windows NT.
'===========================================================================
    
    ' NOTE: OSVERSIONINFO is defined at the module
    ' for imporved performance.
    
    PopVersion
    IsWinNt = OSVersion.dwPlatformId = VER_PLATFORM_WIN32_NT ' return the result
    
End Function
Private Sub PopVersion()
  If OSVersion.dwOSVersionInfoSize = 0 Then               ' this is our first time making this call
        OSVersion.dwOSVersionInfoSize = Len(OSVersion)      ' initialize so API knows which version being used
        GetVersionEx OSVersion                              ' make the call once & then save/re-use it
    End If
End Sub
Public Function IsVistaOrLater() As Boolean
    PopVersion

    IsVistaOrLater = OSVersion.dwMajorVersion >= 6 ' return the result


End Function
Public Function IsXPOrLater() As Boolean
PopVersion
    IsXPOrLater = OSVersion.dwMajorVersion > 5 Or (OSVersion.dwMajorVersion = 5 And OSVersion.dwMinorVersion >= 1)
End Function
Public Function Is2KOrLater() As Boolean
    PopVersion
    Is2KOrLater = OSVersion.dwMajorVersion >= 5
End Function
'Public Const S_OK As Long = &H0
'Public Const S_FALSE As Long = &H1

' For some reason VarType() has a habit of returning the wrong type when
' it holds an object reference that has a default property, so this way checks
' to see if it actually holds a valid object reference rather than just looking
' at the first two bytes of the Variant. ObjPtr will return an error if you
' attempt to use it on a non-object, so it will tell us if 'var' holds
' a valid object reference.
Public Function VariantIsObject(ByVal var As Variant) As Boolean
    On Error Resume Next
    
    ObjPtr var
    VariantIsObject = (Err.Number = 0)
End Function

Public Function GetRefCount(obj As IUnknown) As Long
   If obj Is Nothing Then Exit Function
   CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
   GetRefCount = GetRefCount - 2
End Function
