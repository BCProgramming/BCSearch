Attribute VB_Name = "MdlSecurity"
Option Explicit
'Security APIs and wrappers.
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, ByRef TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeDisplayName Lib "advapi32.dll" Alias "LookupPrivilegeDisplayNameA" (ByVal lpSystemName As String, ByVal lpName As String, ByVal lpDisplayName As String, ByRef cbDisplayName As Long, ByRef lpLanguageID As Long) As Long

Private Declare Function LookupPrivilegeName Lib "advapi32.dll" Alias "LookupPrivilegeNameA" (ByVal lpSystemName As String, ByRef lpLuid As LARGE_INTEGER, ByVal lpName As String, ByRef cbName As Long) As Long

Private Declare Function PrivilegeCheck Lib "advapi32.dll" (ByVal ClientToken As Long, ByRef RequiredPrivileges As PRIVILEGE_SET, ByVal pfResult As Long) As Long
Private Declare Sub SetLastErrorAPI Lib "kernel32.dll" Alias "SetLastError" (ByVal dwErrCode As Long)


'Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByRef TokenInformationClass As Integer, ByRef TokenInformation As Any, ByVal TokenInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, ByRef TokenInformation As Any, ByVal TokenInformationLength As Long, ByRef ReturnLength As Long) As Long

Public Enum PrivilegeAttributes
  SE_PRIVILEGE_ENABLED = &H2
  SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
  SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000
End Enum

Private Enum TokenPrivAction
  GetPrivilege = 0
  SetPrivilege
End Enum

Public Enum TOKEN_INFORMATION_CLASS
  TokenUser = 1
  TokenGroups
  TokenPrivileges
  TokenOwner
  TokenPrimaryGroup
  TokenDefaultDacl
  TokenSource
  TokenType
  TokenImpersonationLevel
  TokenStatistics
  TokenRestrictedSids
  TokenSessionId
End Enum

Public Enum TokenAccessRights
  ' // Token Specific Access Rights.
  READ_CONTROL = &H20000
  STANDARD_RIGHTS_REQUIRED = &HF0000
  STANDARD_RIGHTS_READ = READ_CONTROL
  STANDARD_RIGHTS_WRITE = READ_CONTROL
  STANDARD_RIGHTS_EXECUTE = READ_CONTROL
  
  TOKEN_ASSIGN_PRIMARY = &H1&
  TOKEN_DUPLICATE = &H2&
  TOKEN_IMPERSONATE = &H4&
  TOKEN_QUERY = &H8&
  TOKEN_QUERY_SOURCE = &H10&
  TOKEN_ADJUST_PRIVILEGES = &H20&
  TOKEN_ADJUST_GROUPS = &H40&
  TOKEN_ADJUST_DEFAULT = &H80&
  TOKEN_ADJUST_SESSIONID = &H100&
  TOKEN_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_SESSIONID Or TOKEN_ADJUST_DEFAULT&
  TOKEN_READ = STANDARD_RIGHTS_READ Or TOKEN_QUERY
  TOKEN_WRITE = STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT
  TOKEN_EXECUTE = STANDARD_RIGHTS_EXECUTE
End Enum
Public Const ERROR_NO_TOKEN       As Long = 1008&
Public Const NAME_SIZE            As Long = 52
Public Const ERROR_NOT_ALL_ASSIGNED = 1300&
Public Const PRIVILEGE_SET_ALL_NECESSARY = (1)

' // NT Defined Privileges (24 privileges defined in NT4)
Public Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege"
Public Const SE_TCB_NAME = "SeTcbPrivilege"
Public Const SE_SECURITY_NAME = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP_NAME = "SeBackupPrivilege"
Public Const SE_RESTORE_NAME = "SeRestorePrivilege"
Public Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Public Const SE_DEBUG_NAME = "SeDebugPrivilege"
Public Const SE_AUDIT_NAME = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege"
Public Const SE_UNDOCK_NAME = "SeUndockPrivilege"
Public Const SE_SYNC_AGENT_NAME = "SeSyncAgentPrivilege"
Public Const SE_ENABLE_DELEGATION_NAME = "SeEnableDelegationPrivilege"
Public Const SE_MANAGE_VOLUME_NAME = "SeManageVolumePrivilege"


'Private Type LUID
'    LowPart As Long
'    HighPart As Long
'End Type
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LARGE_INTEGER
    Attributes As Long
End Type


Public Const MaxPrivsHardCode     As Long = 40
Public Type TOKEN_PRIVILEGES
  PrivilegeCount As Long
  Privileges(MaxPrivsHardCode)  As LUID_AND_ATTRIBUTES
End Type

Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpUid As LARGE_INTEGER) As Long


Private Type PRIVILEGE_SET
    PrivilegeCount As Long
    Control As Long
    Privilege(0) As LUID_AND_ATTRIBUTES
End Type


'Private Const NAME_SIZE As Long = 52

Private Function InvalidHandleValue(ByVal Handle As Long) As Boolean

    Const INVALID_HANDLE_VALUE = -1
    
    Select Case Handle
    Case 0, INVALID_HANDLE_VALUE
        InvalidHandleValue = True
    End Select
    
End Function
Public Function securityAPItest()
Dim testStr() As String
Dim countget As Long
testStr = GetPrivilegeNames(countget)

Stop

End Function

Public Function GetPrivilegeNames(ByRef Count As Long) As String()

  Dim TokenPrivs      As TOKEN_PRIVILEGES
  Dim cbTokenPrivs    As Long
  Dim PrivNames()       As String
  Dim PrivName        As String
  Dim ccPrivName      As Long
  Dim Success         As Boolean
  Dim hToken          As Long
  Dim nErr            As Long
  Dim sErr            As String
  Dim i               As Long
  Count = 0
  On Error GoTo ErrHandler
    
  ' NOTE: Caller should call IsWinNt()

  ' New list to store priv names
  'Set PrivNames = New collection
  
  ' Get handle for the current thread token.
  Success = OpenThreadToken(GetCurrentThread(), TOKEN_READ, 1&, hToken)
  If InvalidHandleValue(hToken) Then
  
    ' No thread token; Try for a handle to the current process token.
    Success = OpenProcessToken(GetCurrentProcess(), TOKEN_READ, hToken)
    If InvalidHandleValue(hToken) Then
      ' notify caller
      'ApiRaise DescPrefix:="OpenProcessToken failed."
      RaiseAPIError Err.LastDllError, "GetPrivilegeNames"
    End If
    
  End If

  ' Fetch info about token privileges.
  Success = GetTokenInformation(hToken, TokenPrivileges, TokenPrivs, LenB(TokenPrivs) * 2, cbTokenPrivs)
  If Not Success Then                             ' call failed
    RaiseAPIError Err.LastDllError, "GetPrivilegeNames"
  End If

  For i = LBound(TokenPrivs.Privileges) To TokenPrivs.PrivilegeCount - 1
    With TokenPrivs.Privileges(i)
    
      ' Size buffer to fit name of privilege
      PrivName = String(NAME_SIZE, 0&)
      ccPrivName = NAME_SIZE
      
      ' Lookup the name for this priv. Throw on failure.
      Success = LookupPrivilegeName(vbNullString, .pLuid, PrivName, ccPrivName)
      If Not Success Then
         RaiseAPIError Err.LastDllError, "GetPrivilegeNames"
      End If
      
      ' Trim nulls & add it to the list
      PrivName = Left$(PrivName, ccPrivName)
      'PrivNames.Add PrivName, PrivName
      Count = Count + 1
      ReDim Preserve PrivNames(1 To Count)
      PrivNames(Count) = PrivName
      
      
    End With
  Next i
  
  ' Return the list of priv names
  GetPrivilegeNames = PrivNames
  'Set PrivNames = Nothing
  
ExitLabel:
  
  ' Cleanup
  If Not InvalidHandleValue(hToken) Then
    CloseHandle hToken
  End If
  
  ' Throw error if raised
  If nErr <> 0 Then
    On Error GoTo 0
    Err.Raise nErr, "MdlSecurity", sErr
  End If
  
  Exit Function
  Resume
ErrHandler:
  
  ' Save error info
  nErr = Err.Number
  sErr = Err.Description
  Debug.Assert 0
  Resume ExitLabel
  
End Function


Public Function GetTokenPrivilege( _
  ByVal PrivilegeName As String, _
  Optional ByVal SystemName As String) As Boolean
  
'===========================================================================
' GetTokenPrivilege - Gets a named privilege for the current thread or
' process token.
'
' PrivilegeName   A named privilege such as "SeBackupPrivilege"
' SystemName      The name of the machine on which to
  
' RETURNS         True on success; or,
' ERRORS          Thrown to the caller upon failure.
'===========================================================================

  Dim Success         As Boolean
  Dim hToken          As Long
  Dim nErr            As Long
  Dim sErr            As Long
  Dim PrivSet         As PRIVILEGE_SET
  Dim pfResult        As Long
  
  On Error GoTo ErrHandler

  ' NOTE: Caller should call IsWinNt()

  ' Set to null if nothing supplied
  If Len(SystemName) = 0 Then
    SystemName = vbNullString
  End If

  ' Get handle for the current thread token.
  Success = OpenThreadToken(GetCurrentThread(), TOKEN_READ, False, hToken)
  If InvalidHandleValue(hToken) Then

    ' No thread token; Try for a handle to the current process token.
    Success = OpenProcessToken(GetCurrentProcess(), TOKEN_READ, hToken)
    If InvalidHandleValue(hToken) Then
      ' notify caller
      'ApiRaise DescPrefix:="OpenProcessToken failed."
      RaiseAPIError Err.LastDllError, "MdlSecurity::GetTokenPrivilege"
      
    End If

  End If

  With PrivSet

    ' Only 1 element in PRIVILEGE_SET
    .PrivilegeCount = 1
    .Control = PRIVILEGE_SET_ALL_NECESSARY

    ' Fill the LUID for the named privilege on SystemName. Throw on failure.
    SetLastErrorAPI 0
    Success = LookupPrivilegeValue(SystemName, PrivilegeName, .Privilege(0).pLuid)
    If Not Success Then
     ' ApiRaise DescPrefix:="LookupPrivilegeValue failed."
     RaiseAPIError Err.LastDllError, "GetTokenPrivilege"
    End If

    ' Test that the privilege is enabled.
    Success = PrivilegeCheck(hToken, PrivSet, pfResult)
    If Not Success Then
      'ApiRaise DescPrefix:="PrivilegeCheck failed."
      RaiseAPIError Err.LastDllError, "MdlSecurity::GetTokenPrivilege"
    End If
    
    GetTokenPrivilege = CBool(pfResult)
  
  End With
  
ExitLabel:
  
  ' Cleanup
  If Not InvalidHandleValue(hToken) Then
    CloseHandle hToken
    hToken = 0
  End If
  
  ' Throw trapped errors, if any
  If nErr <> 0 Then
    On Error GoTo 0
    'Err.Raise nErr, Module, sErr
    Err.Raise nErr, "MdlSecurity", sErr
  End If
  
  Exit Function
  Resume
ErrHandler:
  
  nErr = Err.Number
  sErr = Err.Description
  Resume ExitLabel
  
End Function
