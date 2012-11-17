VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExecutive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IAsyncProcess
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Enum ProcessCreationFlags
    CREATE_BREAKAWAY_FROM_JOB = &H1000000
    CREATE_DEFAULT_ERROR_MODE = &H4000000
    CREATE_NEW_CONSOLE = &H10
    CREATE_NEW_PROCESS_GROUP = &H200
    CREATE_NO_WINDOW = &H8000000
    CREATE_SEPARATE_WOW_VDM = &H800
    CREATE_SHARED_WOW_VDM = &H1000
    CREATE_SUSPENDED = &H4
    CREATE_UNICODE_ENVIRONMENT = &H400
    DEBUG_ONLY_THIS_PROCESS = &H2
    DEBUG_PROCESS = &H1
    DETACHED_PROCESS = &H8
    CREATE_PROTECTED_PROCESS = &H40000
    CREATE_PRESERVE_CODE_AUTHZ_LEVEL = &H2000000
    EXTENDED_STARTUPINFO_PRESENT = &H80000
    INHERIT_PARENT_AFFINITY = &H10000
End Enum
Public Enum StartFlagsEnum
 STARTF_FORCEOFFFEEDBACK = &H80
 STARTF_FORCEONFEEDBACK = &H40
 STARTF_RUNFULLSCREEN = &H20
 STARTF_USECOUNTCHARS = &H8
 STARTF_USEFILLATTRIBUTE = &H10
 STARTF_USEHOTKEY = &H200
 STARTF_USEPOSITION = &H4
 STARTF_USESHOWWINDOW = &H1
 STARTF_USESIZE = &H2
 STARTF_USESTDHANDLES = &H100
End Enum

Private Declare Function CreateProcessA Lib "kernel32.dll" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CreateProcessW Lib "kernel32.dll" (ByRef lpApplicationName As Long, ByRef lpCommandLine As Long, _
ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByRef lpCurrentDirectory As Long, _
ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CreateProcessANSI Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long


Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long




Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type


Private mStartupInfo As STARTUPINFO
Private mProcInfo As PROCESS_INFORMATION
Private mAsync As CAsyncExec


Private Function CreateProcess(ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
ByVal dwCreationFlags As ProcessCreationFlags, ByRef lpEnvironment As String, ByVal lpCurrentDirectory As String, _
ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long


Dim tatr As SECURITY_ATTRIBUTES
Dim patr As SECURITY_ATTRIBUTES
tatr.lpSecurityDescriptor = 0
patr.lpSecurityDescriptor = 0
tatr.bInheritHandle = 1
patr.bInheritHandle = 1
Dim useEnvironment As Long, useCmdLine As Long, useCurdir As Long
If lpEnvironment = "" Then useEnvironment = vbNullString Else useEnvironment = lpEnvironment

mStartupInfo.cb = Len(mStartupInfo)

'If IsWinNt Then
'    CreateProcess = CreateProcessW(StrPtr(lpApplicationName), useCmdLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, useEnvironment, useCurdir, lpStartupInfo, lpProcessInformation)
'    Debug.Print Err.LastDllError
'
'Else
    CreateProcess = CreateProcessANSI(vbNullString, lpApplicationName & lpCommandLine, patr, tatr, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation)
    Debug.Print Err.LastDllError
'End If
    
    Set mAsync = New CAsyncExec
    mAsync.Start Me, "WaitForProcess"
    
End Function


Public Sub Launch(ByVal AppName As String)

Dim Ret As Long
Ret = CreateProcess(AppName, "", 0, 0, 0, INHERIT_PARENT_AFFINITY, "", "C:\windows\system32\", mStartupInfo, mProcInfo)
Debug.Print "return:" & Ret
Stop

End Sub
'Public Function Redirect(cmdLine As String) As String
'    Dim i%, t$
'    Dim pa As SECURITY_ATTRIBUTES
'    Dim pra As SECURITY_ATTRIBUTES
'    Dim tra As SECURITY_ATTRIBUTES
'    Dim pi As PROCESS_INFORMATION
'    Dim sui As STARTUPINFO
'    Dim hRead As Long
'    Dim hWrite As Long
'    Dim bRead As Long
'    Dim hstdInPipeWrite As Long, hStdInPipeRead As Long
'    Dim lpBuffer As String, wholestr As String
'    pa.nLength = Len(pa)
'    pa.lpSecurityDescriptor = 0
'    pa.bInheritHandle = True
'
'    pra.nLength = Len(pra)
'    tra.nLength = Len(tra)
'    CreatePipe hStdInPipeRead, hstdInPipeWrite, pa, 0
'    WriteFile hstdInPipeWrite, ByVal Chr$(0), 1, 0, ByVal 0
'    If CreatePipe(hRead, hWrite, pa, 0) <> 0 Then
'        sui.cb = Len(sui)
'        GetStartupInfo sui
'        sui.hStdInput = hStdInPipeRead
'        sui.hStdOutput = hWrite
'        sui.hStdError = hWrite
'        sui.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
'        sui.wShowWindow = SW_HIDE
'        If CreateProcess(vbNullString, cmdLine, pra, tra, 1, 0, ByVal 0&, vbNullString, sui, pi) <> 0 Then
'            'SetWindowText objTarget.hwnd, ""
'            'If Not GetThreadDesktop(pi.hThread) = 0 Then
'            'insert waitforsingleobject loop here...
'            Dim rtret As Long
'            Do
'                rtret = WaitForSingleObject(pi.hProcess, 100)
'                If rtret <> WAIT_TIMEOUT Then
'                    Exit Do
'                End If
'                DoEvents
'
'            Loop
'                Do
'                    lpBuffer = Space(1024)
'                    If ReadFile(hRead, ByVal lpBuffer, 1023, bRead, ByVal 0&) Then
'                '    SendMessage objTarget.hwnd, EM_SETSEL, -1, 0
'                '    SendMessage objTarget.hwnd, EM_REPLACESEL, False, lpBuffer(0)
'                    wholestr = wholestr & Replace$(Trim$(lpBuffer), vbNullChar, "")
'                    If bRead < 1023 Then
'                        Exit Do
'                    End If
'                    DoEvents
'                    Else
'                    Exit Do
'                    End If
'
'                Loop
''            Else
''                wholestr = "Error executing """ & cmdLine & """."
''            End If
'            CloseHandle hWrite
'            CloseHandle hRead
'            CloseHandle hStdInPipeRead
'            CloseHandle hstdInPipeWrite
'        End If
'    End If
'    Redirect = wholestr
'End Function
'
Private Sub IAsyncProcess_ExecAsync(ByVal Vdata As Variant)

End Sub
