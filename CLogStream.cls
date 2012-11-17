VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLogStream: using a IOutputStream, writes log data.
Public Enum LogStreamMessageSeverityConstants
    LSMSC_Information
    MSMSC_Warning
    MSMSC_Error
    MSMSC_Critical
    MSMSC_Fatal


End Enum

'BOOL WINAPI IsWow64Process(
'  __in   HANDLE hProcess,
'  __out  PBOOL Wow64Process
');
'
Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByVal Wow64Process As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long


Private mOut As IOutputStream
Private mStrMode As StringReadMode
Private Function isWow64() As Boolean

    Dim pbool As Long
    Call IsWow64Process(GetCurrentProcess(), pbool)
    isWow64 = pbool <> 0
    


End Function
Friend Sub Create(useOutstream As IOutputStream, Optional ByVal Stringmode As StringReadMode = StrRead_ANSI)

    Set mOut = useOutstream
    mStrMode = Stringmode
    WriteHeader

End Sub

Private Sub WriteHeader()



End Sub
Private Sub WriteString(ByVal StrWrite As String)
mOut.WriteString StrWrite, mStrMode
End Sub
Public Sub Log(ByVal Message As String, ByVal Severity As LogStreamMessageSeverityConstants)

Dim messagemake As String

messagemake = FormatDateTime(Now, vbShortDate) & "@" & FormatDateTime(Now, vbShortTime)



End Sub
