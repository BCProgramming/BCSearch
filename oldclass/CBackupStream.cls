VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBackupStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function BackupRead Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef mvarcontext As Any) As Long
Private Declare Function BackupSeek Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef mvarcontext As Long) As Long
Private Declare Function BackupWrite Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef mvarcontext As Long) As Long
Private mvarcontext As Long, mvaropenedfile As String
Private mvarhStream As Long
Public Event StreamOpen(OnPath As String)
Public Event StreamClose(ClosePath As String)
Public Event StreamFileIOCompleted(buffer() As Byte)


Friend Sub OpenStream(OpenPath As String, ByVal dwDesiredAccess As DesiredAccessFlags, ByVal dwShareMode As FileShareModeConstants, ByVal dwCreationDisposition As CreationDispositionConstants)

    'open the stream.
    'If Not mvarOpenedFile Is Nothing Then CloseStream
    If mvarhStream <> 0 Then CloseStream
    mvaropenedfile = OpenPath
    
    mvarhStream = CreateFile(OpenPath, dwDesiredAccess, dwShareMode, ByVal 0&, dwCreationDisposition, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If mvarhStream <= 0 Then
        'Set mvarOpenedFile = Nothing
        mvarhStream = 0
        mvaropenedfile = ""
        
        RaiseAPIError Err.LastDllError, "FileStream::OpenStream"
    End If
    mvarOpenMode = dwDesiredAccess
    
   ' Call BackupRead(mvarhStream, ByVal 0, 0, 0, 0, 1, ByVal mvarcontext)
    
    
    RaiseEvent StreamOpen(mvaropenedfile)


End Sub
Public Sub CloseStream()
    If mvarcontext <> 0 And mvarstream <> 0 Then
        BackupRead mvarhStream, ByVal 0, 0, 0, 1, 1, mvarcontext
    ElseIf mvarstream <> 0 Then
        CloseHandle mvarhStream
    End If
    mvarhStream = 0
    mvarcontext = 0
    RaiseEvent StreamClose(mvaropenedfile)
End Sub
Public Function ReadBytes(ByVal NumBytes As Long, Optional ByRef numread As Long, Optional ByRef HitEOF As Boolean) As Byte()
    Dim retVal As Long
    
    ReDim ReadBuffer(0 To NumBytes - 1)
    If Not OpenForRead Then
        Err.Raise 54, "FileStream::ReadBytes", "Cannot Read from a OutputStream."
    
    End If
'    If Async Then
'        moverlapped.offset = 0
'        moverlapped.OffsetHigh = 0
'        moverlapped.hEvent = AddAsyncStream(Me)
'        mAsyncState = ASync_Read
'
'         retVal = ReadFileEx(mvarHStream, ByVal VarPtr(ReadBuffer(0)), NumBytes, ByVal VarPtr(moverlapped), AddressOf FileIOCompletionRoutine)
'         Debug.Print "current error:" & GetAPIErrStr(Err.LastDllError)
'         ReadBytes = ReadBuffer
'        Exit Function
'
'    Else
    
        retVal = BackupRead(mvarhStream, ReadBuffer(0), NumBytes, numread, ByVal 0&, 1, ByVal mvarcontext)
        If retVal = 0 Then
            RaiseAPIError Err.LastDllError, "CBackupStream::ReadBytes"
        End If
        If NumBytes > numread Then
            'we have hit the end of the file/stream...
            HitEOF = True
        End If
 '   End If
 '   ReadBytes = ReadBuffer


End Function
Public Function ReadString(ByVal StrAmount As Long, Optional ByVal StringMode As StringReadMode = StrRead_Unicode, Optional ByRef HitEOF As Boolean) As String
Dim readcast As String, retVal As Long
readcast = ReadBytes(StrAmount, retVal, HitEOF)
'If StringMode = StrRead_Unicode Then
';    readcast = StrConv(readcast, vbUnicode)
'End If
ReadString = readcast






End Function

