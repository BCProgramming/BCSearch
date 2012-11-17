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
Option Explicit


Implements IInputStream

Private Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Private Declare Function BackupWrite Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Private Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef lpContext As Long) As Long
Private mvarcontext As Long, mvaropenedfile As String
Private mvarhStream As Long, mvaropenmode As DesiredAccessFlags, mCurrPos As Double
Dim hitEOF As Boolean
Public Event StreamOpen(OnPath As String)
Public Event StreamClose(ClosePath As String)
Public Event StreamFileIOCompleted(Buffer() As Byte)

Private Function OpenForWrite() As Boolean
    OpenForWrite = ((mvaropenmode And GENERIC_WRITE) = GENERIC_WRITE) Or mvaropenmode = GENERIC_ALL
End Function
Private Function OpenForRead() As Boolean
    OpenForRead = ((mvaropenmode And GENERIC_READ) = GENERIC_READ) Or mvaropenmode = GENERIC_ALL
End Function

Friend Sub OpenStream(OpenPath As String, ByVal dwDesiredAccess As DesiredAccessFlags, ByVal dwShareMode As FileShareModeConstants, ByVal dwCreationDisposition As CreationDispositionConstants)

    'open the stream.
    'If Not mvarOpenedFile Is Nothing Then CloseStream
    Dim badata() As Byte
    ReDim badata(1)
    If mvarhStream <> 0 Then CloseStream
    mvaropenedfile = OpenPath
    
    mvarhStream = CreateFile(OpenPath, dwDesiredAccess, dwShareMode, ByVal 0&, dwCreationDisposition, FILE_FLAG_BACKUP_SEMANTICS, 0)
    If mvarhStream <= 0 Then
        'Set mvarOpenedFile = Nothing
        mvarhStream = 0
        mvaropenedfile = ""
        
        RaiseAPIError Err.LastDllError, "FileStream::OpenStream"
    End If
    mvaropenmode = dwDesiredAccess
    Dim dwbytestoread As Long, dwbytesread As Long, ret As Long
    mvarcontext = 0
    
    'Call BackupRead(mvarhStream, ByVal 0, 0, Numread, 0, 1, ByVal mvarcontext)
    'ret = BackupRead(mvarhStream, badata(0), dwbytestoread, dwbytesread, False, False, ByVal mvarcontext)
    mCurrPos = 1
    
    RaiseEvent StreamOpen(OpenPath)


End Sub
Public Sub CloseStream()
    If mvarcontext <> 0 And mvarhStream <> 0 Then
         BackupRead mvarhStream, 0, 0, 0, 1, 0, mvarcontext
    End If
    If mvarhStream <> 0 Then
        CloseHandle mvarhStream
    End If
    mvarhStream = 0
    mvarcontext = 0
    RaiseEvent StreamClose(mvaropenedfile)
End Sub
Public Sub WriteBytes(Writem() As Byte)
    Dim retval As Long
    Dim WriteBuffer() As Byte
    
    Dim numwritten As Long
    If Not OpenForWrite Then
        Err.Raise 54, "FileStream::WriteBytes", "Cannot Write to Input stream"
    End If
    WriteBuffer = Writem
    'retval = WriteFile(mvarhStream, WriteBuffer(LBound(WriteBuffer)), UBound(WriteBuffer) - LBound(WriteBuffer) + 1, numwritten, ByVal &H0)
    retval = BackupWrite(mvarhStream, VarPtr(WriteBuffer(0)), UBound(WriteBuffer) + 1, numwritten, 0, 0, mvarcontext)
    mCurrPos = mCurrPos + numwritten
    If retval = 0 Then
        RaiseAPIError Err.LastDllError, "FileStream::WriteBytes"
    
    End If

End Sub
Public Function readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long, Optional ByRef hitEOF As Boolean) As Byte()
    Dim retval As Long
    Dim ReadBuffer() As Byte
    ReDim ReadBuffer(0 To NumBytes - 1)
    'If Not OpenForRead Then
    '    Err.Raise 54, "FileStream::ReadBytes", "Cannot Read from a OutputStream."
    
    'End If
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

    retval = BackupRead(mvarhStream, VarPtr(ReadBuffer(0)), NumBytes, Numread, False, False, mvarcontext)
        mCurrPos = mCurrPos + Numread
        If retval = 0 Then
            RaiseAPIError Err.LastDllError, "CBackupStream::ReadBytes"
        End If
        If NumBytes > Numread Then
            'we have hit the end of the file/stream...
            hitEOF = True
        End If
 '   End If
 '   ReadBytes = ReadBuffer


End Function
Public Function SeekTo(pPosition As Double, Optional ByVal Seekbase As FileSeekBase = STREAM_BEGIN) As Double
    Dim lopart As Long, hipart As Long
    Dim loseeked As Long, hiseeked As Long
    Dim returndbl As Double, ret As Long
    If mvarhStream = 0 Then
        Err.Raise 9, "CBackupStream::SeekTo", "Seek Operation not valid when stream is not open."
    End If
    DBL2LI pPosition, lopart, hipart
    ret = BackupSeek(mvarhStream, lopart, hipart, loseeked, hiseeked, mvarcontext)
    returndbl = LI2DBL(loseeked, hiseeked)
    
    mCurrPos = mCurrPos + returndbl

    SeekTo = mCurrPos
End Function

Private Sub IInputStream_CloseStream()
'
    Me.CloseStream
End Sub

Private Function IInputStream_EOF() As Boolean
'
    IInputStream_EOF = hitEOF
End Function

Private Function IInputStream_GetPos() As Double
'
IInputStream_GetPos = mCurrPos
End Function

Private Function IInputStream_ReadAllStr(Optional ByVal Stringmode As StringReadMode = -1&) As String
'
    'IInputStream_ReadAllStr=modstreamfunc.ReadString(me,
End Function

Private Function IInputStream_readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long) As Byte()
'
End Function

Private Function IInputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Double
'
End Function

Private Function IInputStream_Size() As Double
'
End Function

Private Function IInputStream_Valid() As Boolean
'
End Function
