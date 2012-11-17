VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPipeStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'A stream via a pipe.
'As a "convenience" it also creates the pipe.



Public Enum PipeAccessMode
    PIPE_ACCESS_DUPLEX = &H3
    PIPE_ACCESS_OUTBOUND = &H2
    PIPE_ACCESS_INBOUND = &H1
    PIPE_FLAG_FIRST_PIPE_INSTANCE = &H80000
    PIPE_FLAG_WRITE_THROUGH = FILE_FLAG_WRITE_THROUGH
    
End Enum
Private Const PIPE_TYPE_BYTE As Long = &H0
Private Const PIPE_TYPE_MESSAGE As Long = &H4
Private Const PIPE_READMODE_BYTE As Long = &H0
Private Const PIPE_READMODE_MESSAGE As Long = &H2
Private Const PIPE_WAIT As Long = &H0
Private Const PIPE_NOWAIT As Long = &H1
Private Const PIPE_UNLIMITED_INSTANCES As Long = 255

Public Event WriteComplete(ByVal Cookie As Long)
Public Event ReadComplete(ByVal Cookie As Long)

Private Declare Function CreateNamedPipe Lib "kernel32.dll" Alias "CreateNamedPipeA" (ByVal lpName As String, ByVal dwOpenMode As Long, ByVal dwPipeMode As Long, _
ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, _
ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private mPipeMode As PipeAccessMode
Private mHPipe As Long
Private Sub ClosePipe()
    CloseHandle mHPipe
End Sub
Friend Sub InitializePipe(ByVal PipeName As String, ByVal AccessMode As PipeAccessMode)
    Dim Usesa As SECURITY_ATTRIBUTES
    Usesa.bInheritHandle = 1
    Usesa.lpSecurityDescriptor = 1
    'Usesa.nLength = Len(Usesa)
    If mHPipe <> 0 Then
        ClosePipe
    
    End If
    mPipeMode = AccessMode
    mHPipe = CreateNamedPipe(PipeName, AccessMode, PIPE_TYPE_BYTE, PIPE_UNLIMITED_INSTANCES, 256, 256, 0, Usesa)
    If mHPipe = INVALID_HANDLE_VALUE Then
        'curses...
        RaiseAPIError Err.LastDllError, "CPipeStream::InitializePipe"
    
    
    End If


End Sub
Private Function CanWrite() As Boolean
    CanWrite = ((mPipeMode And PIPE_ACCESS_OUTBOUND) = PIPE_ACCESS_OUTBOUND) Or ((mPipeMode And PIPE_ACCESS_DUPLEX) = PIPE_ACCESS_DUPLEX)
    
End Function
Private Function CanRead() As Boolean
    CanRead = ((mPipeMode And PIPE_ACCESS_INBOUND) = PIPE_ACCESS_INBOUND) Or ((mPipeMode And PIPE_ACCESS_DUPLEX) = PIPE_ACCESS_DUPLEX)
End Function
Public Function WriteBytes(BytesWrite() As Byte, Optional ByVal Async As Boolean) As Long

'returns the Async action cookie.




    'writes the given bytes to mHPipe.
    If CanWrite() Then
    
    
    End If
    
End Function
