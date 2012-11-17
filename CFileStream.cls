VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDataStream




Private Sub EnsureOutput()
    If Not OpenForWrite Then
        Err.Raise 13, "FileStream::EnsureOutput", "FileStream must be opened for output for this operation."
    End If


End Sub
Private Sub EnsureInput()
If Not OpenForRead Then
        Err.Raise 13, "FileStream::EnsureOutput", "FileStream must be opened for Input for this operation."
    End If
End Sub


Private Function OpenForWrite() As Boolean
    OpenForWrite = ((mvarOpenMode And GENERIC_WRITE) = GENERIC_WRITE) Or mvarOpenMode = GENERIC_ALL
End Function
Private Function OpenForRead() As Boolean
    OpenForRead = ((mvarOpenMode And GENERIC_READ) = GENERIC_READ) Or mvarOpenMode = GENERIC_ALL
End Function


Friend Sub OpenStream(OpenPath As String, ByVal dwDesiredAccess As DesiredAccessFlags, ByVal dwShareMode As FileShareModeConstants, ByVal dwCreationDisposition As CreationDispositionConstants, ByVal dwFlagsAndAttributes As FileAttributeConstants)

    'open the stream.
    'If Not mvarOpenedFile Is Nothing Then CloseStream
    If mvarhStream <> 0 Then CloseStream
    mvaropenedfile = OpenPath
    
    mvarhStream = CreateFile(OpenPath, dwDesiredAccess, dwShareMode, ByVal 0&, dwCreationDisposition, dwFlagsAndAttributes, 0)
    If mvarhStream <= 0 Then
        'Set mvarOpenedFile = Nothing
        mvarhStream = 0
        mvaropenedfile = ""
        RaiseAPIError Err.LastDllError, "FileStream::OpenStream"
    End If
    mvarOpenMode = dwDesiredAccess
    RaiseEvent StreamOpen(mvaropenedfile)


End Sub
Friend Sub SetStream(ByVal Handle As Long, ByVal OpenMode As DesiredAccessFlags)
    mvarhStream = Handle
    mvarOpenMode = OpenMode
    mvaropenedfile = ""
    
End Sub
Public Sub CloseStream()
    If mvarhStream <> 0 Then
        CloseHandle mvarhStream
        RaiseEvent StreamClose(mvaropenedfile)
        mvarhStream = 0
        mvaropenedfile = ""
    End If
End Sub
Public Function EOF() As Boolean
    'Eof = (GetFilePos =
    Dim CurrPos As Double
    CurrPos = Me.GetFilePos
    Me.SetFilePos 0, FILE_END
    
    EOF = (CurrPos >= Me.GetFilePos)
    'move back to start...
    Debug.Print CurrPos
    Me.SetFilePos CurrPos, FILE_BEGIN

End Function
Public Property Get AtEndOfStream() As Boolean
    AtEndOfStream = Me.EOF
End Property
Public Function ReadBytes(ByVal NumBytes As Long, Optional ByRef Numread As Long, Optional ByRef HitEOF As Boolean) As Byte()
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
    
        retVal = ReadFile(mvarhStream, ReadBuffer(0), NumBytes, Numread, ByVal 0&)
        If retVal = 0 Then
            RaiseAPIError Err.LastDllError, "FileStream::ReadBytes"
        End If
        If NumBytes > Numread Then
            'we have hit the end of the file/stream...
            HitEOF = True
        End If
 '   End If
    ReadBytes = ReadBuffer


End Function


Public Sub WriteBytes(Writem() As Byte)
    Dim retVal As Long
    Dim WriteBuffer() As Byte
    
    Dim numwritten As Long
    If Not OpenForWrite Then
        Err.Raise 54, "FileStream::WriteBytes", "Cannot Write to Input stream"
    End If
    WriteBuffer = Writem
    retVal = WriteFile(mvarhStream, WriteBuffer(LBound(WriteBuffer)), UBound(WriteBuffer) - LBound(WriteBuffer) + 1, numwritten, ByVal &H0)

    If retVal = 0 Then
        RaiseAPIError Err.LastDllError, "FileStream::WriteBytes"
    
    End If

End Sub

Friend Function SeekTo(ByVal NewPosition As Double, Optional ByVal SeekMode As FileSeekBase) As Long

    SeekTo = Me.GetFilePos

    SetFilePos NewPosition, SeekMode


End Function
Friend Sub SetFilePos(ByVal NewFilePos As Double, Optional ByVal SeekMode As FileSeekBase)
    Dim LoPart As Long, hipart As Long
    DBL2LI NewFilePos, LoPart, hipart
    SetFilePointer mvarhStream, LoPart, hipart, SeekMode
    
    
End Sub

Private Function IDataStream_GetSeekPos() As Variant
'
End Function

Private Function IDataStream_ReadBytes(ByVal NumBytes As Long, Optional Numread As Long) As Byte()
'
IDataStream_ReadBytes = ReadBytes(NumBytes, Numread)
End Function

Private Function IDataStream_SeekTo(ByVal NewPosition As Variant, SeekBase As StreamSeekBase) As Variant
'
End Function

Private Property Get IDataStream_Supported() As ESupportedStreamOperations
'
End Property

Private Sub IDataStream_WriteBytes(BytesWrite() As Byte)
'
End Sub
