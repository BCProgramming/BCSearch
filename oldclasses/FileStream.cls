VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'TODO: Make "FileStream" a "base" reader; that is, it will implement ReadBytes() and WriteBytes() and nothing else.
'Basically, IOutputStream  have WriteBytes; IInputStream will have "readBytes" and I will have helper classes, such as "BinaryStream" that handle the details (whereas here I have ReadLong Implemented and so forth".
'This will allow more flexibility with the Memory Streams and so forth- a memory stream could be accessed as a text file using, say, a "TextStream" Object; or as a binary file with BinaryStream. It also allows the logic to be centralized for such things as reading text.
'right now both this class and the memory stream class implement a




Implements IOutputStream
Implements IInputStream
Implements IAsyncProcess
'Generic FileStream Class.
Public Enum Async_State_Constants
    Async_NoAsync 'no async operations.
    ASync_Read = 2
    ASync_Write = 4
End Enum
Private mAsyncState As Async_State_Constants
Public Event StreamOpen(OnPath As String)
Public Event StreamClose(ClosePath As String)
Public Event StreamFileIOCompleted(buffer() As Byte)
Private ReadBuffer() As Byte
Private WriteBuffer() As Byte
Private moverlapped As OVERLAPPED
Private mvarOpenMode As DesiredAccessFlags
Private mvarhStream As Long
Private mvaropenedfile As String
Private mvarStringReadMode As StringReadMode

Private Function GetReadMode(ByVal ParamInput As StringReadMode)
'called by functions that use a String read mode parameter. contains the logic that determines which mode to use.
    If ParamInput = StrRead_Default Then
        If mvarStringReadMode = StrRead_Default Then
            GetReadMode = strread_unicode
        Else
            GetReadMode = mvarStringReadMode
        End If
    
    
    Else
        GetReadMode = ParamInput
    
    End If



End Function
Public Property Let StringMode(ByVal vdata As StringReadMode)
    mvarStringReadMode = vdata
End Property
Public Property Get StringMode() As StringReadMode
    StringMode = mvarStringReadMode
End Property

'Error raising routines for output only or input only methods...
Friend Sub FileIOCompleted(bufferRaise() As Byte)
    'called by module when it recieves info on the async call.
    RaiseEvent StreamFileIOCompleted(bufferRaise())
End Sub

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


Public Sub WriteObject(ObjWrite As Object)
    'uses a property bag, writes the object to the property bag, copies the byte array and saves the length and the data to the file.
    Dim PropBag As PropertyBag
    Dim BytesWrite() As Byte
    Dim bytelength As Long
    Set PropBag = New PropertyBag
    On Error GoTo Nopersist
    PropBag.WriteProperty "OBJECT", ObjWrite, Nothing
    On Error GoTo 0
    BytesWrite = PropBag.Contents
    bytelength = UBound(BytesWrite) - LBound(BytesWrite) + 1
     
    Me.WriteLong bytelength
    Me.WriteBytes BytesWrite
    'success.
    Exit Sub
    
Nopersist:
    Err.Raise 9, "FileStream::WriteObject", "Object does not implement proper persistence interfaces."
    


End Sub
Public Function ReadObject() As Object
'reads an object from the file.
Dim PropBag As PropertyBag
Dim BytesRead() As Byte
Dim ByteLen As Long


'first read in length...
    ByteLen = Me.ReadLong
    'ReDim bytesread(1 To bytelen)
    BytesRead = Me.ReadBytes(ByteLen)
    Set PropBag = New PropertyBag
    PropBag.Contents = BytesRead
    Set ReadObject = PropBag.ReadProperty("OBJECT", Nothing)

End Function

Public Sub WriteInteger(ByVal IntWrite As Integer)
    Dim CopyTo(1 To Len(IntWrite)) As Byte
    CopyMemory CopyTo(1), IntWrite, Len(IntWrite)

    WriteBytes CopyTo
End Sub
Public Function ReadInteger() As Integer
    Dim intread As Integer
    Dim ReadIt() As Byte
    ReadIt = ReadBytes(Len(intread))
    CopyMemory intread, ReadIt(1), UBound(ReadIt)
    ReadInteger = intread
    
End Function

Public Sub WriteLong(ByVal LngWrite As Long)
    
    Dim CopyTo(1 To Len(LngWrite)) As Byte
    CopyMemory CopyTo(1), LngWrite, Len(LngWrite)
    WriteBytes CopyTo
End Sub
Public Function ReadLong() As Long
    Dim lngread As Long
    Dim ReadIt() As Byte
    
    ReadIt = ReadBytes(Len(lngread))
    CopyMemory lngread, ReadIt(LBound(ReadIt)), Len(lngread)
    ReadLong = lngread
    
End Function



Public Sub WriteSingle(ByVal SngWrite As Long)
    Dim CopyTo(1 To Len(SngWrite)) As Byte
    CopyMemory CopyTo(1), SngWrite, Len(SngWrite)
    WriteBytes CopyTo
End Sub
Public Sub WriteDouble(ByVal DblWrite As Double)
    Dim CopyTo(1 To Len(DblWrite)) As Byte
    CopyMemory CopyTo(1), DblWrite, Len(DblWrite)
    WriteBytes CopyTo
End Sub
Public Function ReadDouble() As Double
    Dim retVal As Double
    Dim CopyTo() As Byte
    CopyTo = ReadBytes(Len(retVal))
    CopyMemory retVal, CopyTo(LBound(CopyTo)), Len(retVal)
    ReadDouble = retVal
End Function

Public Function ReadSingle() As Single
    Dim retVal As Single
    Dim CopyTo() As Byte
    CopyTo = ReadBytes(Len(retVal))
    CopyMemory retVal, CopyTo(LBound(CopyTo)), Len(retVal)
    ReadSingle = retVal
End Function

'Public Function ReadLines(ByVal NumLines As Long) As String
'    'reads the specified number of lines from the file.
'    Dim currcount As Long
'    Dim ReadBytes As Long
'    Dim StrEntire As String, StrValue As String
'
'    StrValue = Me.ReadString(1)
'
'    StrEntire = StrEntire & StrValue
'    If StrValue = vbCr Then
'        currcount = currcount + 1
'        If currcount = NumLines Then
'            ReadLines = StrEntire & vbLf
'        End If
'    End If
'
'End Function
Public Function ReadLine(Optional ByVal StringMode As StringReadMode = strread_unicode) As String
    ReadLine = ReadUntil(vbCrLf, True, StringMode, vbTextCompare)
End Function
Public Function ReadUntil(ByVal SearchFor As String, Optional ByVal SeekAfterToken As Boolean = True, Optional ByVal StringMode As StringReadMode = strread_unicode, Optional ByVal Comparemethod As VbCompareMethod = vbTextCompare, Optional ByVal chunksize As Integer = 64) As String

    
    
    Dim buildstr As String, chunkread As Long, HitEOF As Boolean
    Dim currentPosition As Long
    Dim readValue() As Byte, posfound As Long
    Dim convert As String
    If chunksize <= 0 Then chunksize = 1 'minimum.
    currentPosition = Me.GetFilePos
    Do
        'readValue = ReadBytes(ChunkSize, chunkread, hiteof)
        'convert = readValue
        'convert = StrConv(convert, vbUnicode)
        convert = ReadString(chunksize, StringMode, HitEOF)
        buildstr = buildstr & convert
        
        If HitEOF Then
            'end of file...
            ReadUntil = buildstr
            
        
        
        End If
        'look for the string....
        posfound = InStr(1, buildstr, SearchFor, Comparemethod)
        If posfound > 0 Then
            'it was found...
            
            'what we want to do here is get the length of the data before the found string and  then seek to it.
            'oh wait... we DO have that value...heh
            posfound = posfound - 1
            buildstr = Mid$(buildstr, 1, posfound)
            If SeekAfterToken Then posfound = posfound + Len(SearchFor)
            Me.SeekTo currentPosition + posfound, FILE_BEGIN
            ReadUntil = buildstr
        
            Exit Function
        End If
    
    Loop
    
    


End Function
Public Function ReadString(ByVal StrAmount As Long, Optional ByVal StringMode As StringReadMode = StrRead_Default, Optional ByRef HitEOF As Boolean)
Dim readcast() As Byte, retVal As Long
StringMode = GetReadMode(StringMode)
readcast = ReadBytes(StrAmount * IIf(StringMode = strread_unicode, 2, 1), retVal, HitEOF)
If StringMode = StrRead_ANSI Then
    ReadString = StrConv(readcast, vbUnicode)
Else
    ReadString = readcast
End If





End Function
Public Sub WriteString(ByVal StrWrite As String, Optional ByVal StringMode As StringReadMode = StrRead_Default)
    Dim cast() As Byte
    If StrWrite = "" Then Exit Sub
'    If StringMode = strread_unicode Then
'        StrWrite = StrConv(StrWrite, vbUnicode)
'    End If

    
    StringMode = GetReadMode(StringMode)
    If StringMode = StrRead_ANSI Then
        StrWrite = StrConv(StrWrite, vbFromUnicode)
    End If
    cast = StrWrite
    
    WriteBytes cast
    Erase cast
    



End Sub


Public Function GetFilePos() As Double

    Dim loWord As Long, hiWord As Long
    
    '    #define GetVLFilePointer(hFile, lpPositionHigh) \
     '       (*lpPositionHigh = 0, \
     '       SetFilePointer(hFile, 0, lpPositionHigh, FILE_CURRENT))
    loWord = SetFilePointer(mvarhStream, 0, hiWord, FILE_CURRENT)
    If loWord > 0 Then
        GetFilePos = LI2DBL(loWord, hiWord)
    Else
        Debug.Print GetAPIErrStr(Err.LastDllError)
        GetFilePos = 0
    End If

End Function
Public Function SeekTo(ByVal NewPosition As Double, Optional ByVal SeekMode As FileSeekBase) As Long

    SeekTo = Me.GetFilePos

    SetFilePos NewPosition, SeekMode


End Function
Public Sub SetFilePos(ByVal NewFilePos As Double, Optional ByVal SeekMode As FileSeekBase)
    Dim LoPart As Long, hipart As Long
    DBL2LI NewFilePos, LoPart, hipart
    SetFilePointer mvarhStream, LoPart, hipart, SeekMode
    
    
End Sub
Private Sub ChkHandle()
    If mvarhStream = 0 Then
        Err.Raise 54, "BCFile.FileStream", "Operation not valid on uninitialized FileStream."
    End If



End Sub
Public Property Get hstream() As Long
    hstream = mvarhStream
End Property
Public Sub Flush()
    FlushFileBuffers mvarhStream
End Sub
Public Function ReadAllStr(Optional ByVal StringMode As StringReadMode = StrRead_Default) As String
    'read it all .
    Dim allread As String
    allread = ReadString(GetStreamSize, StringMode)
  
        ReadAllStr = allread
   
End Function

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
mvarStringReadMode = StrRead_Default
End Sub

Private Sub Class_Terminate()
    CloseStream
    mTotalObjectCount = mTotalObjectCount - 1
End Sub


Private Sub IAsyncProcess_ExecAsync(ByVal vdata As Variant)
If mAsyncState = ASync_Read Then
    FileIOCompleted ReadBuffer()
ElseIf mAsyncState = ASync_Write Then
    FileIOCompleted WriteBuffer()
End If
End Sub

Private Sub IInputStream_CloseStream()
    Me.CloseStream
End Sub

Private Function IInputStream_EOF() As Boolean
'
IInputStream_EOF = Me.EOF
End Function

Private Function IInputStream_GetPos() As Double
IInputStream_GetPos = Me.GetFilePos
End Function



Private Function IInputStream_ReadAllStr(Optional ByVal StringMode As StringReadMode = StrRead_Default) As String
IInputStream_ReadAllStr = Me.ReadAllStr(StringMode)
End Function




Private Function IInputStream_ReadBytes(ByVal NumBytes As Long, Optional Numread As Long) As Byte()
IInputStream_ReadBytes = Me.ReadBytes(NumBytes, Numread)
End Function

Private Function IInputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Double
    IInputStream_SeekTo = Me.SeekTo(ToPos, SeekMode)
End Function

Private Function IInputStream_size() As Double
    IInputStream_size = Me.size
End Function

Private Function IInputStream_Valid() As Boolean
IInputStream_Valid = OpenForRead
End Function

Private Sub IOutputStream_CloseStream()
Me.CloseStream
End Sub

Private Function IOutputStream_GetPos() As Double
'
IOutputStream_GetPos = Me.GetFilePos()
End Function



Private Function IOutputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Variant
IOutputStream_SeekTo = Me.SeekTo(ToPos, SeekMode)
End Function

Private Function IOutputStream_Valid() As Boolean
IOutputStream_Valid = OpenForWrite
End Function

Public Function size() As Double

    Dim CurrPos As Double
    CurrPos = Me.GetFilePos
    Me.SetFilePos 0, FILE_END
    
    size = Me.GetFilePos
    'move back to start...
    Me.SetFilePos CurrPos, FILE_BEGIN

    
End Function

Friend Function GetStreamSize() As Double
'Will fail if the stream is, for example- a pipe.
    Dim hf As Long
    Dim fslo As Long, fshi As Long
    Dim sizeo As Double
    hf = mvarhStream
    fslo = GetFileSize(hf, fshi)
    sizeo = LI2DBL(fslo, fshi)

GetStreamSize = sizeo

End Function

Private Function IOutputStream_WriteBytes(Bytes() As Byte) As Boolean
WriteBytes Bytes()
End Function

Private Sub IOutputStream_WriteString(Stringwrite As String, Optional ByVal StringMode As StringReadMode = StrRead_Default)
Me.WriteString Stringwrite, StringMode
End Sub
Public Sub WriteStream(ReadFrom As FileStream, Optional ByVal chunksize As Long = 64& * 1024&)


    CopyEntireStream ReadFrom, Me, chunksize
End Sub
