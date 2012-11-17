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
Public Event StreamFileIOCompleted(Buffer() As Byte)
Private mAsync As Boolean
Private moverlapped As OVERLAPPED
Private ReadBuffer() As Byte
Private WriteBuffer() As Byte
'Private moverlapped As OVERLAPPED
Private mvaropenmode As DesiredAccessFlags
Private mvarhStream As Long
Private mvaropenedfile As String
Private mvarStringReadMode As StringReadMode

Private Function GetReadMode(ByVal ParamInput As StringReadMode)
'called by functions that use a String read mode parameter. contains the logic that determines which mode to use.
    If ParamInput = StrRead_Default Then
        If mvarStringReadMode = StrRead_Default Then
            GetReadMode = StrRead_unicode
        Else
            GetReadMode = mvarStringReadMode
        End If
    
    
    Else
        GetReadMode = ParamInput
    
    End If



End Function
Public Property Let Stringmode(ByVal Vdata As StringReadMode)
    mvarStringReadMode = Vdata
End Property
Public Property Get Stringmode() As StringReadMode
    Stringmode = mvarStringReadMode
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
    OpenForWrite = ((mvaropenmode And GENERIC_WRITE) = GENERIC_WRITE) Or mvaropenmode = GENERIC_ALL
End Function
Private Function OpenForRead() As Boolean
    OpenForRead = ((mvaropenmode And GENERIC_READ) = GENERIC_READ) Or mvaropenmode = GENERIC_ALL
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
    mvaropenmode = dwDesiredAccess
    RaiseEvent StreamOpen(mvaropenedfile)


End Sub
Friend Sub SetStream(ByVal Handle As Long, ByVal OpenMode As DesiredAccessFlags)
    mvarhStream = Handle
    mvaropenmode = OpenMode
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
    Dim Currpos As Double
    Currpos = Me.GetFilePos
    Me.SetFilePos 0, STREAM_END
    
    EOF = (Currpos >= Me.GetFilePos)
    'move back to start...
    Debug.Print Currpos
    Me.SetFilePos Currpos, STREAM_BEGIN

End Function
Public Property Get AtEndOfStream() As Boolean
    AtEndOfStream = Me.EOF
End Property

Public Function ReadPicture() As StdPicture
    Set ReadPicture = ModStreamFunc.ReadPicture(Me)
End Function
Public Sub WritePicture(writeme As StdPicture)
    ModStreamFunc.WritePicture Me, writeme
    
    
End Sub

Public Function readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long, Optional ByRef hitEOF As Boolean) As Byte()
    Dim retval As Long
    
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
    
        retval = ReadFile(mvarhStream, ReadBuffer(0), NumBytes, Numread, ByVal 0&)
        If retval = 0 Then
            RaiseAPIError Err.LastDllError, "FileStream::ReadBytes"
        End If
        If NumBytes > Numread Then
            'we have hit the end of the file/stream...
            hitEOF = True
        End If
 '   End If
    readbytes = ReadBuffer


End Function


Public Sub WriteBytes(Writem() As Byte)
    Dim retval As Long
    Dim WriteBuffer() As Byte
    
    Dim numwritten As Long
    If Not OpenForWrite Then
        Err.Raise 54, "FileStream::WriteBytes", "Cannot Write to Input stream"
    End If
    WriteBuffer = Writem
  
        retval = WriteFile(mvarhStream, WriteBuffer(LBound(WriteBuffer)), UBound(WriteBuffer) - LBound(WriteBuffer) + 1, numwritten, ByVal &H0)
  
    If retval = 0 Then
        RaiseAPIError Err.LastDllError, "FileStream::WriteBytes"
    
    End If

End Sub


Public Sub WriteObject(ObjWrite As Object)
  ModStreamFunc.WriteObject Me, ObjWrite
End Sub
Public Function ReadObject() As Object
   Set ReadObject = ModStreamFunc.ReadObject(Me)
End Function

Public Sub WriteInteger(ByVal IntWrite As Integer)
    ModStreamFunc.WriteInteger Me, IntWrite
End Sub
Public Function ReadInteger() As Integer
  ReadInteger = ModStreamFunc.ReadInteger(Me)
End Function

Public Sub WriteLong(ByVal LngWrite As Long)
    ModStreamFunc.WriteLong Me, LngWrite
End Sub
Public Function ReadLong() As Long
ReadLong = ModStreamFunc.ReadLong(Me)
    
End Function



Public Sub WriteSingle(ByVal SngWrite As Long)
    ModStreamFunc.WriteSingle Me, SngWrite
End Sub
Public Sub WriteDouble(ByVal DblWrite As Double)
    ModStreamFunc.WriteDouble Me, DblWrite
End Sub
Public Function ReadDouble() As Double
    ReadDouble = ModStreamFunc.ReadDouble(Me)
End Function

Public Function ReadSingle() As Single
   ReadSingle = ModStreamFunc.ReadSingle(Me)
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
'Public Function ReadLine(Optional ByVal Stringmode As StringReadMode = StrRead_unicode) As String
'    ReadLine = ReadUntil(vbCrLf, True, Stringmode, vbTextCompare)
'End Function
Public Function ReadUntil(ByVal SearchFor As String, Optional ByVal SeekAfterToken As Boolean = True, Optional ByVal Stringmode As StringReadMode = StrRead_unicode, Optional ByVal Comparemethod As VbCompareMethod = vbTextCompare, Optional ByVal Chunksize As Integer = 64) As String

    'Bugfix: added missing Routine Exits (OOPS!)
    
    Dim buildstr As String, chunkread As Long, hitEOF As Boolean
    Dim currentPosition As Long
    Dim readValue() As Byte, posfound As Long
    Dim convert As String
    If Chunksize <= 0 Then Chunksize = 1 'minimum.
    currentPosition = Me.GetFilePos
    Do
        'readValue = ReadBytes(ChunkSize, chunkread, hiteof)
        'convert = readValue
        'convert = StrConv(convert, vbUnicode)
        convert = ReadString(Chunksize, Stringmode, hitEOF)
        buildstr = buildstr & convert
        
        If hitEOF Then
            'end of file...
            ReadUntil = buildstr
            Exit Function
        
        
        End If
        'look for the string....
        posfound = InStr(1, buildstr, SearchFor, Comparemethod)
        If posfound > 0 Or Me.GetFilePos() > Me.Size Then
            'it was found...
            
            'what we want to do here is get the length of the data before the found string and  then seek to it.
            'oh wait... we DO have that value...heh
            posfound = posfound - 1
            buildstr = Mid$(buildstr, 1, posfound)
            If SeekAfterToken Then posfound = posfound + Len(SearchFor)
            Me.SeekTo currentPosition + posfound, STREAM_BEGIN
            ReadUntil = buildstr
        
            Exit Function
        End If
    
    Loop
    
    


End Function
Public Function ReadStringAuto(ByVal CharCount As Long) As String
    ReadStringAuto = ModStreamFunc.ReadStringAuto(Me, CharCount)
End Function
Public Function ReadString(ByVal StrAmount As Long, Optional ByVal Stringmode As StringReadMode = StrRead_Default, Optional ByRef hitEOF As Boolean)
Dim readcast() As Byte, retval As Long
'Stringmode = GetReadMode(Stringmode)
ReadString = ModStreamFunc.ReadString(Me, StrAmount, Stringmode, hitEOF)

End Function
Public Function ReadLine(Optional ByVal Stringmode As StringReadMode = StrRead_Default)

    'read one char at a time until we find a cr or an lf.
    Dim buildit As cStringBuilder
    Set buildit = New cStringBuilder
    Dim CurrChar As String
    Dim exitflag As Boolean
    Dim GetPos As Long
    Dim waseol As Boolean
    Do Until exitflag = True
        CurrChar = ReadString(1, Stringmode, exitflag)
        waseol = CurrChar = vbCr Or CurrChar = vbLf
        If CurrChar = vbCr And Not exitflag Then
           'get current position
           GetPos = Me.GetFilePos
           '
           Dim checknext As String
           checknext = ReadString(1, Stringmode, exitflag)
           If checknext <> vbLf Then SetFilePos GetPos, STREAM_BEGIN
           
        End If
        exitflag = waseol
        If Not waseol Then buildit.Append CurrChar
    
    Loop
    ReadLine = buildit.ToString()



End Function
Public Sub WriteString(ByVal StrWrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
    Dim cast() As Byte
    If StrWrite = "" Then Exit Sub
'    If StringMode = strread_unicode Then
'        StrWrite = StrConv(StrWrite, vbUnicode)
'    End If

    
    'Stringmode = GetReadMode(Stringmode)
    ModStreamFunc.WriteString Me, StrWrite, Stringmode
    



End Sub
Public Sub WriteLine(ByVal StrWrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
   StrWrite = StrWrite & vbCrLf
   ModStreamFunc.WriteString Me, StrWrite, Stringmode
   
End Sub


Public Function GetFilePos() As Double

    Dim LoWord As Long, HiWord As Long
    
    '    #define GetVLFilePointer(hFile, lpPositionHigh) \
     '       (*lpPositionHigh = 0, \
     '       SetFilePointer(hFile, 0, lpPositionHigh, FILE_CURRENT))
    LoWord = SetFilePointer(mvarhStream, 0, HiWord, STREAM_CURRENT)
    If LoWord <> 0 Then
        GetFilePos = LI2DBL(LoWord, HiWord)
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
    Dim lopart As Long, hipart As Long
    DBL2LI NewFilePos, lopart, hipart
    SetFilePointer mvarhStream, lopart, hipart, SeekMode
    
    
End Sub
Private Sub ChkHandle()
    If mvarhStream = 0 Then
        Err.Raise 54, "BCFile.FileStream", "Operation not valid on uninitialized FileStream."
    End If



End Sub
Public Property Get hstream() As Long
    hstream = mvarhStream
End Property
Friend Property Let hstream(ByVal Vdata As Long)
    If mvarhStream <> 0 Then
        CloseStream
    
    End If
    mvarhStream = Vdata
End Property
Public Sub Flush()
    FlushFileBuffers mvarhStream
End Sub
Public Function ReadAllStr(Optional ByVal Stringmode As StringReadMode = StrRead_Default) As String
    'read it all .
    Dim allread As String
    allread = ReadString(GetStreamSize, Stringmode)
  
        ReadAllStr = allread
   
End Function

Private Sub Class_Initialize()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
mvarStringReadMode = StrRead_Default
End Sub

Private Sub Class_Terminate()
    CloseStream
    MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub


Private Sub IAsyncProcess_ExecAsync(ByVal Vdata As Variant)
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



Private Function IInputStream_ReadAllStr(Optional ByVal Stringmode As StringReadMode = StrRead_Default) As String
IInputStream_ReadAllStr = Me.ReadAllStr(Stringmode)
End Function




Private Function IInputStream_readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long) As Byte()
IInputStream_readbytes = Me.readbytes(NumBytes, Numread)
End Function

Private Function IInputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Double
    IInputStream_SeekTo = Me.SeekTo(ToPos, SeekMode)
End Function

Private Function IInputStream_Size() As Double
    IInputStream_Size = Me.Size
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

Public Property Get Size() As Double

    Dim Currpos As Double
    Currpos = Me.GetFilePos
    Me.SetFilePos 0, STREAM_END
    
    Size = Me.GetFilePos
    'move back to start...
    Me.SetFilePos Currpos, STREAM_BEGIN

    
End Property
Public Property Let Size(ByVal Vdata As Double)
    Dim Currpos As Long
    Currpos = Me.GetFilePos
    Me.SetFilePos Vdata, STREAM_BEGIN
    Me.SeekTo Currpos, STREAM_BEGIN




End Property
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

Private Sub IOutputStream_WriteString(Stringwrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
Me.WriteString Stringwrite, Stringmode
End Sub
Public Sub WriteStream(ReadFrom As IInputStream, Optional ByVal Chunksize As Long = 64& * 1024&)


    CopyEntireStream ReadFrom, Me, Chunksize
End Sub
