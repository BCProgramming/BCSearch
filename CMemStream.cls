VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CMemStream
'Stream based in memory.

'rewrite this to do reallocations in powers of two rather then multiples of chunksize.

Implements IOutputStream
Implements IInputStream
Private mStreamContents() As Byte
Private mPosition As Long
Private mStreamSize As Long
Private mStreamDataSize As Long
Private mChunksize As Long
Private mvarStringReadMode As StringReadMode

Public Property Let Chunksize(ByVal Vdata As Long)
    mChunksize = Vdata
End Property
Public Property Get Chunksize() As Long
    Chunksize = mChunksize
End Property



Private Sub Class_Initialize()
    mPosition = 1
    mStreamSize = 0
    mChunksize = 32768 \ 4
    Erase mStreamContents
End Sub
Public Function EOF() As Boolean
 EOF = mPosition >= mStreamDataSize
End Function
Public Function GetPosition() As Long
    GetPosition = mPosition
End Function
Public Function readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long = 0, Optional ByRef hitEOF As Boolean) As Byte()
Dim bytesget() As Byte
If mPosition <= 0 Then mPosition = 1
If mPosition + NumBytes > mStreamDataSize Then
    hitEOF = True
    Numread = mStreamDataSize - mPosition
    NumBytes = Numread
Else
    Numread = NumBytes
    hitEOF = True
End If
If NumBytes = 0 Then
    hitEOF = True
    Numread = 0
    Exit Function

End If
ReDim bytesget(0 To NumBytes - 1)
CopyMemory bytesget(0), mStreamContents(mPosition), NumBytes

readbytes = bytesget

'advance the pointer.
mPosition = mPosition + NumBytes


End Function
Public Function ReadDouble() As Double
    Dim retval As Double
    Dim CopyTo() As Byte
    CopyTo = readbytes(Len(retval))
    CopyMemory retval, CopyTo(LBound(CopyTo)), Len(retval)
    ReadDouble = retval
End Function
Public Function ReadInteger() As Integer
    Dim intread As Integer
    Dim ReadIt() As Byte
    ReadIt = readbytes(Len(intread))
    CopyMemory intread, ReadIt(1), UBound(ReadIt)
    ReadInteger = intread
    
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
Public Function ReadLine(Optional ByVal Stringmode As StringReadMode = StrRead_unicode) As String
    ReadLine = ReadUntil(vbCrLf, True, Stringmode, vbTextCompare)
End Function
Public Function ReadLong() As Long
    Dim lngread As Long
    Dim ReadIt() As Byte
    
    ReadIt = readbytes(Len(lngread))
    CopyMemory lngread, ReadIt(LBound(ReadIt)), Len(lngread)
    ReadLong = lngread
    
End Function
Public Function ReadObject() As Object
'reads an object from the file.
Dim PropBag As PropertyBag
Dim bytesread() As Byte
Dim ByteLen As Long


'first read in length...
    ByteLen = Me.ReadLong
    'ReDim bytesread(1 To bytelen)
    bytesread = Me.readbytes(ByteLen)
    Set PropBag = New PropertyBag
    PropBag.Contents = bytesread
    Set ReadObject = PropBag.ReadProperty("OBJECT", Nothing)

End Function

Public Function ReadSingle() As Single
    Dim retval As Single
    Dim CopyTo() As Byte
    CopyTo = readbytes(Len(retval))
    CopyMemory retval, CopyTo(LBound(CopyTo)), Len(retval)
    ReadSingle = retval
End Function

Public Function ReadString(ByVal StrAmount As Long, Optional ByVal Stringmode As StringReadMode = StrRead_Default, Optional ByRef hitEOF As Boolean)
Dim readcast() As Byte, retval As Long
Stringmode = GetReadMode(Stringmode)
readcast = readbytes(StrAmount * IIf(Stringmode = StrRead_unicode, 2, 1), retval, hitEOF)
If Stringmode = StrRead_ANSI Then
    ReadString = StrConv(readcast, vbUnicode)
Else
    ReadString = readcast
End If





End Function
Public Function ReadUntil(ByVal SearchFor As String, Optional ByVal SeekAfterToken As Boolean = True, Optional ByVal Stringmode As StringReadMode = StrRead_unicode, Optional ByVal Comparemethod As VbCompareMethod = vbTextCompare, Optional ByVal Chunksize As Integer = 64) As String

    
    
    Dim buildstr As String, chunkread As Long, hitEOF As Boolean
    Dim currentPosition As Long
    Dim readValue() As Byte, posfound As Long
    Dim convert As String
    If Chunksize <= 0 Then Chunksize = 1 'minimum.
    currentPosition = Me.GetPosition
    Do
        'readValue = ReadBytes(ChunkSize, chunkread, hiteof)
        'convert = readValue
        'convert = StrConv(convert, vbUnicode)
        convert = ReadString(Chunksize, Stringmode, hitEOF)
        buildstr = buildstr & convert
        
        If hitEOF Then
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
            Me.SeekTo currentPosition + posfound, STREAM_BEGIN
            ReadUntil = buildstr
        
            Exit Function
        End If
    
    Loop
    
    


End Function
Public Function SeekTo(ByVal NewPosition As Double, Optional ByVal SeekMode As FileSeekBase) As Long
    Dim useposition As Long
    SeekTo = Me.GetPosition

If SeekMode = STREAM_BEGIN Then
    useposition = NewPosition
ElseIf SeekMode = STREAM_CURRENT Then
    useposition = NewPosition + mPosition
ElseIf SeekMode = STREAM_END Then
    useposition = mStreamDataSize - NewPosition
End If
If useposition > mStreamDataSize Then useposition = mStreamDataSize

mPosition = useposition
End Function

Public Sub WriteBytes(BytesWrite() As Byte)

    'write the given bytes to the current position.
    Dim Writesize As Long

    Writesize = UBound(BytesWrite) - LBound(BytesWrite) + 1

    If mPosition + Writesize > mStreamSize Then
        'in this case, we cannot write the given byte array without overflowing the array.
        'therefore we need to resize the array to be the smallest multiple of mchunksize that mposition+writesize fits within.
        mStreamSize = ((mPosition + Writesize + mChunksize) Mod mChunksize) * mChunksize
        
        ReDim Preserve mStreamContents(1 To mStreamSize)
    
      
    End If
    'mStreamDataSize = mStreamDataSize + Writesize
    mStreamDataSize = mPosition + Writesize
    CopyMemory mStreamContents(mPosition), BytesWrite(LBound(BytesWrite)), Writesize
    mPosition = mPosition + Writesize
End Sub
Public Function Size() As Long
Size = mStreamDataSize
    
End Function
Public Sub WriteDouble(ByVal DblWrite As Double)
    Dim CopyTo(1 To Len(DblWrite)) As Byte
    CopyMemory CopyTo(1), DblWrite, Len(DblWrite)
    WriteBytes CopyTo
End Sub

Public Sub WriteInteger(ByVal IntWrite As Integer)
    Dim CopyTo(1 To Len(IntWrite)) As Byte
    CopyMemory CopyTo(1), IntWrite, Len(IntWrite)

    WriteBytes CopyTo
End Sub

Public Sub WriteLong(ByVal LngWrite As Long)
    
    Dim CopyTo(1 To Len(LngWrite)) As Byte
    CopyMemory CopyTo(1), LngWrite, Len(LngWrite)
    WriteBytes CopyTo
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
    Err.Raise 9, "MemStream::WriteObject", "Object does not implement proper persistence interfaces."
    


End Sub



Public Sub WriteSingle(ByVal SngWrite As Long)
    Dim CopyTo(1 To Len(SngWrite)) As Byte
    CopyMemory CopyTo(1), SngWrite, Len(SngWrite)
    WriteBytes CopyTo
End Sub
Public Sub WriteString(ByVal StrWrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
    Dim cast() As Byte
    If StrWrite = "" Then Exit Sub
'    If StringMode = strread_unicode Then
'        StrWrite = StrConv(StrWrite, vbUnicode)
'    End If

    
    Stringmode = GetReadMode(Stringmode)
    If Stringmode = StrRead_ANSI Then
        StrWrite = StrConv(StrWrite, vbFromUnicode)
    End If
    cast = StrWrite
    
    WriteBytes cast
    Erase cast
    



End Sub

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
Public Property Get AtEndOfStream() As Boolean
 AtEndOfStream = EOF
End Property
Public Property Get Stringmode() As StringReadMode
    Stringmode = mvarStringReadMode
End Property
Public Property Let Stringmode(ByVal Vdata As StringReadMode)
    mvarStringReadMode = Vdata
End Property
Public Function ReadAllStr(Optional ByVal Stringmode As StringReadMode = StrRead_Default) As String
    'read it all .
    Dim allread As String
    allread = ReadString(mStreamDataSize, Stringmode)
  
        ReadAllStr = allread
   
End Function
Public Sub CloseStream()
    Class_Initialize
End Sub
Private Sub IInputStream_CloseStream()
'
CloseStream
End Sub

Private Function IInputStream_EOF() As Boolean
IInputStream_EOF = Me.EOF
End Function

Private Function IInputStream_GetPos() As Double
IInputStream_GetPos = Me.GetPosition
End Function

Private Function IInputStream_ReadAllStr(Optional ByVal Stringmode As StringReadMode = -1&) As String
IInputStream_ReadAllStr = Me.ReadAllStr(Stringmode)
End Function

Private Function IInputStream_readbytes(ByVal NumBytes As Long, Optional ByRef Numread As Long) As Byte()
IInputStream_readbytes = Me.readbytes(NumBytes, Numread, 0)
End Function

Private Function IInputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Double
 Me.SeekTo ToPos, SeekMode
End Function

Private Function IInputStream_Size() As Double
IInputStream_Size = mStreamDataSize
End Function

Private Function IInputStream_Valid() As Boolean
IInputStream_Valid = True
End Function

Private Sub IOutputStream_CloseStream()
Me.CloseStream
End Sub

Private Function IOutputStream_GetPos() As Double
IOutputStream_GetPos = Me.GetPosition
End Function

Private Function IOutputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Variant
IOutputStream_SeekTo = Me.SeekTo(ToPos, SeekMode)
End Function

Private Function IOutputStream_Valid() As Boolean
IOutputStream_Valid = True
End Function

Private Function IOutputStream_WriteBytes(Bytes() As Byte) As Boolean
 Me.WriteBytes Bytes
End Function

Private Sub IOutputStream_WriteString(Stringwrite As String, Optional ByVal Stringmode As StringReadMode = -1&)
Me.WriteString Stringwrite, Stringmode
End Sub
Public Sub WriteStream(ReadFrom As IInputStream, Optional ByVal Chunksize As Long = 64& * 1024&)


    CopyEntireStream ReadFrom, Me, Chunksize
End Sub

