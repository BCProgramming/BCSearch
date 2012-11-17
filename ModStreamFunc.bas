Attribute VB_Name = "ModStreamFunc"
Option Explicit
Public Declare Function IsTextUnicode Lib "advapi32.dll" (ByRef lpBuffer As Any, ByVal cb As Long, ByRef lpi As Long) As Long
Public Const IS_TEXT_UNICODE_ASCII16 As Long = &H1
Private Const IS_TEXT_UNICODE_CONTROLS As Long = &H4
Private Const IS_TEXT_UNICODE_DBCS_LEADBYTE As Long = &H400
Private Const IS_TEXT_UNICODE_ILLEGAL_CHARS As Long = &H100
Private Const IS_TEXT_UNICODE_NOT_ASCII_MASK As Long = &HF000&
Private Const IS_TEXT_UNICODE_NOT_UNICODE_MASK As Long = &HF00&
Private Const IS_TEXT_UNICODE_NULL_BYTES As Long = &H1000
Private Const IS_TEXT_UNICODE_ODD_LENGTH As Long = &H200
Private Const IS_TEXT_UNICODE_REVERSE_ASCII16 As Long = &H10
Private Const IS_TEXT_UNICODE_REVERSE_CONTROLS As Long = &H40
Private Const IS_TEXT_UNICODE_REVERSE_MASK As Long = &HF0&
Private Const IS_TEXT_UNICODE_REVERSE_SIGNATURE As Long = &H80
Private Const IS_TEXT_UNICODE_REVERSE_STATISTICS As Long = &H20
Private Const IS_TEXT_UNICODE_SIGNATURE As Long = &H8
Private Const IS_TEXT_UNICODE_STATISTICS As Long = &H2
Private Const IS_TEXT_UNICODE_UNICODE_MASK As Long = &HF&







'---------------------------------------------------------------------------------------
' Procedure : SaveImage
' Purpose   : Saves a StdPicture object in a byte array.
'---------------------------------------------------------------------------------------
'
Public Sub WritePicture(ToStream As IOutputStream, ByVal image As StdPicture)
Dim abData() As Byte
Dim oPersist As IPersistStream
Dim oStream As iStream
Dim lSize As Long
Dim tStat As STATSTG

   ' Get the image IPersistStream interface
   Set oPersist = image
   
   ' Create a stream on global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Save the picture in the stream
   oPersist.Save oStream, True
      
   ' Get the stream info
   oStream.Stat tStat, STATFLAG_NONAME
      
   ' Get the stream size
   lSize = tStat.cbSize * 10000
   
   ' Initialize the array
   ReDim abData(0 To lSize - 1)
   
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0, STREAM_SEEK_SET
   
   ' Read all the stream in the array
   oStream.Read abData(0), lSize
   
   ' Return the array
   'SaveImage = abData
   'write to the FileStream...
   'write out a "header" consisting of the data length...
   WriteLong ToStream, UBound(abData) + 1
   ToStream.WriteBytes abData()
   
   ' Release the stream object
   Set oStream = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadImage
' Purpose   : Creates a StdPicture object from a byte array.
'---------------------------------------------------------------------------------------
'
Public Function ReadPicture(FromStream As IInputStream) As StdPicture
Dim oPersist As IPersistStream
Dim oStream As iStream
Dim lSize As Long
Dim imagebytes() As Byte
   'read in the data size...
   'lSize = UBound(ImageBytes) - LBound(ImageBytes) + 1
   lSize = ReadLong(FromStream)
   
   ' Create a stream object
   ' in global memory
   Set oStream = CreateStreamOnHGlobal(0, True)
   
   ' Write the header to the stream
   oStream.Write &H746C&, 4&
   
   ' Write the array size
   oStream.Write lSize, 4&
   
   ' Write the image data
   imagebytes() = FromStream.readbytes(lSize)
   oStream.Write imagebytes(LBound(imagebytes)), lSize
   
   ' Move the stream position to
   ' the start of the stream
   oStream.Seek 0, STREAM_SEEK_SET
      
   ' Create a new empty picture object
   'Set LoadImage = New StdPicture
   Set ReadPicture = New StdPicture
   
   ' Get the IPersistStream interface
   ' of the picture object
   Set oPersist = ReadPicture
   
   ' Load the picture from the stream
   oPersist.Load oStream
      
   ' Release the streamobject
   Set oStream = Nothing
   
End Function





Public Sub WriteObject(StreamWrite As IOutputStream, ObjWrite As Object)
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
     
    WriteLong StreamWrite, bytelength
    StreamWrite.WriteBytes BytesWrite
    'success.
    Exit Sub
    
Nopersist:
    Err.Raise 9, "modStreamFunc::WriteObject", "Object does not implement proper persistence interfaces."
    


End Sub
Public Function ReadObject(FromStream As IInputStream) As Object
'reads an object from the file.
Dim PropBag As PropertyBag
Dim bytesread() As Byte
Dim ByteLen As Long


'first read in length...
    ByteLen = ReadLong(FromStream)
    'ReDim bytesread(1 To bytelen)
    bytesread = FromStream.readbytes(ByteLen)
    Set PropBag = New PropertyBag
    PropBag.Contents = bytesread
    Set ReadObject = PropBag.ReadProperty("OBJECT", Nothing)

End Function

Public Sub WriteInteger(ToStream As IOutputStream, ByVal IntWrite As Integer)
    Dim CopyTo(1 To Len(IntWrite)) As Byte
    CopyMemory CopyTo(1), IntWrite, Len(IntWrite)

    ToStream.WriteBytes CopyTo
End Sub
Public Function ReadInteger(FromStream As IInputStream) As Integer
    Dim intread As Integer
    Dim ReadIt() As Byte
    ReadIt = FromStream.readbytes(Len(intread))
    CopyMemory intread, ReadIt(1), UBound(ReadIt)
    ReadInteger = intread
    
End Function

Public Sub WriteLong(ToStream As IOutputStream, ByVal LngWrite As Long)
    
    Dim CopyTo(1 To Len(LngWrite)) As Byte
    CopyMemory CopyTo(1), LngWrite, Len(LngWrite)
    ToStream.WriteBytes CopyTo
End Sub
Public Function ReadLong(FromStream As IInputStream) As Long
    Dim lngread As Long
    Dim ReadIt() As Byte
    
    ReadIt = FromStream.readbytes(Len(lngread))
    CopyMemory lngread, ReadIt(LBound(ReadIt)), Len(lngread)
    ReadLong = lngread
    
    
End Function



Public Sub WriteSingle(ToStream As IOutputStream, ByVal SngWrite As Long)
    Dim CopyTo(1 To Len(SngWrite)) As Byte
    CopyMemory CopyTo(1), SngWrite, Len(SngWrite)
    ToStream.WriteBytes CopyTo
End Sub

Public Sub WriteDouble(ToStream As IOutputStream, ByVal DblWrite As Double)
    Dim CopyTo(1 To Len(DblWrite)) As Byte
    CopyMemory CopyTo(1), DblWrite, Len(DblWrite)
    ToStream.WriteBytes CopyTo
End Sub
Public Function ReadDouble(FromStream As IInputStream) As Double
    Dim retval As Double
    Dim CopyTo() As Byte
    CopyTo = FromStream.readbytes(Len(retval))
    CopyMemory retval, CopyTo(LBound(CopyTo)), Len(retval)
    ReadDouble = retval
End Function

Public Function ReadSingle(FromStream As IInputStream) As Single
    Dim retval As Single
    Dim CopyTo() As Byte
    CopyTo = FromStream.readbytes(Len(retval))
    CopyMemory retval, CopyTo(LBound(CopyTo)), Len(retval)
    ReadSingle = retval
End Function

Public Function ReadString(FromStream As IInputStream, ByVal StrAmount As Long, Optional ByVal Stringmode As StringReadMode = StrRead_Default, Optional ByRef hitEOF As Boolean)
Dim readcast() As Byte, retval As Long


If Stringmode = StrRead_Default Then
    ReadString = ReadStringAuto(FromStream, StrAmount)


ElseIf Stringmode = StrRead_ANSI Then
    readcast = FromStream.readbytes(StrAmount * IIf(Stringmode = StrRead_unicode, 2, 1), retval)
    ReadString = StrConv(readcast, vbUnicode)
Else
    readcast = FromStream.readbytes(StrAmount * IIf(Stringmode = StrRead_unicode, 2, 1), retval)
    ReadString = readcast
End If
If retval > StrAmount Then hitEOF = True




End Function
Public Sub WriteString(ToStream As IOutputStream, ByVal StrWrite As String, Optional ByVal Stringmode As StringReadMode = StrRead_Default)
    Dim cast() As Byte
    
    If StrWrite = "" Then Exit Sub
'    If StringMode = strread_unicode Then
'        StrWrite = StrConv(StrWrite, vbUnicode)
'    End If

    
    
    If Stringmode = StrRead_ANSI Then
        StrWrite = StrConv(StrWrite, vbFromUnicode)
    End If
    cast = StrWrite
    
    ToStream.WriteBytes cast
    Erase cast
    



End Sub
Public Sub WriteStream(writeto As IOutputStream, FromStream As IInputStream, Optional ByVal Chunksize As Long = 32768)


    CopyEntireStream FromStream, writeto, Chunksize


End Sub
Public Function ReadUntil(FromStream As IInputStream, ByVal SearchFor As String, Optional ByVal SeekAfterToken As Boolean = True, Optional ByVal Stringmode As StringReadMode = StrRead_unicode, Optional ByVal Comparemethod As VbCompareMethod = vbTextCompare, Optional ByVal Chunksize As Integer = 64) As String

    
    
    Dim buildstr As String, chunkread As Long, hitEOF As Boolean
    Dim currentPosition As Long
    Dim readValue() As Byte, posfound As Long
    Dim convert As String
    If Chunksize <= 0 Then Chunksize = 1 'minimum.
    currentPosition = FromStream.GetPos
    Do
        'readValue = ReadBytes(ChunkSize, chunkread, hiteof)
        'convert = readValue
        'convert = StrConv(convert, vbUnicode)
        convert = ReadString(FromStream, Chunksize, Stringmode, hitEOF)
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
            FromStream.SeekTo currentPosition + posfound, STREAM_BEGIN
            ReadUntil = buildstr
        
            Exit Function
        End If
    
    Loop
    
    


End Function


Public Function ReadUntil_Auto(FromStream As IInputStream, ByVal SearchFor As String, Optional ByVal SeekAfterToken As Boolean = True, Optional ByVal Comparemethod As VbCompareMethod = vbTextCompare, Optional ByVal Chunksize As Integer = 64) As String

    
    
    Dim buildstr As String, chunkread As Long, hitEOF As Boolean
    Dim currentPosition As Long
    Dim readValue() As Byte, posfound As Long
    Dim convert As String
    If Chunksize <= 0 Then Chunksize = 1 'minimum.
    currentPosition = FromStream.GetPos
    Do
        'readValue = ReadBytes(ChunkSize, chunkread, hiteof)
        'convert = readValue
        'convert = StrConv(convert, vbUnicode)
        convert = ReadStringAuto(FromStream, Chunksize)
        buildstr = buildstr & convert
        
        If FromStream.EOF Then
            'end of file...
            ReadUntil_Auto = buildstr
            
        
        
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
            FromStream.SeekTo currentPosition + posfound, STREAM_BEGIN
            ReadUntil_Auto = buildstr
        
            Exit Function
        End If
    
    Loop
    
    


End Function


'Public Function testUnicode()
'    Dim strconvert As String
'    Dim bytes() As Byte
'    strconvert = "this is a unicode string"
'    bytes = strconvert
'    Debug.Print IsByteStreamUnicode(bytes)
'
'End Function
Public Function IsByteStreamUnicode(Bytes() As Byte) As Boolean
    'returns true if byte stream is unicode; false otherwise.
    
    'select the first 32 bytes from the array, or the highest even number of bytes.
    
    
    Dim sampleSize As Long
    Dim samples() As Byte
    If (UBound(Bytes) - LBound(Bytes) + 1) > 32 Then
        sampleSize = 32
    Else
        sampleSize = ((UBound(Bytes) - LBound(Bytes) + 1)) - ((UBound(Bytes) - LBound(Bytes) + 1) Mod 2)
    
    End If
    ReDim samples(1 To sampleSize)
    CopyMemory samples(1), Bytes(LBound(Bytes)), sampleSize
    
    
    
    
   ' now use IsTextUnicode() to determine wether the Sample() byte array is indeed unicode.
    If IsTextUnicode(samples(1), sampleSize, IS_TEXT_UNICODE_NULL_BYTES Or IS_TEXT_UNICODE_ASCII16 Or IS_TEXT_UNICODE_CONTROLS Or IS_TEXT_UNICODE_STATISTICS) > 0 Then
    
        IsByteStreamUnicode = True
    
    End If



End Function

Public Function ReadStringAuto(OnStream As IInputStream, ByVal StringLength As Long, Optional ByRef hitEOF As Boolean = False) As String
    'Reads a String automatically using the proper encoding... at least, that's the desired effect.
    
    
    'StringLength denotes the number of characters we want.
    
    'First: store the current position.
    
    Dim Currpos As Double
    Dim BytesTest() As Byte, bytesread As Long
    Currpos = OnStream.GetPos
    If StringLength = 0 Then Exit Function
    'Alright; now, read in a sampling of data to test for  Unicode-ness; 32 bytes is sufficient.
    BytesTest = OnStream.readbytes(64, bytesread)
    If bytesread = 0 Then
        ReadStringAuto = ""
    ElseIf bytesread = 1 Then
        'no reason to continue... we only got one byte so return that as an ASCII character.
        ReadStringAuto = Chr$(BytesTest(0))
        
    ElseIf bytesread > 1 Then
        '2 or more is enough to test for Unicode-ness.
        If IsByteStreamUnicode(BytesTest()) Then
            'it is, indeed, unicode.
            'seek back and use the standard readString function.
            OnStream.SeekTo Currpos, STREAM_BEGIN
            ReadStringAuto = ReadString(OnStream, StringLength, StrRead_unicode)
        Else
            'it's ASCII. seek back and use standard readstring.
            OnStream.SeekTo Currpos, STREAM_BEGIN
            ReadStringAuto = ReadString(OnStream, StringLength, StrRead_ANSI)
        End If
    End If
    
    



End Function


'Functions Specifically for reading And writing text files.
Public Function ReadLine(FromStream As IInputStream, Optional ByVal LineDelimiter As String = vbCrLf) As String
    'reads a single line from the text file.
    
    Dim strread As String
    strread = ReadUntil(FromStream, LineDelimiter, True)
    ReadLine = strread
    
    



End Function



'Public Function ReadStringAuto(ByVal Length As Long) As String
'    'auto-detects unicode or ANSI.
'
'
'    'Length is the desired length of the string.
'    'First, read length bytes of characters in.
'
'    Dim FirstChunk() As Byte, Firstchunksize As Long
'    Dim secondChunk() As Byte, SecondChunksize As Long
'    Dim passbuffersize As Long
'    Dim stringread As String
'    FirstChunk = mInputStream.ReadBytes(Length, Firstchunksize)
'
'    passbuffersize = Firstchunksize
'    'if the chunk size was an odd number, reduce the passed buffer size variable to allow for proper detection of Unicode.
'    If (Firstchunksize Mod 2) <> 0 Then
'        passbuffersize = Firstchunksize + 1
'    End If
'
'    If IsTextUnicode(FirstChunk(LBound(FirstChunk)), passbuffersize, IS_TEXT_UNICODE_ASCII16 + &H10 + &H2) > 0 Then
'        'text is unicode...
'        'read in the second chunk, resize the first chunk buffer and copymemory to the first chunk from the second chunk.
'        secondChunk = mInputStream.ReadBytes(Length, SecondChunksize)
'
'        ReDim Preserve FirstChunk(0 To (Length * 2) - 1)
'        CopyMemory FirstChunk(Firstchunksize), secondChunk(0), UBound(secondChunk) + 1
'        Erase secondChunk
'        stringread = FirstChunk
'    Else
'        stringread = FirstChunk
'
'
'    End If
'    ReadString = stringread
'End Function
