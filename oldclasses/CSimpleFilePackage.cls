VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimpleFilePackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CSimpleFilePackage- allows the reading/writing of a set of packaged files.
'Note that this creates temporary files for reading of the files.


'format of the "pure" stream:

'HEADER: 4 bytes: "BCZP"

'Count of Files in package
'file index records- each one is:
'Length of filename (Long)
'Filename (string) *this may contain a path*
'Offset in the package where the file data begins (Double)

Private Type BCZPFileIndexRecord
    FilenameLength As Long
    FullSourcePath As String
    filename As String
    Dataoffset As Long
    DataLength As Long
  
  

End Type
Const MAGIC_NUMBER = 3945035
Private Type BCZPHeader
    Dataoffset As Long  'offset to the data...
    FileCount As Long
    FileRecords() As BCZPFileIndexRecord
End Type
Private misInitialized As Boolean
Private mHeader As BCZPHeader
Private mFilename As String
Private mCachedStream As FileStream
'Private mFileStream As FileStream
Public Sub ReadPackage(ByVal mFilename As String)
    Dim openedstream As FileStream
    Set openedstream = FileSystem.OpenStream(mFilename)
    ReadPackageStream openedstream
End Sub
Public Sub ReadPackageStream(ByVal useStream As FileStream)

    Erase mHeader.FileRecords
    'mFilename = Filename
    mHeader.FileCount = 0
    mHeader.Dataoffset = 0
    'Set useStream = FileSystem.OpenStream(Filename)
    Set mCachedStream = useStream
    Dim readNumber As Long
    readNumber = useStream.ReadLong
    If readNumber <> MAGIC_NUMBER Then
        Err.Raise 9, "CSimpleFilePackage::ReadPackage", "Invalid File Format. Magic number not found."
    End If
    'file record count, dataoffset.
    mHeader.FileCount = useStream.ReadLong
    mHeader.Dataoffset = useStream.ReadLong
    Dim I As Long
    'read each file.
    ReDim mHeader.FileRecords(1 To mHeader.FileCount)
    '     mfilestream.WriteLong CLng(0) data offset.
'        mfilestream.WriteLong Len(mHeader.FileRecords(I).Filename)
'        mfilestream.WriteString mHeader.FileRecords(I).Filename, strread_unicode
'        mfilestream.WriteLong CLng(0) 'placeholder for file size.
Dim templength As Long
        For I = 1 To mHeader.FileCount
            With mHeader.FileRecords(I)
                .Dataoffset = useStream.ReadLong
                templength = useStream.ReadLong
                .filename = useStream.ReadString(templength, strread_unicode)
                .FullSourcePath = ""
                .DataLength = useStream.ReadLong
            
            End With
        
        
        
        Next I
   ' useStream.CloseStream
    
    
    'hooray, it's initialized...
End Sub
Private Function FindFileIndex(ByVal filename As String) As Long
    Dim I As Long
    For I = 1 To mHeader.FileCount
        If StrComp(mHeader.FileRecords(I).filename, filename, vbTextCompare) = 0 Then
            FindFileIndex = I
            Exit Function
        
        End If
    
    
    Next I
    FindFileIndex = -1


End Function
Public Sub ExtractFile(ByVal vntNameIndex As Variant, ToStream As FileStream)
    'Extracts the given file or record index data into the passed stream.
    
    
    'First step: locate the given filename if it is a string- or use the given index and check bounds.
    Dim UseIndex As Long
    
    If varType(vntNameIndex) = vbString Then
        UseIndex = FindFileIndex(CStr(vntNameIndex))
    ElseIf IsNumeric(vntNameIndex) Then
        UseIndex = vntNameIndex
    End If
    If UseIndex < 1 Or UseIndex > mHeader.FileCount Then
        Err.Raise 9, "CSimpleFilePackage::ExtractFile", "Index out of Range or nonexistent filename specified."
    
    End If
    'OK, open the source package, in mfilename...
    With mHeader.FileRecords(UseIndex)
        Dim useIn As FileStream
        'Set useIn = FileSystem.OpenStream(mFilename)
        Set useIn = mCachedStream
        useIn.SeekTo .Dataoffset
        ToStream.WriteBytes useIn.ReadBytes(.DataLength)
    
    
    End With
    'useIn.CloseStream
    
End Sub



Public Sub WritePackageFiltered(ByVal filename As String, Filteruse As IStreamFilter)
    'first, create a temporary file and write the package to it normally...
    Dim tempFile As String
    Dim InputTemp As FileStream, outtemp As FileStream
    Dim compressor As IStreamFilter
   ' Dim filter As BCFile.CCoreFilters
    'Set filter = New CCoreFilters
    tempFile = GetTempFileNameAndPathEx()
    WritePackage tempFile
    
    Set compressor = Filteruse
    'OK, good, good.
    'now open that temp file...
    Set InputTemp = FileSystem.OpenStream(tempFile)
    Set outtemp = FileSystem.CreateStream(filename)
    'filter.FilterType = BCFile.CompressionConstants.Huffman_Compress
    'Set compressor = filter
    compressor.FilterStream InputTemp, outtemp
    
    
    outtemp.CloseStream
    InputTemp.CloseStream
    
    
    
    



End Sub
Public Sub ReadPackageFiltered(ByVal filename As String, Filteruse As IStreamFilter)
    Dim expander As IStreamFilter
    'Dim filterobj As CCoreFilters
    'Set filterobj = New CCoreFilters
    'filterobj.FilterType = Huffman_Expand
    Set expander = Filteruse
    Dim Tempout As String
    Tempout = GetTempFileNameAndPathEx()
    Dim inputstream As FileStream, outputstream As FileStream
    
    Set inputstream = FileSystem.OpenStream(filename)
    Set outputstream = FileSystem.CreateStream(Tempout)
    
    
    expander.FilterStream inputstream, outputstream
    
    inputstream.CloseStream
    outputstream.CloseStream
    Me.ReadPackage Tempout
End Sub
Public Sub WritePackage_Huffman(ByVal filename As String)
    Dim IFilteruse As IStreamFilter
    Dim CoreFilter As CCoreFilters
    Set CoreFilter = New CCoreFilters
    CoreFilter.FilterType = Huffman_Compress
    WritePackageFiltered filename, CoreFilter
    
    




End Sub
Public Sub ReadPackage_Huffman(ByVal filename As String)
   Dim IFilteruse As IStreamFilter
    Dim CoreFilter As CCoreFilters
    Set CoreFilter = New CCoreFilters
    CoreFilter.FilterType = Huffman_Expand
    ReadPackageFiltered filename, CoreFilter
End Sub
Public Sub WritePackage(ByVal mFilename As String)
    Dim openedstream As FileStream
    Set openedstream = FileSystem.CreateStream(mFilename)
    WritePackageStream openedstream
End Sub
Public Sub WritePackageStream(ByVal mFileStream As FileStream)
    'Writes the File Package.
    Dim OpenedFile As FileStream
    'Dim mFileStream As FileStream
    'mFilename = Filename
    'Set mFileStream = FileSystem.CreateStream(Filename)
    
    'FILE FORMAT: BASeCamp SimpleFile Package
    
    'the very first thing in the file is the MAGIC_NUMBER, equal to the MAGIC_NUMBER constant:
    'Then the count of file records.
    'then the offset to the location where file data is stored.
    'MAGIC_NUMBER :Long, 4 bytes.
    'File_records: Long, 4 bytes.
    'DataOffset : Long, 4 bytes.
    
    'then, for every file record in the package, we have the following:
    
    'Filedata offset: Long, 4 bytes. Offset to the data of this file.
    'filename length: long, 4 bytes. Length of filename, in ANSI.
    'Filename, string, length equal to previous filename length field, save is forced to ANSI.
'           mfilestream.WriteLong mHeader.FileRecords(I).Dataoffset
'        mfilestream.WriteLong Len(mHeader.FileRecords(I).Filename)
'        mfilestream.WriteLong FileLen(mHeader.FileRecords(I).FullSourcePath)
'        mfilestream.WriteString mHeader.FileRecords(I).Filename, strread_unicode
    
    
    
    
    Dim I As Long, dataoffsetLoc As Long, temp As Long
    mFileStream.WriteLong MAGIC_NUMBER
    mFileStream.WriteLong mHeader.FileCount
    dataoffsetLoc = mFileStream.GetFilePos
    
    mFileStream.WriteLong mHeader.Dataoffset
    'write out some preliminary header information....
    For I = 1 To mHeader.FileCount
        mFileStream.WriteLong CLng(0)
        mFileStream.WriteLong Len(mHeader.FileRecords(I).filename)
        mFileStream.WriteString mHeader.FileRecords(I).filename, strread_unicode
        mFileStream.WriteLong CLng(0) 'placeholder for file size.
        'mfilestream.WriteLong Len(mHeader.FileRecords(I).FullSourcePath)
        'mfilestream.WriteString mHeader.FileRecords(I).FullSourcePath, strread_unicode


        
        
    Next I
    'temp = mfilestream.GetFilePos()
    mHeader.Dataoffset = mFileStream.GetFilePos
    'mfilestream.SeekTo dataoffsetLoc, FILE_BEGIN
    'write out the saved location. it is where the actual data begins.
    'mfilestream.WriteLong temp
    'seek back to the position so we may write out the files.
    'mfilestream.SeekTo temp, FILE_BEGIN
    For I = 1 To mHeader.FileCount
        'Set OpenedFile = FileSystem.OpenStream(mHeader.FileRecords(I).FullSourcePath)
        'now we can write out the file contents of the specified files, and store the offset data.
        'open this file...
        
        Set OpenedFile = FileSystem.OpenStream(mHeader.FileRecords(I).FullSourcePath)
        'save this location...
        mHeader.FileRecords(I).Dataoffset = mFileStream.GetFilePos
        'now, empty out openedfile into this stream...
        
        mFileStream.WriteStream OpenedFile
        OpenedFile.CloseStream
        
        
        
        
    
    Next I
    
    'Seek back to the start of the file structures....
    'Dim temp As Long
    
    mFileStream.SeekTo dataoffsetLoc, file_BEGIN
    mFileStream.WriteLong mHeader.Dataoffset
    
    For I = 1 To mHeader.FileCount
        mFileStream.WriteLong mHeader.FileRecords(I).Dataoffset
        mFileStream.WriteLong Len(mHeader.FileRecords(I).filename)
        mFileStream.WriteString mHeader.FileRecords(I).filename, strread_unicode
        mFileStream.WriteLong FileLen(mHeader.FileRecords(I).FullSourcePath)
        

    Next I
'     mfilestream.WriteLong CLng(0)
'        mfilestream.WriteLong Len(mHeader.FileRecords(I).Filename)
'        mfilestream.WriteString mHeader.FileRecords(I).Filename, strread_unicode
'        mfilestream.WriteLong CLng(0) 'placeholder for file size.
'    mFileStream.CloseStream
    

    
    
    
End Sub
Public Sub AddFileToPackage(ByVal filename As String)
    Dim gotfile As CFile, I As Long
    Set gotfile = FileSystem.GetFile(filename)
    mHeader.FileCount = mHeader.FileCount + 1
    ReDim Preserve mHeader.FileRecords(1 To mHeader.FileCount)
    With mHeader.FileRecords(mHeader.FileCount)
        .FullSourcePath = filename
         .filename = FileSystem.GetFilenamePart(filename)
        .DataLength = FileLen(filename)
     
    
    End With
    'tada!
End Sub

Private Sub Class_Terminate()
If Not mCachedStream Is Nothing Then
    mCachedStream.CloseStream
End If
End Sub
