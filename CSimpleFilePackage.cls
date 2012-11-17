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

Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long


'12/13/2009 12:14:36 TERATRON
'Changed WritePackageStream and ReadPackageStream to accept and use the more generic IInputStream and IOutputStream interfaces.


'format of the "pure" stream:

'HEADER: 4 bytes: "BCZP"

'Count of Files in package
'file index records- each one is:
'Length of filename (Long)
'Filename (string) *this may contain a path*
'Offset in the package where the file data begins (Double)
Public Enum PackageFormatEnum
    Format_BCZP
    Format_KSGRP
End Enum
Private Type BCZPFileIndexRecord
    FilenameLength As Long
    PackedFolder As String 'Folder it was packed in. empty string means the "root" of the package.
    Filename As String
    DataOffset As Long
    DataLength As Long
    FullSourcePath As String
  

End Type

Const MAGIC_NUMBER = 3945035
Private Type BCZPHeader
    DataOffset As Long  'offset to the data...
    FileCount As Long
    FileRecords() As BCZPFileIndexRecord
End Type
Private misInitialized As Boolean
Private mHeader As BCZPHeader
Private mFilename As String
Private mCachedStream As Object
Private mFileFormat As PackageFormatEnum

Public Property Let FileFormat(ByVal Vdata As PackageFormatEnum)
    mFileFormat = Vdata
End Property
Public Property Get FileFormat() As PackageFormatEnum
    FileFormat = mFileFormat
End Property
'Private mFileStream As FileStream
Public Sub ReadPackage(ByVal mFilename As String)
    Dim openedstream As FileStream
    Set openedstream = FileSystem.OpenStream(mFilename)
    ReadPackageStream openedstream
End Sub
Public Sub ReadPackageStream(ByVal usestream As IInputStream)

    If mFileFormat = Format_KSGRP Then
        ReadPackageStreamGRP usestream
        Exit Sub
    ElseIf mFileFormat = Format_BCZP Then
        ReadPackageStreamBCZP usestream
    End If
End Sub
Public Sub ReadPackageStreamBCZP(usestream As IInputStream)

    Erase mHeader.FileRecords
    'mFilename = Filename
    mHeader.FileCount = 0
    mHeader.DataOffset = 0
    'Set useStream = FileSystem.OpenStream(Filename)
    Set mCachedStream = usestream
    Dim readNumber As Long
    readNumber = ModStreamFunc.ReadLong(usestream)
    If readNumber <> MAGIC_NUMBER Then
        Err.Raise 9, "CSimpleFilePackage::ReadPackage", "Invalid File Format. Magic number not found."
    End If
    'file record count, dataoffset.
    mHeader.FileCount = ReadLong(usestream)
    mHeader.DataOffset = ReadLong(usestream)
    Dim I As Long
    'read each file.
    ReDim mHeader.FileRecords(1 To mHeader.FileCount)

Dim templength As Long
        For I = 1 To mHeader.FileCount
            With mHeader.FileRecords(I)
                .DataOffset = ReadLong(usestream)
                templength = ReadLong(usestream)
                .Filename = ReadString(usestream, templength, StrRead_unicode)
                .FullSourcePath = ""
                .DataLength = ReadLong(usestream)
            
            End With
        
        
        
        Next I
   ' useStream.CloseStream
    
    
    'hooray, it's initialized...
End Sub
Private Function FindFileIndex(ByVal Filename As String) As Long
    Dim I As Long
    For I = 1 To mHeader.FileCount
    'If Trim$(mHeader.FileRecords(I).filename) = "E1L1.MAP" Then Stop
        If StrComp(Trim$(mHeader.FileRecords(I).Filename), Filename, vbTextCompare) = 0 Then
            FindFileIndex = I
            Exit Function
        
        End If
    
    
    Next I
    FindFileIndex = -1


End Function
'Public Sub CreatePackedFolderTree(ByVal BaseFolder As String)
'    'recreates the folder tree. for example, if we have a file whose name is "\files\filename.txt" then calling this routine with createPackedFolderTree("D:\extract") will create a D:\extract\files folder.
'    'Note that no actual Files are unzipped.
'    Dim DirBaseFolder As Directory, i As Long, createrelative As String
'    Set DirBaseFolder = FileSystem.CreateDirectory(BaseFolder, Nothing, True)
'
'    For i = 1 To mHeader.FileCount
'        With mHeader.FileRecords(i)
'            createrelative = .PackedFolder
'            If Left$(createrelative, 1) = "\" Then
'                createrelative = Mid$(createrelative, 2)
'            End If
'            If createrelative <> "" Then
'                Call FileSystem.CreateDirectory(DirBaseFolder.Path & createrelative, Nothing, True)
'
'
'            End If
'        End With
'
'
'
'
'
'    Next i
'
'
'
'End Sub
Public Sub ExtractFile(ByVal vntNameIndex As Variant, tostream As FileStream)
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
        useIn.SeekTo .DataOffset
        tostream.WriteBytes useIn.readbytes(.DataLength)
    
    
    End With
    'useIn.CloseStream
    
End Sub



Public Sub WritePackageFiltered(ByVal Filename As String, Filteruse As IStreamFilter)
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
    Set outtemp = FileSystem.CreateStream(Filename)
    'filter.FilterType = BCFile.CompressionConstants.Huffman_Compress
    'Set compressor = filter
    compressor.FilterStream InputTemp, outtemp
    
    
    outtemp.CloseStream
    InputTemp.CloseStream
    
    
    
    



End Sub
Public Sub ReadPackageFiltered(ByVal Filename As String, Filteruse As IStreamFilter)
    Dim expander As IStreamFilter
    'Dim filterobj As CCoreFilters
    'Set filterobj = New CCoreFilters
    'filterobj.FilterType = Huffman_Expand
    Set expander = Filteruse
    Dim Tempout As String
    Tempout = GetTempFileNameAndPathEx()
    Dim inputstream As FileStream, outputstream As FileStream
    
    Set inputstream = FileSystem.OpenStream(Filename)
    Set outputstream = FileSystem.CreateStream(Tempout)
    
    
    expander.FilterStream inputstream, outputstream
    
    inputstream.CloseStream
    outputstream.CloseStream
    Me.ReadPackage Tempout
End Sub
Public Sub WritePackage_Huffman(ByVal Filename As String)
    Dim IFilteruse As IStreamFilter
    Dim CoreFilter As CCoreFilters
    Set CoreFilter = New CCoreFilters
    CoreFilter.FilterType = Huffman_Compress
    WritePackageFiltered Filename, CoreFilter
    
    




End Sub
Public Sub ReadPackage_Huffman(ByVal Filename As String)
   Dim IFilteruse As IStreamFilter
    Dim CoreFilter As CCoreFilters
    Set CoreFilter = New CCoreFilters
    CoreFilter.FilterType = Huffman_Expand
    ReadPackageFiltered Filename, CoreFilter
End Sub
Public Sub WritePackage(ByVal mFilename As String)
    Dim openedstream As FileStream
    Set openedstream = FileSystem.CreateStream(mFilename)
    WritePackageStream openedstream
End Sub
Private Sub WritePackageStreamGRP(ByVal mFileStream As IOutputStream)
'
'The ".grp" file format is just a collection of a lot of files stored into 1 big one.
'I tried to make the format as simple as possible: The first 12 bytes contains my name,'
'"KenSilverman". The next 4 bytes is the number of files that were compacted into the group file.
'Then for each file, there is a 16 byte structure, where the first 12 bytes are the filename,
'and the last 4 bytes are the file's size.
'The rest of the group file is just the raw data packed one after the other in the same order as the list of files.
Dim strheader As String * 12, shortpathname As String, shortfilename As String * 12
Dim Numfiles As Long, LoopFile As Long
strheader = "KenSilverman"
Numfiles = mHeader.FileCount

    mFileStream.WriteString strheader, StrRead_ANSI
    WriteLong mFileStream, Numfiles

    'for every file, a 8.3 file name, and a 4-byte filesize.
    For LoopFile = 1 To mHeader.FileCount
        shortpathname = Space$(1024)
        Call GetShortPathName(StrPtr(mHeader.FileRecords(LoopFile).Filename), StrPtr(shortpathname), Len(shortpathname))
        shortfilename = FileSystem.GetFilenamePart(shortpathname)
        
        'write the filename... ANSI, of course.
        mFileStream.WriteString shortfilename, StrRead_ANSI
        'write the file size.
        WriteLong mFileStream, mHeader.FileRecords(LoopFile).DataLength
        
    
    
    
    
    
    Next LoopFile
    
    
    'ok, the Header portion has been written. Now, write out each files data.
    Dim StreamOpen As FileStream
    For LoopFile = 1 To mHeader.FileCount
        Set StreamOpen = FileSystem.OpenStream(mHeader.FileRecords(LoopFile).Filename)
        ModStreamFunc.WriteStream mFileStream, StreamOpen
        StreamOpen.CloseStream
    Next
    
    
    
    
  




End Sub

'Reads a GRP format package.

Private Sub ReadPackageStreamGRP(ByVal InStream As IInputStream)
'
'
'The ".grp" file format is just a collection of a lot of files stored into 1 big one.
'I tried to make the format as simple as possible: The first 12 bytes contains my name,'
'"KenSilverman". The next 4 bytes is the number of files that were compacted into the group file.
'Then for each file, there is a 16 byte structure, where the first 12 bytes are the filename,
'and the last 4 bytes are the file's size.
'The rest of the group file is just the raw data packed one after the other in the same order as the list of files.

Dim nameheader As String * 12
Dim Numfiles As Long
Dim shortname As String * 12
Set mCachedStream = InStream

nameheader = ReadString(InStream, Len(nameheader), StrRead_ANSI)
If nameheader <> "KenSilverman" Then
    Err.Raise 9, "CSimpleFilePackage::ReadPackageStreamGRP", "GRP file missing Proper header!"
    Exit Sub

End If

Numfiles = ReadLong(InStream)
Dim currfile As Long
mHeader.FileCount = Numfiles
ReDim mHeader.FileRecords(1 To mHeader.FileCount)

'iterate through each one and populate the proper structures.
For currfile = 1 To mHeader.FileCount
    With mHeader.FileRecords(currfile)
        .Filename = ReadString(InStream, 12, StrRead_ANSI)
        .Filename = Replace$(.Filename, vbNullChar, " ")
        .DataLength = ReadLong(InStream)
        .FilenameLength = Len(.Filename)



    End With
Next currfile

'now that we have the length of the data, we can loop back through the file names and read the appropriate amount... or, to be more precise, populate the proper dataoffset members.
'This way the extraction routines will work.
For currfile = 1 To mHeader.FileCount
    With mHeader.FileRecords(currfile)
    'populate the data offsets.
        .DataOffset = InStream.GetPos
        'Call InStream.ReadBytes(.DataLength)  'read and discard.
        InStream.SeekTo CDbl(.DataLength), STREAM_CURRENT
        
        
    End With

Next

'ok, all the proper data should now be saved.

End Sub
Public Sub WritePackageStream(ByVal mFileStream As IOutputStream)
    'Writes the File Package.
    
    
    
    
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
    If mFileFormat = Format_KSGRP Then
        WritePackageStreamGRP mFileStream
        Exit Sub
    ElseIf mFileFormat = Format_BCZP Then
        WritePackageStreamBCZP mFileStream
    End If
 End Sub
Public Sub WritePackageStreamBCZP(mFileStream As IOutputStream)

    Dim OpenedFile As FileStream
    Dim I As Long, dataoffsetLoc As Long, temp As Long
    ModStreamFunc.WriteLong mFileStream, MAGIC_NUMBER
    ModStreamFunc.WriteLong mFileStream, mHeader.FileCount
    dataoffsetLoc = mFileStream.GetPos
    
    ModStreamFunc.WriteLong mFileStream, mHeader.DataOffset
    'write out some preliminary header information....
    For I = 1 To mHeader.FileCount
        'mFileStream.WriteLong CLng(0)
        ModStreamFunc.WriteLong mFileStream, CLng(0)
        ModStreamFunc.WriteLong mFileStream, Len(mHeader.FileRecords(I).Filename)
        ModStreamFunc.WriteString mFileStream, mHeader.FileRecords(I).Filename, StrRead_unicode
        
        'folder in the package
        ModStreamFunc.WriteLong mFileStream, Len(mHeader.FileRecords(I).PackedFolder)
        ModStreamFunc.WriteString mFileStream, mHeader.FileRecords(I).PackedFolder, StrRead_unicode
        
        ModStreamFunc.WriteLong mFileStream, CLng(0) 'placeholder for file size.
        'mfilestream.WriteLong Len(mHeader.FileRecords(I).FullSourcePath)
        'mfilestream.WriteString mHeader.FileRecords(I).FullSourcePath, strread_unicode


        
        
    Next I
    'temp = mfilestream.GetFilePos()
    mHeader.DataOffset = mFileStream.GetPos
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
        mHeader.FileRecords(I).DataOffset = mFileStream.GetPos
        'now, empty out openedfile into this stream...
        ModStreamFunc.WriteStream mFileStream, OpenedFile
        'mFileStream.WriteStream OpenedFile
        OpenedFile.CloseStream
        
        
        
        
    
    Next I
    
    'Seek back to the start of the file structures....
    'Dim temp As Long
    
    mFileStream.SeekTo CDbl(dataoffsetLoc), STREAM_BEGIN
    ModStreamFunc.WriteLong mFileStream, mHeader.DataOffset
    
    For I = 1 To mHeader.FileCount
        ModStreamFunc.WriteLong mFileStream, mHeader.FileRecords(I).DataOffset
        ModStreamFunc.WriteLong mFileStream, Len(mHeader.FileRecords(I).Filename)
        ModStreamFunc.WriteString mFileStream, mHeader.FileRecords(I).Filename, StrRead_unicode
        ModStreamFunc.WriteLong mFileStream, FileLen(mHeader.FileRecords(I).FullSourcePath)
        

    Next I

    

    
    
    
End Sub
Public Sub AddFileToPackage(ByVal Filename As String, Optional ByVal AddToPackedFolder As String = "\")
    Dim gotfile As CFile, I As Long
    'adds to the "root" directory of the package by default.
    Set gotfile = FileSystem.GetFile(Filename)
    mHeader.FileCount = mHeader.FileCount + 1
    ReDim Preserve mHeader.FileRecords(1 To mHeader.FileCount)
    With mHeader.FileRecords(mHeader.FileCount)
        .FullSourcePath = Filename
         .Filename = FileSystem.GetFilenamePart(Filename)
        .DataLength = FileLen(Filename)
     
    
    End With
    'tada!
End Sub
Private Function GetPackedFolderIndex(ByVal StrFoldername As String)
    'iterates through our packed folder list. if strfoldername is found, the index is returned.
    Dim Curr As Long
    For Curr = 1 To mHeader.FileCount
        If StrComp(mHeader.PackedFolders(Curr).FolderFullName, StrFoldername, vbTextCompare) = 0 Then
            GetPackedFolderIndex = Curr
        Else
            GetPackedFolderIndex = -1
        End If
    Next
    
End Function
Private Sub Class_Terminate()
If Not mCachedStream Is Nothing Then
    mCachedStream.CloseStream
End If
End Sub
