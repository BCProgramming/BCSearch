VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHeirarchalFilePackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Heirarchal File package

Private Type BCZP_File_Record
    FilenameLength As Long
    Filename As String
    datasize As Long
    DataOffset As Long
    SeekToOffset As Long 'temp used during write.
    SourceFile As String
End Type

Private Type BCZP_Folder_Record
   FolderNameLength As Long
   FolderName As String
   FileCount As Long
   FileRecords() As BCZP_File_Record
End Type
Const MAGIC_NUMBER = 3945036
Private Type BCZP_Heirarchal_Header
    FolderCount As Long
    FilterMethodNameLength As Long
    FilterMethod As String 'compression (LZW,huffman) or encryption (AES)
    FilterParam As String 'parameter to filter. usually password for encrypted archives.
    FolderRecords() As BCZP_Folder_Record
End Type
Private mHeader As BCZP_Heirarchal_Header

Private mCachedInputStream As IInputStream 'cached, for extraction of files.
Public Property Let FilterType(ByVal Vdata As String)
    mHeader.FilterMethod = Vdata
    mHeader.FilterMethodNameLength = Len(Vdata)
End Property
Public Property Get FilterType() As String
    FilterType = mHeader.FilterMethod
End Property

Private Function FindFolderRecordIndex(ByVal ForFolder As String) As Long

    Dim i As Long
    For i = 1 To mHeader.FolderCount
        If StrComp(mHeader.FolderRecords(i).FolderName, ForFolder, vbBinaryCompare) = 0 Then
            FindFolderRecordIndex = i
            Exit Function
        End If
    
    Next i
    FindFolderRecordIndex = -1

End Function
Private Function FindFileRecordIndex(ByVal InFolder As Long, ByVal Filename As String) As Long
    Dim i As Long
    For i = 1 To mHeader.FolderRecords(InFolder).FileCount
        If StrComp(mHeader.FolderRecords(InFolder).FileRecords(i).Filename, Filename, vbTextCompare) = 0 Then
            FindFileRecordIndex = i
            Exit Function
        End If
    Next i
    FindFileRecordIndex = -1



End Function
Private Function AddFileRecord(ByVal FolderRecordIndex As Long, ByVal Filename As String) As Long
    With mHeader.FolderRecords(FolderRecordIndex)
        .FileCount = .FileCount + 1
        ReDim Preserve .FileRecords(1 To .FileCount)
        With .FileRecords(.FileCount)
            .Filename = Filename
            .FilenameLength = Len(.Filename)
        End With
    AddFileRecord = .FileCount
    End With
    
End Function
Private Function AddFolderRecord(ByVal FolderAdd As String) As Long
    mHeader.FolderCount = mHeader.FolderCount + 1
    ReDim Preserve mHeader.FolderRecords(1 To mHeader.FolderCount)
    With mHeader.FolderRecords(mHeader.FolderCount)
    
        .FolderName = FolderAdd
        .FolderNameLength = Len(.FolderName)
    End With
    AddFolderRecord = mHeader.FolderCount
End Function
Public Sub AddFile(ByVal SourceFile As String, ByVal Filename As String, ByVal ToFolder As String)
    'adds this file to the package in ToFolder.
    'First: find the folder:
    
    Dim folderindex As Long, FileIndex As Long
    folderindex = FindFolderRecordIndex(ToFolder)
    If folderindex = -1 Then
        folderindex = AddFolderRecord(ToFolder)
    
    End If
    'now that we are certain the folder record exists, check to see if the file is already in there.
    FileIndex = FindFileRecordIndex(folderindex, Filename)
    If FileIndex = -1 Then
        FileIndex = AddFileRecord(folderindex, Filename)
    
        With mHeader.FolderRecords(folderindex).FileRecords(FileIndex)
            .SourceFile = SourceFile
            
            
        'no actual "packing" or anything yet.
        End With
    Else
        Err.Raise 58, "CHeirarchal", "File """ & Filename & """ already exists in package folder """ & ToFolder & """."
    
    End If
    
    
 
End Sub
Public Sub CreateFolder(ByVal FolderName As String)

    Dim searchfold As Long
    searchfold = FindFolderRecordIndex(FolderName)
    If searchfold <> 0 Then
        Err.Raise 58, "CHeirarchalFilePackage", "Folder " & FolderName & " already exists."
    End If
    mHeader.FolderCount = mHeader.FolderCount + 1
    ReDim Preserve mHeader.FolderRecords(1 To mHeader.FolderCount)
    mHeader.FolderRecords(mHeader.FolderCount).FolderName = FolderName
    mHeader.FolderRecords(mHeader.FolderCount).FolderNameLength = Len(FolderName)




End Sub
Private Function InstrCount(ByVal SearchIn As String, ByVal SearchFor As String) As Long
    InstrCount = (Len(SearchIn) - Len(Replace$(SearchIn, SearchFor, ""))) / Len(SearchFor)
End Function
Public Function GetSubFolderIndices(ByVal ParentFolderIndex As Long, ByRef lcount As Long) As Long()
    Dim Ret() As Long
    Dim currcount As Long, parentstr As String
    Dim i As Long, currfolder As String, currsplit() As String
    For i = 1 To mHeader.FolderCount
        currfolder = mHeader.FolderRecords(i).FolderName
        parentstr = mHeader.FolderRecords(ParentFolderIndex).FolderName
        If Left$(currfolder, Len(parentstr)) = parentstr Then
            If InstrCount(currfolder, "\") = (InstrCount(parentstr, "\") - 1) Then
                currcount = currcount + 1
                ReDim Preserve Ret(1 To currcount)
                Ret(currcount) = i
            End If
        End If
        
    
    Next i
    
    
    
    
    
    




End Function


Public Sub RemoveFile(ByVal FolderName As String, ByVal Filename As String)
    Dim folderremove As Long, fileremove As Long
    folderremove = FindFolderRecordIndex(FolderName)
    fileremove = FindFileRecordIndex(folderremove, Filename)
    RemoveFileByIndex folderremove, fileremove
End Sub
Public Sub RemoveFileByIndex(ByVal folderindex As Long, FileIndex As Long)
    Dim i As Long
    With mHeader.FolderRecords(folderindex)
    'copy over the old...
        For i = FileIndex To .FileCount - 1
            .FileRecords(i) = .FileRecords(i + 1)
        Next
        .FileCount = .FileCount - 1
        If .FileCount = 0 Then
            Erase .FileRecords
        Else
            ReDim Preserve .FileRecords(1 To .FileCount)
        End If
    End With
End Sub
Public Sub RemoveFolderByIndex(ByVal folderindex As Long)
    Dim i As Long
    With mHeader
    'copy the item above to the current item; overwrite the folder structure we are removing entirely.
        For i = folderindex To .FolderCount - 1
            .FolderRecords(i) = .FolderRecords(i + 1)
        Next
    
        .FolderCount = .FolderCount - 1
        If .FolderCount = 0 Then
            Erase .FolderRecords
        Else
            ReDim Preserve .FolderRecords(1 To .FolderCount)
        End If
    
    End With

End Sub
Private Sub ReadFromStreamInt(InStream As IInputStream)
    Dim magicnum As Long, LoopFolder As Long, LoopFile As Long
    'erase ALL entries...
    Erase mHeader.FolderRecords
    magicnum = ReadLong(InStream)
    If magicnum <> MAGIC_NUMBER Then
        Err.Raise 321, "CHeirarchalFilePackage::ReadFromStream", "Magic Number not found."
    End If
    mHeader.FolderCount = ReadLong(InStream)
    ReDim mHeader.FolderRecords(1 To mHeader.FolderCount)
    For LoopFolder = 1 To mHeader.FolderCount
        With mHeader.FolderRecords(LoopFolder)
            .FolderNameLength = ReadLong(InStream)
            .FolderName = ReadString(InStream, .FolderNameLength, StrRead_unicode)
            'read the file count...
            .FileCount = ReadLong(InStream)
            ReDim .FileRecords(1 To .FileCount)
            For LoopFile = 1 To .FileCount
                With .FileRecords(LoopFile)
                'dataoffset, datasize.
                .DataOffset = ReadLong(InStream)
                .datasize = ReadLong(InStream)
               .FilenameLength = ReadLong(InStream)
               .Filename = ReadString(InStream, .FilenameLength, StrRead_unicode)
               End With
            Next LoopFile
        End With
    Next LoopFolder
    Set mCachedInputStream = InStream
End Sub
Public Sub ReadFromStream(InStream As IInputStream)
    '//read in header data here, to determine how to proceed with the remainder.
'     ModStreamFunc.WriteLong Outstream, Len(mHeader.FilterMethod)
'    ModStreamFunc.WriteString Outstream, mHeader.FilterMethod
'    'length of data...
'    ModStreamFunc.WriteLong Outstream, tempfiltered.Size
Dim datasize As Long, StreamreadTo As CMemStream, Streamfiltered As CMemStream

'    ModStreamFunc.WriteLong Outstream, Len(mHeader.FilterMethod)
'
'    ModStreamFunc.WriteString Outstream, mHeader.FilterMethod
'    ModStreamFunc.WriteLong Outstream, Len(mHeader.FilterParam)
'    ModStreamFunc.WriteString Outstream, mHeader.FilterParam
'    'length of data...
'    ModStreamFunc.WriteLong Outstream, tempfiltered.Size

    Dim filterparamlen As Long
    mHeader.FilterMethodNameLength = ReadLong(InStream)
    If mHeader.FilterMethodNameLength > 0 Then
        mHeader.FilterMethod = ReadStringAuto(InStream, mHeader.FilterMethodNameLength)
    End If
    '"Filterparam" field.
    
    'mheader.FilterParam
    filterparamlen = ReadLong(InStream)
    mHeader.FilterParam = ReadStringAuto(InStream, filterparamlen)
    
    
    'hooray...
    datasize = ReadLong(InStream)
    
    
    'now, read "datasize" bytes from Instream, construct a Corefilters object, set the appropriate parameters based on the values now in mheader,
    'perform the filter, seek to the beginning of the resulting stream and send it off to the internal read function.
    Set StreamreadTo = New CMemStream
    StreamreadTo.WriteBytes InStream.readbytes(datasize)
    'seek to the start.
    StreamreadTo.SeekTo 0, STREAM_BEGIN
    
    Dim FilterObject As CCoreFilters, FilterIf As IStreamFilter
    
    
    Set FilterObject = New CCoreFilters
    Set Streamfiltered = New CMemStream
    If UCase$(mHeader.FilterMethod) = "NONE" Or UCase$(mHeader.FilterMethod) = "" Then
    
        'no filter... just write it directly....
        Set Streamfiltered = StreamreadTo
    Else
        If UCase$(mHeader.FilterMethod) = "AES" Then
            FilterObject.FilterType = Decrypt_CryptoAPI
            FilterObject.password = mHeader.FilterParam
            
        
        ElseIf UCase$(mHeader.FilterMethod) = "LZW" Then
            FilterObject.FilterType = LZW_Expand
        ElseIf UCase$(mHeader.FilterMethod) = "LZSS" Then
            FilterObject.FilterType = LZSS_Expand
        ElseIf UCase$(mHeader.FilterMethod) = "HUFF" Then
            FilterObject.FilterType = Huffman_Expand
    
        End If
        'cast to interface type...
        Set FilterIf = FilterObject
        'lastly... perform the operation... finally..
        FilterIf.FilterStream StreamreadTo, Streamfiltered
        
        
    End If
    
    ReadFromStreamInt Streamfiltered
    

    
End Sub
Public Sub WriteToStream(Outstream As IOutputStream)
    'write to a memory stream first...
    Dim Tempstream As CMemStream, tempfiltered As CMemStream
    Dim FilterObj As CCoreFilters, FilterStream As IStreamFilter
    Set Tempstream = New CMemStream
    Set FilterObj = New CCoreFilters
    
    '//write out the header data here.
    '//basically,
    
    
    If UCase$(mHeader.FilterMethod) = "NONE" Or mHeader.FilterMethod = "" Then
    'write it directly...
    Set tempfiltered = Tempstream
     WriteToStreamInt tempfiltered
    Else
        If UCase$(mHeader.FilterMethod) = "LZW" Then
            FilterObj.FilterType = LZW_Compress
        ElseIf UCase$(mHeader.FilterMethod) = "LZSS" Then
            FilterObj.FilterType = LZSS_Compress
        ElseIf UCase$(mHeader.FilterMethod) = "AES" Then
            FilterObj.FilterType = Encrypt_CryptoAPI
            FilterObj.password = mHeader.FilterParam
        ElseIf UCase$(mHeader.FilterMethod) = "HUFF" Then
            FilterObj.FilterType = Huffman_Compress
        
        End If
        'now, create a temporary output... oh, we did! use the memory stream.
        'write to the memory stream with our write routine...
        WriteToStreamInt Tempstream
        Tempstream.SeekTo 0, STREAM_BEGIN
        
        
        'filterobj's parameters are set above... now, we simply "engage" the filter, heh.
        Set FilterStream = FilterObj
        
        Set tempfiltered = New CMemStream
        FilterStream.FilterStream Tempstream, tempfiltered
        Tempstream.CloseStream
        
    End If
    'for the other types, we store a short descriptor and the compressed/encrypted stream data.
    'that will be read into another stream during the read operation before the internal read routine is called.
    
    'write header....
    
    
    ModStreamFunc.WriteLong Outstream, Len(mHeader.FilterMethod)
    
    ModStreamFunc.WriteString Outstream, mHeader.FilterMethod
    ModStreamFunc.WriteLong Outstream, Len(mHeader.FilterParam)
    ModStreamFunc.WriteString Outstream, mHeader.FilterParam
    'length of data...
    ModStreamFunc.WriteLong Outstream, tempfiltered.Size
    
    'and finally,the filtered stream data.
    ModStreamFunc.WriteStream Outstream, tempfiltered

End Sub
Private Sub WriteToStreamInt(Outstream As IOutputStream)
    Dim LoopFolder As Long, LoopFile As Long
    
    WriteLong Outstream, MAGIC_NUMBER
    WriteLong Outstream, mHeader.FolderCount
    For LoopFolder = 1 To mHeader.FolderCount
    
    'write out the folder data first:
    With mHeader.FolderRecords(LoopFolder)
        WriteLong Outstream, .FolderNameLength
        Outstream.WriteString .FolderName, StrRead_unicode
        'write out the file count... and then the files...
        WriteLong Outstream, .FileCount
        For LoopFile = 1 To .FileCount
            With .FileRecords(LoopFile)
                'we can't write the offset, since we have to freaking clue where it's going to be stored.
                'write a "placeholder" that we will seek to later.
                'I only put it first to make the seeking easier later. less room for error.
                .SeekToOffset = Outstream.GetPos
                WriteLong Outstream, 0
                WriteLong Outstream, 0
                WriteLong Outstream, .FilenameLength
                WriteString Outstream, .Filename, StrRead_unicode
                
                
            
            
            End With
        Next LoopFile
        
        
        
    End With
    
    Next LoopFolder
    
Dim StreamSlurp As FileStream, temppos As Long
'Now, we loop through again. this time, we open the source file, and slurp it out.
For LoopFolder = 1 To mHeader.FolderCount
    With mHeader.FolderRecords(LoopFolder)
        'loop through all the files.
        For LoopFile = 1 To .FileCount
            'for every one- we open the source file stream.
            'if the source file is "", then chances are this is an opened package file (in which case the original source might not exist anyway).
            Set StreamSlurp = FileSystem.OpenStream(.FileRecords(LoopFile).SourceFile)
            
            
            
            
            .FileRecords(LoopFile).DataOffset = Outstream.GetPos
            .FileRecords(LoopFile).datasize = StreamSlurp.Size
            WriteStream Outstream, StreamSlurp
            'alright, save the current position...
            temppos = Outstream.GetPos
            'seek to the location to save the data offset and size...
            Outstream.SeekTo CDbl(.FileRecords(LoopFile).SeekToOffset)
            WriteLong Outstream, .FileRecords(LoopFile).DataOffset
            WriteLong Outstream, .FileRecords(LoopFile).datasize
            Outstream.SeekTo CDbl(temppos)
    
    
        Next LoopFile
    End With

Next






End Sub
Public Function GetSubFolderCount(ByVal StrParentFolder As String) As Long

Dim GotIndices() As Long, folderindex As Long, lcount As Long
folderindex = FindFolderRecordIndex(StrParentFolder)
GotIndices = GetSubFolderIndices(folderindex, lcount)

GetSubFolderCount = lcount

End Function
Public Sub Expand(ByVal PackageFolder As String, ByVal PackedFile As String, ByVal DestinationFile As String)
    Dim foundindex As Long
    Dim foundfileindex As Long, readbytes() As Byte
    foundindex = FindFolderRecordIndex(PackageFolder)
    If foundindex = -1 Then
        Err.Raise 53, "CHeirarchalFilePackage::Expand", "Invalid Folder, """ & PackageFolder & """."
    End If
    foundfileindex = FindFileRecordIndex(foundindex, PackedFile)
    If foundfileindex = -1 Then
        Err.Raise 53, "CHeirarchalFilePackage::Expand", "Invalid packed filename, """ & PackedFile & """."
    End If
    If mCachedInputStream Is Nothing Then
        'chances are... the source file still exists, so simply copy it to the destination file.
        FileCopy mHeader.FolderRecords(foundindex).FileRecords(foundfileindex).SourceFile, DestinationFile
    Else
        Dim outputstream As FileStream
        With mHeader.FolderRecords(foundindex).FileRecords(foundfileindex)
            mCachedInputStream.SeekTo CDbl(.DataOffset)
            Set outputstream = FileSystem.CreateStream(DestinationFile)
            readbytes = mCachedInputStream.readbytes(.datasize)
            outputstream.WriteBytes readbytes
            outputstream.CloseStream
        End With
        
    
    End If




End Sub

Private Sub Class_Initialize()
mHeader.FilterMethod = ""
mHeader.FilterMethodNameLength = 0
End Sub

Private Sub Class_Terminate()
If Not mCachedInputStream Is Nothing Then
mCachedInputStream.CloseStream
End If
End Sub
