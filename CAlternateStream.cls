VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlternateStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CAlternateStream.cls
Implements IStreamable
'represents a single Stream within a file/folder.

Private mvarStreamName As String 'name of this stream.
Private mvarFilePath As String 'path to the file this stream resides in.
Private mvarStreamAttributes As Long
Private mvarStreamSizeLow As Long, MvarStreamSizeHigh As Long
Private mInvalidated As Boolean

Private mStreamInfo As BY_HANDLE_FILE_INFORMATION


Private Sub Checkvalid()
        If mInvalidated Then
        Err.Raise 5, "CalternateStream:Checkvalid", "Stream """ & mvarStreamName & """ in file """ & mvarFilePath & """ has been deleted."
    
    End If
End Sub
Friend Sub Init(ByVal Infile As String, ByVal StreamName As String, ByVal StreamSizeLow As Long, ByVal StreamSizeHigh As Long, Optional ByVal StreamAttributes As Long)
    mvarFilePath = Infile
    mvarStreamName = StreamName
    mvarStreamSizeLow = StreamSizeLow
    MvarStreamSizeHigh = StreamSizeHigh
    mvarStreamAttributes = StreamAttributes
    
    
    Dim hstream As Long
    hstream = CreateFile(Infile & ":" & StreamName, GENERIC_READ, FILE_SHARE_READ + FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If hstream <= 0 Then
    
    Else
    
        GetFileInformationByHandle hstream, mStreamInfo
    
    'added:closehandle call.
    
        CloseHandle hstream
    End If
    
    
End Sub
Public Property Get Attributes()
Attribute Attributes.VB_Description = "Returns the attributes of this stream."
    Checkvalid
    Attributes = mvarStreamAttributes
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "The name of the stream."
    Checkvalid
    Name = mvarStreamName

End Property
Public Property Get FilePath() As String
Attribute FilePath.VB_Description = "Returns the path of the file this stream resides in."
    Checkvalid
    FilePath = mvarFilePath
End Property
Public Property Get Size() As Double
Attribute Size.VB_Description = "the size, in bytes, of this stream."
    Checkvalid
    Size = LI2DBL(mvarStreamSizeLow, MvarStreamSizeHigh)
End Property
Public Function OpenInputStream(Optional ByVal SharingDisposition As FileShareModeConstants = FILE_SHARE_READ, Optional ByVal Fileflags As FileAttributeConstants = 0) As FileStream
    
    Set OpenInputStream = Me.OpenAsBinaryStream(GENERIC_READ, SharingDisposition, OPEN_EXISTING, Fileflags)
End Function

Public Function OpenAsBinaryStream(Optional ByVal dwDesiredAccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal dwShareMode As FileShareModeConstants = FILE_SHARE_READ + FILE_SHARE_WRITE, Optional ByVal dwCreationDisposition As CreationDispositionConstants = OPEN_EXISTING, Optional ByVal dwFlagsAndAttributes As FileAttributeConstants) As FileStream
    Dim retthis As FileStream
    Set retthis = New FileStream
    retthis.OpenStream IStreamable_Fullpath, dwDesiredAccess, dwShareMode, dwCreationDisposition, dwFlagsAndAttributes
    Set OpenAsBinaryStream = retthis


End Function

Private Sub Class_Initialize()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub

Private Sub Class_Terminate()




MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub

Private Property Get IStreamable_Fullpath() As String
If Left$(mvarStreamName, 1) <> ":" Then
    IStreamable_Fullpath = mvarFilePath & ":" & mvarStreamName
Else
    IStreamable_Fullpath = mvarFilePath & mvarStreamName
End If
End Property
Public Sub Delete()
Attribute Delete.VB_Description = "Deletes this stream from the file."
        Dim retval As Long
        'Deletes this stream. Invalidates future access to the stream with "this object has been deleted" error.
        retval = DeleteFile(IStreamable_Fullpath)
        If retval = 0 Then
            RaiseAPIError Err.LastDllError, "CAlternateStream:Delete"
        Else
            'success
            mInvalidated = True
        
        
        End If
End Sub

Private Property Get IStreamable_Name() As String
'
IStreamable_Name = Me.Name
End Property

Private Function IStreamable_OpenInputStream() As IInputStream
    '
    Dim fOpen As FileStream
    Set fOpen = Me.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_DELETE + FILE_SHARE_READ + FILE_SHARE_WRITE)
    Set IStreamable_OpenInputStream = fOpen
End Function

Private Function IStreamable_OpenOutputStream() As IOutputStream
'
    Dim fOpen As FileStream
    Set fOpen = Me.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_DELETE)
    Set IStreamable_OpenOutputStream = fOpen

End Function

Private Property Get IStreamable_Size() As Long
IStreamable_Size = Me.Size
End Property
