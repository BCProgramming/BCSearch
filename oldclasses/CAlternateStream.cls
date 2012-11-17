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
        Err.Raise 5, "CalternateStream:Attributes", "Stream """ & mvarStreamName & """ in file """ & mvarFilePath & """ has been deleted."
    
    End If
End Sub
Friend Sub Init(ByVal InFile As String, ByVal StreamName As String, ByVal StreamSizeLow As Long, ByVal StreamSizeHigh As Long, Optional ByVal StreamAttributes As Long)
    mvarFilePath = InFile
    mvarStreamName = StreamName
    mvarStreamSizeLow = StreamSizeLow
    MvarStreamSizeHigh = StreamSizeHigh
    mvarStreamAttributes = StreamAttributes
    
    
    Dim hstream As Long
    hstream = CreateFile(InFile & ":" & StreamName, GENERIC_READ, FILE_SHARE_READ + FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
    If hstream <= 0 Then
    
    Else
    
        GetFileInformationByHandle hstream, mStreamInfo
    
    
    End If
    
    
End Sub
Public Property Get Attributes()
Attribute Attributes.VB_Description = "Returns the attributes of this stream."
    Checkvalid
    Attributes = mvarStreamAttributes
End Property
Public Property Get name() As String
Attribute name.VB_Description = "The name of the stream."
    Checkvalid
    name = mvarStreamName

End Property
Public Property Get FilePath() As String
Attribute FilePath.VB_Description = "Returns the path of the file this stream resides in."
    Checkvalid
    FilePath = mvarFilePath
End Property
Public Property Get size() As Double
Attribute size.VB_Description = "the size, in bytes, of this stream."
    Checkvalid
    size = LI2DBL(mvarStreamSizeLow, MvarStreamSizeHigh)
End Property
Public Function OpenAsBinaryStream(Optional ByVal dwDesiredAccess As DesiredAccessFlags = GENERIC_ALL, Optional ByVal dwShareMode As FileShareModeConstants = FILE_SHARE_READ + FILE_SHARE_WRITE, Optional ByVal dwCreationDisposition As CreationDispositionConstants = OPEN_EXISTING, Optional ByVal dwFlagsAndAttributes As FileAttributeConstants) As FileStream
Attribute OpenAsBinaryStream.VB_Description = "Opens this alternate Stream as a Binary Stream."
    Dim retthis As FileStream
    Set retthis = New FileStream
    retthis.OpenStream IStreamable_Fullpath, dwDesiredAccess, dwShareMode, dwCreationDisposition, dwFlagsAndAttributes
    Set OpenAsBinaryStream = retthis
 
 
End Function

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
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
        Dim retVal As Long
        'Deletes this stream. Invalidates future access to the stream with "this object has been deleted" error.
        retVal = DeleteFile(IStreamable_Fullpath)
        If retVal = 0 Then
            RaiseAPIError Err.LastDllError, "CAlternateStream:Delete"
        Else
            'success
            mInvalidated = True
        
        
        End If
End Sub
Private Property Get IStreamable_Size() As Long
IStreamable_Size = Me.size
End Property
