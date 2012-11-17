VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlternateStreams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'"AlternateStreams" collection.
Private mOwnerItem As String 'Path to file/dir that these streams belong to.
Private mCol As Collection
Public Function count() As Long
    count = mCol.count
End Function
Friend Property Let Owner(ByVal Vdata As String)
    mOwnerItem = Vdata
End Property
Public Property Get Owner() As String
    Owner = mOwnerItem
End Property
Public Function CreateStream(ByVal ADSName As String) As CAlternateStream
    'Creates and opens a new stream inside this object/file...
    Dim newads As CAlternateStream, fs As FileStream
    Set newads = New CAlternateStream
    newads.Init mOwnerItem, ADSName, 0, 0, 0
    'open it and close it right away...
    Set fs = newads.OpenAsBinaryStream(GENERIC_WRITE, FILE_SHARE_READ + FILE_SHARE_DELETE, CREATE_NEW, 0)
    fs.CloseStream
    
    'ok...
    'now return the ads object.
    Add newads
    Set CreateStream = newads



End Function

Public Function Create(ByVal StreamName As String) As CAlternateStream

'hee hee... use CreateFile API...
Dim buildPath As String, hFile As Long
Dim newaltstream As CAlternateStream
If mOwnerItem = "" Then Exit Function
buildPath = mOwnerItem
buildPath = mOwnerItem & ":" & StreamName

hFile = CreateFile(buildPath, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_NEW, 0, 0)
If hFile <= 0 Then
    RaiseAPIError Err.LastDllError, "CAlternateStreams::Create()"
Else
    CloseHandle hFile
    Set newaltstream = New CAlternateStream
    newaltstream.Init mOwnerItem, StreamName, 0, 0, 0
    Set Create = newaltstream

End If



End Function

Friend Sub Add(AlternateStream As CAlternateStream)

    mCol.Add AlternateStream, "KEY" & AlternateStream.Name



End Sub
Public Function Item(ByVal vntIndexkey As Variant) As CAlternateStream
Attribute Item.VB_UserMemId = 0
'to access a stream, for example, named "hidden" one would pass in:
If IsNumeric(vntIndexkey) Then
    Set Item = mCol.Item(vntIndexkey)
Else
    '.Item (":hidden:$DATA") to retrieve the CAlternateStream class instance.
    Set Item = mCol.Item("KEY" & vntIndexkey)
End If
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function
Private Sub Class_Initialize()
Set mCol = New Collection
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub
