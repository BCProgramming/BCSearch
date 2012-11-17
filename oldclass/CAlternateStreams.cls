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
Private mCol As collection
Public Function Count() As Long
    Count = mCol.Count
End Function
Friend Property Let Owner(ByVal vdata As String)
    mOwnerItem = vdata
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
buildPath = mOwnerItem
buildPath = mOwnerItem & ":" & StreamName

hFile = CreateFile(StreamName, GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_NEW, 0, 0)
If hFile <= 0 Then
    RaiseAPIError Err.LastDllError, "CAlternateStreams::Create()"
Else
    CloseHandle hFile
    Set newaltstream = New CAlternateStream
    newaltstream.Init buildPath, StreamName, 0, 0, 0
    Set Create = newaltstream

End If



End Function

Friend Sub Add(AlternateStream As CAlternateStream)

    mCol.Add AlternateStream, "KEY" & AlternateStream.Name



End Sub
Public Function Item(ByVal VntIndexKey As String) As CAlternateStream
Attribute Item.VB_UserMemId = 0
'to access a stream, for example, named "hidden" one would pass in:

'.Item (":hidden:$DATA") to retrieve the CAlternateStream class instance.
Set Item = mCol.Item("KEY" & VntIndexKey)

End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Function
Private Sub Class_Initialize()
Set mCol = New collection
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub
