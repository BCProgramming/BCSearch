VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Change()
'Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
'Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long

Private mObjectPath As String
Private mOwner As Object

Public Function ToString(Optional Longform As Boolean = False) As String


    ToString = GetAttributeString(Me.Attributes, Longform)



End Function

Public Property Get Owner() As Object
    Set Owner = mOwner
End Property



Friend Sub Init(Owner As Object)
'
    Dim strpath As String
    Set mOwner = Owner
    If TypeOf mOwner Is CFile Then
    strpath = mOwner.Fullpath
    Else
    strpath = mOwner.Path
    End If
    mObjectPath = strpath
End Sub
Public Property Get Attributes() As FileAttributeConstants
Attribute Attributes.VB_UserMemId = 0
'

    Attributes = MdlFileSystem.GetFileAttributes(mObjectPath)
End Property
Public Property Let Attributes(ByVal Vdata As FileAttributeConstants)
    
    MdlFileSystem.SetFileAttributes mObjectPath, Vdata
    RaiseEvent Change
End Property

'Helper routines: one for each attribute.
Private Sub RemoveAttributeInt(AttribRemove As FileAttributeConstants)
    'removes the specified attribute from the file if present.
    If HasAttribute(AttribRemove) Then Me.Attributes = Me.Attributes Xor AttribRemove
End Sub
Private Function HasAttribute(AttribTest As FileAttributeConstants)
    HasAttribute = ((Me.Attributes And AttribTest) = AttribTest)
End Function
Private Sub SetAttribute(ByVal AttributeMask As FileAttributeConstants, ByVal Vdata As Boolean)
    Dim gotattr As FileAttributeConstants
    gotattr = Me.Attributes
    If (gotattr And AttributeMask) = AttributeMask Then
        Me.Attributes = gotattr - AttributeMask
    
    End If
End Sub

Public Property Get System() As Boolean
    System = HasAttribute(FILE_ATTRIBUTE_SYSTEM)
End Property
Public Property Let System(Vdata As Boolean)
    SetAttribute FILE_ATTRIBUTE_SYSTEM, Vdata
End Property

Public Property Get Hidden() As Boolean
    Hidden = HasAttribute(FILE_ATTRIBUTE_HIDDEN)
End Property
Public Property Let Hidden(Vdata As Boolean)
    SetAttribute FILE_ATTRIBUTE_HIDDEN, Vdata

End Property
Public Property Get Encrypted() As Boolean
    Encrypted = HasAttribute(FILE_ATTRIBUTE_ENCRYPTED)
End Property
Public Property Let Encrypted(ByVal Vdata As Boolean)
    If Vdata And Not Me.Encrypted Then
        EncryptFile mObjectPath
    ElseIf Me.Encrypted Then
        DecryptFile mObjectPath, 0
    End If
    
End Property
Public Property Get Compressed() As Boolean
    Compressed = HasAttribute(FILE_ATTRIBUTE_COMPRESSED)
End Property
Public Property Let Compressed(ByVal Vdata As Boolean)
    'SetAttribute FILE_ATTRIBUTE_COMPRESSED, vdata
    Dim FSCTL_SET_COMPRESSION As Long
    Dim lhfile As Long, lbytesrtn As Long
    Dim devret As Long
    Const FILE_READ_DATA As Long = &H1&
    Const FILE_WRITE_DATA As Long = &H2&
    Const FILE_DEVICE_FILE_SYSTEM As Long = &H9&
    Const METHOD_BUFFERED As Long = 0&
    Const COMPRESSION_FORMAT_NONE As Long = 0&
    Const COMPRESSION_FORMAT_DEFAULT As Long = 1&
    '639040
    FSCTL_SET_COMPRESSION = (CLng(FILE_DEVICE_FILE_SYSTEM) * (2 ^ 16)) Or _
    (CLng(FILE_READ_DATA Or FILE_WRITE_DATA) * (2 ^ 14)) Or (CLng(16) * (2 ^ 2)) Or _
    METHOD_BUFFERED

    lhfile = CreateFile(mObjectPath, GENERIC_ALL, FILE_SHARE_WRITE And _
    FILE_SHARE_READ, 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)

    If lhfile <> INVALID_HANDLE_VALUE Then
        If Vdata Then
            'devret = DeviceIoControlAny(lhfile, FSCTL_SET_COMPRESSION, ByVal COMPRESSION_FORMAT_DEFAULT, &H2, ByVal 0&, 0, bytesreturn, ByVal 0&)
                        devret = DeviceIoControlAny(lhfile, FSCTL_SET_COMPRESSION, _
            COMPRESSION_FORMAT_DEFAULT, 2&, 0&, 0&, lbytesrtn, ByVal 0&)

        Else
            devret = DeviceIoControlAny(lhfile, FSCTL_SET_COMPRESSION, ByVal COMPRESSION_FORMAT_NONE, &H2, ByVal 0&, 0, lbytesrtn, ByVal 0&)
        End If
            
'                        CompressFile = DeviceIoControl(lFileHwnd, FSCTL_SET_COMPRESSION, _
            COMPRESSION_FORMAT_DEFAULT, 2&, 0&, 0&, lBytesRtn, 0&)
    CloseHandle lhfile
    
    End If
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = HasAttribute(FILE_ATTRIBUTE_READONLY)
End Property
Public Property Let ReadOnly(Vdata As Boolean)
    SetAttribute FILE_ATTRIBUTE_READONLY, Vdata

End Property
Public Property Get Archive() As Boolean
    Archive = HasAttribute(FILE_ATTRIBUTE_ARCHIVE)
End Property
Public Property Let Archive(Vdata As Boolean)
    SetAttribute FILE_ATTRIBUTE_ARCHIVE, Vdata
End Property
