VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirWalker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'CFileWalker
'Directory Walker Class

'The Directory Walker Class can be used to walk all the objects in a Directory, just subdirectories, or just files. It also supports the use of File Specifications.
'will be used for implementation of custom IEnumVARIANT of Files and Directories collections.
Implements IVariantEnumerable

Private mFindData As WIN32_FIND_DATA  'last find data

Private attributemask As FileAttributeConstants
Private mExcludeAttr As FileAttributeConstants
Private mFindHandle As Long
    
Private mDirConnect As Directory 'the directory we are Enumerating.
Public Event ItemFound(FoundItem As Object, ByRef Exclude As Boolean)   'event to allow filtering.
'Private data used during the "walk"
Private mUseAttr As FileAttributeConstants, FileSpec As String
Friend Property Get FSpec() As String
    FSpec = FileSpec
End Property
'Friend Function GetIEnum() As stdole.IEnumVARIANT
'    'Dim En As EnumVARIANT
'    Dim tmp As IVariantEnumerable
'    Set tmp = Me
'    Set En = New EnumVARIANT
'    En.Init tmp
'    Set GetIEnum = En
'
'    Set tmp = Nothing
'End Function

'Public Sub InitWalker(ByVal InDirectory As Directory, Optional ByVal MustHaveAttr As FileAttributeConstants, Optional ByVal FileSpec As String = "*")
Friend Sub InitWalker(ByVal InDirectory As Directory, Optional ByVal FSpec As String = "*.*", Optional ByVal Attributes As FileAttributeConstants = FILE_ATTRIBUTE_NORMAL, _
    Optional ByVal ExcludeAttr As FileAttributeConstants)
    CDebug.Post "InitWalker:" & InDirectory.Path
    'Initializes this walker object to begin searching.
    If Not mDirConnect Is Nothing Then
        If Not mFindHandle = 0 Then
            CDebug.Post "closing existing findhandle," & mFindHandle
            FindClose mFindHandle
            
            
        End If
        Set mDirConnect = Nothing
    End If
    'zero it out- erase any previous search.
    Set mDirConnect = InDirectory
    CDebug.Post "ZeroMemory mFindData, Len(mFindData)"
    ZeroMemory mFindData, Len(mFindData)
    FileSpec = FSpec
    mExcludeAttr = ExcludeAttr
    'mUseAttr = MustHaveAttr
    'FileSpec = FileSpec
    'mFindData.cFileName = InDirectory.Path & FileSpec & vbNullChar & vbNullChar
    attributemask = Attributes
    
    'FindFirstFile InDirectory.Path & FileSpec, mFindData
    CDebug.Post "Leaving InitWalker."
End Sub

'Public Function GetNext() As Object
'    'returns: a Directory or a CFile, could be either.
'    'a return value of Nothing means there are no more items.
'    'Items can be filtered by sinking events.
'    Dim findnext As Long, flMeetsCondition
'    Dim FullFileName As String, fexclude As Boolean, grabobject As Object
'    ZeroMemory mFindData, Len(mFindData)
'    On Error GoTo ReportError
'FindAnother:
'    'hey- goto works here, what the hell.
'    ZeroMemory mFindData, Len(mFindData)
'    mFindData.dwFileAttributes = attributemask
'    If mFindHandle = 0 Then
'        mFindHandle = FindFirstFileA(mDirConnect.Path & FileSpec & vbNullChar & vbNullChar, mFindData)
'    Else
'         findnext = FindNextFileA(mFindHandle, mFindData)
'         If findnext = ERROR_NO_MORE_FILES Then
'            'no more files
'            'FindClose mFindHandle
'            Exit Function
'
'         End If
'    End If
'
'    '///TASK:
'    'Convert mfinddata structure data into appropriate filename/file path, and create a CFile or CDirectory Object, depending on it's actual type.
'    FullFileName = mFindData.cFileName
'    FullFileName = Trim$(mFindData.cFileName)
'
'
'
'
'    FullFileName = Replace$(FullFileName, vbNullChar, "")
'    If FullFileName = "" Then Exit Function
'    FullFileName = mDirConnect.Path & FullFileName
'    If InStr(FullFileName, "config.msi") Then Debug.Assert False
'    '///convert to appropriate object-
'    'use attributes to determine wether it is a directory or file- also mask out those not wanted by the client(via attributemask)
'    If (mFindData.dwFileAttributes And mExcludeAttr) = mExcludeAttr And mExcludeAttr <> 0 Then
'        Debug.Print "excluding " & FullFileName, " matches attribute " & mExcludeAttr
'        GoTo FindAnother
'    Else
'        'determine wether to return it as well as wether it is a folder or file...
'        If (mFindData.dwFileAttributes And attributemask) = attributemask Or (attributemask = 0) Then
'            If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
'                'it's a dir. return a dir.
'                If Right$(FullFileName, 1) = "." Or Right$(FullFileName, 2) = ".." Then
'                    GoTo FindAnother    'go find another one.
'                Else
'                    Set grabobject = FileSystem.GetDirectory(FullFileName)
'
'                End If
'            Else
'                'it is not a dir.
'                If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DEVICE) <> FILE_ATTRIBUTE_DEVICE Then
'                    'if it isn't a device, then...
'                    Set grabobject = FileSystem.GetFile(FullFileName)
'                End If
'            End If
'
'
'        End If
'
'
'
'
'
''        If ((mFindData.dwFileAttributes And attributemask) = attributemask Or (attributemask <> 0)) And _
''           Right$(FullFileName, 1) <> "." And Right$(FullFileName, 2) <> ".." Then
''            If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
''                'it is a directory, Convert to CDirectory.
''                Set grabobject = FileSystem.GetDirectory(FullFileName)
''            Else
''                'OK: if our attribute mask include directory, we can (hopefully) assume that  there will be no more directories. and as such break out.
''                If attributemask And FILE_ATTRIBUTE_DIRECTORY = FILE_ATTRIBUTE_DIRECTORY Then
''
''                End If
''                If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DEVICE) <> FILE_ATTRIBUTE_DEVICE Then
''                    'ignore device stuff here...
''                    'convert to CFile.
''                    Set grabobject = FileSystem.GetFile(FullFileName)
''
''
''
''                End If
''            End If
''        Else
''            'Debug.Print FullFileName
''            GoTo FindAnother
''        End If
'
'    End If
'    fexclude = False
'    RaiseEvent ItemFound(grabobject, fexclude)
'    If fexclude Then
'        GoTo FindAnother
'    Else
'        'Debug.Print "returning object type " & TypeName(grabobject)
'        Set GetNext = grabobject
'        'woopee!
'    End If
'    Exit Function
'ReportError:
'    DebugMsg Err.Description & " in GetNext."
'    'FindClose mFindHandle
'End Function
'
'
'
'
'
Public Function GetNext(Optional ByRef SetThis As Object) As Object
'rewritten March 23rd 2009 to remove GOTO abuse.
'DebugMsg "GetNext"
CDebug.Post "GetNext"


    'returns: a Directory or a CFile, could be either.
    'a return value of Nothing means there are no more items.
    'Items can be filtered by sinking events.
    Dim findnext As Long, flMeetsCondition
    Dim FullFileName As String, grabobject As Object
    Dim FoundMatch As Boolean
    ZeroMemory mFindData, Len(mFindData)
    'FindAnother:
    On Error GoTo showError
    Do Until FoundMatch

        ZeroMemory mFindData, Len(mFindData)
        mFindData.dwFileAttributes = attributemask
        If mFindHandle = 0 Then
            'Debug.Print "finding " & FileSpec & " in " & mDirConnect.Path
            mFindHandle = FindFirstFile(mDirConnect.Path & FileSpec & vbNullChar & vbNullChar, mFindData)
        Else

            findnext = FindNextFile(mFindHandle, mFindData)
            If findnext = ERROR_NO_MORE_FILES Then
                'no more files
                FindClose mFindHandle
                mFindHandle = 0
                Exit Function

            End If

        End If

        '///TASK:
        'Convert mfinddata structure data into appropriate filename/file path, and create a CFile or CDirectory Object, depending on it's actual type.
        FullFileName = mFindData.cFileName
        FullFileName = Trim$(mFindData.cFileName)




        FullFileName = Replace$(FullFileName, vbNullChar, "")
        If FullFileName = "" Then Exit Function
        FullFileName = mDirConnect.Path & FullFileName
        '///convert to appropriate object-
        'use attributes to determine wether it is a directory or file- also mask out those not wanted by the client(via attributemask)
        If (mFindData.dwFileAttributes And mExcludeAttr) = mExcludeAttr And mExcludeAttr <> 0 Then
            'Debug.Print "excluding " & FullFileName, " matches attribute " & mExcludeAttr
            DebugMsg "excluding " & FullFileName & " matches attribute " & mExcludeAttr
            FoundMatch = False
        Else
            'determine wether to return it as well as wether it is a folder or file...
            If ((mFindData.dwFileAttributes And attributemask) = attributemask) Or (attributemask = 0) Then
                If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    'it's a dir. return a dir.
                    If Right$(FullFileName, 1) = "." Or Right$(FullFileName, 2) = ".." Then
                         'GoTo FindAnother    'go find another one.
                    Else
                        'Debug.Print "directory"
                        On Error Resume Next
                        CDebug.Post "using GetDirectory on " & FullFileName
                        Set grabobject = FileSystem.GetDirectory(FullFileName)
                        If Err.Number <> 0 Then
                        FoundMatch = True
                        End If
                    End If
                Else
                    'it is not a dir.
                    If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DEVICE) <> FILE_ATTRIBUTE_DEVICE Then
                        'if it isn't a device, then...
                        On Error Resume Next
                        CDebug.Post "using Getfile on " & FullFileName
                        Set grabobject = FileSystem.GetFile(FullFileName)
                        If Err.Number <> 0 Then
                            FoundMatch = True
                        End If
                        
                    End If
                End If '(mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY


            End If '(mFindData.dwFileAttributes And mExcludeAttr) = mExcludeAttr And mExcludeAttr <> 0

        End If



    '        If ((mFindData.dwFileAttributes And attributemask) = attributemask Or (attributemask <> 0)) And _
    '           Right$(FullFileName, 1) <> "." And Right$(FullFileName, 2) <> ".." Then
    '            If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
    '                'it is a directory, Convert to CDirectory.
    '                Set grabobject = FileSystem.GetDirectory(FullFileName)
    '            Else
    '                'OK: if our attribute mask include directory, we can (hopefully) assume that  there will be no more directories. and as such break out.
    '                If attributemask And FILE_ATTRIBUTE_DIRECTORY = FILE_ATTRIBUTE_DIRECTORY Then
    '
    '                End If
    '                If (mFindData.dwFileAttributes And FILE_ATTRIBUTE_DEVICE) <> FILE_ATTRIBUTE_DEVICE Then
    '                    'ignore device stuff here...
    '                    'convert to CFile.
    '                    Set grabobject = FileSystem.GetFile(FullFileName)
    '
    '
    '
    '                End If
    '            End If
    '        Else
    '            'Debug.Print FullFileName
    '            GoTo FindAnother
    '        End If


     'Until Foundmatch
        Dim fexclude As Boolean
        If Not grabobject Is Nothing Then
            fexclude = False
            RaiseEvent ItemFound(grabobject, fexclude)
            If fexclude Then
            FoundMatch = False
            Else
            'Debug.Print "returning object type " & TypeName(grabobject)
            Set GetNext = grabobject
            Set SetThis = grabobject
            FoundMatch = True
            'woopee!
            End If
        End If
    Loop
    CDebug.Post "Exiting GetNext Routine"
    Exit Function
showError:
    CDebug.Post "Error in GetNext:" & Err.Description, Err.Number
    'FindClose mFindHandle
End Function

'Public Function Clone() As CDirWalker
'
'Clone, hmm- that's a toughie in this instance.
'Set Clone = Me
'End Function
Public Sub Reset()
    FindClose mFindHandle
    mFindHandle = 0
End Sub

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
If mFindHandle <> 0 Then Me.Reset
End Sub

Private Function IVariantEnumerable_Clone() As stdole.IEnumVARIANT
'not implemented either.
End Function

Private Function IVariantEnumerable_GetNext(GetVar As Variant) As Long
    Set GetVar = Me.GetNext()
    If GetVar Is Nothing Then IVariantEnumerable_GetNext = False Else IVariantEnumerable_GetNext = True
End Function

Private Sub IVariantEnumerable_Reset()
'not implemented.
End Sub

Private Sub IVariantEnumerable_Skip(ByVal celt As Long)
'not implemented.
End Sub
Private Sub DebugMsg(ByVal strdebug As String)
'Open "C:\templog.tmp" For Output As #1
'Print #1, strdebug
'Close #1
End Sub
