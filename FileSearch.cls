VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'"Core" file search class- requires a Callback object that will also be responsible for saving each returned value.

Implements IAsyncProcess
Implements IFileSearchCallback
'File/Directory Searcher
Public Type AsyncSearchData
    FileSpec As String
    StartDir As String
    Callback As IFileSearchCallback
    

End Type

'uses FindFile APIs to search for files.
'Basically an extended and slightly slower version of the code used in the CDirWalker Class, but with a callback and recursion to boot.
Private mSearchData As WIN32_FIND_DATA
Private mFileSpec As String
Private SubSearcherCount As Long
Private CurrCallback As IFileSearchCallback
Private SearchHandle As Long
Private mvarcancelled As Boolean
'Private mAsyncCall As Boolean
Public Function Hascancelled() As Boolean
    Hascancelled = mvarcancelled Or CurrCallback.Cancelled
End Function
Public Sub Cancel()
    mvarcancelled = True
End Sub
Friend Function GetCurrCallback() As IFileSearchCallback
    Set CurrCallback = CurrCallback
End Function
Public Sub SearchAsync(ByVal FileSpec As String, ByVal StartDir As String, Callback As IFileSearchCallback)
'
Dim AsyncExec As CAsyncExec
Dim asyncinfo As AsyncSearchData
Set AsyncExec = New CAsyncExec

 '   mFileSpec = FileSpec
 asyncinfo.FileSpec = FileSpec
 asyncinfo.StartDir = StartDir
 Set asyncinfo.Callback = Callback
  Set CurrCallback = Callback
  AsyncExec.Start Me, asyncinfo
End Sub
Public Sub Search(ByVal FileSpec As String, ByVal StartDir As String, Callback As IFileSearchCallback, Optional ByRef Cancelled As Boolean)
    Dim fname As String, FullName As String, ret As Long
    Dim DirsRecurse() As String, dircount As Long  'Array of folders to recurse into.
    Set CurrCallback = Callback
    mvarcancelled = False
    If Right$(StartDir, 1) <> "\" And Right(StartDir, 1) <> "/" Then StartDir = StartDir & "\"
    mFileSpec = FileSpec
    If Callback.Cancelled Then Exit Sub
    ZeroMemory mSearchData, Len(mSearchData)
    SearchHandle = FindFirstFile(StartDir & FileSpec, mSearchData)
    If SearchHandle = ERROR_NO_MORE_FILES Then
        'no files....
        Callback.ExecuteComplete Me
        Set CurrCallback = Nothing
    ElseIf SearchHandle = -1 Then
        'error
        RaiseAPIError Err.LastDllError, "FileSearch::Search"
    Else
        'SearchHandle is good.
        Do
            If Hascancelled Then
                Cancelled = True
                Exit Sub
            End If
            fname = mSearchData.cFileName
            fname = Left$(fname, InStr(fname, vbNullChar) - 1)
            FullName = StartDir & fname
            If fname = "" Then Exit Do
            If fname <> "." And fname <> ".." Then
            
             '   If Not ((mSearchData.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT) = FILE_ATTRIBUTE_REPARSE_POINT) Then
                    'Call the callback routine...
                    'The callback routine is responsible for storing the data and using it.
                    'different paths for Directories and files...
                    Dim flcancel As Boolean
                    Callback.Found Me, FullName, flcancel
                    If flcancel Then
                        'Exit the loop. Search has been cancelled.
                        mvarcancelled = True
                        Exit Do
                    End If
                    
                    If (mSearchData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                        'attempt to recurse.... But only if allowed :)
                        If Callback.AllowRecurse(FullName) Then
                            If Right$(FullName, 1) <> "\" Then FullName = FullName & "\"
                            dircount = dircount + 1
                            ReDim Preserve DirsRecurse(1 To dircount)
                            DirsRecurse(dircount) = FullName
                        
                        End If
                          
                            
                        
                        
                    Else
                        'file...
                        'Callback.Found Sender, Found, Cancel, filterfound
                        DoEvents
                        If Hascancelled Then Exit Do
                    End If
              '  End If
            End If 'Dir "." and ".."
            ZeroMemory mSearchData, Len(mSearchData)
            ret = FindNextFile(SearchHandle, mSearchData)
            If ret = ERROR_NO_MORE_FILES Then
                Exit Do
            End If
        Loop
        FindClose SearchHandle
        'OK- now we recurse:
        Dim I As Long
        For I = 1 To dircount
          Dim SubSearch As FileSearch
          If Callback.Cancelled Then Exit Sub
        'recursion:
        'create a new FileSearch class.
        'set it's Callback to Us
        'start it searching. (note we will add a param that allows to call it Asynchronously from here, but for debugging purposes that won't be used)
        Set SubSearch = New FileSearch
        'Debug.Print "recursing into dir, " & FullName
        SubSearcherCount = SubSearcherCount + 1
        
        SubSearch.Search FileSpec, DirsRecurse(I), Me
        
        
        Next I
        
        
        
    End If
    CurrCallback.ExecuteComplete Me
End Sub

Private Sub Class_Initialize()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub

Private Sub IAsyncProcess_ExecAsync(ByVal Vdata As Variant)
    'start the Search.
    Dim cast As AsyncSearchData
    cast.FileSpec = Vdata.FileSpec
    cast.StartDir = Vdata.StartDir
    Set cast.Callback = Vdata.Callback
    
    
    Search cast.FileSpec, cast.StartDir, cast.Callback
End Sub

Private Function IFileSearchCallback_AllowRecurse(InDir As String) As Boolean
 IFileSearchCallback_AllowRecurse = CurrCallback.AllowRecurse(InDir)
End Function

Private Function IFileSearchCallback_Cancelled() As Boolean
    IFileSearchCallback_Cancelled = Hascancelled
End Function

Private Sub IFileSearchCallback_ExecuteComplete(Sender As Object)
    SubSearcherCount = SubSearcherCount - 1
End Sub





Private Sub IFileSearchCallback_Found(Sender As Object, Found As String, Optional Cancel As Boolean, Optional FiltersFound As Variant)
    CurrCallback.Found Sender, Found, Cancel, FiltersFound
End Sub

Private Sub IFileSearchCallback_ProgressMessage(ByVal StrMessage As String)
    CurrCallback.ProgressMessage StrMessage
End Sub

Private Sub IFileSearchCallback_SearchError(ErrCode As Long, ErrDesc As String, Cancel As Boolean)
    CurrCallback.SearchError ErrCode, ErrDesc, Cancel
End Sub
