VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSearchEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Extended(!) File search Class.
'uses the same interface to return data to the caller,
'but allows more customization of results before they get to the callback.
Implements IFileSearchCallback
Implements IAsyncProcess
Private mCallback As IFileSearchCallback
Private mSearcher As FileSearch
Private WithEvents mTimerUse As CTimer
Attribute mTimerUse.VB_VarHelpID = -1
Private mSearchFilters As SearchFilters
Private mvarCancel As Boolean
Private mMatched As Collection



Public Property Get Matched() As Object
    Set Matched = mMatched
End Property
Public Sub Cancel()
mvarCancel = True
mSearcher.Cancel
End Sub
Public Property Set Callback(Vdata As IFileSearchCallback)
    Set mCallback = Vdata
End Property
Public Property Get Callback() As IFileSearchCallback
    Set Callback = mCallback
End Property

Public Property Get Filters() As SearchFilters

    Set Filters = mSearchFilters


End Property
Public Property Set Filters(Vdata As SearchFilters)
    Set mSearchFilters = Vdata
    
End Property
Private Sub Class_Initialize()
    Dim firstfilter As CSearchFilter
    Set mSearchFilters = New SearchFilters
    Set mSearcher = New FileSearch
    Set firstfilter = New CSearchFilter
    Set firstfilter.Callback = New IFilterChangeCallback
    'this is a FILE search class- so no directories allowed ;)
    

    firstfilter.Attributes = FILE_ATTRIBUTE_DIRECTORY
    firstfilter.SearchOperation = Filter_Exclude
    firstfilter.Name = "Directory Excluder"
    mSearchFilters.Add firstfilter
    MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub

Private Sub IAsyncProcess_ExecAsync(ByVal Vdata As Variant)
'
End Sub

Private Function IFileSearchCallback_AllowRecurse(InDir As String) As Boolean
'
IFileSearchCallback_AllowRecurse = mCallback.AllowRecurse(InDir)
End Function

Private Function IFileSearchCallback_Cancelled() As Boolean
Dim Accumboolean As Boolean
Accumboolean = mvarCancel
If Not mCallback Is Nothing Then
    Accumboolean = Accumboolean Or mCallback.Cancelled
    If Not mSearcher Is Nothing Then
        If Not mSearcher.GetCurrCallback() Is Nothing Then
            Accumboolean = Accumboolean Or mSearcher.GetCurrCallback().Cancelled()
        End If
    End If
End If


IFileSearchCallback_Cancelled = Accumboolean

End Function

Private Sub IFileSearchCallback_ExecuteComplete(Sender As Object)
'
mCallback.ExecuteComplete Sender
End Sub



Private Sub IFileSearchCallback_Found(Sender As Object, Found As String, Optional Cancel As Boolean, Optional FiltersFound As Variant)
  '
    'Run through every filter...
    Dim I As Long, mFilter As CSearchFilter
    Dim CurrFound As Boolean, madefile As CFile, madeDir As Directory
    Dim prevfound As Boolean
    Dim foundweight As Long, passFound As Boolean
     Dim matchedNames As String
     Dim matchedfilters As Collection, newstr As String, regmatches As Collection
     Set matchedfilters = New Collection
     Set regmatches = New Collection
     Dim containsexpcol() As Object, containsrecount As Long, matchedf As Long
     
     Dim currfiltername As String
    'Dim weightAdd As Integer
    On Error GoTo breakout
    'If InStr(1, Found, ".txt", vbTextCompare) > 0 Then
    '    Debug.Assert False
    
    'End If
    If (GetFileAttributes(Found) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
    'change this.... later... to support folders...
        Exit Sub
    Else
        Set madefile = FileSystem.GetFile(Found)
        Debug.Print "madefile reference count checkpoint 1=" & GetRefCount(madefile)
    End If
    CurrFound = True: prevfound = True
    CDebug.Post madefile.Fullpath
     CDebug.Post "madefile reference count checkpoint 2=" & GetRefCount(madefile)
    matchedNames = ""
    For I = 0 To mSearchFilters.count
        
        Set mFilter = mSearchFilters.Item(I)
        currfiltername = mFilter.Name
        If currfiltername = "" Then
            currfiltername = "Unnamed Filter #" & I
        End If
        prevfound = CurrFound
        'If (madefile.Fileattributes And FILE_ATTRIBUTE_HIDDEN) Then Stop
        
        CurrFound = mFilter.FilterResult(madefile)
        'If StrComp(madefile.Extension, "LOG", vbTextCompare) = 0 Then Stop
        
        If mFilter.SearchOperation = Filter_Exclude Then
            'if mfilter matched, we want to exclude this file from the search.
            '    weightAdd = weightAdd + (((CInt(CurrFound) * 2) + 1))
            If CurrFound Then Exit Sub
            
            'excludes no matter what.
        ElseIf mFilter.SearchOperation = Filter_Include Then
            'true add one,
            'false subtract 1.
            'weightAdd = weightAdd + (((CInt(CurrFound) * 2) + 1) * -1)
            newstr = currfiltername
            'includes no matter what on match
            If Not CurrFound Then Exit Sub
        ElseIf mFilter.SearchOperation = Filter_Or Then
            
            If (CurrFound Or prevfound) And (Not prevfound And I = 0) Then
                If CurrFound Then newstr = currfiltername
                'CurrFound = True'
                passFound = True
            End If
            
        '
        ElseIf mFilter.SearchOperation = Filter_And Then
        '
            If (CurrFound And prevfound) Then
               ' CurrFound = True
                newstr = currfiltername
                passFound = True
            End If
        End If
        
        If CurrFound Then
            matchedfilters.Add mFilter
            'add the filter's cached regexp matches here as well...
            If Not mFilter.REContainMatchCol Is Nothing Then
                regmatches.Add mFilter.REContainMatchCol
            End If
        End If
        If newstr <> "" Then
        matchedNames = matchedNames & ";" & newstr
        End If
        CurrFound = passFound Or CurrFound
    Next I
    'If weightAdd > 0 Then
    'TODO:// reference count for madefile goes up by one in Filter loop for some reason. Find out why.
     CDebug.Post "madefile reference count checkpoint 3=" & GetRefCount(madefile)
    DoEvents
    If Not mvarCancel Then
        If CurrFound Then
            Set mMatched = matchedfilters
            Set madefile = Nothing
            mCallback.Found Me, Found, Cancel, mMatched
        End If
    Else
        Cancel = True
        Sender.Cancel
        Exit Sub
    End If
    
    
    'End If
    Set madefile = Nothing
    Exit Sub
breakout:
    
    mCallback.SearchError Err.Number, Err.Description, Cancel
    
    

End Sub

Private Sub IFileSearchCallback_ProgressMessage(ByVal StrMessage As String)
'
mCallback.ProgressMessage StrMessage
End Sub
Private Function SplitFileSpecs(ByVal FileSpecs As String) As String()
    Dim currcount As Long
    Dim ret() As String
    Dim inquote As Boolean, currspec As String
    Dim CurrChar As String, Currpos As Long
    currcount = -1
    currspec = ""
    For Currpos = 1 To Len(FileSpecs)
    CurrChar = Mid$(FileSpecs, Currpos, 1)
        If CurrChar = """" Then
            inquote = Not inquote
        ElseIf CurrChar = ";" Then
            currcount = currcount + 1
            ReDim Preserve ret(currcount)
            ret(currcount) = currspec
            currspec = ""
        Else
            currspec = currspec & CurrChar
        End If
        
    Next Currpos
    ReDim Preserve ret(currcount + 1)
    ret(currcount + 1) = currspec

    SplitFileSpecs = ret

End Function
Public Sub Search(ByVal Spec As String, ByVal InDir As String, Callback As IFileSearchCallback, Optional ByVal Asynchronous As Boolean = False)
    Dim SearchDirs() As String, I As Long, useextradata As Variant
    mvarCancel = False
    Set mCallback = Callback
    
    If Asynchronous Then
        Set mTimerUse = New CTimer
        mTimerUse.Interval = 1
        ReDim useextradata(0 To 3)
        Set useextradata(0) = Me
        useextradata(1) = Spec
        useextradata(2) = InDir
        Set useextradata(3) = Callback
        
        mTimerUse.Item = useextradata


    Else
    
        If InDir = "" Then Exit Sub
        If InStr(InDir, ";") Then
            
            SearchDirs = SplitFileSpecs(InDir)
            For I = 0 To UBound(SearchDirs)
            Dim fcancelled As Boolean
                
                    mSearcher.Search Spec, SearchDirs(I), Me, fcancelled
              
                If fcancelled Then Exit For
            Next I
            Callback.ExecuteComplete Me
        Else
          
                mSearcher.Search Spec, InDir, Me
                Callback.ExecuteComplete Me
        
        End If

    End If
End Sub

Private Sub IFileSearchCallback_SearchError(ErrCode As Long, ErrDesc As String, Cancel As Boolean)
    mCallback.SearchError ErrCode, ErrDesc, Cancel
End Sub

Private Sub mTimerUse_ThatTime()
'Start the async operation...
'first, disable the timer...
 'Array(Me, Spec, InDir, Callback)
    mTimerUse.Interval = 0
    Dim usesearch As CFileSearchEx, gotcallback As IFileSearchCallback
    Set usesearch = mTimerUse.Item(0)
    Set gotcallback = mTimerUse.Item(3)
    usesearch.Search mTimerUse.Item(1), mTimerUse.Item(2), gotcallback, False
End Sub
