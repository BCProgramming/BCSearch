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
Private mCallback As IFileSearchCallback
Private mSearcher As FileSearch

Private mSearchFilters As SearchFilters
Private mvarCancel As Boolean
Private mMatched As collection
Public Property Get Matched() As Object
    Set Matched = mMatched
End Property
Public Sub Cancel()
mvarCancel = True
mSearcher.Cancel
End Sub
Public Property Set Callback(vdata As IFileSearchCallback)
    Set mCallback = vdata
End Property
Public Property Get Callback() As IFileSearchCallback
    Set Callback = mCallback
End Property

Public Property Get Filters() As SearchFilters

    Set Filters = mSearchFilters


End Property
Public Property Set Filters(vdata As SearchFilters)
    Set mSearchFilters = vdata
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
    firstfilter.name = "Directory Excluder"
    mSearchFilters.Add firstfilter
    mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub

Private Function IFileSearchCallback_AllowRecurse(InDir As String) As Boolean
'
IFileSearchCallback_AllowRecurse = mCallback.AllowRecurse(InDir)
End Function

Private Function IFileSearchCallback_Cancelled() As Boolean
IFileSearchCallback_Cancelled = mvarCancel Or mCallback.Cancelled
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
 Dim matchedfilters As collection, newstr As String
 Set matchedfilters = New collection
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
End If
CurrFound = True: prevfound = True
Debug.Print madefile.Fullpath
matchedNames = ""
For I = 0 To mSearchFilters.Count
    
    Set mFilter = mSearchFilters.Item(I)
    currfiltername = mFilter.name
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
        
        If (CurrFound Or prevfound) Then
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
    End If
    If newstr <> "" Then
    matchedNames = matchedNames & ";" & newstr
    End If
    CurrFound = passFound Or CurrFound
Next I
'If weightAdd > 0 Then
DoEvents
If Not mvarCancel Then
    If CurrFound Then
        Set mMatched = matchedfilters
        mCallback.Found Me, Found, Cancel
    End If
Else
    Cancel = True
    Sender.Cancel
    Exit Sub
End If


'End If

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
    Dim CurrChar As String, CurrPos As Long
    currcount = -1
    currspec = ""
    For CurrPos = 1 To Len(FileSpecs)
    CurrChar = Mid$(FileSpecs, CurrPos, 1)
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
        
    Next CurrPos
    ReDim Preserve ret(currcount + 1)
    ret(currcount + 1) = currspec

    SplitFileSpecs = ret

End Function
Public Sub Search(ByVal Spec As String, ByVal InDir As String, Callback As IFileSearchCallback)
    Dim SearchDirs() As String, I As Long
    mvarCancel = False
    Set mCallback = Callback
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


End Sub

Private Sub IFileSearchCallback_SearchError(ErrCode As Long, ErrDesc As String, Cancel As Boolean)
    mCallback.SearchError ErrCode, ErrDesc, Cancel
End Sub
