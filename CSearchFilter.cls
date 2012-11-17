VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearchFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CSearchFilter

'Moving the file filtering logic from a single IFileSearchCallback call into a "chain" of IFileSearchCallbacks for the FileSearchEx class (The standard "FileSearch" class will still use the standard method)


'This allows much more control over results.
Public Event Change()
Const DefaultVBScript As String = "Public Function DoFilter(FileObj)" & vbCrLf & vbTab & "'work with FileObj here. use properties of Fileobj, filename, extension, attributes, etc and return true or false." & vbCrLf & "DoFilter=True 'return true by default." & vbCrLf & "End Function"
Private mCallback As IFilterChangeCallback
Public Enum DateSpecConstants  'Bitmasks. Which Ones should be taken into account?
    datespec_create = 2
    datespec_Accessed = 4
    datespec_modified = 8
End Enum
Public Type DateRange
    DateStart As Date
    DateEnd As Date
End Type
Public Type DateSpecifications
    DateSpecUse As DateSpecConstants
    DateCreated As DateRange
    '---
    DateAccessed As DateRange
    '---
    DateModified As DateRange
End Type
Public Enum SearchResultFilterType
    'Filter_And_Previous 'And the search state with the previous result.
    
    Filter_Include  'Include if matched...
    Filter_Exclude  'exclude if matched...
    Filter_Or 'add weight if this or the previous matched.
    Filter_And
    'Filter_Or_Previous 'OR the search state with the previous result.
    'Filter_XOr_Previous 'XOR the search state with the previous result.
    'Filter_EQV_Previous 'EQV the search state with the previous result.
    'Filter_IMP_Previous 'IMP the search state with the previous result.


    'Filter_IncludeRegardless  'set current search state on the found item to true. and then break out. no further filters are handled.
    'Filter_ExcludeRegardless  'set current search state on the found item to false, and then break out. no further filters are handled.
End Enum
Public Type AlternateStreamSearchSpecifications
    mincount As Long  'minimum...
    maxcount As Long  'and maximum counts of streams.
    'AltFilterType As SearchResultFilterType
    NameSpecisRegExp As Boolean
    nameSpec As String  'handled with "Like"...
    ContainsIsRegExp As Boolean
    ContainsStr As String
    ContainsRegExpMinmatches As Long
    ContainsRegExpMaxMatches As Long
    SizeLargerThan As Double
    SizeSmallerThan As Double
        'no "scripts" for alternate streams.
    


End Type
Public Type RenameData
    DoRename As Boolean
    NewNameMask As String
End Type
Public Type SearchSpecifications
    SearchFiltertype As SearchResultFilterType
    FileSpecIsRegExp As Boolean
    FileSpec As String 'Specification...
    FileAttribExactMatch As Boolean
    FileAttributes As FileAttributeConstants 'Attributes to include...
    ContainsFileSizeLimit As Long
    ContainsIsRegExp As Boolean
    ContainsStr As String   'search files for this string....
    ContainsRegExpMinmatches As Long
    ContainsRegExpMaxMatches As Long
    'Date specifications.
    DateSpecs As DateSpecifications
    'Size Lower and higher values.
    FileSizeLargerThan As Double 'larger then this... converted to Long integer.
    FileSizeSmallerThan As Double 'smaller then this... Note that these two values are normalized.
    mScriptLanguage As String
    mScriptCode As String
        '^for script filters!
    mSearchAlternateStreams As Boolean 'wether to perform the alternate stream filter operation.
   AlternateStreamSpecs As AlternateStreamSearchSpecifications
   RenameInfo As RenameData
End Type
'Private mFileFindLog As FileStream
Private mScriptObject As ScriptControl
Private mvarname As String
Private mSearchSpec As SearchSpecifications
Private mTag As Variant
Private mHidden As Boolean
Private mRegexpContainMatch As MatchCollection
Public Property Let Hidden(ByVal Vdata As Boolean)
    mHidden = Vdata
End Property
Public Property Get Hidden() As Boolean
    Hidden = mHidden
End Property
Public Property Get REContainMatchCol() As Object
    Set REContainMatchCol = mRegexpContainMatch
End Property


'   AltFilterType As SearchResultFilterType
'    NameSpecIsRegexp As Boolean
'    NameSpec As String  'handled with "Like"...
'    ContainsIsRegExp As Boolean
'    ContainsStr As String
'    ContainsRegExpMinmatches As Long
'    SizeLargerThan As Double
'    SizeSmallerThan As Double
'        'no "scripts" for alternate streams.
'
Public Sub SetAlternateStreamSearchData(ByVal mincount As Long, ByVal maxcount As Long, ByVal AltFilterType As SearchResultFilterType, ByVal NameSpecisRegExp As Boolean, _
ByVal nameSpec As String, ByVal ContainsIsRegExp As Boolean, ByVal ContainsStr As String, ByVal ContainsRegExpMinmatches As Long, ByVal SizeLargerThan As Double, ByVal SizeSmallerThan As Double)

    With mSearchSpec.AlternateStreamSpecs
        
'        .AltFilterType = AltFilterType
        .NameSpecisRegExp = NameSpecisRegExp
        .nameSpec = nameSpec
        .ContainsIsRegExp = ContainsIsRegExp
        .ContainsStr = ContainsStr
        .ContainsRegExpMinmatches = ContainsRegExpMinmatches
        .SizeLargerThan = SizeLargerThan
        .SizeSmallerThan = SizeSmallerThan
        .mincount = mincount
        .maxcount = maxcount
        
    End With


End Sub


Public Property Let ScriptLanguage(ByVal Vdata As String)
    mSearchSpec.mScriptLanguage = Vdata
    Set mScriptObject = Nothing
End Property
Public Property Get ScriptLanguage() As String
    ScriptLanguage = mSearchSpec.mScriptLanguage
End Property

Public Property Let ScriptCode(ByVal Vdata As String)
    mSearchSpec.mScriptCode = Vdata
    Set mScriptObject = Nothing
End Property
Public Property Get ScriptCode() As String
    ScriptCode = mSearchSpec.mScriptCode
End Property
Public Property Get Name() As String
    Name = mvarname
End Property
Public Property Let Name(ByVal Vdata As String)
    mvarname = Vdata
End Property
Public Property Get Tag() As Variant
    If IsObject(mTag) Then Set Tag = mTag Else Tag = mTag
    
End Property
Public Property Let Tag(ByVal Vdata As Variant)
    mTag = Vdata
    
End Property
Public Property Set Tag(Vdata As Variant)
    Set mTag = Vdata
End Property
Public Property Set Callback(Vdata As IFilterChangeCallback)
    Set mCallback = Vdata
End Property
Public Property Get Callback() As IFilterChangeCallback
    Set Callback = mCallback
End Property

Public Sub initFilterIndirect(Struct As SearchSpecifications)
    mSearchSpec = Struct


End Sub
'INIT method... faster way ;)

'use setAlternateStreamSearchData to set that information...

Public Sub InitFilter(ByVal FilterType As SearchResultFilterType, Optional ByVal FileSpec As String = "*.*", _
                Optional ByVal FileAttributes As FileAttributeConstants, Optional ByVal ContainsStr As String = "", _
                Optional ByVal ContainsIsRegExp As Boolean = False, Optional ByVal ContainsMinmatches As Long = 1, _
                Optional ByVal FileSizeLargerThan As Double = -1, Optional ByVal FileSizeSmallerThan As Double = -1, _
                Optional ByVal DateSpecsUse As DateSpecConstants = 0, _
                Optional ByVal DateCreatedStart As Date, Optional ByVal DateCreatedEnd As Date, _
                Optional ByVal DateAccessedStart As Date, Optional ByVal DateAccessedEnd As Date, _
                Optional ByVal DateModifiedStart As Date, Optional ByVal DateModifiedEnd As Date, _
                Optional ByVal ScriptLanguage As String = "", Optional ByVal ScriptCode As String = "")
                
                With mSearchSpec
                    .FileSpec = FileSpec
                    .SearchFiltertype = FilterType
                    .FileAttributes = FileAttributes
                    .ContainsStr = ContainsStr
                    .ContainsIsRegExp = ContainsIsRegExp
                    .ContainsRegExpMinmatches = ContainsMinmatches
                    .FileSizeLargerThan = FileSizeLargerThan
                    .FileSizeSmallerThan = FileSizeSmallerThan
                    .DateSpecs.DateSpecUse = DateSpecsUse
                    With .DateSpecs
                        .DateCreated.DateStart = DateCreatedStart
                        .DateCreated.DateEnd = DateCreatedEnd
                        .DateModified.DateStart = DateModifiedStart
                        .DateModified.DateEnd = DateModifiedEnd
                        .DateAccessed.DateStart = DateAccessedStart
                        .DateAccessed.DateEnd = DateAccessedEnd
                        
                    
                    End With
                    .mScriptLanguage = ScriptLanguage
                    .mScriptCode = ScriptCode
                
                End With
                
                
                
End Sub



Public Property Let DatesCheck(ByVal Vdata As DateSpecConstants)
    mSearchSpec.DateSpecs.DateSpecUse = Vdata
    mCallback.Change Me
End Property
Public Property Get DatesCheck() As DateSpecConstants
    DatesCheck = mSearchSpec.DateSpecs.DateSpecUse
End Property
Public Property Let DateStart(ByVal WhichDate As DateSpecConstants, ByVal Vdata As Date)
    With mSearchSpec.DateSpecs
        If (WhichDate And datespec_create) = datespec_create Then
            .DateCreated.DateStart = Vdata
        End If
        If (WhichDate And datespec_Accessed) = datespec_Accessed Then
            .DateAccessed.DateStart = Vdata
        End If
        If (WhichDate And datespec_modified) = datespec_modified Then
            .DateModified.DateStart = Vdata
        End If
    
    End With
    mCallback.Change Me

End Property
Public Property Get DateStart(ByVal WhichDate As DateSpecConstants) As Date
    With mSearchSpec.DateSpecs
        Select Case .DateSpecUse
            Case datespec_create
                DateStart = .DateCreated.DateStart
            Case datespec_Accessed
                DateStart = .DateAccessed.DateStart
            Case datespec_modified
                DateStart = .DateModified.DateStart
        End Select
    End With



End Property


Public Property Let DateEnd(ByVal WhichDate As DateSpecConstants, ByVal Vdata As Date)
    With mSearchSpec.DateSpecs
        If (WhichDate And datespec_create) = datespec_create Then
            .DateCreated.DateEnd = Vdata
        End If
        If (WhichDate And datespec_Accessed) = datespec_Accessed Then
            .DateAccessed.DateEnd = Vdata
        End If
        If (WhichDate And datespec_modified) = datespec_modified Then
            .DateModified.DateEnd = Vdata
        End If
    
    End With
    mCallback.Change Me

End Property
Public Property Get DateEnd(ByVal WhichDate As DateSpecConstants) As Date
    With mSearchSpec.DateSpecs
        Select Case WhichDate
            Case datespec_create
                DateEnd = .DateCreated.DateEnd
            Case datespec_Accessed
                DateEnd = .DateAccessed.DateEnd
            Case datespec_modified
                DateEnd = .DateModified.DateEnd
        End Select
    End With



End Property


Public Property Let SizeLargerThan(ByVal Vdata As Double)
    mSearchSpec.FileSizeLargerThan = Vdata 'really the "smaller" value...
    mCallback.Change Me
End Property
Public Property Get SizeLargerThan() As Double
    SizeLargerThan = mSearchSpec.FileSizeLargerThan
End Property
Public Property Let SizeSmallerThan(ByVal Vdata As Double)
    mSearchSpec.FileSizeSmallerThan = Vdata
    mCallback.Change Me
End Property
Public Property Get SizeSmallerThan() As Double
    SizeSmallerThan = mSearchSpec.FileSizeSmallerThan
End Property

Public Property Let ContainsMinmatches(ByVal Vdata As Long)
    mSearchSpec.ContainsRegExpMinmatches = Vdata
    mCallback.Change Me
End Property
Public Property Get ContainsRegExpMinmatches() As Long
    ContainsRegExpMinmatches = mSearchSpec.ContainsRegExpMinmatches
End Property
Public Property Let ContainsRegExpMinmatches(ByVal Vdata As Long)
    mSearchSpec.ContainsRegExpMinmatches = Vdata
End Property

Public Property Let ContainsStr(ByVal StrFind As String)
    mSearchSpec.ContainsStr = StrFind
    mCallback.Change Me
End Property
Public Property Get ContainsStr() As String
    ContainsStr = mSearchSpec.ContainsStr
End Property
Public Property Let ContainsIsRegExp(ByVal Vdata As Boolean)
    mSearchSpec.ContainsIsRegExp = Vdata
    mCallback.Change Me
End Property
Public Property Get ContainsIsRegExp() As Boolean
    ContainsIsRegExp = mSearchSpec.ContainsIsRegExp
End Property
Public Property Let SearchOperation(ByVal Vdata As SearchResultFilterType)
    mSearchSpec.SearchFiltertype = Vdata
    If Not mCallback Is Nothing Then
    mCallback.Change Me
    End If
End Property
Public Property Get SearchOperation() As SearchResultFilterType
    SearchOperation = mSearchSpec.SearchFiltertype
End Property
Public Property Let FileSpec(ByVal Vdata As String)
    mSearchSpec.FileSpec = Vdata
    If Not mCallback Is Nothing Then
    mCallback.Change Me
    End If
End Property
Public Property Get FileSpec() As String
    FileSpec = mSearchSpec.FileSpec
End Property
Public Property Let AttributesMatchExact(ByVal Vdata As Boolean)
'
mSearchSpec.FileAttribExactMatch = Vdata
End Property
Public Property Get AttributesMatchExact() As Boolean
    AttributesMatchExact = mSearchSpec.FileAttribExactMatch
End Property
Public Property Let Attributes(ByVal Fileattr As FileAttributeConstants)
    mSearchSpec.FileAttributes = Fileattr
    If Not mCallback Is Nothing Then
    mCallback.Change Me
    End If
End Property
Public Property Get Attributes() As FileAttributeConstants
    Attributes = mSearchSpec.FileAttributes
End Property
Public Function FilterResult(FileResult As Object) As Boolean
    'FileResult: Found file to be included/excluded based on the specifications in this class.
    Static mRegExp As Object
    Dim currmatchstate As Boolean, specs() As String, I As Long
    Dim matchedfilters As Collection
    Dim matchRE As MatchCollection
    'If FileResult.Fileattributes And FILE_ATTRIBUTE_HIDDEN Then Stop
    'If Left$(FileResult.Extension, 1) = "L" Then Stop
    If mRegExp Is Nothing Then
    
        Set mRegExp = CreateObject("VBScript.RegExp")
        mRegExp.MultiLine = True
        mRegExp.Global = True
    
    End If
    With mSearchSpec
        If .FileAttributes > 0 Then
            
            If (FileResult.FileAttributes And .FileAttributes) = .FileAttributes Then
                currmatchstate = True
            Else
                currmatchstate = False
                Exit Function
            End If
    
    
        End If
    If .FileSpec <> "" Then
    
        If .FileSpecIsRegExp Then
            Debug.Print "testing filename against regular expression pattern, " & .FileSpec
            mRegExp.Pattern = .FileSpec
            mRegExp.IgnoreCase = True
            
            If mRegExp.test(FileSystem.GetFilenamePart(FileResult.Fullpath)) Then
                currmatchstate = True
            Else
                currmatchstate = False
            End If
        
        Else '.FileSpecIsRegEXP
            If InStr(.FileSpec, ";") > 0 Then
                specs = Split(.FileSpec, ";")
            
            Else
                ReDim specs(0)
                specs(0) = UCase$(.FileSpec)
            End If
            For I = 0 To UBound(specs)
                If UCase$(FileSystem.GetFilenamePart(FileResult.Fullpath)) Like UCase$(specs(I)) Then
                    currmatchstate = True
                    Exit For
                Else
                    'currmatchstate = False
                    'Exit Function
                End If
            Next I
            If Not currmatchstate Then Exit Function
        End If '.FileSpecIsRegEXP
    
   
        'File spec matching completed....
        
        End If
        
        
        'file size....
        '//////////////////////////////////////////////
        If .FileSizeSmallerThan > 0 Then
        
            If FileResult.Size < .FileSizeSmallerThan Then
                currmatchstate = True
            Else
                currmatchstate = False
                Exit Function
            End If
            
        End If
    
        If .FileSizeLargerThan > 0 Then
            If FileResult.Size > .FileSizeLargerThan Then
                currmatchstate = True
            Else
                currmatchstate = False
                Exit Function
            End If
        End If '.FileSizeLargerThan > 0
        'end file size compares.
        '/////////////////////////////////////////////////
        
        'If Not currmatchstate Then
        'again- no need to continue here if we have "False"
        'Dates, this is the PITA...
        'so much so I've made it a separate procedure...
        
        If MeetsDateSpecs(FileResult, currmatchstate) And currmatchstate Then
        
            currmatchstate = True
        
        Else
        
            currmatchstate = False
            Exit Function
        
        End If 'MeetsDateSpecs(FileResult, currmatchstate)
        
        
        
        
        
        'Now... Script....
        If mSearchSpec.mScriptLanguage <> "" Then
            If mScriptObject Is Nothing Then
                Set mScriptObject = CreateObject("MSScriptControl.ScriptControl")
            Else
        'mScriptObject.Reset
        
            End If
        
            mScriptObject.Language = mSearchSpec.mScriptLanguage
            mScriptObject.AddCode mSearchSpec.mScriptCode
            If mScriptObject.CodeObject.DoFilter(FileResult) Then
                'The script is confined to methods and properties provided by the CFile object- should be enough to filter intelligently...
                Debug.Print "Script Filter <PASSED>"
                currmatchstate = True
            
            Else
                
                Debug.Print "Script Filter FAILED"
                currmatchstate = False
                Exit Function
            
        
        
            End If
        End If
        
        'End If 'currmatchstate
        Dim contentStream As FileStream, contentstr As String
        '//TODO: "contains" code. Likely will use good ol' Filestream...
        
        On Error Resume Next
        If .ContainsStr <> "" And ((FileResult.FileAttributes And FILE_ATTRIBUTE_OFFLINE) <> FILE_ATTRIBUTE_OFFLINE) Then
        Err.Clear
            Set contentStream = FileResult.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_DELETE, OPEN_EXISTING)
            If Err.Number <> 0 Then
            'error opening file...
            'assume a match failure in this case.
            
            
            Else
                Dim matches As Object
                contentstr = contentStream.ReadAllStr(StrRead_Default)
                If Err.Number <> 0 Then
            'error reading file...
                Else
                    contentStream.CloseStream
                    If Len(contentstr) >= Len(.ContainsStr) Then
                        If mSearchSpec.ContainsIsRegExp Then
                            mRegExp.Pattern = mSearchSpec.ContainsStr
                            If InStr(contentstr, "Static") Then Debug.Assert False
                            If mRegExp.test(contentstr) Then
                              Set matches = mRegExp.Execute(contentstr)
                                'must be at least minmatches...
                                If matches.count >= .ContainsRegExpMinmatches And matches.count <= .ContainsRegExpMaxMatches Then
                                    Debug.Print "matchcount meets range..."; matches.count; " >= "; .ContainsRegExpMinmatches
                                    currmatchstate = True
                                    Set mRegexpContainMatch = matches
                                Else
                                    Set mRegexpContainMatch = Nothing
                                    currmatchstate = False
                                  
                                End If
                             End If
                        
                        Else
                            'not regexp- use Instr.
                            If InStr(1, contentstr, mSearchSpec.ContainsStr, vbTextCompare) > 0 Then
                                currmatchstate = True
                            Else
                                currmatchstate = False
                            End If
                        
                        End If
                    End If
            
            
                End If
            
            
            End If
    
        End If
    
  
    
    'now handle alternate streams...
    Dim AltStreams As CAlternateStreams, loopstream As CAlternateStream
    
    With .AlternateStreamSpecs
        If .mincount > 0 And .maxcount > 0 And .nameSpec <> "" And currmatchstate = True Then
        Set AltStreams = FileResult.AlternateStreams
        
            'alternate Stream Specifications entered.
            currmatchstate = False
            If AltStreams.count > .mincount Then
                'so far: larger then mincount.
                If AltStreams.count < .maxcount Then
                    'so far, within mincount and maxcount
                    
                    
                    For Each loopstream In AltStreams
                    
                    
                        
                        If .NameSpecisRegExp Then
                                mRegExp.Pattern = mSearchSpec.ContainsStr
                                If mRegExp.test(contentstr) Then
                                    currmatchstate = True
                                    'Exit For
                                
                                Else
                                    currmatchstate = False
                                    GoTo NextIteration
                                
                                
                                End If
                        Else
                            If loopstream.Name Like .nameSpec Then
                                currmatchstate = True
                                'Exit
                            Else
                                currmatchstate = False
                                GoTo NextIteration
                            End If
                        End If
                        
                        
                        'Insert more checks
                        '.SizeLargerThan,.SizeSmallerThan,.ContainsStr,.ContainsRegExpMaxMatches ,.ContainsRegExpMinmatches,.ContainsIsRegExp, etc.
                        
                        
                        If .SizeLargerThan > .SizeSmallerThan Then
                            If loopstream.Size > .SizeSmallerThan And loopstream.Size < .SizeLargerThan Then
                                'meets size specs...
                                currmatchstate = True
                            Else
                                currmatchstate = False
                                GoTo NextIteration
                            End If
                        End If
                        
                        
                        
                        If .ContainsStr <> "" Then
                            '
                            'read the contents of the stream.
                            Dim StreamContents As String
                            Dim readStream As FileStream
                            Set readStream = loopstream.OpenAsBinaryStream(GENERIC_READ, FILE_SHARE_DELETE + FILE_SHARE_WRITE, OPEN_EXISTING)
                            StreamContents = readStream.ReadAllStr
                            readStream.CloseStream
                            
                            
                            
                            
                            
                            
                            If .ContainsIsRegExp Then
                            
                                'regular expression match code.
                                mRegExp.Pattern = .ContainsStr
                                If mRegExp.test(StreamContents) Then
                                    Set matches = mRegExp.Execute(StreamContents)
                                    If matches.count > .ContainsRegExpMinmatches And matches.count < .ContainsRegExpMaxMatches Then
                                        currmatchstate = True
                                    Else
                                        currmatchstate = False
                                        GoTo NextIteration
                                    End If
                                End If
                            End If
                        End If
                        
NextIteration:
                    Next
                    
                    
                End If
            End If
        
        
        
        
        
        
        
        
        
        
        
        End If
    
    
    End With
    
    
    '        Select Case .SearchFiltertype
    '            Case Filter_Exclude 'exclude if this one matches....
    '                currmatchstate = PreviousFoundState And Not currmatchstate
    '            Case Filter_And_Previous
    '                currmatchstate = currmatchstate And PreviousFoundState
    '            Case Filter_Or_Previous
    '                currmatchstate = currmatchstate Or PreviousFoundState
    '            Case Filter_XOr_Previous
    '                currmatchstate = currmatchstate Xor PreviousFoundState
    '            Case Filter_EQV_Previous
    '                currmatchstate = currmatchstate Eqv PreviousFoundState
    '            Case Filter_IMP_Previous
    '                currmatchstate = currmatchstate Imp PreviousFoundState
    '
    '
    '            Case Filter_IncludeRegardless
    '            'NOP
    '            Case Filter_ExcludeRegardless
    '            'NOP.
    '
    '
    '        End Select
returnIt:
    FilterResult = currmatchstate
    
    End With
ERROROPEN:
    'error occured opening file to filter on containing text...
End Function
Public Function GetMatchTypeString() As String
Select Case mSearchSpec.SearchFiltertype

    Case Filter_Include
        GetMatchTypeString = "Include"
    Case Filter_Exclude
        GetMatchTypeString = "Exclude"
    Case Filter_Or
        GetMatchTypeString = "Or"
    Case Filter_And
        GetMatchTypeString = "And"
'    Case Filter_Or_Previous
'    GetMatchTypeString = "OR"
'    Case Filter_XOr_Previous
'        GetMatchTypeString = "XOR"
'    Case Filter_EQV_Previous
'        GetMatchTypeString = "EQV"
'    Case Filter_IMP_Previous
'        GetMatchTypeString = "IMP"
'    Case Filter_IncludeRegardless
'        GetMatchTypeString = "Include_Always"
'    Case Filter_ExcludeRegardless
'        GetMatchTypeString = "Exclude_Always"

End Select





End Function
Private Function MeetsDateSpecs(FileCheck As CFile, RunningResult As Boolean) As Boolean
MeetsDateSpecs = RunningResult
With mSearchSpec.DateSpecs
    If (.DateSpecUse And datespec_create) = datespec_create Then
        If Not DateWithinRange(FileCheck.DateCreated, .DateCreated.DateStart, .DateCreated.DateEnd) Then
            Exit Function
        End If
            
    
    End If
    If (.DateSpecUse And datespec_modified) = datespec_modified Then
       If Not DateWithinRange(FileCheck.DateModified, .DateModified.DateStart, .DateModified.DateEnd) Then
       'Stop
            Exit Function
        Else
            'Stop
        End If
    End If
    If (.DateSpecUse And datespec_Accessed) = datespec_Accessed Then
        If Not DateWithinRange(FileCheck.DateLastAccessed, .DateAccessed.DateStart, .DateAccessed.DateEnd) Then
            Exit Function
        End If
    
    End If
    




    MeetsDateSpecs = True



End With




End Function
Private Function DateWithinRange(ByVal DateCheck As Date, StartDate As Date, EndDate As Date)
    Dim temp As Date
    If EndDate < StartDate Then
        'swap...
        temp = StartDate
        StartDate = EndDate
        EndDate = temp
    End If
    Debug.Print DateCheck, StartDate, EndDate
    
    
'    DateWithinRange = (DateCheck > StartDate) And (DateCheck > EndDate)
    DateWithinRange = Sgn(DateCheck - StartDate) = 1 And Sgn(EndDate - DateCheck) = 1
End Function

Private Sub Class_Initialize()
    With mSearchSpec.DateSpecs
        .DateAccessed.DateStart = Now - Time
        .DateAccessed.DateEnd = Now - Time
    
        .DateModified.DateStart = Now - Time
        .DateModified.DateEnd = Now - Time
        
        .DateCreated.DateStart = Now - Time
        .DateCreated.DateEnd = Now - Time
        
    
    End With
    mSearchSpec.ContainsFileSizeLimit = 32768 * 1024
    mSearchSpec.ContainsRegExpMaxMatches = 99999
    mSearchSpec.mScriptLanguage = "VBScript"
    mSearchSpec.SearchFiltertype = Filter_Or
    mSearchSpec.mScriptCode = "Public Function DoFilter(FileObj)" & vbCrLf & vbTab & "'work with FileObj here. use properties of Fileobj, filename, extension, attributes, etc and return true or false." & vbCrLf & "DoFilter=True 'return true by default." & vbCrLf & "End Function"
End Sub
Private Function GetDefaultScriptCode(ByVal CodeLanguage As String)
    Dim ret As String
    Select Case UCase$(CodeLanguage)
        Case "VBSCRIPT"
        ret = "Public Function DoFilter(FileObj)" & vbCrLf & _
        vbTab & "'Work with FileObj, Which is a CFile Object, here. Return True to ""Match""(Accept) the file, and false to reject it." & vbCrLf & _
        vbTab & "DoFilter=True   'Accept it by default." & vbCrLf & _
        "End Function"
        Case "JAVASCRIPT", "JSCRIPT"
        ret = "function DoFilter(FileObj){" & vbCrLf & _
        vbTab & "// work with FileObj, Which is a CFile Object, here. Return true to ""Match""(Accept) the file, and false to reject it." & vbCrLf & _
        vbTab & "return true; // return true by default." & vbCrLf & "}"
        Case "PERL"
'        #!/usr/bin/perl
'use strict;
'sub DoFilter {
'    my($FileObj) = @_;
'    #work with $FileObj here.
'    return true;
'}
        ret = "#!/usr/bin/perl" & vbCrLf & _
        "use strict;" & vbCrLf & _
        "sub DoFilter {" & vbCrLf & vbTab & _
        "my($FileObj) = @_;" & vbCrLf & vbTab & _
        "#work with $FileObj here." & vbCrLf & vbTab & _
        "return true;" & vbCrLf & _
        "}"
        Case "BCSCRIPT"
            'ret = "
        Case "PYTHON"

       'default python code here...

    End Select
End Function
Private Sub Class_InitProperties()
'



End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
  With PropBag
    mvarname = .ReadProperty("Name", "")
    If IsObject(.ReadProperty("Tag")) Then Set mTag = .ReadProperty("Tag") Else mTag = .ReadProperty("Tag")
    mSearchSpec.SearchFiltertype = .ReadProperty("FilterType", Filter_Include)
    mSearchSpec.FileSpec = .ReadProperty("Filter", "*.*")
    mSearchSpec.FileSpecIsRegExp = .ReadProperty("FileSpecRegExp", False)
    mSearchSpec.FileAttributes = .ReadProperty("Attributes", vbNormal)
    mSearchSpec.FileAttribExactMatch = .ReadProperty("ExactAttributes", False)
    mSearchSpec.FileSizeLargerThan = .ReadProperty("SizeLargerThan", 0)
    mSearchSpec.FileSizeSmallerThan = .ReadProperty("SizeSmallerThan", 0)
    mSearchSpec.ContainsStr = .ReadProperty("ContainsStr", "")
    mSearchSpec.ContainsIsRegExp = CBool(.ReadProperty("ContainsIsRegExp", False))
    mSearchSpec.ContainsRegExpMinmatches = .ReadProperty("ContainsRegExpMinMatches", 0)
    mSearchSpec.ContainsRegExpMaxMatches = .ReadProperty("ContainsRegExpMaxMatches", 9999)
    mSearchSpec.mScriptLanguage = .ReadProperty("ScriptLanguage", "VBScript")
    mSearchSpec.mScriptCode = .ReadProperty("ScriptCode", "")
    mSearchSpec.ContainsStr = .ReadProperty("ContainsStr", "")
    mSearchSpec.ContainsIsRegExp = .ReadProperty("ContainsIsRegExp", False)
    mSearchSpec.ContainsRegExpMinmatches = .ReadProperty("ContainsRegExpMinmatches", 0)
    mSearchSpec.ContainsRegExpMaxMatches = .ReadProperty("ContainsRegExpMaxmatches", 9999)
    mSearchSpec.DateSpecs.DateSpecUse = .ReadProperty("DateSpecUse", 0)
    
    'Dates...
    mSearchSpec.DateSpecs.DateCreated.DateStart = .ReadProperty("CreatedStart", #12/12/1980#)
    mSearchSpec.DateSpecs.DateCreated.DateEnd = .ReadProperty("CreatedEnd", #12/12/9999#)
    mSearchSpec.DateSpecs.DateModified.DateStart = .ReadProperty("ModifiedStart", #12/12/1980#)
    mSearchSpec.DateSpecs.DateModified.DateEnd = .ReadProperty("ModifiedEnd", #12/12/9999#)

    mSearchSpec.DateSpecs.DateAccessed.DateStart = .ReadProperty("AccessedStart", #12/12/1980#)
    mSearchSpec.DateSpecs.DateAccessed.DateEnd = .ReadProperty("AccessedEnd", #12/12/9999#)
    
    'read the alternate stream data...
        'mSearchSpec.AlternateStreamSpecs.AltFilterType = .ReadProperty("ADSSpecs.AltFilterType", Filter_Include)
        mSearchSpec.AlternateStreamSpecs.ContainsIsRegExp = .ReadProperty("ADSSpecs.ContainsIsRegexp", False)
        mSearchSpec.AlternateStreamSpecs.ContainsRegExpMaxMatches = .ReadProperty("ADSSpecs.RegExpMaxMatches", -1)
        mSearchSpec.AlternateStreamSpecs.ContainsRegExpMinmatches = .ReadProperty("ADSSpecs.RegExpMinMatches", -1)
        mSearchSpec.AlternateStreamSpecs.ContainsStr = .ReadProperty("ADSSpecs.ContainsStr", "")
        mSearchSpec.AlternateStreamSpecs.maxcount = .ReadProperty("ADSSpecs.MaxCount", -1) 'max count of streams...
        mSearchSpec.AlternateStreamSpecs.mincount = .ReadProperty("ADSSpecs.MinCount", -1) 'min count of streams...
        mSearchSpec.AlternateStreamSpecs.nameSpec = .ReadProperty("ADSSpecs.NameSpec", "*")
        mSearchSpec.AlternateStreamSpecs.NameSpecisRegExp = .ReadProperty("ADSSpecs.NameSpecIsRegExp", -1)
        mSearchSpec.AlternateStreamSpecs.SizeLargerThan = .ReadProperty("ADSSpecs.SizeLargerThan")
        mSearchSpec.AlternateStreamSpecs.SizeSmallerThan = .ReadProperty("ADSSpecs.SizeSmallerThan", -1)
    
    End With
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Name", mvarname, ""
        .WriteProperty "Tag", Me.Tag
        .WriteProperty "FilterType", mSearchSpec.SearchFiltertype, SearchResultFilterType.Filter_Include
        .WriteProperty "Filter", mSearchSpec.FileSpec, "*.*"
        .WriteProperty "FileSpecRegExp", mSearchSpec.FileSpecIsRegExp, False
        .WriteProperty "Attributes", mSearchSpec.FileAttributes, vbNormal
        .WriteProperty "ExactAttribute", mSearchSpec.FileAttribExactMatch, False
        .WriteProperty "SizeLargerThan", mSearchSpec.FileSizeLargerThan, 0
        .WriteProperty "SizeSmallerThan", mSearchSpec.FileSizeSmallerThan, 0
        'msearchspec.ContainsStr
        'msearchspec.ContainsIsRegExp
        'msearchspec.ContainsRegExpMinmatches
        'msearchspec.mScriptLanguage
        'msearchspec.mScriptCode
        .WriteProperty "ContainsStr", mSearchSpec.ContainsStr, ""
        .WriteProperty "ContainsIsRegExp", mSearchSpec.ContainsIsRegExp, False
        .WriteProperty "ContainsRegExpMinmatches", mSearchSpec.ContainsRegExpMinmatches, False
        .WriteProperty "ContainsRegExpMaxMatches", mSearchSpec.ContainsRegExpMaxMatches, 9999
        
        .WriteProperty "ScriptLanguage", mSearchSpec.mScriptLanguage, ""
        .WriteProperty "ScriptCode", mSearchSpec.mScriptCode, ""
        
        'date crap....
        .WriteProperty "DateSpecUse", mSearchSpec.DateSpecs.DateSpecUse, 0
        
        .WriteProperty "CreatedStart", mSearchSpec.DateSpecs.DateCreated.DateStart, #12/12/1980#
        .WriteProperty "CreatedEnd", mSearchSpec.DateSpecs.DateCreated.DateEnd, #12/12/9999#
        
        .WriteProperty "AccessedStart", mSearchSpec.DateSpecs.DateAccessed.DateStart, #12/12/1980#
        .WriteProperty "AccessedEnd", mSearchSpec.DateSpecs.DateAccessed.DateEnd, #12/12/9999#
        .WriteProperty "ModifiedStart", mSearchSpec.DateSpecs.DateModified.DateStart, #12/12/1980#
        .WriteProperty "ModifiedEnd", mSearchSpec.DateSpecs.DateModified.DateEnd, #12/12/9999#
        
        'Alternate Data Stream Specs....
        
        '.WriteProperty "ADSSpecs.AltFilterType", mSearchSpec.AlternateStreamSpecs.AltFilterType, Filter_Include
        .WriteProperty "ADSSpecs.ContainsIsRegexp", mSearchSpec.AlternateStreamSpecs.ContainsIsRegExp, False
        .WriteProperty "ADSSpecs.RegExpMaxMatches", mSearchSpec.AlternateStreamSpecs.ContainsRegExpMaxMatches, -1
        .WriteProperty "ADSSpecs.RegExpMinMatches", mSearchSpec.AlternateStreamSpecs.ContainsRegExpMinmatches, -1
        .WriteProperty "ADSSpecs.ContainsStr", mSearchSpec.AlternateStreamSpecs.ContainsStr
        .WriteProperty "ADSSpecs.MaxCount", mSearchSpec.AlternateStreamSpecs.maxcount   'max count of streams...
        .WriteProperty "ADSSpecs.MinCount", mSearchSpec.AlternateStreamSpecs.mincount   'min count of streams...
        .WriteProperty "ADSSpecs.NameSpec", mSearchSpec.AlternateStreamSpecs.nameSpec
        .WriteProperty "ADSSpecs.NameSpecIsRegExp", mSearchSpec.AlternateStreamSpecs.NameSpecisRegExp
        .WriteProperty "ADSSpecs.SizeLargerThan", mSearchSpec.AlternateStreamSpecs.SizeLargerThan
        .WriteProperty "ADSSpecs.SizeSmallerThan", mSearchSpec.AlternateStreamSpecs.SizeSmallerThan
        
        
    End With
End Sub
Public Property Get FileSpecIsRegExp() As Boolean
    FileSpecIsRegExp = mSearchSpec.FileSpecIsRegExp
End Property
Public Property Let FileSpecIsRegExp(ByVal Vdata As Boolean)
    mSearchSpec.FileSpecIsRegExp = Vdata
End Property
Public Property Get SearchStruct() As SearchSpecifications
    SearchStruct = mSearchSpec
End Property
