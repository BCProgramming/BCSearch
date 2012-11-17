VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CINIData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CIniData: Saves/Loads values from INI files.

Public Enum SettingsTypeConstants
    Setting_System
    Setting_User
End Enum
Public Type INIValue
    Precomment As String 'any comment that appears before a Value
    Name As String
    Value As Variant
End Type

Public Type INISection
    Precomment As String 'Any comment that appears before a Section.
    SectionName As String
    ValueCount As Long
    Values() As INIValue
End Type

Public Type INIFileData
    
    SectionCount As Long
    sections() As INISection

End Type


'July 19th, 2008, 6:35 PM

'Implements ISettingsProvider
Private mconfigset As String
Private mINIDate As Date, mINIfile As String
Private mINIData As INIFileData
Private mPropCount As Long
Private Function FindSectionIndex(ByVal Section As String, InStruct As INIFileData) As Long
    'find the index of the given section.
    Dim CurrSec As Long
    For CurrSec = 1 To InStruct.SectionCount
        If StrComp(InStruct.sections(CurrSec).SectionName, Section, vbTextCompare) = 0 Then
            'Found
            FindSectionIndex = CurrSec
            Exit Function
        End If
    Next CurrSec
    FindSectionIndex = -1



End Function
Public Function FindValueIndex(ByVal ValueName As String, InStruct As INISection) As Long
  'find the index of the given value.
    Dim CurrVal As Long
    For CurrVal = 1 To InStruct.ValueCount
        If StrComp(Trim$(InStruct.Values(CurrVal).Name), ValueName, vbTextCompare) = 0 Then
            'Found
            FindValueIndex = CurrVal
            Exit Function
        End If
    Next CurrVal
    FindValueIndex = -1





End Function

'Private Function CacheValid() As Boolean
'    If mPropertyCache Is Nothing Then Exit Function
'    If Dir$(mINIfile) = "" Or mINIfile = "" Then Exit Function
'    If FileDateTime(mINIfile) > mINIDate Then Exit Function
'    CacheValid = True
'
'
'End Function
Public Function ReadProfileSetting(ByVal Section As String, ByVal ValueName As String) As String
    Dim SectionIndex As Long
    Dim ValueIndex As Long
    If mINIfile = "" Then Err.Raise 9, "CINIData::ReadProfileSetting", "INI file must be loaded before being read."
    SectionIndex = FindSectionIndex(Section, mINIData)
    If SectionIndex = -1 Then
        Err.Raise 9, "CINIData::ReadProfileSetting", "Section """ & Section & """ Not Found."
        Exit Function
    End If
    ValueIndex = FindValueIndex(ValueName, mINIData.sections(SectionIndex))
    If ValueIndex = -1 Then
        Err.Raise 9, "CINIData::ReadProfileSetting", "Value """ & ValueName & """ Not Found."
        Exit Function
    End If
    
        
    
    ReadProfileSetting = mINIData.sections(SectionIndex).Values(ValueIndex).Value

End Function
Public Sub WriteProfileSetting(ByVal Section As String, ByVal ValueName As String, ByVal NewValue As String)
    'use the Index searching functions to find the correct indices of the value and Section- in cases where the returned index is -1 (not found) we create a new item.
    Dim SectionIndex As Long
    Dim ValueIndex As Long
    SectionIndex = FindSectionIndex(Section, mINIData)
        If SectionIndex = -1 Then
            'create a new section.
            mINIData.SectionCount = mINIData.SectionCount + 1
            ReDim Preserve mINIData.sections(1 To mINIData.SectionCount)
            mINIData.sections(mINIData.SectionCount).SectionName = Section
            SectionIndex = mINIData.SectionCount
        End If
    ValueIndex = FindValueIndex(ValueName, mINIData.sections(SectionIndex))
        If ValueIndex = -1 Then
            'create a new value.
            With mINIData.sections(SectionIndex)
                .ValueCount = .ValueCount + 1
                ReDim Preserve .Values(1 To .ValueCount)
                .Values(.ValueCount).Name = ValueName
                ValueIndex = .ValueCount
            End With
            
        End If
        
        
        'all the checks pan out- write the value damn it.
        mINIData.sections(SectionIndex).Values(ValueIndex).Value = NewValue
    
    
    




End Sub


Public Sub test()
    Dim teststruct As INIFileData
    Call LoadINIFile("C:\windows\win.ini", teststruct)
    Stop
    
    
End Sub
Public Sub LoadINI(ByVal FileName As String)
    LoadINIFile FileName, mINIData
    'mINIfile = Filename
End Sub
Public Sub SaveINI()
If mINIfile <> "" Then
    SaveINIFile mINIfile, mINIData
End If
End Sub
Public Sub SaveINIFile(ByVal FileName As String, FromStructure As INIFileData)
    'save the ini file data to the specified file.
    Dim fNum As Long, CurrSection As Long, CurrValue As Long
    fNum = FreeFile
    Reset
    Open FileName For Output As fNum
    'simple, in theory. Just iterate on the sections, write the comment, and the values.
    For CurrSection = 1 To FromStructure.SectionCount
    'only write the comment if it exists.
    If FromStructure.sections(CurrSection).Precomment <> "" Then
        Print #fNum, ";" & FromStructure.sections(CurrSection).Precomment
    End If
        Print #fNum, "[" & FromStructure.sections(CurrSection).SectionName & "]"
        For CurrValue = 1 To FromStructure.sections(CurrSection).ValueCount
            With FromStructure.sections(CurrSection).Values(CurrValue)
            If .Precomment <> "" Then Print #fNum, .Precomment
            Print #fNum, .Name & "=" & .Value
        
            End With
        Next CurrValue
        
    Next CurrSection
    
    
    
    Close #fNum
    
End Sub
Public Sub LoadINIFile(ByVal FileName As String, ByRef ToStructure As INIFileData)

'Load ini data into specified structure.
'the idea here is that it is far easier to load the entire file all at once, rather then simply cache individual values.
'our only problem is retaining comments between loading/saving.
Dim fNum As Long, currline As String, CommentRunner As String
fNum = FreeFile
'zero out passed structure.
On Error GoTo nofile
ToStructure.SectionCount = 0
ReDim ToStructure.sections(1 To 1)  'only first item, "global"
ToStructure.sections(1).SectionName = "Global"
Open FileName For Input As #fNum
    Do Until EOF(fNum)
    
    
    Line Input #fNum, currline
    'each line is one of several types
    'a comment, starting with ";"
    'a section name- value enclosed in square brackets []
    'a value name, to be added to the current section data, which is simply name=value
    currline = Trim$(currline)
    If currline = "" Then currline = " "
    If Asc(currline) = 59 Then
        'ignore it. it's a semicolon.
        CommentRunner = CommentRunner & vbCrLf & currline
    ElseIf Asc(currline) = 91 Then
        '"[" character. section name
        'grab the name between brackets...
        ToStructure.SectionCount = ToStructure.SectionCount + 1
        ReDim Preserve ToStructure.sections(1 To ToStructure.SectionCount)
        ToStructure.sections(ToStructure.SectionCount).Precomment = CommentRunner
        CommentRunner = ""
        ToStructure.sections(ToStructure.SectionCount).SectionName = Mid$(currline, 2, InStr(currline, "]") - 2)
    Else
        'probably a value name.
        With ToStructure.sections(ToStructure.SectionCount)
        .ValueCount = .ValueCount + 1
        
        ReDim Preserve .Values(1 To .ValueCount)
        .Values(.ValueCount).Precomment = CommentRunner
        CommentRunner = ""
        If InStr(currline, "=") <> 0 Then
            '= exists.
            
            .Values(.ValueCount).Value = Mid$(currline, InStr(currline, "=") + 1)
            .Values(.ValueCount).Name = Mid$(currline, 1, InStr(currline, "=") - 1)
        Else
            .Values(.ValueCount).Name = currline
            .Values(.ValueCount).Value = True
        End If
        End With
    End If

    Loop
    mINIfile = FileName
    Close #fNum
    Exit Sub
nofile:
    mINIfile = FileName
End Sub


'Public Function GetProfileSetting(ByVal FromINI As String, ByVal Section As String, ByVal ValueName As String) As String
'    Dim hFile As Long
'    Dim mCurrLine As String
'    Dim mCurrSection As String
'    Dim PropName As String, PropValue As String
'    If Not CacheValid() Then
'        Set mPropertyCache = New PropertyBag
'        'make the cache.
'        hFile = FreeFile
'        Open FromINI For Input As #hFile
'        mCurrSection = "Global"
'        Do Until EOF(hFile)
'        Line Input #hFile, mCurrLine
'            mCurrLine = Trim$(mCurrLine)
'            If Left$(mCurrLine, 1) = ";" Then
'            '
'
'            Else
'                If Left$(mCurrLine, 1) = "[" Then
'                    'section header.
'                    mCurrSection = Mid$(mCurrLine, 2, Len(mCurrLine) - 2)
'                Else
'                    If InStr(mCurrLine, "=") = 0 Then
'                        mPropertyCache.WriteProperty mCurrLine, True
'                    Else
'                        PropName = mCurrSection & "\" & Mid$(mCurrLine, 1, InStr(mCurrLine, "=") - 1)
'                        PropValue = Mid$(mCurrLine, Len(PropName) - Len(mCurrSection) + 1)
'                        mPropertyCache.WriteProperty PropName, PropValue
'                    End If
'
'                End If
'            End If
'        Loop
'
'        Close hFile
'        mINIfile = FromINI
'        mINIDate = FileDateTime(FromINI)
'
'
'
'    End If
'
'    GetProfileSetting = mPropertyCache.ReadProperty(Section & "\" & ValueName)
'
'
'
'
'
'
'
'
'
'
'
'End Function

Private Sub Class_Initialize()
    'Set mPropertyCache = New PropertyBag
End Sub
Private Function GetCountStr(ByVal StringSearch As String, StringFind As String, Optional ByVal Comparemode As VbCompareMethod) As Long
    GetCountStr = (Len(StringSearch) - Len(Replace$(StringSearch, StringFind, ""))) / Len(StringFind)




End Function
Private Function IsSubSectionOf(ByVal Section As String, SectionCheck As String) As Boolean
    'determines wether SectionCheck is a subsection of "Section" via String rules.
    If Right$(Section, 1) <> "\" Then Section = Section & "\"
    If StrComp(Section, Mid$(SectionCheck, 1, Len(Section)), vbTextCompare) <> 0 Then
    'nope- sectioncheck does not start with "section"
        Exit Function
    End If
 

    'now,the number of slashes cannot be more then 1 more in sectioncheck then in section.
    If GetCountStr(SectionCheck, "\") - GetCountStr(Section, "\") > 1 Then
        Exit Function
    End If
    
    IsSubSectionOf = True

End Function

Private Sub Class_Terminate()
'make changes to INI
SaveINI
End Sub

Private Sub ClearAllData()
    Kill mINIfile
End Sub

Public Sub EnumerateSections(ByVal SettingType As SettingsTypeConstants, ByVal InSection As String, sections() As String, SectionCount As Long)
'
'iterate through our array. To determine if a SectionName is a "subSection" of the passed INsection, it must start with "insection"
'and must also have a slash as the next character, which cannot be the last.
Dim CurrSection As Long
Dim FoundSections As Long, DoAdd As Boolean
For CurrSection = 1 To mINIData.SectionCount
    'if "insection" specified "global" then we return all sections, so no check is needed.
    If StrComp(InSection, "Global", vbTextCompare) = 0 Then
        DoAdd = True
    Else
        If IsSubSectionOf(InSection, mINIData.sections(CurrSection).SectionName) Then
            DoAdd = True
        End If
    
    End If
    If DoAdd Then
        FoundSections = FoundSections + 1
        ReDim Preserve sections(1 To FoundSections)
        sections(FoundSections) = mINIData.sections(CurrSection).SectionName
    
    End If


Next
    





End Sub


Public Sub EnumerateValues(ByVal SettingType As SettingsTypeConstants, ByVal InSection As String, Values() As String, ValueCount As Long)
'
'much easier to enumerate values.
Dim SectionIndex As Long, ValueLoop As Long

SectionIndex = FindSectionIndex(InSection, mINIData)
If SectionIndex = -1 Then Err.Raise 9, "CINIData::EnumerateValues", "Section """ & InSection & """ not found."
ReDim Values(1 To mINIData.sections(SectionIndex).ValueCount)
For ValueLoop = 1 To mINIData.sections(SectionIndex).ValueCount
    Values(ValueLoop) = mINIData.sections(SectionIndex).Values(ValueLoop).Name
Next ValueLoop
ValueCount = mINIData.sections(SectionIndex).ValueCount
End Sub

Private Function GetDataValue(ByVal SettingType As SettingsTypeConstants, ByVal Section As String, ByVal ValueName As String) As Variant
'
GetDataValue = Me.ReadProfileSetting(Section, ValueName)
End Function



Private Sub Init(ByVal DataSource As String, Optional ByVal configset As String = "@")

'

'CSIDL_LOCAL_APPDATA for the user- CSIDL_COMMON_APPDATA for "machine" settings.
    Dim DirUse As String
    mconfigset = configset
    If configset = "" Then configset = "@"
    If DataSource = "" Then
        'DirUse = GetSpecialFolder(CSIDL_LOCAL_APPDATA).Path
        'If Right$(DirUse, 1) <> "\" Then DirUse = DirUse & "\"
        'in    \BASeCamp\Settings
        '


        DirUse = DirUse & configset
        Do Until InStr(DirUse, "\\") = 0
            DirUse = Replace$(DirUse, "\\", "\")
        Loop
        If Right$(DirUse, 1) <> "\" Then DirUse = DirUse & "\"

        'ensure that it exists
        'EnsureFolderExists DirUse

        DirUse = DirUse & GetConfigINI(configset)
    
        'diruse is now the file to use.

    Else
        If Not Right$(DirUse, 1) = "\" Then DirUse = DirUse & "\"
        'EnsureFolderExists DirUse

        DirUse = DirUse & GetConfigINI(configset)
        
    End If

         LoadINI DirUse
'

End Sub
Private Function GetConfigINI(ByVal configset As String) As String
GetConfigINI = configset & ".INI"
End Function
Private Property Get ISettingsProvider_Name() As String
ISettingsProvider_Name = "INI"
End Property

Private Function SetDataValue(ByVal SettingType As SettingsTypeConstants, ByVal Section As String, ByVal ValueName As String, ByVal Value As Variant) As Variant
'
Me.WriteProfileSetting Section, ValueName, Value
End Function

'Private Sub ISettingsProvider_SetOwner(Vdata As CSettings)
'    Set mowner = Vdata
'End Sub

