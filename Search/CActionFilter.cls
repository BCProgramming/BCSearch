VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CActionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IProgress
'CActionFilter
'Acts on the results in the search listview to perform things such as moving, renaming, etc.

'when the action filter form is told to "execute" the actions, it does so this way:

'loops through all action filters
'for each one, sends in the listitems collection, or, a collection of selected items. (we either recieve a ListItems collection or a plain ol' collection, either way,
'the items are all cListItem objects.

'whatever changes a action filter performs on a file must be reflected in the listitem objects that were passed in. This is so that not only does the
'listview stay current but also so action filters down the list have valid file data to work with, such as name and extension.

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_RENAME As Long = &H4
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2


Private Type ActionRenameStruct
    NewNameMask As String
End Type
Private Type ActionCopyStruct
    DestinationMask As String
End Type
Private Type ActionScriptedStruct
    ScriptLanguage As String
    ScriptCode As String
End Type
Public Enum AttributeModificationMethod
    attribute_toggle
    attribute_add
    attribute_Remove
    attribute_Set
End Enum
Private Type ActionAttributeStruct
    mAttributeModifyMode As AttributeModificationMethod
    mAttributes As FileAttributeConstants

End Type
Private Type ActionFilterStruct
    Actiontype As ActionFilterTypeEnum
    Description As String
    RenameStruct As ActionRenameStruct
    CopyStruct As ActionCopyStruct
    ScriptedStruct As ActionScriptedStruct
    AttributeStruct As ActionAttributeStruct
End Type

Public Enum ActionFilterTypeEnum
    
    Action_Rename
    Action_Move
    Action_Copy
    Action_Attributes
    
    
    Action_Scripted = 1000
End Enum
Private mActionFilterType As ActionFilterTypeEnum
Private ActionStruct As ActionFilterStruct
Private mkey As String
Private mScriptControl As ScriptControl
'Only rename action supported at this point.
Public Property Let Key(ByVal Vdata As String)
    mkey = Vdata
End Property
Public Property Get Key() As String
    Key = mkey
End Property
Public Property Let ActionScripted_Language(ByVal Vdata As String)
    ActionStruct.ScriptedStruct.ScriptLanguage = Vdata
    'Set mScriptControl = New ScriptControl
    'mScriptControl.Language = vData
End Property
Public Property Get ActionScripted_Language() As String
    ActionScripted_Language = ActionStruct.ScriptedStruct.ScriptLanguage
    
End Property

Public Property Let ActionScripted_Code(ByVal Vdata As String)
    Set mScriptControl = New ScriptControl
    mScriptControl.Language = ActionStruct.ScriptedStruct.ScriptLanguage
    mScriptControl.AddCode ActionStruct.ScriptedStruct.ScriptCode
End Property
Public Property Get ActionScripted_Code() As String
    ActionScripted_Code = ActionStruct.ScriptedStruct.ScriptCode
End Property

Public Property Let ActionRename_Mask(ByVal Vdata As String)
    ActionStruct.RenameStruct.NewNameMask = Vdata
End Property
Public Property Get ActionRename_Mask() As String
    ActionRename_Mask = ActionStruct.RenameStruct.NewNameMask
End Property
Public Property Let ActionMove_Mask(ByVal Vdata As String)
    ActionStruct.CopyStruct.DestinationMask = Vdata
End Property
Public Property Get ActionMove_Mask() As String
    ActionMove_Mask = ActionStruct.CopyStruct.DestinationMask
End Property
Public Property Let ActionAttributes_AttributeModifyMode(ByVal Vdata As AttributeModificationMethod)
    ActionStruct.AttributeStruct.mAttributeModifyMode = Vdata
    
End Property
Public Property Get ActionAttributes_AttributeModifyMode() As AttributeModificationMethod
    ActionAttributes_AttributeModifyMode = ActionStruct.AttributeStruct.mAttributeModifyMode
End Property
Public Property Let ActionAttributes_Attributes(ByVal Vdata As FileAttributeConstants)
    ActionStruct.AttributeStruct.mAttributes = Vdata
End Property
Public Property Get ActionAttributes_Attributes() As FileAttributeConstants
    ActionAttributes_Attributes = ActionStruct.AttributeStruct.mAttributes
End Property
Public Property Let Description(ByVal Vdata As String)
    ActionStruct.Description = Vdata
End Property
Public Property Get Description() As String
    Description = ActionStruct.Description
End Property
Friend Property Let Actiontype(ByVal Vdata As ActionFilterTypeEnum)
    mActionFilterType = Vdata
End Property
Friend Property Get Actiontype() As ActionFilterTypeEnum
    Actiontype = mActionFilterType
End Property
Private Function CreateReplacementDictionary(ForObject As cListItem, Optional n As Long = 1) As Dictionary

'ForObject: a Listitem to create the array for.

'returns:
'a dictionary where each key is a substitution item and each value for those keys is the value of that key.




Dim retobj As Dictionary
Dim FileObject As CFile
Dim dateformat As String
dateformat = "MM-DD-YYYY HH-NN-SS"
If ForObject.Tag = "" Then Debug.Assert False
Set FileObject = bcfile.GetFile(ForObject.Tag)


Set retobj = New Dictionary


retobj.Add "filename", FileObject.Filename
retobj.Add "extension", FileObject.Extension
retobj.Add "basename", FileObject.basename
retobj.Add "filesize", FileObject.size
retobj.Add "datecreated", Format$(FileObject.DateCreated, dateformat)
retobj.Add "datemodified", Format$(FileObject.DateModified, dateformat)
retobj.Add "datelastaccessed", Format$(FileObject.DateLastAccessed, dateformat)
retobj.Add "folder", FileObject.Directory.Path
retobj.Add "size", FileObject.size
retobj.Add "seq", n
'ok.... now add folders based on their location in the path.


Dim pathSplitted() As String

Dim CurrIndex As Long

pathSplitted = SplitPath(FileObject.Directory.Path)
For CurrIndex = 0 To UBound(pathSplitted)
    retobj.Add "folder" & Trim$(CurrIndex + 1), pathSplitted(CurrIndex)
    retobj.Add "folder-" & Trim$(UBound(pathSplitted) - (CurrIndex + 1)), pathSplitted(CurrIndex)


Next CurrIndex





Set CreateReplacementDictionary = retobj








End Function

Private Function Substitute(ByVal StrMask As String, Replacementdict As Dictionary) As String

    If Replacementdict Is Nothing Then
        
        Exit Function
    End If

    Dim FoundPos As Long
    Dim nextpos As Long
    Dim dictkeys() As Variant, currkey As Long
    Dim usekey As String, CurrPos As Long
    Dim searchfor As String
    dictkeys = Replacementdict.Keys()
    
    For currkey = LBound(dictkeys) To UBound(dictkeys)
        usekey = UCase$(dictkeys(currkey))
        
        
        
        'first, look for a percent sign followed by the key...
        CurrPos = 1
        
        
        searchfor = "%" & usekey
        CurrPos = InStr(CurrPos, StrMask, searchfor, vbTextCompare)
        Do While CurrPos > 0
        
        
        'we found an instance of this mask.
        'if the next character is a percent sign, just replace the entire "variable" with the appropriate value from the dictionary.
        If Mid$(StrMask, CurrPos + Len(searchfor), 1) = "%" Then
            StrMask = Left$(StrMask, CurrPos - 1) & Replacementdict.Item(dictkeys(currkey)) & Mid$(StrMask, CurrPos + Len(searchfor) + 1)
        ElseIf Mid$(StrMask, CurrPos + Len(searchfor), 1) = ":" Then
        
            'I know. it's awful and stupid and annoying.
            'basic manipulation support.
            
            '%variable:left:<num>%  take the leftmost <num> characters.
            
            '%variable:right:<num>% take the rightmost <num> characters.
            'split at the colon... (HAHAHA)
            Dim SplitContents() As String, substring As String
            Dim splitstring As String
            Dim NextPercent As Long
            NextPercent = InStr(CurrPos + 1, StrMask, "%")
            splitstring = Mid$(StrMask, CurrPos + 1, NextPercent - 2)
            SplitContents = Split(splitstring, ":")
            
                        'Exit Function
            '0=mask name, 1= function, and anything after is an "argument" to that function.
            'This part has been split into a separate function.
            
            substring = PerformSubstitutionFunc(SplitContents(), Replacementdict.Item(dictkeys(currkey)))
            
            'plop substring back into the strmask...
            StrMask = Left$(StrMask, CurrPos - 1) & substring & Mid$(StrMask, NextPercent + 1)
            
            

        
        
        
        Else
            'doesn't refer to this variable at all, probably.
            'move the currpos pointer to the next percentage sign...
            CurrPos = InStr(CurrPos + Len(searchfor), StrMask, "%") + 1
            
        
        End If
        
        
        
        
        
        
        
        CurrPos = InStr(CurrPos, StrMask, searchfor, vbTextCompare)
        
        Loop
    
    
    Next currkey
    
    
    Substitute = StrMask
    


End Function
Private Function PerformSubstitutionFunc(SubstParams() As String, VariableValue As String) As String

    Dim FuncPerform As String
    Dim FuncArguments() As String
    Dim I As Long
    Dim paramcount As Long
    Dim Ret As String
    
    
    'setup...
    If IsNumeric(SubstParams(1)) Then
        FuncPerform = ""
        FuncArguments = SubstParams
    Else
        FuncPerform = SubstParams(1)
        ReDim FuncArguments(1 To UBound(SubstParams) - 1)
        
        For I = 2 To UBound(SubstParams)
            FuncArguments(I - 1) = SubstParams(I)
        Next
        
    
    
    
    End If
    paramcount = UBound(FuncArguments)
    Select Case UCase$(FuncPerform)
    Case "LEFT"
        If paramcount < 1 Then GoTo Notoptional
        Ret = Left$(VariableValue, Val(FuncArguments(1)))
    Case "RIGHT"
        If paramcount < 1 Then GoTo Notoptional
        Ret = Right$(VariableValue, Val(FuncArguments(1)))
    Case "MID"
        If paramcount < 1 Then GoTo Notoptional
        If paramcount = 1 Then
            Ret = Mid$(VariableValue, Val(FuncArguments(1)))
        Else
        
            Ret = Mid$(VariableValue, Val(FuncArguments(1)), Val(FuncArguments(2)))
        End If
    End Select
    PerformSubstitutionFunc = Ret


    Exit Function
Notoptional:
    CDebug.Post "Argument not optional for CActionFilter..."
    


End Function
Private Sub DoScriptedAction(items As cListItems, ProgressObj As IProgress)
'
'script needs to have a function:
'DoAction(item)
'or
'DoActions(items)


Dim LoopProcedure As Procedure, itemaction As Object
Dim hasPluralFunc As Boolean, hasSingularFunc As Boolean
For Each LoopProcedure In mScriptControl.Procedures
    If LoopProcedure.Name = "DoActions" And LoopProcedure.NumArgs = 1 Then
        hasPluralFunc = True
    ElseIf LoopProcedure.Name = "DoAction" And LoopProcedure.NumArgs = 1 Then
        hasSingularFunc = True
    
    End If

Next
Dim I As Long
If Not mScriptControl Is Nothing Then
    If hasPluralFunc Then
        mScriptControl.CodeObject.DoActions items
    ElseIf hasSingularFunc Then
        For I = 1 To items.Count
            Set itemaction = items.Item(I)
            mScriptControl.CodeObject.DoAction itemaction
        Next
    'mScriptControl.CodeObject.DoAction itemaction

    End If
End If


End Sub

Private Sub DoRenameAction(itemsrename As Object, ProgressObj As IProgress)
'


    '
    Dim newname As String, I As Long, currItem As cListItem, passname As String
    Dim fullnew As String, fullcurr As String, gotfile As CFile
    'build a Null delimited list of source files.
    'for the FO_MOVE operation, the first
    With ActionStruct.RenameStruct
    For I = 1 To itemsrename.Count
        Set currItem = itemsrename.Item(I)
        Set gotfile = GetFile(currItem.Tag)
        
        newname = DoSubstitute(.NewNameMask, currItem, I)
        'rename curritem.tag to newname.
        
        gotfile.Rename newname, CurrApp.MainForm.hWnd, False
        FrmSearch.RefreshItemData currItem, gotfile
     Next I
     End With



End Sub
Private Sub DoCopyAction(ItemsMove As Object, ProgressObj As IProgress)
   '
    Dim newname As String, I As Long, currItem As cListItem
    Dim fullnew As String, fullcurr As String
    'build a Null delimited list of source files.
    'for the FO_COPY operation, the first
    With ActionStruct.CopyStruct
    For I = 1 To ItemsMove.Count
        
        Set currItem = ItemsMove.Item(I)
        newname = DoSubstitute(.DestinationMask, currItem)

     Next I
     End With

    
End Sub
Private Sub DoMoveAction(ItemsMove As Object, ProgressObj As IProgress)
    '
    Dim newname As String, I As Long, currItem As cListItem
    Dim FileGot As CFile

    With ActionStruct.CopyStruct
    For I = 1 To ItemsMove.Count
        Set currItem = ItemsMove.Item(I)
        Set FileGot = GetFile(currItem.Tag)
        newname = DoSubstitute(.DestinationMask, currItem)
        
    
    
    
     Next I
     End With

    
End Sub
Public Sub DoAction(OnCollection As Object, Optional ProgressObj As IProgress = Nothing)
'
Dim currItem As Long
Dim useitem As cListItem
If ProgressObj Is Nothing Then Set ProgressObj = Me

For currItem = 1 To OnCollection.Count
    Set useitem = OnCollection.Item(currItem)
    Select Case ActionStruct.Actiontype
        Case Action_Scripted
            DoScriptedAction OnCollection, ProgressObj
        Case Action_Rename
            DoRenameAction OnCollection, ProgressObj
        Case Action_Move
            DoMoveAction OnCollection, ProgressObj
        Case Action_Copy
            DoCopyAction OnCollection, ProgressObj
    End Select
    


Next currItem


End Sub
Public Function DoSubstitute(ByVal StrMask As String, currItem As cListItem, Optional n As Long = 1) As String
'public- I use it elsewhere, even though I don't have the actionFilters dialog enabled as of yet.
Dim userep As Dictionary
Set userep = CreateReplacementDictionary(currItem, n)
DoSubstitute = Substitute(StrMask, userep)

End Function
Public Function GetActiontypeString() As String
Select Case True
    Case mActionFilterType = Action_Rename
        GetActiontypeString = "Rename"
    Case mActionFilterType = Action_Move
        GetActiontypeString = "Move"
    Case mActionFilterType = Action_Copy
        GetActiontypeString = "Copy"
    Case mActionFilterType = Action_Attributes
        GetActiontypeString = "Attributes"
    Case mActionFilterType = Action_Scripted
        GetActiontypeString = "Scripted"
    
    End Select
End Function

Private Sub Class_Initialize()
    Set mScriptControl = New ScriptControl
End Sub

Private Sub IProgress_UpdateUI(ByVal PercentComplete As Double, ByVal StatusMessage As String)
    'much easier this way...
End Sub
