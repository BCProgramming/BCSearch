VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListViewGrouper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CListViewGrouper: Class for "grouping" file items.

Implements IFilterChangeCallback
'Groups are a fairly simple concept: a set of "groups" is created that group the files based on some attribute, such as file size.


'to make the grouping capabilities as easy as possible to implement, I will use SearchFilters to identify when a item conforms to a specific group.

'Basically, each grouping type, such as "size" and "name" will have a specific set of filters that will be used to match against the file properties.
Public Enum GroupingTypeEnum
    GroupbyAttributes
    GroupBySize
End Enum
Private Type GroupItemData
    GroupKey As String 'a "group key" used as the key to the group added to the collection in the listview.
                        'also used as a key into a local collection variable where the same group objects are added.
    GroupCaption As String 'For example; say, "read-only" for a group in the group by attributes.
    GroupMatchFilter As CSearchFilter 'if a file matches this filter, then it will be part of this group. filters should be constructed so that
                            'any file only conforms to one of the match filters in a specific group set; for example, the file attributes group will set the filter to match the attributes exactly.
End Type
Private Type GroupingDataStruct
    GroupName As String 'For example- "Size" if it groups by size.
    GroupItemsCount As Long
    GroupItems() As GroupItemData
End Type
Private Type GroupSetData
    GroupsetsCount As Long
    GroupSets() As GroupingDataStruct
End Type
Private mGroupInfo As GroupSetData


Private Sub InitStructure(groupdata As GroupSetData, Onlvw As vbalListViewCtl)

'    With groupdata
'        .GroupsetsCount = 1
'        ReDim .GroupSets(1 To 2)
'
'
'        InitAttributeFilters .GroupSets(1)
'        InitSizeFilters .GroupSets(2)
'
'
'    End With



End Sub
Private Sub InitFilenameFilters(ToStructure As GroupingDataStruct)
'Pretty basic; group by their first letter.



End Sub
Private Sub InitFileTypeFilters(ToStructure As GroupingDataStruct, Onlvw As vbalListViewCtl)
    'add one group for each unique file type.
End Sub
Private Sub InitFileExtensionFilters(ToStructure As GroupingDataStruct, Onlvw As vbalListViewCtl)
    'add one group for each unique file extension.
End Sub
'
Private Sub InitSizeFilters(ToStructure As GroupingDataStruct, Onlvw As vbalListViewCtl)

    Dim ArraySizes() As Variant, I As Long
    
    '1K,2K,4K,8K,16K,
    'start from 16 bytes.
    
    'double for each following item up to a maximum of 1TB.
    ToStructure.GroupName = "SIZE"
    Dim Currbytecount As Variant, curritems As Long
    
    Currbytecount = 16
    Do Until Currbytecount > 1024 ^ 4
        curritems = curritems + 1
        ReDim Preserve ArraySizes(1 To curritems)
        ArraySizes(curritems) = Currbytecount
        Currbytecount = Currbytecount * 2
        
    
    
    Loop
    curritems = curritems + 1
    ReDim Preserve ArraySizes(1 To curritems)
    ArraySizes(curritems) = ArraySizes(curritems - 1)
    'duplicate the last item.
    ToStructure.GroupItemsCount = curritems
    ReDim ToStructure.GroupItems(1 To ToStructure.GroupItemsCount)
    
    For I = LBound(ArraySizes) To UBound(ArraySizes)
    With ToStructure.GroupItems(I)
        Set .GroupMatchFilter = New CSearchFilter
        .GroupMatchFilter.SearchOperation = Filter_Include
        .GroupKey = "SIZEGROUP" & Trim$(Str(I))
        Set .GroupMatchFilter.Callback = Me
        If I = LBound(ArraySizes) Then
            'First item, make "smaller then" filter for first item.
            .GroupCaption = "Smaller than " & bcfile.FormatSize(ArraySizes(I), Format_ShowThousandsSeparator)
            
            .GroupMatchFilter.SizeSmallerThan = ArraySizes(I)
            
        ElseIf I = UBound(ArraySizes) Then
            'last item, make "larger then" filter for this last item.
            .GroupCaption = "Larger than " & bcfile.FormatSize(ArraySizes(I), Format_ShowThousandsSeparator)
            .GroupMatchFilter.SizeLargerThan = ArraySizes(I)
            '
        Else
            'in between. label as "larger than" the item before us and smaller then this item,
            'but only use "smaller then" this item in the filter itself.
            .GroupCaption = "Between " & bcfile.FormatSize(ArraySizes(I - 1)) & " and " & bcfile.FormatSize(ArraySizes(I))
            .GroupMatchFilter.SizeSmallerThan = ArraySizes(I)
            
        
        
        End If
    End With
    
    Next I










End Sub
'The following is the Old procedure- it attempted to dynamically decide what filesizes to use to group with based on the sizes of the files in the listview.

'Private Sub InitSizeFilters(ToStructure As GroupingDataStruct, Onlvw As vbalListViewCtl)
'    'group by file size ranges.
'    'Note, size ranges will be taken from the actual file listing.
'    'we will go for a number of groups about 1 for every 20 or so items.
'
'    Dim Numgroups As Long
'
'    'if the listview has less then 20 items, then we have 4 specified groups.
'    Select Case Onlvw.ListItems.Count
'    Case Is < 20
'        Numgroups = 4
'    Case Is < 50
'    'between 20 and 50.
'    Numgroups = 10
'    Case Is < 100
'    '50-100
'    Numgroups = 20
'    Case Is < 500
'    '40 groups
'    Numgroups = 40
'    Case Is > 500
'        'larger then 500
'        'go with 50 groups.
'        Numgroups = 50
'    End Select
'    'if there are 20-50 items, we group by tens (number of items \10)= number of groups.
'
'    '50>100 items, we group via 20's.
'
'    'higher then 100 and we group by 50's.
'
'    'once we have the number of groups, we calculate increments based on the largest and the smallest file in the listview, and create the "filters" as appropriate.
'
'    'Alright... Create an Array of all the file sizes...
'    Dim Filesizes() As Double
'    Dim I As Long, tempfileobj As CFile, loopItem As cListItem
'    Dim runningtotal As Double, Currmax As Double, CurrMin As Double, Average As Double
'    On Local Error Resume Next
'    ReDim Filesizes(1 To Onlvw.ListItems.Count)
'    For I = 1 To Onlvw.ListItems.Count
'        Set loopItem = Onlvw.ListItems.Item(I)
'        If loopItem.Tag <> "" Then
'            Set tempfileobj = bcfile.GetFile(loopItem.Tag)
'            If Err <> 0 Then
'                Filesizes(I) = -1  'means, "don't count me, dammit"
'            Else
'                Filesizes(I) = tempfileobj.size
'                runningtotal = runningtotal + Filesizes(I)
'                If Filesizes(I) > Currmax Then Currmax = Filesizes(I)
'                If Filesizes(I) < CurrMin Or CurrMin = 0 Then CurrMin = Filesizes(I)
'            End If
'        End If
'    Next I
'
'    Average = runningtotal / Onlvw.ListItems.Count
'    'we have currmin,currmax, and average... although I suppose average is kinda pointless.
'
'
'    'So what do we do? well,we create "numgroups" sizing groups between the minimum and maximum values:
'
'    Dim RangeSizeAccurate As Double
'    Dim CurrRangeStart As Double, CurrRangeEnd As Double, CurrIndex As Long
'    RangeSizeAccurate = CurrMin + ((Currmax - CurrMin) / Numgroups)
'    'rangesizeaccurate: the specific size, in bytes, that each group will be.
'    CurrRangeStart = CurrMin
'    CurrIndex = 1
'    Do
'        ToStructure.GroupItemsCount = ToStructure.GroupItemsCount + 1
'        ReDim Preserve ToStructure.GroupItems(1 To ToStructure.GroupItemsCount)
'        With ToStructure.GroupItems(ToStructure.GroupItemsCount)
'            Set .GroupMatchFilter = New CSearchFilter
'            Set .GroupMatchFilter.Callback = Me
'            .GroupKey = "SIZERNUM" & Trim$(Str(ToStructure.GroupItemsCount))
'        'CurrRangeEnd = CurrRangeStart + RangeSizeAccurate
'        .GroupMatchFilter.SearchOperation = Filter_Include
'        If CurrRangeEnd = 0 Then
'            'this is the first item; this means the filter will be for any size less than the calculated range ending.
'            CurrRangeEnd = CurrRangeStart + RangeSizeAccurate
'            .GroupCaption = "less than " & bcfile.FormatSize(CurrRangeEnd, Format_ShowTotalInBrackets)
'
'            .GroupMatchFilter.SizeSmallerThan = CurrRangeEnd
'        ElseIf ToStructure.GroupItemsCount = Numgroups Or ((CurrRangeEnd + RangeSizeAccurate) > Currmax) Then
'            'this is the last item;
'            'therefore just use a "larger then" filter...
'            .GroupCaption = "Larger than " & bcfile.FormatSize(CurrRangeEnd, Format_ShowThousandsSeparator)
'            .GroupMatchFilter.SizeLargerThan = CurrRangeEnd
'            Exit Do
'
'        Else
'            'interesting tidbit:
'            'while the apparent result will be that files will be displayed in the appropriate groups, each filter is only a "sizesmallerthan" filter (well, except the last one). because the code that creates the groups
'            'goes through the array sequentially and stops when the first match is encountered, it will appear to be in the proper group. Also, our captions are used to reflect that, even though the filters used for
'            'the group do not conform to the group caption.
'            .GroupCaption = "between " & bcfile.FormatSize(CurrRangeEnd, Format_ShowThousandsSeparator) & " and " & bcfile.FormatSize(CurrRangeEnd + RangeSizeAccurate, Format_ShowThousandsSeparator)
'            .GroupMatchFilter.SizeSmallerThan = CurrRangeEnd + RangeSizeAccurate
'
'
'
'
'        End If
'        End With
'        'CurrRangeEnd = CurrRangeEnd + RangeSizeAccurate
'
'        'instead of adding "rangesizeaccurate" directly, we want the increments to be fewer closer to the average.
'
'        'the closer to the average "CurrrangeEnd" is, the smaller the increments should be, based on a percentage of some sort.
'        If Average <> 0 Then
'        Dim absdiff As Double
'        Dim diffpct As Double
'        absdiff = Abs(Average - CurrRangeEnd)
'        'now, we divide this difference by the average itself, to get a percentage...
'
'        diffpct = absdiff / Average
'
'        'we use this diffpct to multiply the used increment range...
'        CurrRangeEnd = CurrRangeEnd + (RangeSizeAccurate * diffpct)
'
'
'
'
'
'
'
'        End If
'    Loop
'
'
'    'Debug.Assert False
'
'
'End Sub
'
Private Sub InitAttributeFilters(ToStructure As GroupingDataStruct)

    Dim fattributes As FileAttributeConstants
    Dim Allattributes As Variant
    Dim createdfilter As CSearchFilter
    
    ToStructure.GroupName = "Attributes"
    

    Dim Attribs() As FileAttributeConstants, AttribsStr() As String
    Dim I As Long
    Attribs = bcfile.GetAllFileAttributes()
    
    'iterate through each one, populate attribsstr with the string representation.
    ReDim AttribsStr(1 To UBound(Attribs))
    ToStructure.GroupItemsCount = UBound(Attribs)
    ReDim ToStructure.GroupItems(1 To ToStructure.GroupItemsCount)
    
    For I = 1 To UBound(Attribs)
        AttribsStr(I) = bcfile.GetAttributeString(Attribs(I), True)
    
        'add a new group...
        With ToStructure.GroupItems(I)
            .GroupKey = bcfile.GetAttributeString(Attribs(I), False)
            .GroupCaption = AttribsStr(I)
            Set .GroupMatchFilter = New CSearchFilter
            With .GroupMatchFilter
                Set .Callback = Me
                .AttributesMatchExact = True
                .Attributes = Attribs(I)
                .SearchOperation = Filter_Include
            
            End With
        
        
        
        End With
    
    
    Next I
    
    
    
    
    
    
    
    
    



End Sub
Public Sub Ungroup(Onlvw As vbalListViewCtl)
Onlvw.ItemGroups.enabled = False
Onlvw.ItemGroups.Clear



End Sub
Private Sub PerformGroup(Onlvw As vbaBClListViewLib6.vbalListViewCtl, usegroup As GroupingDataStruct)
'using the data in usegroup assign each item in onlvw to a group based on the properties of the file in it's "tag" property.
Dim loopItem As cListItem, currItem As Long
Dim FileTag As bcfile.CFile, wasgrouped As Boolean
Dim groupsadded As Collection, I As Long, addedgroup As Object
Set groupsadded = New Collection
'First, add the groups...
Onlvw.ItemGroups.Clear
On Error Resume Next
For I = LBound(usegroup.GroupItems) To UBound(usegroup.GroupItems)
'    If usegroup.GroupItems(I).GroupKey = "A" Then Stop
    Set addedgroup = Onlvw.ItemGroups.Add(, usegroup.GroupItems(I).GroupKey, usegroup.GroupItems(I).GroupCaption)
    If Err = 0 Then
    
    'Debug.Print "equal:" & addedgroup.State = eLVGroupHidden
    Debug.Print TypeName(addedgroup)
    Debug.Assert TypeName(addedgroup) = "cItemGroup"
    addedgroup.State = eLVGroupNormal
    addedgroup.Align = eLVGroupHeaderAlignLeft
    groupsadded.Add addedgroup, usegroup.GroupItems(I).GroupKey
    Else
    Debug.Assert False
    Err.Clear
    End If
Next I
'last, add an extra "unspecified" group, in case some items don't match, as well as an "error" field.
Set addedgroup = Onlvw.ItemGroups.Add(, "BCUnspecified", "Unspecified")
groupsadded.Add addedgroup, "BCUnspecified"
Set addedgroup = Onlvw.ItemGroups.Add(, "BCError", "Error")
addedgroup.State = eLVGroupNormal
groupsadded.Add addedgroup


'now that the groups are all added, Loop through each listitem...
Onlvw.Visible = False
On Error GoTo UseErrorGroup 'when an error occurs.... use the error group.
For currItem = 1 To Onlvw.ListItems.Count
    Debug.Print "grouping item " & currItem
    Set loopItem = Onlvw.ListItems(currItem)
    
    If loopItem.Tag = "" Then
        'no tag... this is an error. This will be in the "error" group.
        loopItem.Group = groupsadded.Item("BCError")
    Else
        'there is a tag... let's do some work:
        Set FileTag = bcfile.GetFile(loopItem.Tag)
        'alright, now, iterate through our "groups" structure....
        'set the flag to false...
        wasgrouped = False
        For I = LBound(usegroup.GroupItems) To UBound(usegroup.GroupItems)
            With usegroup.GroupItems(I)
            'Debug.Print .GroupMatchFilter.Attributes, getattributestring(.GroupMatchFilter.Attributes)
                If Not .GroupMatchFilter Is Nothing Then  'sanity check...
                Debug.Print "smaller than:" & .GroupMatchFilter.SizeSmallerThan & " Larger than:" & .GroupMatchFilter.SizeLargerThan
                    If .GroupMatchFilter.SizeSmallerThan > FileTag.size Or (.GroupMatchFilter.SizeLargerThan < FileTag.size And .GroupMatchFilter.SizeLargerThan <> 0) Then
                        'Debug.Assert False
                    End If
                    
                    If .GroupMatchFilter.FilterResult(FileTag) Then
                        'if this tests as "true" then it belongs to usegroup.groupitems(I).
                        loopItem.Group = groupsadded.Item(.GroupKey)
                        wasgrouped = True
                        Exit For
                    End If
                End If
            End With
        Next I
        'if the flag is still false, the item was not placed in one of the groups.
        If Not wasgrouped Then
            'add it to the "unspecified" group.
            loopItem.Group = groupsadded.Item("BCUnspecified")
        
        End If
    
    End If
nextitem:
Next
Onlvw.ItemGroups.enabled = True
Onlvw.Visible = True
Exit Sub
UseErrorGroup:
Debug.Assert False
loopItem.Group = groupsadded.Item("BCError")
Resume
Resume nextitem

End Sub
Public Sub AssignGroups(Onlvw As vbalListViewCtl, ByVal GroupBy As GroupingTypeEnum)
    Static LastListview As vbalListViewCtl, lastcount As Long
    Static InitializedGroups() As Boolean
    If Not LastListview Is Onlvw Or lastcount <> Onlvw.ListItems.Count Then
    'Bug: will bug somewhat if a search is performed, a grouping is done, and another search is performed that returns the same number of results.
    'not sure how to remedy that without involving the listview itself; we could theoretically test the first listitem if it's the same object; this will be the same with the same set of results.
    
        'InitStructure mGroupInfo, onlvw
        'In this block, we assume this is a new listview or search results altogether, so we must clear ALL cached grouping data from memory and start caching group data for this one.
        ReDim InitializedGroups(1 To 2)
        'currently two items: attributes, and size, respectively.
        'defaults to false, which works just fine.
        
        ReDim mGroupInfo.GroupSets(1 To 2)
        mGroupInfo.GroupsetsCount = 2
        
        
        
        Set LastListview = Onlvw
    End If

If GroupBy = GroupbyAttributes Then
    If Not InitializedGroups(1) Then
        InitializedGroups(1) = True
        InitAttributeFilters mGroupInfo.GroupSets(1)
    End If
    
    
    PerformGroup Onlvw, mGroupInfo.GroupSets(1)
ElseIf GroupBy = GroupBySize Then
    If Not InitializedGroups(2) Then
        InitializedGroups(2) = True
        'InitAttributeFilters mGroupInfo.GroupSets(2)
        InitSizeFilters mGroupInfo.GroupSets(2), Onlvw
    End If
    
    PerformGroup Onlvw, mGroupInfo.GroupSets(2)


   ' PerformGroup onlvw, mGroupInfo.GroupSets(2)

End If


lastcount = Onlvw.ListItems.Count

End Sub







Private Sub Class_Initialize()
   ' InitStructure mGroupInfo
End Sub

Private Sub IFilterChangeCallback_Change(changedObj As bcfile.CSearchFilter)
'
End Sub
