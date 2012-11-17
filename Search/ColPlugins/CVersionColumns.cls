VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVersionColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IColumnPlugin

Dim u As CFileVersionInfo
'u.CompanyName
'u.Description
'u.FileSubType
'u.FileType
'u.Flags
'u.InternalName
'u.LegalCopyright
'u.Major
'u.Minor
'u.Revision
'u.TargetOS
'u.ProductName
'u.ProductMajor
'u.ProductMinor
'u.ProductRevision

Private Sub Class_Initialize()
'Debug.Assert False
End Sub

'File Version (Major.Minor.Revision)
'Product Version (Major.Minor.Revision)

Private Function IColumnPlugin_CanConfig() As Boolean
'
End Function

Private Sub IColumnPlugin_ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
'
Dim x As cColumn
Set x = ClickedColumn
If x.IconIndex = 0 Then
    x.SortOrder = eSortOrderDescending
Else
    If x.SortOrder = eSortOrderAscending Then x.SortOrder = eSortOrderDescending Else x.SortOrder = eSortOrderAscending

End If
x.SortType = eLVSortStringNoCase

End Sub

Private Sub IColumnPlugin_Configure()
'
End Sub

Private Property Get IColumnPlugin_Description() As String
'
IColumnPlugin_Description = "Provides columns about file versions"
End Property

Private Function IColumnPlugin_GetColumnCount() As Long
'
IColumnPlugin_GetColumnCount = 2 'File and Product Versions
End Function

Private Function IColumnPlugin_GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
'
Dim retme As ColumnData

Static fileobj As CFile, VersionObj As CFileVersionInfo
Dim newFileObj As CFile
Dim flGetAgain As Boolean
If Not fileobj Is Nothing Then
    If StrComp(fileobj.FullPath, mlstitem.Tag, vbTextCompare) = 0 Then
        'same object...
    Else
        flGetAgain = True
    End If
Else
    flGetAgain = True
End If

If flGetAgain Then
    Set fileobj = GetFile(mlstitem.Tag)
    Set VersionObj = fileobj.GetVersionInfo

End If
With VersionObj
    If ColumnIndex = 1 Then
        retme.ColumnData = Trim$(.Major) & "." & Trim$(.Minor) & "." & Trim$(.Revision)
    
    ElseIf ColumnIndex = 2 Then
        retme.ColumnData = Trim$(.productMajor) & "." & Trim$(.ProductMinor) & "." & Trim$(.ProductRevision)
    
    End If
End With
IColumnPlugin_GetColumnData = retme
End Function

Private Function IColumnPlugin_GetColumnInfo(ColumnIndex As Long) As ColumnInfo
'
Dim retme As ColumnInfo
If ColumnIndex = 1 Then
    'fileversion
    retme.ColumnDefaultWidth = 50
    retme.ColumnTitle = "File Ver."
    retme.ColumnTooltip = "The File Version of the File."
ElseIf ColumnIndex = 2 Then
    'productversion
    retme.ColumnDefaultWidth = 50
    retme.ColumnTitle = "Product Ver."
    retme.ColumnTooltip = "The Version number of the product this file is for."
End If


IColumnPlugin_GetColumnInfo = retme
End Function

Private Property Get IColumnPlugin_Name() As String
 IColumnPlugin_Name = "File Version Columns"
End Property

Private Function IColumnPlugin_PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function
