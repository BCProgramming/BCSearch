VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtraColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'currently Supports:

'Slack Space: size - size on disk.
'Attributes



Implements IColumnPlugin

Private Function IColumnPlugin_CanConfig() As Boolean
'
IColumnPlugin_CanConfig = False
End Function

Private Sub IColumnPlugin_ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
'
    Dim castcolumn As cColumn
    Set castcolumn = ClickedColumn
    If ColIndex = 1 Then
    
        castcolumn.SortType = eLVSortNumeric
    ElseIf ColIndex = 2 Then
        castcolumn.SortType = eLVSortStringNoCase
    End If

End Sub

Private Sub IColumnPlugin_Configure()
'
End Sub

Private Property Get IColumnPlugin_Description() As String
'
End Property

Private Function IColumnPlugin_GetColumnCount() As Long
'
IColumnPlugin_GetColumnCount = 2
End Function

Private Function IColumnPlugin_GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
'
    Dim gotfile As CFile
    Dim makedata As ColumnData
    Set gotfile = GetFile(mlstitem.Tag)
    If ColumnIndex = 1 Then
    makedata.ColumnData = Str(gotfile.size - gotfile.compressedsize) & " bytes"
    ElseIf ColumnIndex = 2 Then
        'attributes.
        makedata.ColumnData = Replace$(bcfile.GetAttributeString(gotfile.FileAttributes, False), ",", "")
    End If
IColumnPlugin_GetColumnData = makedata
End Function

Private Function IColumnPlugin_GetColumnInfo(ColumnIndex As Long) As ColumnInfo
'
Dim makeinfo As ColumnInfo
If ColumnIndex = 1 Then
    makeinfo.ColumnDefaultWidth = 50
    makeinfo.ColumnTitle = "Slack space"
    makeinfo.ColumnTooltip = "Size of the file minus its size on disk."
    
ElseIf ColumnIndex = 2 Then
    makeinfo.ColumnDefaultWidth = 50
    makeinfo.ColumnTitle = "Attributes"
    makeinfo.ColumnTooltip = "File Attributes"
    
End If
IColumnPlugin_GetColumnInfo = makeinfo
End Function

Private Property Get IColumnPlugin_Name() As String
IColumnPlugin_Name = "Supplemental Columns Plugin"
End Property

Private Function IColumnPlugin_PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function
