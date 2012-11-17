VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAlternateStreamColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IColumnPlugin Implementation for displaying Data on a Files Alternate data streams.




'current Columns implemented: ADS Count, ADS Names




Implements IColumnPlugin


Private Function IColumnPlugin_CanConfig() As Boolean
'
End Function

Private Sub IColumnPlugin_ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
'
End Sub

Private Sub IColumnPlugin_Configure()
'
End Sub

Private Property Get IColumnPlugin_Description() As String
'
IColumnPlugin_Description = "Displays columns about the Alternate data streams of a file"
End Property

Private Function IColumnPlugin_GetColumnCount() As Long
'
IColumnPlugin_GetColumnCount = 2
End Function

Private Function IColumnPlugin_GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
'
Dim Filename As String
Dim gotfile As CFile
Dim coldata As ColumnData
Dim LoopStream As CAlternateStream
Filename = mlstitem.Tag
'the tag is the filename. Always has been.
Set gotfile = bcfile.GetFile(Filename)

If ColumnIndex = 1 Then 'count
    coldata.ColumnData = gotfile.AlternateStreams.Count

ElseIf ColumnIndex = 2 Then 'names
    'names...
    Dim buildstr As String
    'If gotfile.AlternateStreams.Count > 1 Then Stop
    For Each LoopStream In gotfile.AlternateStreams
        buildstr = buildstr & LoopStream.Name & ";"
    
    
    Next
    coldata.ColumnData = buildstr
End If


IColumnPlugin_GetColumnData = coldata

End Function

Private Function IColumnPlugin_GetColumnInfo(ColumnIndex As Long) As ColumnInfo
    '
    Dim colinfo As ColumnInfo
    If ColumnIndex = 1 Then
        colinfo.ColumnTitle = "ADS Count"
        colinfo.ColumnDefaultWidth = 48
        colinfo.ColumnTooltip = "Count of Alternate Data Streams"
    ElseIf ColumnIndex = 2 Then
        colinfo.ColumnTitle = "ADS Names"
        colinfo.ColumnDefaultWidth = 128
        colinfo.ColumnTooltip = "The names of Alternate Data streams in the file"
    End If
    IColumnPlugin_GetColumnInfo = colinfo
End Function

Private Property Get IColumnPlugin_Name() As String
'
IColumnPlugin_Name = "ADS Column Provider"
End Property

Private Function IColumnPlugin_PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function
