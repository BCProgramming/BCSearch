VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMP3Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IColumnPlugin

Private mColumnInfo() As ColumnInfo


Private Sub Class_Initialize()

'modmp3.MP3read_GetTagInfo_ID3v2(mp3filename,songtitle,songartist,songalbum,songyear,songcomment,songtracknumber,songgenre,songcomposer,songoriginalartist,songcopyright,songurl,songencodedby


    ReDim mColumnInfo(1 To 12)
    mColumnInfo(1).ColumnTitle = "Title"
    mColumnInfo(1).ColumnTooltip = "Title of the Song"
    mColumnInfo(1).ColumnDefaultWidth = 50
    
    mColumnInfo(2).ColumnTitle = "Artist"
    mColumnInfo(2).ColumnTooltip = "Song Artist"
    mColumnInfo(2).ColumnDefaultWidth = 50
    
    mColumnInfo(3).ColumnTitle = "Album"
    mColumnInfo(3).ColumnTooltip = "Song Album"
    mColumnInfo(3).ColumnDefaultWidth = 50
    
    mColumnInfo(4).ColumnTitle = "Year"
    mColumnInfo(4).ColumnTooltip = "Song Year"
    mColumnInfo(4).ColumnDefaultWidth = 50
    
    mColumnInfo(5).ColumnTitle = "Comment"
    mColumnInfo(5).ColumnTooltip = "Comment"
    mColumnInfo(5).ColumnDefaultWidth = 50
    
    mColumnInfo(6).ColumnTitle = "Track#"
    mColumnInfo(6).ColumnTooltip = "Track Number"
    mColumnInfo(6).ColumnDefaultWidth = 50
    
    mColumnInfo(7).ColumnTitle = "Genre"
    mColumnInfo(7).ColumnTooltip = "Genre"
    mColumnInfo(7).ColumnDefaultWidth = 50
    
End Sub


Private Function IColumnPlugin_CanConfig() As Boolean
IColumnPlugin_CanConfig = False
End Function

Private Sub IColumnPlugin_ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
Dim castcol As cColumn
    Set castcol = ClickedColumn
    castcol.SortType = eLVSortStringNoCase
    CDebug.Post "MP3 column clicked:" & ColIndex
    
End Sub

Private Sub IColumnPlugin_Configure()
'
End Sub

Private Property Get IColumnPlugin_Description() As String
'
IColumnPlugin_Description = "display columns relating to MP3 ID3 tags"
End Property

'MP3 plugin...

Private Function IColumnPlugin_GetColumnCount() As Long
    IColumnPlugin_GetColumnCount = 7
End Function

'since it can be expensive to retrieve the MP3 ID tags, we only do it once for each listitem, via static variables.

Private Function IColumnPlugin_GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
'
Dim coldata As ColumnData
Static mLastFile As String

Static songtitle As String, songartist As String, songalbum As String
Static songyear As Long, songcomment As String, songtracknumber As Long, songgenre As String, songcomposer As String, songoriginalartist As String, songcopyright As String, songurl As String, songencodedby As String

If mLastFile = mlstitem.Tag Then
    'use the cached values....
Else
    coldata.ColumnData = ""
    If modMP3.MP3read_HasTag_ID3v2(mlstitem.Tag) Then

        modMP3.MP3read_GetTagInfo_ID3v2 mlstitem.Tag, songtitle, songartist, songalbum, songyear, songcomment, songtracknumber, songgenre, songcomposer, songoriginalartist, songcopyright, songurl, songencodedby
    End If
End If
If ColumnIndex = 1 Then
    'colData.ColumnData = IIf(StrComp(bcfile.GetExtension(mlstitem.Tag), "MP3", vbTextCompare) = 0, "Yes", "No")
    coldata.ColumnData = songtitle
ElseIf ColumnIndex = 2 Then
    coldata.ColumnData = songartist
    'colData.ColumnData = IIf(StrComp(Left$(mlstitem.Tag, 1), "P", vbTextCompare) = 0, "Yes", "No")
ElseIf ColumnIndex = 3 Then
    coldata.ColumnData = songalbum
    
ElseIf ColumnIndex = 4 Then
    coldata.ColumnData = songyear
ElseIf ColumnIndex = 5 Then
'comment
coldata.ColumnData = songcomment
ElseIf ColumnIndex = 6 Then
'track
coldata.ColumnData = songtracknumber
ElseIf ColumnIndex = 7 Then
'genre
coldata.ColumnData = songgenre
End If
IColumnPlugin_GetColumnData = coldata
mLastFile = mlstitem.Tag
End Function

Private Function IColumnPlugin_GetColumnInfo(ColumnIndex As Long) As ColumnInfo
    IColumnPlugin_GetColumnInfo = mColumnInfo(ColumnIndex)
End Function

Private Property Get IColumnPlugin_Name() As String
'
IColumnPlugin_Name = "MP3 ID3 Columns"
End Property

Private Function IColumnPlugin_PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function
