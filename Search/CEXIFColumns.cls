VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEXIFColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'EXIF columns provider...
Implements IColumnPlugin




Private mEXIFObject As CExifData
Private mLastFile As String
Private Sub Class_Initialize()
    Set mEXIFObject = New CExifData
End Sub

Private Function IColumnPlugin_CanConfig() As Boolean
IColumnPlugin_CanConfig = False
End Function

Private Sub IColumnPlugin_ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
'
End Sub

Private Sub IColumnPlugin_Configure()
'
End Sub

Private Property Get IColumnPlugin_Description() As String
IColumnPlugin_Description = "Display columns for EXIF data"
End Property

Private Function IColumnPlugin_GetColumnCount() As Long
'



End Function

Private Function IColumnPlugin_GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
    '
    If StrComp(mLastFile, mlstitem.Tag, vbTextCompare) = 0 Then
    'its the same as before...
    
    
    End If


End Function

Private Function IColumnPlugin_GetColumnInfo(ColumnIndex As Long) As ColumnInfo
'
End Function

Private Property Get IColumnPlugin_Name() As String
IColumnPlugin_Name = "EXIF columns plugin"
End Property

Private Function IColumnPlugin_PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function
