VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IColumnPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ColumnPlugin... my own "duplication" of the IColumnProvider Interface provided by the shell.

'Which, for one reason or another refuses to give me the first character of the column or data...

'In either case, I can always implement a IColumnPlugin that uses IColumnProvider, so no big loss there.
Public Type ColumnInfo
    ColumnTitle As String 'Title to be shown in the columnheaders.
    ColumnDefaultWidth As Long 'default width, in pixels.
    ColumnTooltip As String 'extra info... what does it show?
    
End Type
Public Type ColumnData
    ColumnData As String
    ColumnIcon As Long 'if this is non-zero it is used as a icon handle.
End Type
Public Property Get Name() As String
    'return a somewhat descriptive name for the provider.
    
End Property
Public Property Get Description() As String
    'return a description.
End Property
Public Function GetColumnCount() As Long
'return the number of Columns this plugin implements.

End Function
Public Function GetColumnInfo(ColumnIndex As Long) As ColumnInfo
    'should be valid from 1 to "GetColumnCount".
End Function
Public Function GetColumnData(mlstitem As Object, ByVal ColumnIndex As Long) As ColumnData
'retrieve column data, passes in the listItem currently being added.
End Function
Public Function PluginColumnCompare(lstItemA As Object, lstItemB As Object) As Long
'
End Function

Public Sub ColumnClick(LvwClicked As Object, ClickedColumn As Object, ByVal ColIndex As Long)
    'fired when user clicks on the column this provider handles.
End Sub
Public Function CanConfig() As Boolean
    'returns wether this columnplugin can be configured.
End Function
Public Sub Configure()
    'show configuration for this column plugin, if applicable.
End Sub
