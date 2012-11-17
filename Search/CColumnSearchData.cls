VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnSearchData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CColumnSearchData: stores data about a column search. There is a collection of these stored in CExtraFilterData...


Private mPluginProgID As String
Private mColumnIndex As Long

Private mPluginRef As IColumnPlugin

'and the search data....

'since we cannot even guess at the format of the column, we'll just do the basic string type searches.
Private mSearchSpec As String
Private mSearchIsRegExp As Boolean

Public Function DoFilter(ByVal FileObj As CFile) As Boolean
    'returns true if it passes the test, false otherwise.
    
    If mSearchSpec = "" Then
        DoFilter = True
        Exit Function
    End If
    Dim ColData As ColumnData
    'coldata= mpluginref.GetColumnData(



End Function

Public Sub Reinitialize()
  Dim mTemp As Object
   Set mPluginRef = Nothing
    On Error Resume Next
    Set mTemp = CreateObject(mPluginProgID)
    If Err.Number <> 0 Then
        Err.Raise 9, "CColumnSearchData::PluginProgID", "CreateObject() On """ & mPluginProgID & """ Failed.(" & Err.Description & ")"""
        
    End If
    'Now, try casting to IColumnPlugin....
    Dim CastPlug As IColumnPlugin
    Set CastPlug = mTemp
    If Err <> 0 Then
        Err.Raise 9, "CColumnSearchData::PluginProgID", "Failed to cast Create object """ & mPluginProgID & """ to IColumnPlugin interface.(" & Err.Description & ")"
    End If
'otherwise, continue. make sure the columnindex is within range.
    If mColumnIndex > CastPlug.GetColumnCount Then
        Debug.Print "mcolumnIndex was out of range (" & mColumnIndex & " > " & CastPlug.GetColumnCount & ")"
    
    End If
    Set mPluginRef = CastPlug

End Sub
Public Property Get PluginprogID() As String
    PluginprogID = mPluginProgID
    
End Property
Public Property Let PluginprogID(ByVal Vdata As String)
    mPluginProgID = Vdata
   
    
  
    
    
    
    
End Property
Public Property Get ColumnIndex() As Long
    ColumnIndex = mColumnIndex
End Property
Public Property Let ColumnIndex(ByVal Vdata As Long)
    mColumnIndex = Vdata
End Property
Public Property Get SearchSpec()
    SearchSpec = mSearchSpec
End Property
Public Property Let SearchSpec(ByVal Vdata As String)
    mSearchSpec = Vdata
End Property

Public Property Let SearchSpecIsRegExp(ByVal Vdata As Boolean)
    mSearchIsRegExp = Vdata
End Property
Public Property Get SearchSpecIsRegExp() As Boolean
    SearchSpecIsRegExp = mSearchIsRegExp
End Property
'Public Sub Init(ByVal PluginprogID As String, ByVal ColumnIndex As Long)




'End Sub


