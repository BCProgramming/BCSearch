VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SearchFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mArray() As CSearchFilter

Private mCount As Long
Private Sub Class_Initialize()
    mCount = -1
    
    MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount + 1
End Sub
Public Function Add(Vdata As CSearchFilter) As Long
    mCount = mCount + 1
    ReDim Preserve mArray(mCount)
    Set mArray(mCount) = Vdata
    Add = mCount
End Function
Public Function count() As Long
    count = mCount
End Function
Public Function Item(ByVal Index As Long) As CSearchFilter
    Set Item = mArray(Index)
End Function
Public Sub Clear()
    Erase mArray
    mCount = -1
End Sub
Public Function GetIndex(OfObject As Variant) As Long
    Dim I As Long
    For I = 0 To mCount
        If IsObject(OfObject) Then
            If TypeOf OfObject Is CSearchFilter Then
                If mArray(I) Is OfObject Then
                    GetIndex = I
                    Exit Function
                End If
            End If
        ElseIf varType(OfObject) = vbString Then
            If StrComp(mArray(I).Name, OfObject, vbTextCompare) = 0 Then
                GetIndex = I
                Exit Function
            End If
        ElseIf IsNumeric(OfObject) Then
            GetIndex = OfObject
        
        End If
        
    Next I
End Function
Public Function Remove(vntIndexkey As Variant) As CSearchFilter
    Dim indRemove As Long, I As Long
    indRemove = GetIndex(vntIndexkey)
    If indRemove > mCount Then Exit Function
    
    
    'to remove:
    
    'copy all items down one start with the one after this position- then redimension the array.
    
    For I = indRemove To mCount - 1
        Set mArray(I) = mArray(I + 1)
    
    Next I
    ReDim Preserve mArray(mCount - 1)
    mCount = mCount - 1


End Function

Private Sub Class_InitProperties()
'
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
'
Dim CountGet As Long

Dim I As Long

With PropBag
        '.WriteProperty "Count", UBound(mArray) - 1
        
        CountGet = .ReadProperty("Count", 0)
        
        ReDim mArray(0 To CountGet)
        
        For I = LBound(mArray) To CountGet - 1
        '.WriteProperty "Filter(" & Trim$(str$(i)) & ")", mArray(i)
        Set mArray(I) = .ReadProperty("Filter(" & Trim$(str$(I)) & ")")
        
        
        
        Next I
        ReDim Preserve mArray(0 To CountGet - 1)
        mCount = CountGet
End With

End Sub

Private Sub Class_Terminate()
MdlFileSystem.TotalObjectCount = MdlFileSystem.TotalObjectCount - 1
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
'
    Dim I As Long
    With PropBag
        .WriteProperty "Count", UBound(mArray) + 1
        For I = LBound(mArray) To UBound(mArray)
        .WriteProperty "Filter(" & Trim$(str$(I)) & ")", mArray(I)
        
        
        
        
        Next I
    End With


End Sub
