VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumVARIANT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This is the enumerator class itself. When VB initialized a For Loop on a
' class using this enumerator, an instance of this class will be created using
' the NewEnum() function of that class.
'
' The regular interface definition of Next() is not supported by Visual Basic.
' It was redefined in a custom typedef (VBFCustomCollection.IEnumVARIANT) but
' there is still a major limitation. All COM methods must return an HRESULT
' value of S_OK or S_FALSE to determine success or failure. With subs, VB
' handles this internally, which means that the Next() implementation cannot
' possibly return the correct value. To fix this, we have to remap the Vtable
' of the interface so that all calls to Next() are redirected to another
' function in a module in which we don't have to follow the strict interface
' implementation rules - we can return a value.

Option Explicit

'Implements VBFCustomCollection.IEnumVARIANT

Private m_OldNextPtr As Long    ' Used to store the pointer to the default Next() implementation
'Private m_CurrentPos As Long    ' Stores the current position in the enumeration
'Private m_CollectionPtr As Long ' stores a pointer to the actual IEnumVariant Implementation.
                                ' in order to create a weak reference that does not increment the
                                ' ref-count of the collection object.
Private m_Collection As IVariantEnumerable
'Private Property Get m_collectionPtr() As Long
'    m_collectionPtr = VarPtr(m_Collection)
'End Property
Public Sub Init(ByVal collection As IVariantEnumerable)
    Set m_Collection = collection
End Sub
Public Sub Term()
    Set m_Collection = Nothing
End Sub
Private Sub Class_Initialize()
    ' We need to put the current object reference into an interface
    ' reference so the ObjPtr() function returns the Vtable reference.
    Dim tmp As VBFCustomCollection.IEnumVARIANT
    Set tmp = Me

    ' Vtable for IEnumVARIANT:
    '   0 = QueryReference
    '   1 = AddRef
    '   2 = Release
    '   3 = Next
    '   4 = Skip
    '   5 = Reset
    '   6 = Clone
    m_OldNextPtr = VtableSwap(ObjPtr(tmp), 3, AddressOf NEW_IEnumVariant_Next)
    Set tmp = Nothing
End Sub

Private Sub Class_Terminate()
    Dim tmp As VBFCustomCollection.IEnumVARIANT
    Set tmp = Me

    VtableSwap ObjPtr(tmp), 3, m_OldNextPtr
    Set tmp = Nothing
    Set m_Collection = Nothing
End Sub

Private Sub IEnumVARIANT_Clone(ppenum As VBFCustomCollection.IEnumVARIANT)
    ' No implementation because VB will never use this and it would be
    ' more hassle than it's worth to implement it.
    Dim col As IVariantEnumerable
    'RtlMoveMemory ByVal VarPtr(col), m_collectionPtr, 4
    Set col = m_Collection
    
    Set ppenum = col.Clone()
    
End Sub

Private Sub IEnumVARIANT_Next(ByVal celt As Long, rgvar As Variant, ByVal pceltFetched As Long)
    ' This implementation will never be called directly, instead being mapped
    ' to a secondary function in a module to get around VB limitations.
End Sub

Private Sub IEnumVARIANT_Reset()
'    m_CurrentPos = 0
        Dim col As IVariantEnumerable
    
    ' Get a reference to the collection.
     'RtlMoveMemory ByVal VarPtr(col), m_collectionPtr, 4
    Set col = m_Collection
    'ub = col.VariantUBound
    
    col.Reset
    ' Free the reference to the collection.
    'RtlZeroMemory ByVal VarPtr(col), 4
End Sub

Private Sub IEnumVARIANT_Skip(ByVal celt As Long)
    ' In theory we should probably delegate this because it is supposed to
    ' return a value, but since VB never actually calls this to the best
    ' of my knowledge, it shouldn't make much of a difference.
    
    
    Dim col As IVariantEnumerable
    
    ' Get a reference to the collection.
    'RtlMoveMemory ByVal VarPtr(col), m_collectionPtr, 4
    Set col = m_Collection
    
    'ub = col.VariantUBound
    
    col.Skip celt
    ' Free the reference to the collection.
    'RtlZeroMemory ByVal VarPtr(col), 4
End Sub

' This next function is not part of the original IEnumVARIANT
' interface, but this allows the remapped Next function to directly
' get the next item in the collection.
Private Sub IEnumVARIANT_VBNext(rgvar As Variant, retVal As Long)
    Dim nextItem As Variant
    Dim col As IVariantEnumerable
    
    ' Get a reference to the collection.
    'CRASH HERE:!!!
     'RtlMoveMemory ByVal VarPtr(col), m_collectionPtr, 4
    Set col = m_Collection
    
'    If col.VariantGetItem(m_CurrentPos, nextItem) Then
'        If VariantIsObject(nextItem) Then
'            Set rgvar = nextItem
'        Else
'            rgvar = nextItem
'        End If
'
'        m_CurrentPos = m_CurrentPos + 1
'
'        retval = True
'    Else
'        retval = False
'    End If

    retVal = col.GetNext(rgvar)
    ' Free the reference to the collection.
    'RtlZeroMemory ByVal VarPtr(col), 4
End Sub
