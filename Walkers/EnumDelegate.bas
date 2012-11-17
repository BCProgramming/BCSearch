Attribute VB_Name = "EnumDelegate"
Option Explicit
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
' The ObjPtr() of an interface returns a pointer to its Vtable,
' and all IUnknown based interfaces have three entries by default:
' QueryInterface, AddRef, and Release. After that comes the first public
' method of the interface.
Public Function VtableSwap(ByVal objectPtr As Long, ByVal index As Long, ByVal functionPtr As Long) As Long
    Dim offset As Long
    Dim vtablePtr As Long
    Dim tmp As Long
    
    ' Find the pointer to the VTable itself and then calculate the
    ' position of the entry we want to modify.
    RtlMoveMemory vtablePtr, ByVal objectPtr, 4
    offset = vtablePtr + (index * 4)
    
    ' Copy the current function pointer into the return value.
    RtlMoveMemory VtableSwap, ByVal offset, 4
    
    ' The Vtable itself has to be unprotected temporarily so we can
    ' modify the function pointer.
    VirtualProtect offset, 4, PAGE_EXECUTE_READWRITE, tmp
    RtlMoveMemory ByVal offset, functionPtr, 4
    VirtualProtect offset, 4, tmp, tmp
End Function

' This is the remapped function that is called instead of the actual
' Next() implementation of the IEnumVARIANT interfaces. The first
' parameter is a reference to the calling object (VB translates it
' as Me in an actual class implementation).
Public Function NEW_IEnumVariant_Next(ByVal objectPtr As Long, ByVal celt As Long, rgvar As Variant, ByVal pceltFetched As Long) As Long
    On Error GoTo out

    Dim I As Long, fetched As Long
    Dim arr() As Variant
    Dim retVal As Long
    'Dim obj As VBFCustomCollection.IEnumVARIANT
    
    ' Move the pointer into a weak reference.
    RtlMoveMemory ByVal VarPtr(obj), objectPtr, 4
    
    ReDim arr(celt)
    
    For I = 0 To celt - 1
        ' Get the next item in the collection; it will return False if there
        ' are no more or True if it did return one.
        obj.VBNext arr(I), retVal
        If Not retVal Then
            Exit For
        End If

        fetched = fetched + 1
    Next
    
    ' Kill the Variant so the weak reference doesn't try to call Release
    ' and make the enumerator terminate prematurely.
    RtlZeroMemory ByVal VarPtr(obj), 4
    
    ' pceltFetched is a pointer to a long, so we don't want to copy anything
    ' into it if it is a null reference (obviously). Otherwise, we want to
    ' put the number of successfully fetched items into it.
    If Not pceltFetched = 0 Then
        RtlMoveMemory ByVal pceltFetched, fetched, 4
    End If
    
    ' Copy the temporary array into the rgvar out-parameter and erase the
    ' temporary array.
    RtlMoveMemory ByVal VarPtr(rgvar), arr(0), 16 * fetched
    RtlZeroMemory arr(0), 16 * fetched
    
    ' If it actually fetched less than we requested (implying we reached
    ' the end of the collection), set the return value accordingly.
    If fetched < celt Then
        NEW_IEnumVariant_Next = S_FALSE
    Else
        NEW_IEnumVariant_Next = S_OK
    End If
    
    Exit Function
out:
    ' Just in case it crashes for some reason, we should exit gracefully
    ' instead of killing the IDE.
    If Not pceltFetched = 0 Then
        RtlMoveMemory ByVal pceltFetched, ByVal 0, 4
    End If
    NEW_IEnumVariant_Next = S_FALSE
End Function
