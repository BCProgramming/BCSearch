VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVariantEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This interface is not strictly necessary, although I added it so that
' the enumerator will work on any class that implements this interface.
' It exposes two functions that the enumerator needs - one to get the
' upper bound of the collection storage object (usually an array but
' it could be anything in theory, like the number of files in a directory),
' and one to get the Variant representation of any object in the collection
' (again, usually from an array but it could be used to grab a particular
' file from a directory for example if the enumerator was enumerating a
' filesystem).

Option Explicit

Public Function GetNext(GetVar As Variant) As Long
    'should return next item...
End Function
Public Function Clone() As stdole.IEnumVARIANT
    'duplicate the state of this iterator into a new instance.
    
End Function
Public Sub Reset()
    'reset to start of iterated data.
End Sub
Public Sub Skip(ByVal celt As Long)
    'skip items...
End Sub
