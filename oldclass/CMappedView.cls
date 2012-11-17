VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMappedView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CMappedView: Used to map Views of a File mapping.


Private mMappedAddress As Long
Private mMappedFromMapping As CFileMapping
Friend Sub InitToAddress(ByVal AddrInit As Long, FromMapping As CFileMapping)
    'makes this class responsible for deallocating the view.
    mMappedAddress = AddrInit
    Set mMappedFromMapping = FromMapping
End Sub

Private Sub Class_Terminate()

    UnmapViewOfFile ByVal mMappedAddress

End Sub
