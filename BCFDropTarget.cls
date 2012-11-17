VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BCFDropTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDropTarget

Private Sub Class_Initialize()

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObj As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As olelib.DROPEFFECTS)

End Sub

Private Sub IDropTarget_DragLeave()
'
End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As olelib.DROPEFFECTS)
'
End Sub

Private Sub IDropTarget_Drop(ByVal pDataObj As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptX As Long, ByVal ptY As Long, pdwEffect As olelib.DROPEFFECTS)
'
End Sub
