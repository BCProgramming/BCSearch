VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContextCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IContextCallback: implemented by classes that which to change/ modify the Explorer menu as implemented in MdlContextMenu.



Public Sub BeforeShowMenu(ByVal hMenu As Long, Optional ByRef CancelShow As Boolean)
    'fired before the menu is shown.
End Sub
Public Sub AfterShowMenu(ByRef idChosen As Long)
    'fired After the menu is closed; Change to 0 if handled.
End Sub
