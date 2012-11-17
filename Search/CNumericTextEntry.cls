VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNumericTextEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const Numchars = "1234567890"
Private WithEvents HandleTBox As TextBox
Attribute HandleTBox.VB_VarHelpID = -1




Public Sub Init(Tbox As TextBox)

    Set HandleTBox = Tbox
    
End Sub



Private Sub HandleTBox_KeyPress(KeyAscii As Integer)
    If Not InStr(Numchars, Chr$(KeyAscii)) > 0 Then
        VBA.Beep
        KeyAscii = 0
    
    
    End If
End Sub
