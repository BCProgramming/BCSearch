VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBCSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'BCSplitter: basic SplitterBar Class.
Private Declare Function SetCapture Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32.dll" () As Long



Private WithEvents mSplitterPic As PictureBox 'Object used as Pivot.
Attribute mSplitterPic.VB_VarHelpID = -1

Private mUpperControl As Control, mLowerControl As Control


Public Sub Init(ObjPivot As PictureBox, UpperControl As Control, LowerControl As Control)
    Set mSplitterPic = ObjPivot
    
    If Not ObjPivot.Container Is ObjectA.Container And ObjPivot.Container Is ObjectB.Container Then
        Err.Raise 13, "cannot initialize splitter with objects on different containers."
    End If
    Set mUpperControl = UpperControl
    Set mLowerControl = LowerControl
    mSplitterPic.MousePointer = vbSizeNS
End Sub

Private Sub mSplitterPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
SetCapture mSplitterPic.hWnd
End Sub

Private Sub mSplitterPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "mousemove:" & x, y
End Sub

Private Sub mSplitterPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
End Sub
