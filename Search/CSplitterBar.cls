VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' cSplitter
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

' If we used WithEvents on these items, then we not need to
' write any code at all in the form being split. However, this
' means you have a a WithEvents reference to the form here, whilst
' the form also has a WithEvents reference to the splitter.
' This seems to cause immediate crash (VB dev environment disappears!)
' when trying to terminate the splitter class under VB5 (SP2 and above)
Private WithEvents m_picSplitter As PictureBox
Private WithEvents m_frmParent As Form

Private m_bSplitting As Boolean
Private m_lSplitOffset As Long
Private m_lBorder As Long
Private m_eOrientation As ESPLTOrientationConstants

Public Enum ESPLTOrientationConstants
    cSPLTOrientationHorizontal = 1
    cSPLTOrientationVertical = 2
End Enum

Public Event DoSplit(bSplit As Boolean)
Public Event SplitComplete()

Property Let Orientation(eOrientation As ESPLTOrientationConstants)
    m_eOrientation = eOrientation
    If Not (m_picSplitter Is Nothing) Then
        If (eOrientation = cSPLTOrientationHorizontal) Then
            m_picSplitter.MousePointer = vbSizeNS
        Else
            m_picSplitter.MousePointer = vbSizeWE
        End If
    End If
End Property
Property Get Orientation() As ESPLTOrientationConstants
    Orientation = m_eOrientation
End Property
Property Let BorderSize(lSize As Long)
    m_lBorder = lSize
End Property
Property Get BorderSize() As Long
    BorderSize = m_lBorder
End Property
Public Sub Initialise( _
        ByRef picSplitter As PictureBox, _
        ByRef frmParent As Form _
    )
    Set m_picSplitter = picSplitter
    Set m_frmParent = frmParent
    With m_picSplitter
        .BorderStyle = 0
        .ZOrder 1
        .MousePointer = vbSizeWE
        .Visible = True
    End With
End Sub
Public Sub MouseDown( _
        ByVal Pos As Single _
    )
Dim bSplit As Boolean
    bSplit = True
    RaiseEvent DoSplit(bSplit)
    If Not (bSplit) Then Exit Sub
   
    m_bSplitting = True
    m_lSplitOffset = Pos
    With m_picSplitter
        .BackColor = &H80000010
        .ZOrder 0
        .BorderStyle = 1
        .Width = 4 * Screen.TwipsPerPixelX
    End With
    SetCapture m_frmParent.hwnd
End Sub
Public Sub MouseMove( _
        ByVal Pos As Single _
    )
    If (m_bSplitting) Then
        If (m_eOrientation = cSPLTOrientationHorizontal) Then
            ' Horizontal orientation:
            If (Pos < m_frmParent.ScaleHeight - m_lBorder) And (Pos > m_lBorder) Then
                Screen.MousePointer = vbSizeNS
                m_picSplitter.Move m_picSplitter.Left, Pos
            Else
                Screen.MousePointer = vbNoDrop
            End If
        Else
            ' Vertical orientation:
            If (Pos < m_frmParent.ScaleWidth - m_lBorder) And (Pos > m_lBorder) Then
                Screen.MousePointer = vbSizeWE
                m_picSplitter.Move Pos
            Else
                Screen.MousePointer = vbNoDrop
            End If
        End If
    End If
End Sub
Public Function MouseUp( _
        ByRef Pos As Single _
    ) As Boolean
Dim lRealPos As Long

    If (m_bSplitting) Then
        ' End the moving:
        ReleaseCapture
        With m_picSplitter
            .BackColor = &H8000000F
            .BorderStyle = 0
       
            ' Move to a position within bounds
            ' if we are out of bounds:
            If (Pos < m_lBorder) Then
                Pos = m_lBorder
            End If
            If (m_eOrientation = cSPLTOrientationHorizontal) Then
                If (Pos > (m_frmParent.ScaleHeight - m_lBorder)) Then
                    Pos = m_frmParent.ScaleHeight - m_lBorder
                End If
            Else
                If (Pos > (m_frmParent.ScaleWidth - m_lBorder)) Then
                    Pos = m_frmParent.ScaleWidth - m_lBorder
                End If
            End If
           
            ' Now drop the splitter:
            Pos = Pos - m_lSplitOffset
            If (m_eOrientation = cSPLTOrientationHorizontal) Then
                .Move .Left, Pos
            Else
                .Move Pos
            End If
            .ZOrder 1
           
        End With
       
        m_bSplitting = False
        Screen.MousePointer = vbNormal
       
        MouseUp = True
       
        RaiseEvent SplitComplete
    End If
End Function

Private Sub Class_Initialize()
    m_eOrientation = cSPLTOrientationVertical
End Sub


Private Sub Class_Terminate()
    m_bSplitting = False
    Set m_picSplitter = Nothing
    Set m_frmParent = Nothing
End Sub


Private Sub m_frmParent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MouseDown Y
End Sub

Private Sub m_frmParent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MouseMove Y
End Sub

Private Sub m_frmParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MouseUp Y
End Sub

Private Sub m_picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MouseDown Y
End Sub

Private Sub m_picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MouseUp Y
End Sub
