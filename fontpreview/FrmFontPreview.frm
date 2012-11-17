VERSION 5.00
Begin VB.Form FrmFontPreview 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4365
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin Project1.UCFontPreview fontpreview 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7435
   End
End
Attribute VB_Name = "FrmFontPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

 Dim strFilename As String
    ' Get the filename from the Files collection
    strFilename = Data.Files(1)
    ' Load the picture
    fontpreview.ShowPreviewForFont strFilename

End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
  If Not Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectNone    ' Don't allow dropping
    End If

End Sub
