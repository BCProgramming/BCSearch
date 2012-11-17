VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.UserControl UCFontPreview 
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4815
   ScaleWidth      =   5820
   Begin RichTextLib.RichTextBox rtbpreview 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7646
      _Version        =   393217
      ReadOnly        =   -1  'True
      RightMargin     =   32768
      OLEDropMode     =   1
      TextRTF         =   $"UCFontPreview.ctx":0000
   End
End
Attribute VB_Name = "UCFontPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarfontFile As String
Private mPreviewObject As CFontPreview
'C:\windows\fonts\AGENCYB.TTF

Public Property Let FontFile(ByVal Vdata As String)
    mvarfontFile = Vdata
    ShowPreview Vdata
End Property
Public Property Get FontFile() As String
    FontFile = mvarfontFile
End Property
Public Sub ShowPreviewForFont(ByVal FontFile As String)
    ShowPreview FontFile
End Sub

Private Sub DoCreate()
  Dim sfont As StdFont, I As Long
  Dim PreviewSizes As Variant
  If mPreviewObject Is Nothing Then Exit Sub
     PreviewSizes = Array(8, 10, 12, 14, 18, 24, 28)
   Const PreviewText As String = "The quick brown fox jumps over the lazy dog."
    Const PreviewCharacters As String = "1234567890!@#$%^&*()"
  
    rtbpreview.Text = ""
    Set sfont = New StdFont
    'Set rtbpreview.Font = sfont
    With mPreviewObject
        sfont.Name = .FaceName
        sfont.Bold = .Bold
        sfont.Italic = .Italic
    End With
    'PicFontPreview.CurrentX = 10
    'PicFontPreview.CurrentY = 10
    For I = 0 To UBound(PreviewSizes)
        sfont.Size = PreviewSizes(I)
        'PicFontPreview.Print PreviewText
    
        AppendText PreviewText & "(" & PreviewSizes(I) & "pt.)" & vbCrLf, vbBlack, sfont.Name, PreviewSizes(I), sfont.Bold, sfont.Italic
    
    Next I
    
    AppendText vbCrLf, vbBlack, "MS Shell Dlg", UserControl.Font.Size, UserControl.Font.Bold, UserControl.Font.Italic
    
    With mPreviewObject
    
    AppendText "Filename:  " & vbTab & """" & .FontFile & """" & vbCrLf & _
    "Face Name:" & vbTab & """" & .FaceName & """" & vbCrLf & _
        "Family Name:" & vbTab & """" & .FamilyName & """" & vbCrLf & _
        "Full Name:" & vbTab & """" & .FullName & """" & vbCrLf & _
        "Sub Family Name:" & vbTab & """" & .SubFamilyName & """" & vbCrLf & _
        "Trademark:   " & vbTab & """" & .Trademark & """" & vbCrLf & _
        "Copyright:   " & vbTab & """" & .Copyright & """" & vbCrLf & _
        "Version:     " & vbTab & """" & .VersionString & """" & vbCrLf, vbBlack, "MS Shell Dlg", UserControl.Font.Size, False, False
    End With
End Sub
Private Sub AppendText(ByVal Text As String, ByVal FontColour As Long, ByVal FontName As String, ByVal fontSize As Double, ByVal FontBold As Boolean, ByVal FontItalic As Boolean)

        rtbpreview.SelStart = Len(rtbpreview.Text)
        
        rtbpreview.SelFontName = FontName
        rtbpreview.SelBold = FontBold
        rtbpreview.SelFontSize = fontSize
        rtbpreview.SelItalic = FontItalic
        rtbpreview.SelColor = FontColour
        rtbpreview.SelText = Text
    

End Sub

Private Sub ShowPreview(Optional ByVal FontFile As String)
    
   
    If FontFile = "" Then FontFile = mvarfontFile
    
    On Error GoTo ReportError
    
 
    Set mPreviewObject = New CFontPreview
    mPreviewObject.FontFile = FontFile
'    DoDraw
DoCreate
ReportError:
End Sub


Private Sub PicFontPreview_Resize()
'DoDraw
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
Debug.Print "completedrag"
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print "dragdrop"
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Debug.Print "dragover"
End Sub

Private Sub UserControl_Resize()
    rtbpreview.Move 0, 0, ScaleWidth, ScaleHeight
End Sub


