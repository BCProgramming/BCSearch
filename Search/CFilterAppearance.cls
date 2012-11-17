VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExtraFilterData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Encapsulates the data for the appearance of the results of a specific filter, as well as extra
'data related to the filter, namely, for example, Searching by Column entries.

'also the copy to search parameters.

Private mFont As StdFont
Private mForeColor As Long, mBackColor As Long
Private mTag As Variant
Private mREmatchCol As MatchCollection

Public Property Set REMatchCol(Vdata As MatchCollection)
    Set mREmatchCol = Vdata
End Property
Public Property Get REMatchCol() As MatchCollection
    Set REMatchCol = mREmatchCol
End Property
Public Property Let Tag(ByVal Vdata As Variant)
    mTag = Vdata
End Property
Public Property Set Tag(Vdata As Variant)
    Set mTag = Vdata
End Property
Public Property Get Tag() As Variant
    If IsObject(mTag) Then Set Tag = mTag Else Let Tag = mTag
End Property
Public Property Set Font(Vdata As IFontDisp)
    Set mFont = Vdata
End Property
Public Property Get Font() As IFontDisp
    Set Font = mFont
End Property


Public Property Let ForeColor(ByVal Vdata As Long)
    mForeColor = Vdata
End Property
Public Property Get ForeColor() As Long
    ForeColor = mForeColor
End Property
Public Property Let BackColor(ByVal Vdata As Long)
    mBackColor = Vdata
End Property
Public Property Get BackColor() As Long
    BackColor = mBackColor
End Property

Private Sub Class_InitProperties()
    Set mFont = New StdFont
    mFont.Name = GetDefaultUIFont
    mForeColor = vbWindowText
    mBackColor = vbWindowBackground
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
'
With PropBag
    Set mFont = .ReadProperty("Font", New StdFont)
    mForeColor = .ReadProperty("ForeColor", vbWindowText)
    mBackColor = .ReadProperty("BackColor", vbWindowBackground)
End With

End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
'
With PropBag
    .WriteProperty "Font", mFont
    .WriteProperty "ForeColor", mForeColor
    .WriteProperty "BackColor", mBackColor


End With


End Sub
