VERSION 5.00
Begin VB.UserControl jcFrames 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   HitBehavior     =   0  'None
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "jcFrames.ctx":0000
End
Attribute VB_Name = "jcFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'============================================================================================
'   jcFrames v 1.0 Copyright © 2005.All rights reserved.
'   Juan Carlos San Román Arias (sanroman2004@yahoo.com)
'
'   You may use this control in your applications free of charge,
'   provided that you do not redistribute this source code without
'   giving me credit for my work.  Of course, credit in your
'   applications is always welcome.
'
'   Thanks to Jim K for doing the initial idea of the usercontrol using
'   my job posted in PSC
'
'   Thanks to ElectroZ for his frame style used here as TextBox style
'============================================================================================
'
'   Modifications: Paul R. Territo, Ph.D
'
'   The following code is based on the above authors submission which
'   can be found at the follow URL:
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63827&lngWId=1
'
'   29Dec05 - Moved all external API drawing and Type structures into UserControl
'       eliminate the need for external dependancies (i.e. OCX). This provides
'       a single drop in place UserControl which follows the general rules of
'       encapsulation (i.e. self-contained).
'
'============================================================================================
'  -----------------------------
'  Version 1.1.0 - 29 Dec. 2005
'  -----------------------------
'   Thanks to Paul R. Territo, Ph.D for your advices and usercontrol modification.
'   - usercontrol includes now API drawing and type declaration (no more mods in usercontrol)
'   - Added icon alignment (left and right)
'   - caption alignment takes into consideration if icon picture exists and its alignment
'============================================================================================
'  -----------------------------
'  Version 1.2.0 - 04 Jan. 2006
'  -----------------------------
'   - Added different header styles for Windows frame style (txtboxcolor and gradient)
'   - Added different gradient styles for header gardient style for Windows frame style
'     (horizontal, vertical and cilinder)
'   - Caption is trimmed when its width exceeds control width
'============================================================================================
'  -----------------------------
'  Version 2.0.0 - 11 Jan. 2006
'  -----------------------------
'   - 4 new styles have been added: Inner widge, Outer widge, Header and Panel
'   - Header styles have been extended for other frame style (messenger, jcGradient
'     textbox and panel style)
'   - Control structure was reorganized
'   - Gradientframe style was renamed as jcGradient
'   - Added TxtBoxShadow property for textbox style
'   - Added multiline caption for Panel style
'============================================================================================
'  ----------------------------
'  Version 2.0.1 - 8 Feb. 2006
'  ----------------------------
'   - Added enabled property (it enables or disables all the controls in usercontrol)
'   - Added TransBlt from Chameleon button to draw grayscale image when control is disabled
'============================================================================================

Option Explicit

'*************************************************************
'   Required Type Definitions
'*************************************************************

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Private Type POINT
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum jcStyleConst
    XPDefault = 0
    jcGradient = 1
    TextBox = 2
    Windows = 3
    Messenger = 4
    InnerWedge = 5
    OuterWedge = 6
    Header = 7
    Panel = 8
End Enum

'xp theme
Public Enum jcThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Visual2005 = 3
    Norton2004 = 4
    Custom = 5
End Enum

'gradient type
Public Enum jcGradConst
    VerticalGradient = 0
    HorizontalGradient = 1
    VCilinderGradient = 2
    HCilinderGradient = 3
End Enum

'header style
Public Enum jcHeaderConst
    TxtBoxColor = 0
    Gradient = 1
End Enum

'TxtBox style
Public Enum jcShadowConst
    [No shadow] = 0
    Shadow = 1
End Enum

'gradient backstyle
Public Enum jcBackStyleConst
    g_Opaque = 0
    g_Transparent = 1
End Enum

'icon aligment
Public Enum IconAlignConst
    vbLeftAligment = 0
    vbRightAligment = 1
End Enum

'*************************************************************
'   events
'*************************************************************
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************
'   Required API Declarations
'*************************************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private useMask As Boolean, useGrey As Boolean

'*************************************************************
'   Members
'*************************************************************
Private m_FrameColor            As OLE_COLOR
Private m_TextBoxColor          As OLE_COLOR
Private m_BackColor             As OLE_COLOR
Private m_FillColor             As OLE_COLOR

Private m_FrameColorDis         As OLE_COLOR
Private m_TextBoxColorDis       As OLE_COLOR
Private m_FillColorDis          As OLE_COLOR
Private jcColorToDis            As OLE_COLOR
Private jcColorFromDis          As OLE_COLOR
Private jcColorBorderPicDis     As OLE_COLOR

Private m_FrameColorIni         As OLE_COLOR
Private m_TextBoxColorIni       As OLE_COLOR
Private m_FillColorIni          As OLE_COLOR
Private jcColorToIni            As OLE_COLOR
Private jcColorFromIni          As OLE_COLOR
Private jcColorBorderPicIni     As OLE_COLOR

Private m_Caption               As String
Private m_Enabled               As Boolean
Private m_TextBoxHeight         As Long
Private m_TextHeight            As Long
Private m_TextWidth             As Long
Private m_Height                As Long
Private m_TextColor             As Long
Private m_Alignment             As Long
Private m_Font                  As StdFont
Private m_RoundedCorner         As Boolean
Private m_RoundedCornerTxtBox   As Boolean
Private m_Style                 As jcStyleConst
Private m_HeaderStyle           As jcHeaderConst
Private m_GradientHeaderStyle   As jcGradConst
Private m_Icon                  As StdPicture
Private m_IconSize              As Integer
Private m_IconAlignment         As IconAlignConst
Private m_ThemeColor            As jcThemeConst
Private m_ColorTo               As OLE_COLOR
Private m_ColorFrom             As OLE_COLOR
Private m_Indentation           As Integer
Private m_Space                 As Integer
Private m_TxtBoxShadow          As jcShadowConst
Private jcTextBoxCenter         As Long
Private jcTextDrawParams        As Long
Private jcColorTo               As OLE_COLOR
Private jcColorFrom             As OLE_COLOR
Private jcColorBorderPic        As OLE_COLOR
Private jcLpp As POINT
Private Const TEXT_INACTIVE = &H80000011 '&H6A6A6A
Private Const m_Border_Inactive = &H8000000B
Private Const m_BtnFace_Inactive = &H8000000F
Private Const m_BtnFace = &H80000016 '&H8000000F '&H80000016&

'*************************************************************
'   Constants
'*************************************************************
Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_NOCLIP = &H100
Private Const DT_CALCRECT = &H400

Private Const ALTERNATE = 1      ' ALTERNATE and WINDING are
Private Const WINDING = 2        ' constants for FillMode.
Private Const BLACKBRUSH = 4     ' Constant for brush type.
Private Const WHITE_BRUSH = 0    ' Constant for brush type.

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4

'==========================================================================
' Init, Initialize, Read & Write UserControl
'==========================================================================
Private Sub UserControl_InitProperties()
    'Set default properties
    m_Caption = Ambient.DisplayName
    m_BackColor = TranslateColor(Ambient.BackColor)
    m_FillColorIni = TranslateColor(Ambient.BackColor)
    m_RoundedCorner = True
    m_RoundedCornerTxtBox = False
    m_Style = jcGradient
    m_ThemeColor = Blue
    m_TextColor = TranslateColor(vbBlack)
    m_FrameColorIni = TranslateColor(vbBlack)
    m_TextBoxColorIni = TranslateColor(vbWhite)
    m_TxtBoxShadow = [No shadow]
    m_TextBoxHeight = 22
    m_HeaderStyle = Gradient
    m_GradientHeaderStyle = VerticalGradient
    SetjcTextDrawParams
End Sub

Private Sub UserControl_Initialize()
    Set m_Font = New StdFont
    Set UserControl.Font = m_Font
    m_IconSize = 16
    m_ColorFrom = 10395391
    m_ColorTo = 15790335
    m_TxtBoxShadow = [No shadow]
    m_ThemeColor = Blue
    m_Enabled = True
    Call SetDefaultThemeColor(m_ThemeColor)
    m_TextBoxHeight = 22
    m_Alignment = vbCenter
    m_IconAlignment = vbLeftAligment
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_FrameColorIni = .ReadProperty("FrameColor", vbBlack)
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_FillColorIni = .ReadProperty("FillColor", Ambient.BackColor)
        m_TextBoxColorIni = .ReadProperty("TextBoxColor", vbWhite)
        m_TxtBoxShadow = .ReadProperty("TxtBoxShadow", [No shadow])
        m_Style = .ReadProperty("Style", jcGradient)
        m_RoundedCorner = .ReadProperty("RoundedCorner", True)
        m_Enabled = .ReadProperty("Enabled", True)
        m_RoundedCornerTxtBox = .ReadProperty("RoundedCornerTxtBox", False)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        m_TextBoxHeight = .ReadProperty("TextBoxHeight", 22)
        m_TextColor = .ReadProperty("TextColor", vbBlack)
        m_Alignment = .ReadProperty("Alignment", vbCenter)
        m_IconAlignment = .ReadProperty("IconAlignment", vbLeftAligment)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        Set m_Icon = .ReadProperty("Picture", Nothing)
        m_IconSize = .ReadProperty("IconSize", 16)
        m_ThemeColor = .ReadProperty("ThemeColor", Blue)
        m_ColorFrom = .ReadProperty("ColorFrom", 10395391)
        m_ColorTo = .ReadProperty("ColorTo", 15790335)
        m_HeaderStyle = .ReadProperty("HeaderStyle", TxtBoxColor)
        m_GradientHeaderStyle = .ReadProperty("GradientHeaderStyle", VerticalGradient)
    End With
    
    'Add properties
    UserControl.BackColor = TranslateColor(m_BackColor)
    SetjcTextDrawParams
    SetFont m_Font
    Call SetDefaultThemeColor(m_ThemeColor)
    SetDisabledColor
    'Paint control
    PaintFrame
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "FrameColor", m_FrameColorIni, vbBlack
        .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
        .WriteProperty "FillColor", m_FillColorIni, Ambient.BackColor
        .WriteProperty "TextBoxColor", m_TextBoxColorIni, vbWhite
        .WriteProperty "TxtBoxShadow", m_TxtBoxShadow, [No shadow]
        .WriteProperty "Style", m_Style, jcGradient
        .WriteProperty "RoundedCorner", m_RoundedCorner, True
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "RoundedCornerTxtBox", m_RoundedCornerTxtBox, False
        .WriteProperty "Caption", m_Caption, Ambient.DisplayName
        .WriteProperty "TextBoxHeight", m_TextBoxHeight, 22
        .WriteProperty "TextColor", m_TextColor, vbBlack
        .WriteProperty "Alignment", m_Alignment, vbCenter
        .WriteProperty "IconAlignment", m_IconAlignment, vbLeftAligment
        .WriteProperty "Font", m_Font, Ambient.Font
        .WriteProperty "Picture", m_Icon, Nothing
        .WriteProperty "IconSize", m_IconSize, 16
        .WriteProperty "ThemeColor", m_ThemeColor, Blue
        .WriteProperty "ColorFrom", m_ColorFrom, 10395391
        .WriteProperty "ColorTo", m_ColorTo, 15790335
        .WriteProperty "HeaderStyle", m_HeaderStyle, TxtBoxColor
        .WriteProperty "GradientHeaderStyle", m_GradientHeaderStyle, VerticalGradient
    End With
End Sub

'==========================================================================
' Usercontrol events
'==========================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 700 Then UserControl.Width = 700
    If UserControl.Height < 400 Then UserControl.Height = 400
    PaintFrame
End Sub

'==========================================================================
' Properties
'==========================================================================
Public Property Let FrameColor(ByRef new_FrameColor As OLE_COLOR)
    m_FrameColorIni = TranslateColor(new_FrameColor)
    If m_ThemeColor = Custom Then jcColorBorderPic = m_FrameColor
    PropertyChanged "FrameColor"
    PaintFrame
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = m_FrameColorIni
End Property

Public Property Let FillColor(ByRef new_FillColor As OLE_COLOR)
Attribute FillColor.VB_Description = "Returns/Sets the Fill color for TextBox and Windows style"
    m_FillColorIni = TranslateColor(new_FillColor)
    PropertyChanged "FillColor"
    PaintFrame
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColorIni
End Property

Public Property Let RoundedCornerTxtBox(ByRef new_RoundedCornerTxtBox As Boolean)
    m_RoundedCornerTxtBox = new_RoundedCornerTxtBox
    PropertyChanged "RoundedCornerTxtBox"
    PaintFrame
End Property

Public Property Get RoundedCornerTxtBox() As Boolean
    RoundedCornerTxtBox = m_RoundedCornerTxtBox
End Property

Public Property Let Enabled(ByRef new_Enabled As Boolean)
    m_Enabled = new_Enabled
    PropertyChanged "Enabled"
    PaintFrame
    FrameEnabled m_Enabled
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let TxtBoxShadow(ByRef new_TxtBoxShadow As jcShadowConst)
    m_TxtBoxShadow = new_TxtBoxShadow
    PropertyChanged "TxtBoxShadow"
    PaintFrame
End Property

Public Property Get TxtBoxShadow() As jcShadowConst
    TxtBoxShadow = m_TxtBoxShadow
End Property

Public Property Let RoundedCorner(ByRef new_RoundedCorner As Boolean)
    m_RoundedCorner = new_RoundedCorner
    PropertyChanged "RoundedCorner"
    PaintFrame
End Property

Public Property Get RoundedCorner() As Boolean
    RoundedCorner = m_RoundedCorner
End Property

Public Property Let Caption(ByRef new_caption As String)
    m_Caption = new_caption
    PaintFrame
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Alignment(ByRef new_Alignment As AlignmentConstants)
    m_Alignment = new_Alignment
    SetjcTextDrawParams
    PropertyChanged "Alignment"
    PaintFrame
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Style(ByRef new_Style As jcStyleConst)
    m_Style = new_Style
    PropertyChanged "Style"
    SetDefault ' m_ThemeColor
    PaintFrame
End Property

Public Property Get Style() As jcStyleConst
    Style = m_Style
End Property

Public Property Let TextBoxHeight(ByRef new_TextBoxHeight As Long)
    m_TextBoxHeight = new_TextBoxHeight
    PropertyChanged "TextBoxHeight"
    PaintFrame
End Property

Public Property Get TextBoxHeight() As Long
    TextBoxHeight = m_TextBoxHeight
End Property

Public Property Let TextColor(ByRef new_TextColor As OLE_COLOR)
    m_TextColor = TranslateColor(new_TextColor)
    PropertyChanged "TextColor"
    PaintFrame
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextBoxColor(ByRef new_TextBoxColor As OLE_COLOR)
    m_TextBoxColorIni = TranslateColor(new_TextBoxColor)
    PropertyChanged "TextBoxColor"
    PaintFrame
End Property

Public Property Get TextBoxColor() As OLE_COLOR
    TextBoxColor = m_TextBoxColorIni
End Property

Public Property Let BackColor(ByRef new_BackColor As OLE_COLOR)
    m_BackColor = TranslateColor(new_BackColor)
    UserControl.BackColor = m_BackColor
    PropertyChanged "BackColor"
    PaintFrame
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Set Font(ByRef new_font As StdFont)
    SetFont new_font
    PropertyChanged "Font"
    PaintFrame
End Property

Public Property Let Font(ByRef new_font As StdFont)
    SetFont new_font
    PropertyChanged "Font"
    PaintFrame
End Property
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Icon
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Icon = New_Picture
    PropertyChanged "Picture"
    PaintFrame
End Property

Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_Value As Integer)
    m_IconSize = New_Value
    PropertyChanged "IconSize"
    PaintFrame
End Property

Public Property Let IconAlignment(ByRef new_IconAlignment As IconAlignConst)
    m_IconAlignment = new_IconAlignment
    PropertyChanged "IconAlignment"
    PaintFrame
End Property

Public Property Get IconAlignment() As IconAlignConst
    IconAlignment = m_IconAlignment
End Property

Public Property Get ThemeColor() As jcThemeConst
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As jcThemeConst)
    If m_ThemeColor <> vData Then
        m_ThemeColor = vData
        Call SetDefaultThemeColor(m_ThemeColor)
        PaintFrame
        PropertyChanged "ThemeColor"
    End If
End Property

Public Property Get ColorFrom() As OLE_COLOR
Attribute ColorFrom.VB_Description = "Returns/Sets the Start color for gradient"
    ColorFrom = m_ColorFrom
End Property

Public Property Let ColorFrom(ByRef new_ColorFrom As OLE_COLOR)
    m_ColorFrom = TranslateColor(new_ColorFrom)
    If m_ThemeColor = Custom Then jcColorFromIni = m_ColorFrom
    PropertyChanged "ColorFrom"
    PaintFrame
End Property

Public Property Get ColorTo() As OLE_COLOR
Attribute ColorTo.VB_Description = "Returns/Sets the End color for gradient"
    ColorTo = m_ColorTo
End Property

Public Property Let ColorTo(ByRef new_ColorTo As OLE_COLOR)
    m_ColorTo = TranslateColor(new_ColorTo)
    If m_ThemeColor = Custom Then jcColorToIni = m_ColorTo
    PropertyChanged "ColorTo"
    PaintFrame
End Property

Public Property Let HeaderStyle(ByRef new_HeaderStyle As jcHeaderConst)
    m_HeaderStyle = new_HeaderStyle
    PropertyChanged "HeaderStyle"
    PaintFrame
End Property

Public Property Get HeaderStyle() As jcHeaderConst
    HeaderStyle = m_HeaderStyle
End Property

Public Property Let GradientHeaderStyle(ByRef new_GradientHeaderStyle As jcGradConst)
    m_GradientHeaderStyle = new_GradientHeaderStyle
    PropertyChanged "GradientHeaderStyle"
    PaintFrame
End Property

Public Property Get GradientHeaderStyle() As jcGradConst
    GradientHeaderStyle = m_GradientHeaderStyle
End Property

Private Sub SetjcTextDrawParams()
    'Set text draw params using m_Alignment
    If m_Style = Panel Then
        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_WORDBREAK Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_WORDBREAK Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_WORDBREAK Or DT_VCENTER
        End If
    Else
        If m_Alignment = vbLeftJustify Then
            jcTextDrawParams = DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
        ElseIf m_Alignment = vbRightJustify Then
            jcTextDrawParams = DT_RIGHT Or DT_SINGLELINE Or DT_VCENTER
        Else
            jcTextDrawParams = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
        End If
    End If
End Sub

Private Sub SetFont(ByRef new_font As StdFont)
    With m_Font
        .Bold = new_font.Bold
        .Italic = new_font.Italic
        .Name = new_font.Name
        .Size = new_font.Size
    End With
    Set UserControl.Font = m_Font
End Sub

'==========================================================================
' Functions and subroutines
'==========================================================================
Private Sub SetDefaultThemeColor(ThemeType As Long)
    Select Case ThemeType
        Case 0 '"NormalColor"
            jcColorFromIni = RGB(129, 169, 226)
            jcColorToIni = RGB(221, 236, 254)
            jcColorBorderPicIni = RGB(0, 0, 128)
        Case 1 '"Metallic"
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)
        Case 2 '"HomeStead"
            jcColorFromIni = RGB(181, 197, 143)
            jcColorToIni = RGB(247, 249, 225)
            jcColorBorderPicIni = RGB(63, 93, 56)
        Case 3 '"Visual2005"
            jcColorFromIni = RGB(194, 194, 171)
            jcColorToIni = RGB(248, 248, 242)
            jcColorBorderPicIni = RGB(145, 145, 115)
        Case 4 '"Norton2004"
            jcColorFromIni = RGB(217, 172, 1)
            jcColorToIni = RGB(255, 239, 165)
            jcColorBorderPicIni = RGB(117, 91, 30)
        Case 5  'Custom
            jcColorFromIni = m_ColorFrom
            jcColorToIni = m_ColorTo
            jcColorBorderPicIni = m_FrameColor
        Case Else
            jcColorFromIni = RGB(153, 151, 180)
            jcColorToIni = RGB(244, 244, 251)
            jcColorBorderPicIni = RGB(75, 75, 111)
    End Select
End Sub

'==================
' Main drawing sub
'==================
Private Sub PaintFrame()
    Dim R As RECT, R_Caption As RECT, RC As RECT
    Dim Ix As Integer, Iy As Integer
    
    m_Height = 3
    m_Indentation = 15
    m_Space = 6
    Ix = 0
    Iy = 0
    EraseRegion
    
    'Clear user control
    UserControl.Cls
    
    'Set caption height and width
    If Len(m_Caption) <> 0 Then
        m_TextWidth = UserControl.TextWidth(m_Caption)
        m_TextHeight = UserControl.TextHeight(m_Caption)
        jcTextBoxCenter = m_TextHeight / 2
    Else
        jcTextBoxCenter = 0
    End If
    
    'Select colors according to enabled property
    If m_Enabled = False Then
        m_FrameColor = m_FrameColorDis
        m_TextBoxColor = m_TextBoxColorDis
        m_FillColor = m_FillColorDis
        jcColorTo = jcColorToDis
        jcColorFrom = jcColorFromDis
        jcColorBorderPic = jcColorBorderPicDis
    Else
        m_FrameColor = m_FrameColorIni
        m_TextBoxColor = m_TextBoxColorIni
        m_FillColor = m_FillColorIni
        jcColorTo = jcColorToIni
        jcColorFrom = jcColorFromIni
        jcColorBorderPic = jcColorBorderPicIni
    End If
  
    'select frame style
    Select Case m_Style
        Case Is = XPDefault
            Draw_XPDefault R_Caption
        Case Is = jcGradient
            Draw_jcGradient R_Caption, Iy
        Case Is = TextBox
            Draw_TextBox R_Caption, Ix, Iy
        Case Is = Windows
            Draw_Windows R_Caption, Iy
        Case Is = Messenger
            Draw_Messenger R_Caption, Iy
        Case Is = InnerWedge
            Draw_InnerWedge R_Caption, Iy
        Case Is = OuterWedge
            Draw_OuterWedge R_Caption, Iy
        Case Is = Header
            Draw_Header R_Caption
        Case Is = Panel
            Draw_Panel R_Caption, Iy
    End Select
    
    'caption and icon alignments
    If Not (m_Icon Is Nothing Or m_Style = XPDefault) Then
        If m_IconAlignment = vbLeftAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            If m_Style = TextBox Then
                Ix = m_Indentation + m_Space * 2
            Else
                Ix = m_Space
            End If
        ElseIf m_IconAlignment = vbRightAligment Then
            If m_Alignment = vbLeftJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            ElseIf m_Alignment = vbRightJustify Then
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            Else
                R_Caption.Left = R_Caption.Left + m_Space + m_IconSize
                R_Caption.Right = R_Caption.Right - (m_Space + m_IconSize)
            End If
            If m_Style = TextBox Then
                Ix = UserControl.ScaleWidth - m_Space * 2 - m_IconSize - m_Indentation
            Else
                Ix = UserControl.ScaleWidth - m_Space - m_IconSize
            End If
        End If
    End If

    'Draw caption
    If Len(m_Caption) <> 0 Then
        'Set text color
        Dim m_caption_aux As String
        m_caption_aux = TrimWord(m_Caption, R_Caption.Right - R_Caption.Left)
        
        'Draw text
        UserControl.ForeColor = IIf(m_Enabled, m_TextColor, TranslateColor(TEXT_INACTIVE))
        
        If m_Style = Panel Then
            CopyRect RC, R_Caption
            DrawTextEx UserControl.hDC, m_Caption, Len(m_Caption), RC, DT_CALCRECT Or DT_WORDBREAK, ByVal 0&
            OffsetRect RC, (R_Caption.Right - RC.Right) \ 2, (R_Caption.Bottom - RC.Bottom) \ 2
            DrawTextEx UserControl.hDC, m_Caption, Len(m_Caption), RC, jcTextDrawParams, ByVal 0&
        Else
            DrawTextEx UserControl.hDC, m_caption_aux, Len(m_caption_aux), R_Caption, jcTextDrawParams, ByVal 0&
        End If
        
    End If
    
    'draw picture
    If Not (m_Icon Is Nothing Or m_Style = XPDefault Or m_Style = InnerWedge Or m_Style = OuterWedge) Then
        If m_Style = Messenger Then
            If Iy < m_Height * 2 + 2 Then Iy = m_Height * 2 + 2
        ElseIf m_Style = jcGradient Then
            If Iy < m_Height + 2 Then Iy = m_Height + 2
        Else
            If Iy < 0 Then Iy = m_Space / 2
        End If
        
        If m_Enabled = True Then
            UserControl.PaintPicture m_Icon, Ix, Iy, m_IconSize, m_IconSize
            'TransBlt UserControl.hDC, Ix, Iy, m_IconSize, m_IconSize, m_Icon, vbBlack, , , False, False
        Else
            TransBlt UserControl.hDC, Ix, Iy, m_IconSize, m_IconSize, m_Icon, vbBlack, , , True, False
        End If
    End If
End Sub

Private Sub SetDefault()
    Select Case m_Style
        Case Is = XPDefault
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = jcGradient
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            SetjcTextDrawParams
        Case Is = TextBox
            m_TextColor = vbBlack
            m_FrameColorIni = &H6A6A6A
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCornerTxtBox = True
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Windows
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = &HB0EFF0
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_RoundedCorner = True
            m_RoundedCornerTxtBox = False
            m_FillColorIni = &HE0FFFF
            m_GradientHeaderStyle = HorizontalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams
        Case Is = Messenger
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VerticalGradient
            m_HeaderStyle = TxtBoxColor
            SetjcTextDrawParams
        Case Is = InnerWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 192
            m_TextBoxColorIni = 192
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = OuterWedge
            m_TextColor = vbWhite
            m_FrameColorIni = 10878976
            m_TextBoxColorIni = 10878976
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Header
            m_TextColor = &HCF3603
            m_FrameColorIni = RGB(195, 195, 195)
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbLeftJustify
            m_FillColorIni = TranslateColor(Ambient.BackColor)
            SetjcTextDrawParams
        Case Is = Panel
            m_TextColor = vbBlack
            m_FrameColorIni = vbBlack
            m_TextBoxColorIni = vbWhite
            m_TextBoxHeight = 22
            m_Alignment = vbCenter
            m_ThemeColor = Blue
            m_GradientHeaderStyle = VCilinderGradient
            m_HeaderStyle = Gradient
            SetjcTextDrawParams
    End Select
    
End Sub

Private Sub PaintShpInBar(iColorA As Long, iColorB As Long, m_Height As Long)
    Dim I As Integer, x_left As Integer, y_top As Integer, SpaceBtwnShp As Integer, NumShp As Integer
    Dim RectHeight As Long, RectWidth As Long, R As RECT

    SpaceBtwnShp = 2    'space between shapes
    NumShp = 9          'number of points
    RectHeight = 2      'shape height
    RectWidth = 2       'shape width
    
    'x and y shape  coordinates
    x_left = (UserControl.ScaleWidth - NumShp * RectWidth - (NumShp - 1) * SpaceBtwnShp) / 2
    y_top = (m_Height - RectHeight) / 2
    
    For I = 0 To NumShp - 1
        SetRect R, x_left + I * SpaceBtwnShp + I * RectWidth + 1, y_top + 1, 1, 1
        APIRectangle UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, iColorA
        SetRect R, x_left + I * SpaceBtwnShp + I * RectWidth, y_top, 1, 1
        APIRectangle UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, iColorB
    Next I
End Sub

Private Sub Draw_XPDefault(R_Caption As RECT)
    Dim p_left As Long, m_roundedRadius As Long, R As RECT, lpp As POINT
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    If Len(m_Caption) <> 0 Then
        If m_Alignment = vbLeftJustify Then
            p_left = m_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Indentation - m_Space - 1
        Else
            p_left = (UserControl.ScaleWidth - 1 - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor

        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If
End Sub

Private Sub Draw_jcGradient(R_Caption As RECT, Iy As Integer)
    Dim R As RECT, m_roundedRadius As Long
    
    jcTextBoxCenter = m_TextBoxHeight / 2
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R

    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height
    DrawGradientInRectangle UserControl.hDC, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic
    
    If m_HeaderStyle = Gradient Then
         SetRect R, 0, m_Height, UserControl.ScaleWidth - 2, m_TextBoxHeight
         DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
         SetRect R, 0, m_Height, UserControl.ScaleWidth - 1, m_TextBoxHeight + m_Height + 2
         DrawAPIRoundRect False, 0&, m_FillColor, m_FrameColor, R
    End If

    SetRect R, 0, m_Height + m_TextBoxHeight, UserControl.ScaleWidth - 2, m_Height
    DrawGradientInRectangle UserControl.hDC, jcColorTo, jcColorFrom, R, VCilinderGradient, True, jcColorBorderPic

    SetRect R, 1, m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
    DrawGradientInRectangle UserControl.hDC, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
    
    'set caption rect
    SetRect R_Caption, m_Space, m_Height + 1, UserControl.ScaleWidth - 2 - m_Space, m_TextBoxHeight + 2

    'set icon Y coordinate
    Iy = (m_Height * 2 + m_TextBoxHeight - m_IconSize) / 2
End Sub

Private Sub Draw_TextBox(R_Caption As RECT, Ix As Integer, Iy As Integer)
     Dim m_roundedRadius As Long, R As RECT
     
     jcTextBoxCenter = m_TextBoxHeight / 2
     
     'Draw border rectangle
     SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
     DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

     'Draw textbox border rectangle
     If m_HeaderStyle = Gradient Then
         If m_TxtBoxShadow = Shadow Then
            SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight
            OffsetRect R, 2, 2
            DrawAPIRoundRect False, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
         End If
         SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 2 - 2 * m_Indentation, m_TextBoxHeight - 1
         DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor ', 3.08
     Else
         SetRect R, m_Indentation, 0, UserControl.ScaleWidth - 1 - m_Indentation, m_TextBoxHeight
         If m_TxtBoxShadow = Shadow Then
            OffsetRect R, 2, 2
            DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, BlendColors(m_FillColor, &HA7A7A7), BlendColors(m_FillColor, &HA7A7A7), R
            OffsetRect R, -2, -2
         End If
         DrawAPIRoundRect m_RoundedCornerTxtBox, m_TextBoxHeight, m_TextBoxColor, m_FrameColor, R
     End If
    
     'set caption rect
     SetRect R_Caption, m_Indentation + m_Space * 1.5, 0, UserControl.ScaleWidth - 1 - m_Indentation - m_Space * 1.5, m_TextBoxHeight - 1
     
     'set icon coordinates
     Ix = m_Indentation + m_Space * 2
     Iy = (m_TextBoxHeight - m_IconSize) / 2

End Sub

Private Sub Draw_Windows(R_Caption As RECT, Iy As Integer)
     Dim R As RECT, m_roundedRadius As Long
     
     jcTextBoxCenter = m_TextBoxHeight / 2
     
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R

    'Draw text box borders
    If m_HeaderStyle = Gradient Then
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 2, m_TextBoxHeight - 1
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, m_FrameColor ', 3.08
    Else
        SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, m_TextBoxHeight
        DrawAPIRoundRect m_RoundedCornerTxtBox, 10&, m_TextBoxColor, m_FrameColor, R
    End If
     
    'set caption rect
    SetRect R_Caption, m_Space, 0, UserControl.ScaleWidth - m_Space, m_TextBoxHeight '- 1
     
    'set icon coordinates
    Iy = (m_TextBoxHeight - m_IconSize) / 2
End Sub

Private Sub Draw_Messenger(R_Caption As RECT, Iy As Integer)
    Dim R As RECT, m_roundedRadius As Long
    
    jcTextBoxCenter = 0
    
    'Draw border rectangle
    SetRect R, 0&, jcTextBoxCenter, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, BlendColors(jcColorFrom, vbWhite), IIf(m_ThemeColor = Custom, m_FrameColor, jcColorBorderPic), R
    
    'Draw header
    SetRect R, 0, 0, UserControl.ScaleWidth - 2, m_Height * 2
    DrawGradientInRectangle UserControl.hDC, jcColorFrom, vbWhite, R, VerticalGradient, True, jcColorBorderPic, 2.01

    PaintShpInBar vbWhite, BlendColors(vbBlack, jcColorFrom), m_Height * 2

    If m_HeaderStyle = Gradient Or m_Enabled = False Then
        SetRect R, 0&, m_Height * 2, UserControl.ScaleWidth - 2, m_TextBoxHeight + 1
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, True, jcColorBorderPic
    Else
        SetRect R, 0, m_Height * 2 + m_TextBoxHeight + 1, UserControl.ScaleWidth - 2, m_Height * 2 + m_TextBoxHeight + 1
        APILineEx UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, jcColorBorderPic 'vbBlack
    End If

    SetRect R, 1, 1 + m_Height * 2 + m_TextBoxHeight, UserControl.ScaleWidth - 3, UserControl.ScaleHeight - (2 + m_Height * 2 + m_TextBoxHeight) - UserControl.ScaleHeight * 0.2
    DrawGradientInRectangle UserControl.hDC, BlendColors(jcColorFrom, vbWhite), BlendColors(jcColorTo, vbWhite), R, VerticalGradient, False, m_TextBoxColor
    
    'set caption rect
    SetRect R_Caption, m_Space, m_Height * 2 + 2, UserControl.ScaleWidth - 1 - m_Space, m_TextBoxHeight + 6

    'set icon coordinates
    Iy = m_Height * 2 + (m_TextBoxHeight - m_IconSize) / 2

End Sub

Private Sub Draw_InnerWedge(R_Caption As RECT, Iy As Integer)
    Dim txtWidth As Integer, txtHeight As Integer, R As RECT
    Dim m_roundedRadius As Long, hFRgn As Long
    Dim poly(1 To 4) As POINT, NumCoords As Long, hBrush As Long, hRgn As Long
    
    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)

    txtWidth = m_TextWidth + 10
    If txtWidth < 100 Then txtWidth = 100
    txtHeight = m_TextHeight + 5
    NumCoords = 4
    
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    If (txtWidth + R.Left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.Left - m_Indentation - 1
    End If
    
    'Assign values to points.
    poly(1).x = R.Left:                               poly(1).y = R.Top
    poly(2).x = R.Left:                               poly(2).y = R.Top + txtHeight
    poly(3).x = R.Left + txtWidth:                    poly(3).y = R.Top + txtHeight
    poly(4).x = R.Left + txtWidth + txtHeight / 2:    poly(4).y = R.Top
    
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
    
    'fill frame
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, R

    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)
    If hRgn Then FillRgn UserControl.hDC, hRgn, hBrush
    
    'draw frame borders
    APILineEx UserControl.hDC, poly(2).x, poly(2).y, poly(3).x, poly(3).y, m_FrameColor
    APILineEx UserControl.hDC, poly(3).x, poly(3).y, poly(4).x, poly(4).y, m_FrameColor
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FrameColor, R, True
    
    'delete created region
    DeleteObject hRgn
    DeleteObject hBrush
    
    'set caption rectangle
    SetRect R_Caption, poly(1).x + m_Indentation / 2, poly(1).y, txtWidth + poly(1).x, txtHeight + poly(1).y + 2
    
'    'set icon coordinates
'    Iy = (txtHeight - m_IconSize) / 2
    UserControl.FillStyle = 0
End Sub

Private Sub Draw_OuterWedge(R_Caption As RECT, Iy As Integer)
    Dim txtWidth As Integer, txtHeight As Integer, R As RECT, R1 As RECT
    Dim m_roundedRadius As Long, hFRgn As Long
    Dim poly(1 To 4) As POINT, NumCoords As Long, hBrush As Long, hRgn As Long
    
    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 10&)

    txtWidth = m_TextWidth + 10
    If txtWidth < 100 Then txtWidth = 100
    txtHeight = m_TextHeight + 5
    NumCoords = 4
    
    SetRect R, 0&, 0&, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    If (txtWidth + R.Left + txtHeight / 2) >= R.Right - m_Indentation Then
        txtWidth = R.Right - txtHeight / 2 - R.Left - m_Indentation - 1
    End If
    
    'Assign values to points.
    poly(1).x = R.Left + 6:                          poly(1).y = R.Top
    poly(2).x = R.Left + 6:                          poly(2).y = R.Top + txtHeight
    poly(3).x = R.Left + txtWidth + txtHeight / 2:   poly(3).y = R.Top + txtHeight
    poly(4).x = R.Left + txtWidth:                   poly(4).y = R.Top
    
    'Creates first region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    
    'If the creation of the region was successful then color.
    hBrush = CreateSolidBrush(m_TextBoxColor)
    If hRgn Then FillRgn UserControl.hDC, hRgn, hBrush
    
    'fill frame
    SetRect R1, 0&, 0&, txtWidth * 0.9, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 10&, m_TextBoxColor, m_FrameColor, R1
    SetRect R1, txtWidth * 0.9 - 5, 1, txtWidth * 0.9 + 3, txtHeight * 1.3
    DrawAPIRoundRect m_RoundedCorner, 0&, m_TextBoxColor, m_TextBoxColor, R1
    
    SetRect R1, -1, -1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1
    DrawAPIRoundRect m_RoundedCorner, 10&, m_FillColor, m_FillColor, R1, True
    
   'draw frame borders
    UserControl.ForeColor = m_FrameColor
    APILineEx UserControl.hDC, poly(1).x, poly(1).y, poly(4).x, poly(4).y, UserControl.ForeColor
    APILineEx UserControl.hDC, poly(4).x, poly(4).y, poly(3).x, poly(3).y, UserControl.ForeColor
    
    RoundRect UserControl.hDC, R.Left, R.Top + txtHeight, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
    RoundRect UserControl.hDC, R.Left, R.Top + txtHeight, R.Left + 10, R.Top + txtHeight + 10, 0, 0
    
    UserControl.ForeColor = m_FillColor
    RoundRect UserControl.hDC, R.Left + 1, R.Top + txtHeight + 1, R.Left + 10, R.Top + txtHeight + 10, 0, 0
    
    'delete created region
    DeleteObject hRgn
    DeleteObject hBrush
    
    'set caption rectangle
    SetRect R_Caption, poly(1).x + m_Indentation / 2 - 6, poly(1).y, txtWidth + poly(1).x - 6, txtHeight + poly(1).y + 2
    
End Sub

Private Sub Draw_Header(R_Caption As RECT)
    Dim p_left As Long
    
    APILineEx UserControl.hDC, 0&, jcTextBoxCenter, UserControl.ScaleWidth, jcTextBoxCenter, IIf(m_Enabled, TranslateColor(&H80000015), TranslateColor(TEXT_INACTIVE)) 'TranslateColor(&H80000015)&H808080
    APILineEx UserControl.hDC, 0&, jcTextBoxCenter + 1, UserControl.ScaleWidth, jcTextBoxCenter + 1, vbWhite
    
    If Len(m_Caption) <> 0 Then

        If m_Alignment = vbLeftJustify Then
            p_left = 0 'm_Indentation
        ElseIf m_Alignment = vbRightJustify Then
            p_left = UserControl.ScaleWidth - m_TextWidth - m_Space
        Else
            p_left = (UserControl.ScaleWidth - m_TextWidth) / 2
        End If

        'Draw a line
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter, p_left + m_TextWidth + m_Space, jcTextBoxCenter, m_FillColor 'TranslateColor(Ambient.BackColor)
        APILineEx UserControl.hDC, p_left, jcTextBoxCenter + 1, p_left + m_TextWidth + m_Space, jcTextBoxCenter + 1, m_FillColor 'TranslateColor(Ambient.BackColor)
        
        'set caption rect
        SetRect R_Caption, p_left + m_Space / 2, 0, m_TextWidth + p_left + m_Space / 2, m_TextHeight
    End If
End Sub

Private Sub Draw_Panel(R_Caption As RECT, Iy As Integer)
    Dim R As RECT, m_roundedRadius As Long, hFRgn As Long, hRgn As Long
    
    jcTextBoxCenter = m_TextBoxHeight / 2
    
    'Draw border rectangle
    UserControl.FillColor = m_FillColor
    
    If m_ThemeColor = Custom Or m_HeaderStyle = TxtBoxColor Then
        UserControl.ForeColor = m_FrameColor
    Else
        UserControl.ForeColor = jcColorBorderPic
    End If

    'If m_Enabled = False Then UserControl.ForeColor = m_Border_Inactive

    m_roundedRadius = IIf(m_RoundedCorner = False, 0&, 9&)

    SetRect R, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight
    If m_HeaderStyle = Gradient Then
        DrawGradientInRectangle UserControl.hDC, jcColorFrom, jcColorTo, R, m_GradientHeaderStyle, False, UserControl.ForeColor, 2.03
    End If
    
    'Creates first region to fill with color.
    hRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, 0&, 0&)
    'Creates second region to fill with color.
    hFRgn = CreateRoundRectRgn(R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius)
    'Combine our two regions
    CombineRgn hRgn, hRgn, hFRgn, RGN_AND
    'delete second region
    DeleteObject hFRgn
 
    SetWindowRgn UserControl.hwnd, hRgn, True
    
    UserControl.FillStyle = IIf(m_HeaderStyle = Gradient, 1, 0)
    
    If UserControl.ForeColor <> UserControl.BackColor Or m_HeaderStyle = TxtBoxColor Then
        RoundRect UserControl.hDC, R.Left, R.Top, R.Right - 1, R.Bottom - 1, m_roundedRadius, m_roundedRadius
        UserControl.FillStyle = 0
        DrawCorners UserControl.ForeColor
    End If
    
    'set caption rect
    SetRect R_Caption, m_Space, 0&, UserControl.ScaleWidth - m_Space, UserControl.ScaleHeight - 2

    'set icon coordinates
    Iy = (UserControl.ScaleHeight - m_IconSize) / 2
End Sub

'==========================================================================
' API Functions and subroutines
'==========================================================================

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)

    'Use the API LineTo for Fast Drawing
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, pt
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal H As Long, Optional lcolor As OLE_COLOR = -1) As Long
    'Draw an api rectangle
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINT
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(hDC, hPen)
    MoveToEx hDC, x, y, pt
    LineTo hDC, x + w, y
    LineTo hDC, x + w, y + H
    LineTo hDC, x, y + H
    LineTo hDC, x, y
    SelectObject hDC, hPenOld
    DeleteObject hPen
End Function

Private Sub DrawGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal y As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional blnVertical = True)
    
    'Draw a Vertical or horizontal Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
        dR = (sR - eR) / Y2
        dG = (sG - eG) / Y2
        dB = (sB - eB) / Y2
        For ni = 1 To Y2 - 1
            APILineEx lhdcEx, x, y + ni, X2, y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    Else
        dR = (sR - eR) / X2
        dG = (sG - eG) / X2
        dB = (sB - eB) / X2
        For ni = 1 To X2 - 1
            APILineEx lhdcEx, x + ni, y, x + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
    End If
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lcolor As Long) As Long

    If OleTranslateColor(lcolor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
End Function

Private Function TrimWord(strCaption As String, lngWidth As Long) As String
    Dim strCaptionAux As String, lngLenOfText As Long
    TrimWord = strCaption
    If TextWidth(strCaption) > lngWidth Then
        lngLenOfText = Len(strCaption)
        Do Until TextWidth(TrimWord & "...") <= lngWidth Or lngLenOfText = 0
            lngLenOfText = lngLenOfText - 1
            TrimWord = Left(TrimWord, lngLenOfText)
        Loop
        If lngLenOfText = 0 Then TrimWord = Empty Else TrimWord = TrimWord & "..."
    End If
End Function

Private Sub EraseRegion()
    Dim hRgn As Long
    'Creates second region to fill with color.
    hRgn = CreateRoundRectRgn(0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, 0&, 0&)
    SetWindowRgn UserControl.hwnd, hRgn, True
    'delete our elliptical region
    DeleteObject hRgn
    UserControl.FillStyle = 0
End Sub

Private Sub DrawCorners(pencolor As Long)
    With UserControl
        'left top corner
         SetPixel .hDC, 0, 4, pencolor
         SetPixel .hDC, 4, 0, pencolor
         'left bottom corner
         SetPixel .hDC, .ScaleWidth - 5, 0, pencolor
         SetPixel .hDC, .ScaleWidth - 1, 4, pencolor
         'right top corner
         SetPixel .hDC, 0, .ScaleHeight - 5, pencolor
         SetPixel .hDC, 4, .ScaleHeight - 1, pencolor
         'right bottom corner
         SetPixel .hDC, .ScaleWidth - 5, .ScaleHeight - 1, pencolor
         SetPixel .hDC, .ScaleWidth - 1, .ScaleHeight - 5, pencolor
    End With
End Sub

Private Sub DrawAPIRoundRect(blnRounded As Boolean, LngRoundValue As Long, MyFillColor As Long, MyBorderColor As Long, R As RECT, Optional blnTransparent As Boolean = False)
    Dim m_roundedRadius As Long
    
    UserControl.FillColor = MyFillColor
    UserControl.ForeColor = MyBorderColor
    UserControl.FillStyle = IIf(blnTransparent = True, 1, 0)
    
    m_roundedRadius = IIf(blnRounded = False, 0&, LngRoundValue)
    
    RoundRect UserControl.hDC, R.Left, R.Top, R.Right, R.Bottom, m_roundedRadius, m_roundedRadius
    UserControl.FillStyle = 0
End Sub

Private Sub DrawGradientInRectangle(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, R As RECT, GradientType As jcGradConst, Optional blnDrawBorder As Boolean = False, Optional lBorderColor As Long = vbBlack, Optional LightCenter As Double = 2.01)
    Select Case GradientType
        Case VerticalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.Left, R.Top, R.Right + R.Left, R.Bottom, True
        Case HorizontalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.Left, R.Top, R.Right, R.Bottom + R.Top, False
        Case VCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, R, True, LightCenter
        Case HCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, R, False, LightCenter
    End Select
    If blnDrawBorder Then APIRectangle lhdcEx, R.Left, R.Top, R.Right, R.Bottom, lBorderColor
End Sub

Private Sub DrawGradCilinder(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, R As RECT, Optional ByVal blnVertical As Boolean = True, Optional ByVal LightCenter As Double = 2.01)
    If LightCenter <= 1# Then LightCenter = 1.01
    If blnVertical Then
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, R.Left, R.Top, R.Right + R.Left, R.Bottom / LightCenter, True
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.Left, R.Top + R.Bottom / LightCenter - 1, R.Right + R.Left, (LightCenter - 1) * R.Bottom / LightCenter + 1, True
    Else
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, R.Left, R.Top, R.Right / LightCenter, R.Bottom + R.Top, False
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, R.Left + R.Right / LightCenter - 1, R.Top, (LightCenter - 1) * R.Right / LightCenter + 1, R.Bottom + R.Top, False
    End If
End Sub

Private Sub FrameEnabled(blnValor As Boolean)
    Dim C As Control
    On Error Resume Next
    For Each C In UserControl.ContainedControls: C.Enabled = blnValor: Next
End Sub

Private Sub SetDisabledColor()
    m_FrameColorDis = TranslateColor(m_Border_Inactive)
    m_TextBoxColorDis = TranslateColor(m_BtnFace) '_Inactive)
    m_FillColorDis = TranslateColor(Ambient.BackColor)
    jcColorToDis = TranslateColor(m_BtnFace_Inactive)
    jcColorFromDis = TranslateColor(m_BtnFace_Inactive)
    jcColorBorderPicDis = TranslateColor(m_Border_Inactive)
End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)

    If DstW = 0 Or DstH = 0 Then Exit Sub
    
    Dim B As Long, H As Long, F As Long, I As Long, newW As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    Dim hOldOb As Long
    Dim SrcDC As Long, tObj As Long, ttt As Long

    SrcDC = CreateCompatibleDC(hDC)

    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim hBrush As Long
        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(TransColor) 'MaskColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
        DeleteObject hBrush
    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If
    useMask = True
    If Not useMask Then TransColor = -1

    newW = DstW - 1

    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            I = F + B
            If GetNearestColor(hDC, CLng(Data2(I).rgbRed) + 256& * Data2(I).rgbGreen + 65536 * Data2(I).rgbBlue) <> TransColor Then
                With Data1(I)
                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(I).rgbRed) + Data2(I).rgbGreen + Data2(I).rgbBlue) <= 384 Then Data1(I) = BrushRGB
                        Else
                            Data1(I) = BrushRGB
                        End If
                    Else
                        If isGreyscale Then
                            gCol = CLng(Data2(I).rgbRed * 0.3) + Data2(I).rgbGreen * 0.59 + Data2(I).rgbBlue * 0.11
                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                        Else
                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(I).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(I).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(I).rgbBlue * 2) \ 3
                            Else
                                Data1(I) = Data2(I)
                            End If
                        End If
                    End If
                End With
            End If
        Next B
    Next H

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC: DeleteDC Sr2DC
    DeleteObject tObj: DeleteDC SrcDC
End Sub
