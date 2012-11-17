VERSION 5.00
Object = "{AFFDD50D-733B-4E1C-8F98-E88F1ED6980D}#1.0#0"; "vbaListView6BC.ocx"
Begin VB.Form FrmDirSizeAnalyzer 
   Caption         =   "Directory Size Analyzer"
   ClientHeight    =   7260
   ClientLeft      =   11865
   ClientTop       =   3270
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   484
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   Begin VB.PictureBox PicProgress 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6390
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   6
      Top             =   3060
      Width           =   1545
   End
   Begin vbaBClListViewLib6.vbalListViewCtl lvwfiles 
      Height          =   4065
      Left            =   90
      TabIndex        =   5
      Top             =   540
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7170
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.PictureBox picTemp 
      Height          =   330
      Left            =   3060
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   4
      Top             =   225
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox PicsBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   426
      TabIndex        =   3
      Top             =   6930
      Width           =   6390
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   375
      Left            =   4815
      TabIndex        =   2
      Top             =   90
      Width           =   1140
   End
   Begin VB.ComboBox CboDirectory 
      Height          =   315
      Left            =   855
      TabIndex        =   1
      Top             =   90
      Width           =   3300
   End
   Begin VB.Label lblDir 
      AutoSize        =   -1  'True
      Caption         =   "Directory:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   675
   End
End
Attribute VB_Name = "FrmDirSizeAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Private Const GRADIENT_FILL_TRIANGLE = &H2&
Private Declare Function CreateSolidBrush Lib "gdi32" ( _
   ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
   ByVal hdc As Long, lpRect As RECT, _
   ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" ( _
   ByVal OLE_COLOR As Long, _
   ByVal HPALETTE As Long, _
   pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum
Private Declare Function FrameRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
    ) As Long
Private Declare Function DrawTextA Lib "user32" ( _
    ByVal hdc As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long
Private mAnalyzing As Boolean
Private ImlIcons As cVBALImageList


'Private type used to store data about each folder, referenced via itemdata of each listitem.
Private Type DirDataStruct
    DirObject As Directory
    DirTotalSize As Double 'total size of all data in folder.
    DirContentSize As Double 'size of data IN the folder, not counting subfolders.

End Type

Private mDirData() As DirDataStruct
Private mDirDataCount As Long
Private WithEvents mSbar As cNoStatusBar
Attribute mSbar.VB_VarHelpID = -1
Private WithEvents mTimer As CTimer
Attribute mTimer.VB_VarHelpID = -1



Private Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).X = tR.Left
   tTV(0).Y = tR.Top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).X = tR.Right
   tTV(1).Y = tR.Bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   End If
   
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub
Private Sub setTriVertexColorComponent( _
   ByRef iColor As Integer, _
   ByVal lComponent As Long _
   )
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub

Private Function TranslateColor( _
    ByVal oClr As OLE_COLOR, _
    Optional hPal As Long = 0 _
    ) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function




Private Sub cmdAnalyze_Click()
'griddirs.AddRow(
'RefreshListForFolder CboDirectory.Text
Dim tsize As Double, tt As Long
lvwfiles.ListItems.Clear
Set mTimer = New CTimer
mTimer.Interval = 50
mDirDataCount = 0
Erase mDirData
If Not bcfile.Exists(CboDirectory.Text) Then
    MsgBox "Directory not found"
    mTimer.Interval = 0
    Exit Sub
End If

AddDirectoryDataLvw CboDirectory.Text, tsize, Nothing


'now, loop through all items and populate the Totalsize percentages... (Subitem 2)

Dim loopItem As cListItem, I As Long
For I = 1 To lvwfiles.ListItems.count
    Set loopItem = lvwfiles.ListItems.Item(I)
    If loopItem.Tag > 0 Then
    loopItem.SubItems(2).caption = FormatPercent(CDbl(loopItem.Tag) / tsize)
    Else
    loopItem.SubItems(2).caption = "0% <empty>"
    End If


Next I
mTimer.Interval = 0
Set mTimer = Nothing



End Sub

Private Sub Form_Load()
'griddirs.AddRow


'tvwDirs.Columns.Add "SIZE", "Size(%)", , 64
'tvwDirs.Columns.Add "SIZEBYTE", "Size", , 64
lvwfiles.columns.Add , "NAME", "Name"
lvwfiles.columns.Add , "SIZE", "% of parent"
lvwfiles.columns.Add , "PCTTOTAL", "% of Total"
lvwfiles.columns.Add , "SIZEBYTE", "Size"
lvwfiles.columns.Add , "FILEFOLDER", "File/Subfolder"


Set ImlIcons = New cVBALImageList
ImlIcons.IconSizeX = 16
ImlIcons.IconSizeY = 16
ImlIcons.ColourDepth = ILC_COLOR32
ImlIcons.Create

lvwfiles.Imagelist(eLVSmallIcon) = CurrApp.SystemIML(Size_Small).himl
lvwfiles.Imagelist(eLVLargeIcon) = CurrApp.SystemIML(Size_Large).himl

'ImlIcons.AddFromResourceID 200, App.hInstance, IMAGE_ICON, "TVWMINUS"
'ImlIcons.AddFromResourceID 201, App.hInstance, IMAGE_ICON, "TVWPLUS"
ImlIcons.AddFromFile App.Path & "\Minustvw.ico", IMAGE_ICON, "TVWMINUS"
ImlIcons.AddFromFile App.Path & "\Plustvw.ico", IMAGE_ICON, "TVWPLUS"

Set mSbar = New cNoStatusBar
mSbar.Create PicsBar
'mSbar.SimpleMode = True
mSbar.AddPanel estbrStandard, "Ready.", , PicsBar.TextWidth("C:\windows\system32\drivers\etc\hosts_backuphosts_darkly"), , True, , "MESSAGE"
mSbar.AddPanel estbrNoBorders, "", , 64, True, False, , "PROGRESS"

'mSbar.SimpleText = "Ready."
mSbar.SizeGrip = True
mSbar.AllowXPStyles = True


'Set mProgress = New cProgressBar
'mProgress.XpStyle = True
'mProgress.Min = 0
'mProgress.Max = 100
'mProgress.Value = 0
'mProgress.ShowText = True
'Set mProgress.DrawObject = PicProgress
'lvwfiles.ImageList = ImlIcons
'Set GridDirs.OwnerDrawImpl = Me
End Sub

Private Sub AddDirectoryDataLvw(ByVal SFolderName, ByRef TotalSize As Double, Optional ParentItem As cListItem = Nothing, Optional ByRef ThisItem As cListItem = Nothing, Optional ByRef FileCount As Long, Optional ByRef DirCount As Long)
    Dim ThisFolderItem As cListItem
    Dim ThisFolder As Directory, gotitem As cListItem
    Dim ThisFolderFileSize As Double
    Dim LoopFolder As Directory
    Dim Loopsy As Directories
    Dim FolderNames() As String
    Dim FolderCount As Long, usekey As String
    Dim ItemsIterate As Collection, loopItem As cListItem
    Dim thisFolderIndex As Long
    Dim DirCountRunner As Long, FilecountRunner As Long
    
    'First, get a reference to the folder...
    mSbar.SimpleText = SFolderName
    If Right$(SFolderName, 1) <> "\" Then SFolderName = SFolderName & "\"
    Set ThisFolder = bcfile.GetDirectory(SFolderName)
    TotalSize = ThisFolder.size(False)
    'add a listitem for this folder...
    usekey = ThisFolder.Path
    DirCountRunner = ThisFolder.Directories.count
    FilecountRunner = ThisFolder.Files.count
    'If Right$(ThisFolder.Path, 1) = "\" Then usekey = Mid$(ThisFolder.Path, 1, Len(ThisFolder.Path) - 1)
    If Not ParentItem Is Nothing Then
        Set ThisFolderItem = lvwfiles.ListItems.Add(ParentItem.Index + 1, ThisFolder.Path, ThisFolder.Name, CurrApp.SystemIML(Size_Large).ItemIndex(usekey, True), CurrApp.SystemIML(Size_Small).ItemIndex(usekey, True))
        ThisFolderItem.Indent = ParentItem.Indent + 1
    Else
        Set ThisFolderItem = lvwfiles.ListItems.Add(, ThisFolder.Path, ThisFolder.Name, CurrApp.SystemIML(Size_Large).ItemIndex(usekey, True), CurrApp.SystemIML(Size_Small).ItemIndex(usekey, True))
        
    End If
    mDirDataCount = mDirDataCount + 1
    ReDim Preserve mDirData(1 To mDirDataCount)
    Set mDirData(mDirDataCount).DirObject = ThisFolder
    'totalsize is current total size- that is, only the content of the folder.
    mDirData(mDirDataCount).DirContentSize = TotalSize
    thisFolderIndex = mDirDataCount
    ThisFolderItem.ItemData = mDirDataCount
    
    Set Loopsy = ThisFolder.Directories
    If Loopsy.count > 0 Then
        ReDim FolderNames(1 To Loopsy.count)
        
        Dim folderiterator As CDirWalker
        'Set folderiterator = ThisFolder.GetWalker("*", FILE_ATTRIBUTE_DIRECTORY, 0)
        Set folderiterator = ThisFolder.Directories.GetWalker
        With ThisFolder.Directories.GetWalker
        mSbar.PanelText("MESSAGE") = ThisFolder.Name
        '.PanelMinWidth("MESSAGE") = PicsBar.TextWidth(ThisFolder.Name)
        
        DoEvents
        PicsBar.Refresh
        Do Until .GetNext(LoopFolder) Is Nothing

        
        
        'For Each LoopFolder In Loopsy
            FolderCount = FolderCount + 1
            
            FolderNames(FolderCount) = LoopFolder.Path
        Loop
        End With
        'Next LoopFolder
        Set ItemsIterate = New Collection
        For FolderCount = 1 To UBound(FolderNames)
            Dim tmpSize As Double, tmpfc As Long, tmpdc As Long
            AddDirectoryDataLvw FolderNames(FolderCount), tmpSize, ThisFolderItem, gotitem, tmpfc, tmpdc
            TotalSize = TotalSize + tmpSize
            FilecountRunner = FilecountRunner + tmpfc
            DirCountRunner = DirCountRunner + tmpdc
            mDirData(thisFolderIndex).DirTotalSize = TotalSize
            ItemsIterate.Add gotitem
 
        Next FolderCount
        
        
        'Now reiterate through the items, since we have the full size...
        For Each loopItem In ItemsIterate
            'tag is the size of that folder...
            If TotalSize > 0 Then
            loopItem.SubItems(1).caption = FormatPercent(CDbl(loopItem.Tag) / TotalSize, 2, 0)
            Else
                loopItem.SubItems(1).caption = "0% <empty>"
            End If
        
        Next
        
        
        
    End If
    
    DirCount = DirCountRunner
    FileCount = FilecountRunner
    
    ThisFolderItem.SubItems(3).caption = bcfile.FormatSize(TotalSize)
    ThisFolderItem.SubItems(4).caption = FileCount & "/" & DirCount
    ThisFolderItem.Tag = TotalSize
    
    
    Set ThisItem = ThisFolderItem

End Sub


Private Sub Form_Resize()
'
'lvwfiles.Move 0, CboDirectory.Top + CboDirectory.Height + 3, ScaleWidth, ScaleHeight - (CboDirectory.Top + CboDirectory.Height + 3) - PicsBar.Height - Abs(PicInfo.Height * PicInfo.Visible)
'
''mSbar.GetPanelRect(
'PicInfo.Move 0, lvwfiles.Height + lvwfiles.Top, ScaleWidth, ScaleHeight - (lvwfiles.Height + lvwfiles.Top)
Dim gotleft As Long, gotright As Long, gottop As Long, gotbottom As Long
'
'
Dim mleft As Long, MRight As Long, MTop As Long, MBottom As Long
Call mSbar.GetPanelRect("PROGRESS", gotleft, gottop, gotright, gotbottom)
'
'



lvwfiles.Move 0, CboDirectory.Top + CboDirectory.Height + 3, ScaleWidth, ScaleHeight - (CboDirectory.Top + CboDirectory.Height + 3) - PicsBar.Height
mleft = gotleft + PicsBar.Left + 3
MTop = PicsBar.Top + gottop + 3
MRight = PicsBar.Left + gotright - 3
MBottom = PicsBar.Top + gotbottom - 3
PicProgress.Move mleft, MTop, MRight - mleft, MBottom - MTop
PicProgress.Move PicsBar.Left + gotleft, PicsBar.Top + gottop, PicsBar.Left
End Sub

Private Sub lvwfiles_ItemClick(Item As vbaBClListViewLib6.cListItem)
ShowInfoForItem Item
End Sub

Private Sub lvwfiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim hitObj As cListItem
    Dim gotDir As Directory
    
    If Button And vbRightButton Then
    Set hitObj = lvwfiles.HitTest(X, Y)
    If Not hitObj Is Nothing Then
        If hitObj.ItemData <> 0 Then
            Set gotDir = mDirData(hitObj.ItemData).DirObject
            gotDir.ShowExplorerMenu Me.hWnd
        End If
    End If

    End If
End Sub

Private Sub mSbar_OwnerDraw(ByVal hdc As Long, iLeft As Long, iTop As Long, iRight As Long, iBottom As Long, bDoDefault As Boolean)
    '
End Sub

Private Sub mTimer_ThatTime()
 '    mProgress.Draw
    Static MprogressStat As Long
    Dim rectDraw As RECT
  
    Dim mleft As Double
    mleft = (MprogressStat / 100) * PicProgress.ScaleWidth
    MprogressStat = MprogressStat + 5
    'PicProgress.FillColor = vbBlue
    'PicProgress.FillStyle = vbSolid
    'PicProgress.Line (mleft, 0)-(mleft + PicProgress.ScaleWidth / 8, PicProgress.ScaleHeight), , BF
    rectDraw.Left = mleft
    rectDraw.Top = 0
    rectDraw.Right = mleft - PicProgress.ScaleWidth / 6
    rectDraw.Bottom = PicProgress.ScaleHeight
    PicProgress.Cls
    GradientFillRect PicProgress.hdc, rectDraw, StartGradient, EndGradient, GRADIENT_FILL_RECT_H
    If MprogressStat >= 100 Then MprogressStat = 0
    PicProgress.Refresh
    
End Sub
Private Property Get StartGradient() As Long

On Error Resume Next
Static Ret
If IsEmpty(Ret) Then
    Ret = Val(CurrApp.Settings.ReadProfileSetting("SizeAnalyzer", "StartGradient"))
    If Err <> 0 Then
        Ret = vbRed
    End If
    
End If

StartGradient = CLng(Ret)
End Property
Private Property Get EndGradient() As Long
    On Error Resume Next
    Static Ret
    If IsEmpty(Ret) Then
        Ret = Val(CurrApp.Settings.ReadProfileSetting("SizeAnalyzer", "EndGradient"))
        If Err <> 0 Then
            EndGradient = vbBlue
        
        End If
    End If
    EndGradient = Ret


End Property
'Private Sub AddDirectoryDataTvw(ByVal sFolderName As String, ByRef TotalSize As Double, Optional ParentNode As vbalCTreeViewLib6.cCTreeViewNode = Nothing)
''
'Dim ThisFolderNode As cCTreeViewNode
'Dim ThisFolder As Directory
'Dim ThisFolderFileSize As Double
'Dim LoopFolder As Directory
'Dim Loopsy As Directories
'
'Dim FolderNames() As String
'
'Set ThisFolder = bcfile.GetDirectory(sFolderName)
'
''Add an Item for this folder...
'Set ThisFolderNode = tvwDirs.Nodes.Add(ParentNode, etvwChild, ThisFolder.Path, ThisFolder.Name)
'
''we have Name, Size, and SizeByte....
'
''First, we only know the name...
''Size and SizeByte we need to get after adding child nodes.
''Sooo... iterate through each Directory....
'Set Loopsy = ThisFolder.Directories
'
'Dim foldercount As Long
'
'For Each LoopFolder In Loopsy
'    foldercount = foldercount + 1
'    ReDim Preserve FolderNames(1 To foldercount)
'    FolderNames(foldercount) = LoopFolder.Path
'
'Next LoopFolder
'
'If foldercount > 0 Then
'    'OK, NOW... re-iterate through the array...
'    Dim tempsize As Double
'    Dim SizeRunner As Double
'    For foldercount = 1 To UBound(FolderNames)
'        AddDirectoryData FolderNames(foldercount), tempsize, ThisFolderNode
'        SizeRunner = SizeRunner + tempsize
'
'
'    Next
'
'    'add file sizes in this folder...
'End If
'SizeRunner = SizeRunner + ThisFolder.Size(False)
'
''tempsize is total size of thisfolder...
'
'ThisFolderNode.SubItem(1).Text = "Size Percentage..."
'ThisFolderNode.SubItem(2).Text = bcfile.FormatSize(SizeRunner, True)
'ThisFolderNode.Visible = True
'ThisFolderNode.NoCheckBox = True
'
'ThisFolderNode.Expanded = True
'
'
'End Sub

'Private Sub RefreshListForFolder(ByVal FolderName As String)
'    Dim LoopDir As Directory
'    Dim CurrFolder As Long
'    Dim ParentDir As Directory
'    Dim FolderNames() As String, FolderSizes() As Double
'    Dim dirslook As Directories
'    Dim TotalBytes As Double
'    'first, retrieve folder names and folder sizes. We then add them all up (rather then calling the parent folder's size method).
'    Set ParentDir = GetDirectory(FolderName)
'    Set dirslook = ParentDir.Directories
'    ReDim FolderNames(1 To dirslook.Count)
'    ReDim FolderSizes(1 To dirslook.Count)
'
'    For Each LoopDir In dirslook
'        If Trim$(LoopDir.Name) <> "" Then
'            CurrFolder = CurrFolder + 1
'            'ReDim Preserve FolderNames(1 To CurrFolder)
'            'ReDim Preserve FolderSizes(1 To CurrFolder)
'            FolderNames(CurrFolder) = LoopDir.Name
'            FolderSizes(CurrFolder) = LoopDir.Size(True)
'            TotalBytes = TotalBytes + FolderSizes(CurrFolder)
'        End If
'    Next LoopDir
'
'
'
'
'    'done... whew.
'    Set ParentDir = Nothing
'    GridDirs.Clear
'    GridDirs.Redraw = False
'    GridDirs.Editable = False
'    For CurrFolder = 1 To UBound(FolderNames)
'        Dim percentage As Single
'        Dim ByteStr As String
'        percentage = FolderSizes(CurrFolder) / TotalBytes
'        ByteStr = bcfile.FormatSize(FolderSizes(CurrFolder), True)
'        GridDirs.AllowGrouping = True
'
'
'        GridDirs.AddRow , , True, , 1
'        GridDirs.cell(CurrFolder, 1).Text = FolderNames(CurrFolder)
'        GridDirs.cell(CurrFolder, 2).Text = percentage
'        GridDirs.cell(CurrFolder, 3).Text = ByteStr
'
'
'
'
'
'    Next CurrFolder
'    GridDirs.Redraw = True
'
'
'    'doesn't remove cols... nice...
'
'End Sub


Private Sub PicsBar_Click()

End Sub

Private Sub PicsBar_Paint()
    mSbar.Draw
End Sub
Private Sub ShowInfoForItem(ItemShow As cListItem)

'step One: loop from this item's index forwards-save each item that is one indent level higher.
Dim CurrIndex As Long
Dim CurrCount As Long
Dim Cached() As cListItem
Dim currItem As cListItem
CurrIndex = ItemShow.Index + 1
Do
    Set currItem = lvwfiles.ListItems(CurrIndex)
    If currItem.Indent = ItemShow.Indent + 1 Then
        'add it...
        CurrCount = CurrCount + 1
        ReDim Preserve Cached(1 To CurrCount)
        Set Cached(CurrCount) = currItem
    End If
    CurrIndex = CurrIndex + 1
Loop Until currItem.Indent <= ItemShow.Indent Or CurrIndex > lvwfiles.ListItems.count

'lstInfo.Clear
'lstInfo.AddItem "Content size:" & mDirData(ItemShow.ItemData).DirContentSize
Dim I As Long
For I = 1 To CurrCount
'lstInfo.AddItem mDirData(Cached(I).ItemData).DirObject.Name & FormatPercent(mDirData(Cached(I).ItemData).DirTotalSize / mDirData(ItemShow.ItemData).DirTotalSize)
'lstInfo.AddItem mDirData(Cached(i).ItemData).DirObject.Name
Next I



End Sub
