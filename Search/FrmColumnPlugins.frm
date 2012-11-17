VERSION 5.00
Object = "{AFFDD50D-733B-4E1C-8F98-E88F1ED6980D}#1.0#0"; "vbaListView6BC.ocx"
Begin VB.Form FrmColumnPlugins 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Column Plugins"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin vbaBClListViewLib6.vbalListViewCtl LvwColumns 
      Height          =   3435
      Left            =   90
      TabIndex        =   3
      Top             =   810
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   6059
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
      CheckBoxes      =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4320
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2835
      TabIndex        =   1
      Top             =   4320
      Width           =   1050
   End
   Begin VB.PictureBox PicNoPlugins 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   780
      Width           =   4875
   End
   Begin VB.Label lblColumnPlugin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmColumnPlugins.frx":0000
      Height          =   585
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   3975
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmColumnPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()


Dim loopItem As cListItem
Dim currItem As Long
For currItem = 1 To LvwColumns.ListItems.Count
    Set loopItem = LvwColumns.ListItems.Item(currItem)
    CurrApp.Settings.WriteProfileSetting "Column Plugins", loopItem.Tag, IIf(loopItem.checked, -1, 0)

Next currItem


End Sub

Private Sub Form_Load()
'reload ALL progIDs from the INI file.
Dim Values() As String, vcount As Long
Dim CreateObj As Object
Dim AttemptCast As IColumnPlugin
Dim UseProgid As String
Dim I As Long
Dim progvalue As String
Dim newitem As cListItem
Const NoPluginString As String = "No Plugins to load in INI file"
PicNoPlugins.Move LvwColumns.Left, LvwColumns.Top, LvwColumns.Width, LvwColumns.Height
'step one: add columns.
With LvwColumns.columns
    .Add , "NAME", "Name"
    .Add , "DESC", "Description"
    .Add , "COLTEXT", "Columns"

End With
On Error Resume Next

Call CurrApp.Settings.EnumerateValues(Setting_System, "Column Plugins", Values(), vcount)

'enumerate each one; get the ProgID, if it has the setting set to 1 ,check it, otherwise leave the item unchecked.
If vcount = 0 Then
    'no plugins to load.
    PicNoPlugins.ZOrder vbBringToFront
    PicNoPlugins.Cls
    PicNoPlugins.FontSize = PicNoPlugins.FontSize * 2
    PicNoPlugins.FontBold = True
    PicNoPlugins.CurrentX = (PicNoPlugins.ScaleWidth - PicNoPlugins.TextWidth(NoPluginString)) / 2
    PicNoPlugins.CurrentY = (PicNoPlugins.ScaleHeight - PicNoPlugins.TextHeight(NoPluginString)) / 2
End If
For I = 1 To vcount
    UseProgid = Values(I)
    progvalue = CurrApp.Settings.ReadProfileSetting("Column Plugins", UseProgid)
    Set newitem = LvwColumns.ListItems.Add(, UseProgid)
    
    On Error Resume Next
    Set CreateObj = CreateObject(UseProgid)
    If Err = 0 Then
    'attempt to cast to our interface...
        Set AttemptCast = CreateObj
    
    End If
    If Err <> 0 Then
        newitem.checked = False
        newitem.Text = "Failed to Create """ & progvalue & """:" & Err.Description & "(#" & Err.Number & ")"
        newitem.ForeColor = vbRed
        newitem.BackColor = RGB(235, 235, 255)  'light yellow
        
        Err.Clear
        
        
    Else
        'no Error, CreateObject was Successful, as was the casting...
        Dim colloop As Long
        Dim coldata As ColumnInfo
        Dim buildstr As String
        newitem.checked = CBool(Val(progvalue))
        newitem.Tag = UseProgid
        newitem.Text = AttemptCast.Name
        newitem.SubItems(1).caption = AttemptCast.Description
        For colloop = 1 To AttemptCast.GetColumnCount
            coldata = AttemptCast.GetColumnInfo(colloop)
            buildstr = """" & coldata.ColumnTitle & """"
            If colloop < AttemptCast.GetColumnCount Then
                buildstr = buildstr & ","
            End If
        
        
        Next
        newitem.SubItems(2).caption = buildstr
    
    End If
    
    
    



Next I





End Sub

Private Sub lblDe_Click()

End Sub

