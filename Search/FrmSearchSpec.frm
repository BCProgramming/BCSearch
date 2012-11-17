VERSION 5.00
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmSearchSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Specifications Editor"
   ClientHeight    =   6285
   ClientLeft      =   3750
   ClientTop       =   3120
   ClientWidth     =   7605
   Icon            =   "FrmSearchSpec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   28
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5220
      TabIndex        =   29
      Top             =   5820
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6420
      TabIndex        =   30
      Top             =   5820
      Width           =   1155
   End
   Begin vbalTabStrip6.TabControl TabSpecifications 
      Height          =   5595
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   9869
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox PicTabs 
         Height          =   4995
         Index           =   3
         Left            =   8040
         ScaleHeight     =   4935
         ScaleWidth      =   7215
         TabIndex        =   34
         Top             =   720
         Width           =   7275
         Begin VB.TextBox txtLanguage 
            Height          =   285
            Left            =   900
            TabIndex        =   36
            Text            =   "VBScript"
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtSourceCode 
            BeginProperty Font 
               Name            =   "Consolas"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1440
            Width           =   7035
         End
         Begin VB.Label lblScript 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scripting"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   60
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Language:"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   660
            Width           =   765
         End
         Begin VB.Label lblCode 
            BackStyle       =   0  'Transparent
            Caption         =   "Code:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   855
         End
      End
      Begin VB.PictureBox PicTabs 
         Height          =   4995
         Index           =   0
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   7215
         TabIndex        =   1
         Top             =   540
         Width           =   7275
         Begin VB.CheckBox chkregexp 
            Caption         =   "Regular Expression"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "When set, the filter will be matched against filenames by being treated as a regular expression."
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton cmdChangeDate 
            Caption         =   "&Date Specifications..."
            Height          =   435
            Left            =   3900
            TabIndex        =   18
            Top             =   60
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            Caption         =   "&Size"
            Height          =   1095
            Left            =   3720
            TabIndex        =   19
            Top             =   2160
            Width           =   3255
            Begin MSComCtl2.UpDown UDCSize 
               Height          =   315
               Index           =   0
               Left            =   1860
               TabIndex        =   21
               Top             =   240
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   556
               _Version        =   393216
               BuddyControl    =   "txtSize(0)"
               BuddyDispid     =   196622
               BuddyIndex      =   0
               OrigLeft        =   1680
               OrigTop         =   1620
               OrigRight       =   1935
               OrigBottom      =   2175
               Max             =   16777216
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.ComboBox cboSizesuffix 
               Height          =   315
               Index           =   1
               Left            =   2340
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Tag             =   "NoPersist"
               Top             =   660
               Width           =   735
            End
            Begin VB.ComboBox cboSizesuffix 
               Height          =   315
               Index           =   0
               Left            =   2340
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Tag             =   "NoPersist"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtSize 
               Height          =   315
               Index           =   1
               Left            =   1080
               TabIndex        =   24
               Text            =   "0"
               Top             =   660
               Width           =   780
            End
            Begin VB.TextBox txtSize 
               Height          =   315
               Index           =   0
               Left            =   1080
               TabIndex        =   22
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
            Begin MSComCtl2.UpDown UDCSize 
               Height          =   315
               Index           =   1
               Left            =   1860
               TabIndex        =   26
               Top             =   660
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   556
               _Version        =   393216
               BuddyControl    =   "txtSize(1)"
               BuddyDispid     =   196622
               BuddyIndex      =   1
               OrigLeft        =   1680
               OrigTop         =   1620
               OrigRight       =   1935
               OrigBottom      =   2175
               Max             =   16777216
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Smaller than:"
               Height          =   195
               Left            =   180
               TabIndex        =   25
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Larger than:"
               Height          =   195
               Left            =   180
               TabIndex        =   20
               Top             =   300
               Width           =   855
            End
         End
         Begin VB.ComboBox cboOperation 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "nopersist"
            Top             =   60
            Width           =   2715
         End
         Begin VB.PictureBox PicAttributes 
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   60
            ScaleHeight     =   2175
            ScaleWidth      =   2475
            TabIndex        =   9
            Top             =   1860
            Width           =   2475
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "Directory"
               Height          =   195
               Index           =   6
               Left            =   1140
               TabIndex        =   57
               Top             =   1485
               Width           =   1140
            End
            Begin VB.CheckBox chkAttributeExact 
               Caption         =   "&Match Attributes Exactly"
               Height          =   195
               Left            =   60
               TabIndex        =   10
               Top             =   180
               Width           =   2175
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&Encrypted"
               Height          =   375
               Index           =   5
               Left            =   1140
               TabIndex        =   17
               Top             =   1080
               Width           =   1275
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&Compressed"
               Height          =   375
               Index           =   4
               Left            =   1140
               TabIndex        =   16
               Top             =   720
               Width           =   1275
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&Read-Only"
               Height          =   375
               Index           =   0
               Left            =   60
               TabIndex        =   12
               Top             =   720
               Width           =   1095
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&System"
               Height          =   375
               Index           =   3
               Left            =   60
               TabIndex        =   15
               Top             =   1800
               Width           =   915
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&Hidden"
               Height          =   375
               Index           =   2
               Left            =   60
               TabIndex        =   14
               Top             =   1440
               Width           =   915
            End
            Begin VB.CheckBox ChkAttributes 
               Caption         =   "&Archive"
               Height          =   315
               Index           =   1
               Left            =   60
               TabIndex        =   13
               Top             =   1140
               Width           =   915
            End
            Begin VB.Label lblAttributes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Attributes:"
               Height          =   195
               Left            =   60
               TabIndex        =   11
               Top             =   480
               Width           =   705
            End
         End
         Begin VB.TextBox txtFilterName 
            Height          =   315
            Left            =   960
            TabIndex        =   4
            Top             =   660
            Width           =   2595
         End
         Begin VB.ComboBox cboFilter 
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Text            =   "*.*"
            ToolTipText     =   "File Specification Filter. Leave blank if you don't want to filter on a specification."
            Top             =   1080
            Width           =   2595
         End
         Begin VB.Label lblDateInfo 
            Height          =   1635
            Left            =   3900
            TabIndex        =   8
            Top             =   540
            Width           =   3015
         End
         Begin VB.Label lblOperation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Operation:"
            Height          =   195
            Left            =   60
            TabIndex        =   2
            Top             =   120
            Width           =   735
         End
         Begin VB.Label lblFilterName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Filter Name:"
            Height          =   195
            Left            =   60
            TabIndex        =   5
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblFilter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Filter(s):"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1140
            Width           =   540
         End
      End
      Begin VB.PictureBox PicTabs 
         Height          =   4995
         Index           =   1
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   7215
         TabIndex        =   32
         Top             =   540
         Width           =   7275
         Begin VB.PictureBox picRegexp 
            Height          =   4815
            Left            =   5460
            ScaleHeight     =   4755
            ScaleWidth      =   1635
            TabIndex        =   42
            Top             =   60
            Width           =   1695
            Begin VB.TextBox txtMinMatches 
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Top             =   480
               Width           =   825
            End
            Begin MSComCtl2.UpDown UDCminmatch 
               Height          =   315
               Left            =   960
               TabIndex        =   45
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   556
               _Version        =   393216
               BuddyControl    =   "txtMinMatches"
               BuddyDispid     =   196637
               OrigLeft        =   960
               OrigTop         =   480
               OrigRight       =   1215
               OrigBottom      =   1275
               Max             =   9999999
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min. MatchCount:"
               Height          =   195
               Left            =   60
               TabIndex        =   43
               Top             =   120
               Width           =   1260
            End
         End
         Begin VB.CheckBox chkContainsRegExp 
            Caption         =   "Regular Expression"
            Height          =   315
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   1755
         End
         Begin VB.TextBox txtContains 
            Height          =   4155
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   720
            Width           =   5235
         End
      End
      Begin VB.PictureBox PicTabs 
         Height          =   5040
         Index           =   4
         Left            =   135
         ScaleHeight     =   4980
         ScaleWidth      =   7290
         TabIndex        =   46
         Top             =   495
         Width           =   7350
         Begin VB.PictureBox picbackground 
            BackColor       =   &H80000005&
            Height          =   330
            Left            =   1215
            ScaleHeight     =   270
            ScaleWidth      =   675
            TabIndex        =   55
            Tag             =   "NoChangeBG"
            Top             =   4185
            Width           =   735
         End
         Begin VB.PictureBox PicForeground 
            BackColor       =   &H80000008&
            Height          =   330
            Left            =   1215
            ScaleHeight     =   270
            ScaleWidth      =   675
            TabIndex        =   54
            Tag             =   "NoChangeBG"
            Top             =   3645
            Width           =   735
         End
         Begin VB.CommandButton CmdChange 
            Caption         =   "&Change..."
            Height          =   420
            Left            =   675
            TabIndex        =   50
            Top             =   990
            Width           =   1230
         End
         Begin VB.Label lblsample 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sample Font"
            Height          =   195
            Left            =   2475
            TabIndex        =   56
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Background:"
            Height          =   195
            Left            =   225
            TabIndex        =   53
            Top             =   4230
            Width           =   915
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Foreground:"
            Height          =   195
            Left            =   270
            TabIndex        =   52
            Top             =   3690
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sample:"
            Height          =   195
            Left            =   2475
            TabIndex        =   51
            Top             =   1170
            Width           =   570
         End
         Begin VB.Label lblFontSettings 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FontSpec"
            Height          =   1815
            Left            =   180
            TabIndex        =   49
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblfont 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Font:"
            Height          =   195
            Left            =   180
            TabIndex        =   48
            Top             =   1125
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   $"FrmSearchSpec.frx":000C
            Height          =   585
            Left            =   135
            TabIndex        =   47
            Top             =   135
            Width           =   3525
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox PicTabs 
         Height          =   4455
         Index           =   2
         Left            =   225
         ScaleHeight     =   4395
         ScaleWidth      =   7245
         TabIndex        =   33
         Top             =   540
         Width           =   7305
         Begin VB.ComboBox cboADSThanMultiplier 
            Height          =   315
            Index           =   1
            Left            =   4815
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Tag             =   "noPersist"
            Top             =   1980
            Width           =   1050
         End
         Begin VB.ComboBox cboADSThanMultiplier 
            Height          =   315
            Index           =   0
            Left            =   4815
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Tag             =   "noPersist"
            Top             =   1575
            Width           =   1050
         End
         Begin MSComCtl2.UpDown UDCSmallerThan 
            Height          =   285
            Left            =   4530
            TabIndex        =   75
            Top             =   1980
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtADSSmallerThan"
            BuddyDispid     =   196652
            OrigLeft        =   5310
            OrigTop         =   990
            OrigRight       =   5565
            OrigBottom      =   1275
            Max             =   1024
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDCADSLargerThan 
            Height          =   285
            Left            =   4530
            TabIndex        =   74
            Top             =   1575
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtADSLargerThan"
            BuddyDispid     =   196653
            OrigLeft        =   5040
            OrigTop         =   540
            OrigRight       =   5295
            OrigBottom      =   825
            Max             =   1024
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtADSSmallerThan 
            Height          =   285
            Left            =   3870
            TabIndex        =   73
            Text            =   "0"
            Top             =   1980
            Width           =   660
         End
         Begin VB.TextBox txtADSLargerThan 
            Height          =   285
            Left            =   3870
            TabIndex        =   72
            Text            =   "0"
            Top             =   1575
            Width           =   660
         End
         Begin VB.CheckBox chkadsContainsisregexp 
            Caption         =   "Use Regular Expression"
            Height          =   285
            Left            =   225
            TabIndex        =   69
            Top             =   4095
            Width           =   3435
         End
         Begin VB.TextBox txtADScontaining 
            Height          =   1545
            Left            =   180
            TabIndex        =   68
            Top             =   2520
            Width           =   4785
         End
         Begin VB.CheckBox chkADSregexp 
            Caption         =   "Use Regular Expression"
            Height          =   285
            Left            =   225
            TabIndex        =   66
            Top             =   1890
            Width           =   1995
         End
         Begin VB.TextBox txtadsspec 
            Height          =   285
            Left            =   1395
            TabIndex        =   65
            Text            =   "*"
            Top             =   1530
            Width           =   780
         End
         Begin VB.TextBox txtmax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1395
            TabIndex        =   62
            Text            =   "0"
            Top             =   945
            Width           =   555
         End
         Begin VB.TextBox txtmin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1395
            TabIndex        =   59
            Text            =   "0"
            Top             =   585
            Width           =   555
         End
         Begin MSComCtl2.UpDown UDCMinCount 
            Height          =   285
            Left            =   1935
            TabIndex        =   60
            Top             =   585
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtmin"
            BuddyDispid     =   196659
            OrigLeft        =   2655
            OrigTop         =   675
            OrigRight       =   2910
            OrigBottom      =   1455
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDCmax 
            Height          =   285
            Left            =   1935
            TabIndex        =   63
            Top             =   945
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtmax"
            BuddyDispid     =   196658
            OrigLeft        =   2655
            OrigTop         =   675
            OrigRight       =   2910
            OrigBottom      =   1455
            Max             =   99
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Smaller Than:"
            Height          =   195
            Left            =   2880
            TabIndex        =   71
            Top             =   1980
            Width           =   975
         End
         Begin VB.Label lbladsl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Larger Than:"
            Height          =   195
            Left            =   2925
            TabIndex        =   70
            Top             =   1620
            Width           =   915
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Containing:"
            Height          =   195
            Left            =   225
            TabIndex        =   67
            Top             =   2295
            Width           =   795
         End
         Begin VB.Label lblSpec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specification:"
            Height          =   195
            Left            =   225
            TabIndex        =   64
            Top             =   1575
            Width           =   960
         End
         Begin VB.Label lblCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum Count:"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   61
            Top             =   990
            Width           =   1170
         End
         Begin VB.Label lblCount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minimum Count:"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   58
            Top             =   630
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "FrmSearchSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





'feature idea: some way to automatically copy items to specific folders, for example, all files to a folder with the same name as their extension, or their type, or something along those lines, or by date, etc.













'For Edge drawing...

Public Enum EDEDBorderStyle
    BDR_RAISEDOUTER = 1
    BDR_SUNKENOUTER = 2
    BDR_RAISEDINNER = 4
    BDR_SUNKENINNER = 8
    
    BDR_BUTTON = BDR_RAISEDINNER Or BDR_RAISEDOUTER
    BDR_CONTROL = BDR_SUNKENINNER Or BDR_SUNKENOUTER
    BDR_THINBUTTON = BDR_RAISEDOUTER
    BDR_THINCONTROL = BDR_SUNKENOUTER
    
    BDR_ETCHRAISE = BDR_RAISEDOUTER Or BDR_SUNKENINNER
    BDR_ETCHINSET = BDR_SUNKENOUTER Or BDR_RAISEDINNER
    
    BDR_ALL = BDR_BUTTON Or BDR_CONTROL
End Enum
Public Enum EDEDBorderParts
    BF_LEFT = 1
    BF_TOP = 2
    BF_RIGHT = 4
    BF_BOTTOM = 8
    BF_TOPLEFT = BF_LEFT Or BF_TOP
    Bf_BOTTOMRIGHT = BF_RIGHT Or BF_BOTTOM
    BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM
    BF_MIDDLE = &H800
    BF_SOFT = &H1000
    BF_ADJUST = &H2000
    BF_FLAT = &H4000
    BF_MONO = &H8000&
    BF_ALL = BF_RECT Or BF_MIDDLE Or BF_SOFT Or BF_ADJUST Or BF_FLAT Or BF_MONO
End Enum
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private mOwner As FrmSearch
Private mremovefrom As Object
Private mRemovekey As Variant
Private mdeleteOnCancel As Boolean
Private mFilterautocomplete As CAutoCompleteCombo
Private mTboxHandlers As Collection
Private WithEvents Dateform As FrmDateEdit
Attribute Dateform.VB_VarHelpID = -1
Private mFilterEdit As CSearchFilter
Private mFormSaver As CFormDataSaver
'Private TT As clsTooltip
Public Property Get SaveControlStates()
    SaveControlStates = False
End Property

'Moved from Change event to "Validate" event.
Private Sub cboFilter_Validate(Cancel As Boolean)
Dim grabExtension As String
'if it starts with *., then try to load the association data for the given extension...
If Left$(cboFilter.Text, 2) = "*." Then
    grabExtension = Mid$(cboFilter.Text, 3)
    
    
    Dim gottype As String
    gottype = bcfile.GetFileTypeFromExtension(grabExtension)
    If gottype <> "" And txtFilterName.Text = "" Then
        txtFilterName.Text = gottype
    
    End If
    

End If

End Sub

Private Sub CmdChange_Click()
    Dim currfont As StdFont
    Dim newfont As StdFont
    Dim sappearance As CExtraFilterData
    Set sappearance = mFilterEdit.Tag
    Set currfont = lblsample.Font
    Set newfont = SelectFont(currfont, Me.hWnd, CF_BOTH)
    'Set lblsample.Font = newfont
    Set sappearance.Font = newfont
    ReloadAppearance
End Sub

Private Sub Form_Initialize()
    Set mFormSaver = New CFormDataSaver
    mFormSaver.Initialize Me, CurrApp.Settings


End Sub

Public Sub EditFilter(EditThis As CSearchFilter, OwnerForm As Form, Optional ByVal mFormCaption As String = "Search Specifications Editor", Optional DeleteOnCancel As Boolean = False)
    Set mFilterEdit = EditThis
    ReloadFilterData
    mdeleteOnCancel = DeleteOnCancel
    Me.caption = mFormCaption
    If TypeOf OwnerForm Is FrmSearch Then
        Set mOwner = OwnerForm
    End If
    Me.Show , OwnerForm
End Sub


Public Sub ReloadFilterData()

    txtFilterName.Text = mFilterEdit.Name
    cboFilter.Text = mFilterEdit.FileSpec
    With mFilterEdit
        ChkAttributes(0).Value = Abs((.Attributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY)
        ChkAttributes(1).Value = Abs((.Attributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE)
        ChkAttributes(2).Value = Abs((.Attributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN)
        ChkAttributes(3).Value = Abs((.Attributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM)
        ChkAttributes(4).Value = Abs((.Attributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED)
        ChkAttributes(5).Value = Abs((.Attributes And FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_ENCRYPTED)
        ChkAttributes(6).Value = Abs((.Attributes And FILE_ATTRIBUTE_ENCRYPTED) = FILE_ATTRIBUTE_DIRECTORY)
        chkAttributeExact.Value = Abs(.AttributesMatchExact)
        cboOperation.ListIndex = .SearchOperation
        chkregexp.Value = Abs((mFilterEdit.FileSpecIsRegExp))
        txtContains.Text = mFilterEdit.ContainsStr
        chkContainsRegExp.Value = Abs(mFilterEdit.ContainsIsRegExp)
        txtMinMatches.Text = mFilterEdit.ContainsRegExpMinmatches
        
        
        With .searchstruct.AlternateStreamSpecs
        txtmin.Text = .mincount
        txtmax.Text = .maxcount
        
        chkadsContainsisregexp.Value = IIf(.NameSpecisRegExp, vbChecked, vbUnchecked)
        txtADScontaining.Text = .ContainsStr
        chkadsContainsisregexp.Value = IIf(.ContainsIsRegExp, vbChecked, vbUnchecked)
        txtadsspec.Text = .nameSpec
        '
        End With
        
    End With
    
    
    
    
    txtLanguage.Text = mFilterEdit.ScriptLanguage
    'txtScriptFile.Text = ""
    txtSourceCode.Text = mFilterEdit.ScriptCode
    
    
    
    
    
    
    
    ReloadSizeData
    ReloadAppearance
    refreshDateCaption
End Sub
Private Sub ReloadAppearance()
    Dim appearancedata As CExtraFilterData
    Set appearancedata = mFilterEdit.Tag
    With appearancedata
    
        lblFontSettings.caption = FontToString(.Font)
        PicForeground.BackColor = .ForeColor
        picbackground.BackColor = .BackColor
        'fontsettings get's the font's string info,clear and fill lvwsample with sample text
        Set lblsample.Font = .Font
    End With

    
    
    
End Sub
Public Sub SaveAppearance()
Dim appearancedata As CExtraFilterData
Set appearancedata = mFilterEdit.Tag



End Sub
Private Sub ReloadSizeData()
    'load size data from mfilteredit.
    
    Dim SizeValues(3) As Double
    Dim SizeIndex(3) As Integer, I As Long
    SizeValues(0) = mFilterEdit.SizeLargerThan
    SizeValues(1) = mFilterEdit.SizeSmallerThan
    
    SizeValues(2) = mFilterEdit.searchstruct.AlternateStreamSpecs.SizeLargerThan
    SizeValues(3) = mFilterEdit.searchstruct.AlternateStreamSpecs.SizeSmallerThan
    For I = 0 To 3
        Do Until SizeValues(I) < 1024
            If (SizeValues(I) \ 1024) <> (SizeValues(I) \ 1024) Then
                Exit Do
            Else
                SizeValues(I) = SizeValues(I) \ 1024
        
            End If
            SizeIndex(I) = SizeIndex(I) + 1
        
        Loop
    Next
    cboSizesuffix(0).ListIndex = SizeIndex(0)
    cboSizesuffix(1).ListIndex = SizeIndex(1)
    
    txtSize(0).Text = SizeValues(0)
    txtSize(1).Text = SizeValues(1)
    
    
    txtADSLargerThan.Text = SizeValues(2)
    txtADSSmallerThan.Text = SizeValues(3)
    
    cboADSThanMultiplier(0).ListIndex = SizeIndex(2)
    cboADSThanMultiplier(1).ListIndex = SizeIndex(3)



End Sub
Private Function GetAttributeMaskValue() As FileAttributeConstants
'readonly,archive,hidden,system
    Dim runner As FileAttributeConstants
    If ChkAttributes(0).Value = vbChecked Then
        runner = FILE_ATTRIBUTE_READONLY
    End If
    If ChkAttributes(1).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_ARCHIVE
    End If
    
    If ChkAttributes(2).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_HIDDEN
    End If
    
    If ChkAttributes(3).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_SYSTEM
    End If
    If ChkAttributes(4).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_COMPRESSED
    End If
    If ChkAttributes(5).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_ENCRYPTED
    End If
    If ChkAttributes(6).Value = vbChecked Then
        runner = runner + FILE_ATTRIBUTE_DIRECTORY
    End If
    GetAttributeMaskValue = runner


End Function
Private Sub CmdApply_Click()

    mFilterEdit.FileSpec = cboFilter.Text
    mFilterEdit.Name = txtFilterName.Text
    mFilterEdit.AttributesMatchExact = chkAttributeExact.Value = vbChecked
    mFilterEdit.Attributes = GetAttributeMaskValue()
    mFilterEdit.SearchOperation = cboOperation.ListIndex
    mFilterEdit.SizeLargerThan = Val(txtSize(0).Text) * ((1024 * Abs(CBool(cboSizesuffix(0).ListIndex))) ^ (cboSizesuffix(0).ListIndex))
    mFilterEdit.SizeSmallerThan = Val(txtSize(1).Text) * ((1024 * Abs(CBool(cboSizesuffix(1).ListIndex))) ^ (cboSizesuffix(1).ListIndex))
    mFilterEdit.ScriptCode = txtSourceCode.Text
    mFilterEdit.ScriptLanguage = txtLanguage.Text
    mFilterEdit.FileSpecIsRegExp = chkregexp.Value = vbChecked
    mFilterEdit.ContainsStr = txtContains.Text
    mFilterEdit.ContainsIsRegExp = chkContainsRegExp.Value = vbChecked
    mFilterEdit.ContainsRegExpMinmatches = Val(txtMinMatches.Text)
    mFilterEdit.ContainsMinmatches = Val(txtMinMatches.Text)

    
    CDebug.Post mFilterEdit.SizeLargerThan
    CDebug.Post mFilterEdit.SizeSmallerThan
    'txtmin,txtmax, and txtadsspec
    
    
    
    mFilterEdit.SetAlternateStreamSearchData txtmin.Text, txtmax.Text, Filter_Include, chkADSregexp.Value = vbChecked, txtadsspec.Text, chkadsContainsisregexp.Value = vbChecked, _
    txtADScontaining.Text, 0, _
    txtADSLargerThan.Text * ((1024 * Abs(CBool(cboADSThanMultiplier(0).ListIndex))) ^ (cboADSThanMultiplier(0).ListIndex)), _
    txtADSSmallerThan.Text * ((1024 * Abs(CBool(cboADSThanMultiplier(1).ListIndex))) ^ (cboADSThanMultiplier(1).ListIndex))
    
End Sub


Private Sub cmdCancel_Click()
Dim gotitem As cListItem
    If mdeleteOnCancel Then
    On Error Resume Next
     Set gotitem = mOwner.ListItemFromSearchFilter(mFilterEdit)
     mOwner.lvwfilters.ListItems.Remove gotitem.Index
     mOwner.mFileSearch.Filters.Remove mFilterEdit
    End If
    
    
    Unload Me
End Sub

Private Sub cmdChangeDate_Click()
    Dateform.EditDates mFilterEdit
End Sub


Private Sub cmdOK_Click()
    CmdApply_Click
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub
Private Sub refreshDateCaption()
    Dim StrCap As String
    Dim StrSpecNames(1 To 3) As String
    Dim I As Long, currspecname
    I = 2
    StrSpecNames(1) = "Created"
    StrSpecNames(2) = "Accessed"
    StrSpecNames(3) = "Modified"
    Do
        currspecname = StrSpecNames(Log(I) / Log(2))
        If (mFilterEdit.DatesCheck And I) = I Then
            StrCap = StrCap & currspecname & " After " & FormatDateTime(mFilterEdit.DateStart(I), vbLongDate) & " but before " & FormatDateTime(mFilterEdit.DateEnd(I), vbLongDate) & vbCrLf
        Else
            StrCap = StrCap & currspecname & " (Any)," & vbCrLf
        
        End If
    
    
    
    I = I * 2
    Loop Until I > 8
    
    lblDateInfo.caption = StrCap



End Sub

Private Sub Dateform_EditComplete()
    refreshDateCaption
End Sub

Private Sub Form_Load()
Dim I As Long, loopctl As Control
    Set Dateform = New FrmDateEdit
    TabSpecifications.RemoveAllTabs
    TabSpecifications.AddTab "Standard"
    TabSpecifications.AddTab "Containing"
    TabSpecifications.AddTab "ADS"
    TabSpecifications.AddTab "Scripting"
    TabSpecifications.AddTab "Appearance"
    
    
    '0 Filter_Include  'Include if matched...
    '1 Filter_Exclude  'exclude if matched...
    '2 Filter_Or 'add weight if this or the previous matched.
    '3 Filter_And
    cboOperation.Clear
    cboOperation.AddItem "Include"
    cboOperation.AddItem "Exclude"
    cboOperation.AddItem "Or"
    cboOperation.AddItem "And"
    
    For I = 0 To 1
        cboSizesuffix(I).Clear
        cboSizesuffix(I).AddItem "Bytes"
        cboSizesuffix(I).AddItem "KB"
        cboSizesuffix(I).AddItem "MB"
        cboSizesuffix(I).AddItem "GB"
        cboSizesuffix(I).AddItem "TB"
        
        cboADSThanMultiplier(I).AddItem "Bytes"
        cboADSThanMultiplier(I).AddItem "KB"
        cboADSThanMultiplier(I).AddItem "MB"
        cboADSThanMultiplier(I).AddItem "GB"
        cboADSThanMultiplier(I).AddItem "TB"
        
    Next I
    TabSpecifications_TabClick 1
    'CurrApp.Settings.WriteProfileSetting "FILTERFORM", "Left", Me.left
    'CurrApp.Settings.WriteProfileSetting "FILTERFORM", "Top", Me.top
'load tooltips...
Dim mToolTipControlNames As Variant
Dim mTooltipStrings As Variant, newhandler As CNumericTextEntry
'this only occurs once, so what the heckleson.
mToolTipControlNames = Array("txtfiltername")
mTooltipStrings = Array("Enter then name of this Search filter here. This does not affect the search but is used when " & vbCrLf & "referring to specific filters in other parts of the program.")

Set mFilterautocomplete = New CAutoCompleteCombo
mFilterautocomplete.Init cboFilter, False
Set mTboxHandlers = New Collection
With mTboxHandlers

Set newhandler = New CNumericTextEntry
newhandler.Init txtMinMatches
.Add newhandler
On Error Resume Next
Set newhandler = New CNumericTextEntry
'newhandler.Init txtSize(0)
'If Err <> 0 Then .Add newhandler


Set newhandler = New CNumericTextEntry
'newhandler.Init txtSize(1)
'If Err <> 0 Then .Add newhandler


End With
'Set TT = New clsTooltip
Dim loopcontrol As Control
On Error Resume Next
For Each loopcontrol In Me.Controls
   ' If TypeOf LoopControl Is PictureBox Then
        If loopcontrol.Tag <> "NoChangeBG" Then
            loopcontrol.BackColor = vbWindowBackground
            loopcontrol.ForeColor = vbWindowText
        End If
    
    
  '  End If


Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'save form position...
    
 
    
    
End Sub

Private Sub PicAttributes_Paint()
    Dim attrrect As RECT
    attrrect.Left = 3
    attrrect.Top = 3
    attrrect.Right = PicAttributes.ScaleX(PicAttributes.ScaleWidth, PicAttributes.ScaleMode, vbPixels) - 3
    attrrect.Bottom = PicAttributes.ScaleY(PicAttributes.ScaleHeight, PicAttributes.ScaleMode, vbPixels) - 3
    'DrawEdge PicAttributes.hdc, attrrect, BDR_ETCHINSET, BF_RECT
End Sub

Private Sub picbackground_Click()
    Dim getcolor As Long
    getcolor = ShowColor()
    If getcolor > -1 Then
        picbackground.BackColor = getcolor
        mFilterEdit.Tag.BackColor = getcolor
    End If
End Sub

Private Sub PicForeground_Click()
    Dim getcolor As Long
    getcolor = ShowColor()
    If getcolor > -1 Then
        PicForeground.BackColor = getcolor
        mFilterEdit.Tag.ForeColor = getcolor
    End If
End Sub

Private Sub TabSpecifications_TabClick(ByVal lTab As Long)
Dim I As Long
Dim InvisibleTabs(), VisibleTab
Dim inviscount As Long
    For I = PicTabs.LBound To PicTabs.UBound
    If I <> (lTab - 1) Then
        PicTabs(I).Visible = False
        ReDim Preserve InvisibleTabs(inviscount)
        Set InvisibleTabs(inviscount) = PicTabs(I)
        inviscount = inviscount + 1
    End If
    Next I
'change visibility to avoid messing up taborder...
    PicTabs(lTab - 1).Visible = True
    PicTabs(lTab - 1).ZOrder 0
    PicTabs(lTab - 1).BorderStyle = 0
    PicTabs(lTab - 1).Move TabSpecifications.ClientLeft, TabSpecifications.ClientTop, TabSpecifications.ClientWidth - TabSpecifications.ClientLeft, TabSpecifications.ClientHeight - TabSpecifications.ClientTop
    
    
    
    SetTabStop PicTabs(lTab - 1), InvisibleTabs
        
    
    
    
    CDebug.Post lTab
End Sub



