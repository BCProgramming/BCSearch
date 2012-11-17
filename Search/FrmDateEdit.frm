VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FrmDateEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Date Specs"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4200
   Icon            =   "FrmDateEdit.frx":0000
   LinkTopic       =   "FrmDateEdit"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkDateSpecs 
      Caption         =   "&Modified After:"
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   26
      Top             =   3480
      Width           =   1395
   End
   Begin VB.CheckBox ChkDateSpecs 
      Caption         =   "&Accessed After:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   25
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CheckBox ChkDateSpecs 
      Caption         =   "&Created After:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   180
      Width           =   1395
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1035
   End
   Begin MSComCtl2.DTPicker DTPickerTimeBegin 
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateBegin 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateEnd 
      Height          =   315
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerTimeEnd 
      Height          =   315
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateBegin 
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   2220
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerTimeBegin 
      Height          =   315
      Index           =   4
      Left            =   2400
      TabIndex        =   9
      Top             =   2220
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateEnd 
      Height          =   315
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerTimeEnd 
      Height          =   315
      Index           =   4
      Left            =   2400
      TabIndex        =   13
      Top             =   2880
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateBegin 
      Height          =   315
      Index           =   8
      Left            =   0
      TabIndex        =   14
      Top             =   3780
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerTimeBegin 
      Height          =   315
      Index           =   8
      Left            =   2400
      TabIndex        =   16
      Top             =   3780
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerDateEnd 
      Height          =   315
      Index           =   8
      Left            =   0
      TabIndex        =   18
      Top             =   4440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941121
      CurrentDate     =   39868
   End
   Begin MSComCtl2.DTPicker DTPickerTimeEnd 
      Height          =   315
      Index           =   8
      Left            =   2400
      TabIndex        =   20
      Top             =   4440
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   215941122
      CurrentDate     =   39868
   End
   Begin VB.Line LneSeparator 
      BorderColor     =   &H8000000E&
      Index           =   3
      X1              =   -120
      X2              =   4080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line LneSeparator 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   2
      X1              =   -120
      X2              =   4080
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line LneSeparator 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   -120
      X2              =   4080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line LneSeparator 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   -120
      X2              =   4080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   19
      Top             =   4500
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "But Before:"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   15
      Top             =   3840
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   2940
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "But Before:"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   1380
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "But Before:"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   900
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "At"
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   540
      Width           =   150
   End
End
Attribute VB_Name = "FrmDateEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event EditComplete()

Private WithEvents mEditDatesOf As CSearchFilter
Attribute mEditDatesOf.VB_VarHelpID = -1
Private mFormSaver As CFormDataSaver

Private Sub Form_Initialize()
    Set mFormSaver = New CFormDataSaver
    mFormSaver.Initialize Me, CurrApp.Settings


End Sub


Private Sub ReloadFilterData()
    'load the date Specifications from the given filter...
    'first, populate the checkboxes....
    Dim I As Long
    With mEditDatesOf
        I = 2
        Do
        'the checkbox....
            ChkDateSpecs(I).Value = IIf((mEditDatesOf.DatesCheck And I), vbChecked, vbUnchecked)
            
        'Beginning dates...
            DTPickerDateBegin(I).Value = mEditDatesOf.DateStart(I)
            DTPickerTimeBegin(I).Value = mEditDatesOf.DateStart(I)
            
        
        'Ending Dates...
            DTPickerDateEnd(I).Value = mEditDatesOf.DateEnd(I)
            DTPickerTimeEnd(I).Value = mEditDatesOf.DateEnd(I)
            I = I * 2
        Loop Until I > 8
    
    
    End With



End Sub
Public Sub EditDates(ByVal OfFilter As CSearchFilter)
    Set mEditDatesOf = OfFilter
    
    
    ReloadFilterData
    
    Me.Show
End Sub

Private Sub ChkDateSpecs_Click(Index As Integer)
    'enable/disable based in current checked value.
    
    Dim checkedVal As Boolean
    checkedVal = ChkDateSpecs(Index).Value = vbChecked
    
    DTPickerDateBegin(Index).enabled = checkedVal
    DTPickerDateEnd(Index).enabled = checkedVal
    DTPickerTimeBegin(Index).enabled = checkedVal
    DTPickerTimeEnd(Index).enabled = checkedVal
    
    
End Sub

Private Sub CmdApply_Click()
'apply changes in this form to mEditDatesOf

Dim buildspecs As DateSpecConstants, I As Long, dtsubtract As Date
With mEditDatesOf
        I = 2
        Do
        'the checkbox....
            'ChkDateSpecs(i).Value = Abs((mEditDatesOf.DatesCheck And i) = i)
            buildspecs = buildspecs + Abs(ChkDateSpecs(I).Value * I)
            
        'Beginning dates...
        dtsubtract = DateSerial(Year(DTPickerTimeBegin(I).Value), _
                                Month(DTPickerTimeBegin(I).Value), Day(DTPickerTimeBegin(I).Value))
                                
            mEditDatesOf.DateStart(I) = DTPickerDateBegin(I).Value '+ (DTPickerTimeBegin(I).Value)
            
            
        
        'Ending Dates...
            dtsubtract = DateSerial(Year(DTPickerTimeBegin(I).Value), _
                                Month(DTPickerTimeBegin(I).Value), Day(DTPickerTimeBegin(I).Value))
            mEditDatesOf.DateEnd(I) = DTPickerDateEnd(I).Value '+ (DTPickerTimeEnd(I).Value)
            
            I = I * 2
        Loop Until I > 8
    
    
    End With
    mEditDatesOf.DatesCheck = buildspecs



RaiseEvent EditComplete
End Sub

Private Sub cmdCancel_Click()
RaiseEvent EditComplete
End Sub

Private Sub cmdOK_Click()
CmdApply_Click
RaiseEvent EditComplete
Unload Me

End Sub

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub Form_Load()
'BlurForm Me
 ExtendFrame Me
End Sub
