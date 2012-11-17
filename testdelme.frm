VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1230
      Left            =   2565
      TabIndex        =   0
      Top             =   1305
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents taskdlg As CTaskDialog
Attribute taskdlg.VB_VarHelpID = -1

Private Sub Command1_Click()
    Set taskdlg = New CTaskDialog
    taskdlg.flags = TDF_ENABLE_HYPERLINKS + TDF_USE_COMMAND_LINKS_NO_ICON
    taskdlg.CommonButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
    taskdlg.WindowTitle = "satan visits."
    taskdlg.AddButton 1, "Yes" & vbCrLf & "I'd like to sell my soul to lucifer. And Face everlasting damnation.", False, True
    taskdlg.AddButton 2, "No" & vbCrLf & "I have other plans, you know, meetings and whatnot.", False
    taskdlg.AddButton 3, "Tuesday" & vbCrLf & "I'm free tuesday... is that good?", False
    'taskdlg.AddButton IDOK, "OK", False, False
    taskdlg.MainInstruction = "Satan would like to make an offer."
    taskdlg.content = "Hell's keeper would like to add your sould to his collection. If you agree, you can have an entire bag of marshmallows. He won't even open it for you to eat one. he promises."
    taskdlg.Show Me.hWnd
End Sub

Private Sub taskdlg_ButtonClick(Interactor As BCFile.ITaskDialogInteractor, ByVal ButtonID As Long, ByVal RadioButton As Boolean)
    Debug.Print "button clicked " & ButtonID
End Sub

Private Sub taskdlg_Created(Interactor As BCFile.ITaskDialogInteractor)
Debug.Print "taskdlg_created"
End Sub
