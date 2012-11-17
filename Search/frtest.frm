VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrtest 
      Interval        =   100
      Left            =   1380
      Top             =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Dim fnum As Long
Private Sub Form_Load()
    fnum = FreeFile
    Open "D:\activewnd.txt" For Output As fnum
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #fnum

End Sub

Private Sub tmrtest_Timer()
    Dim hwnd As Long
    Dim strgrab As String
    hwnd = GetActiveWindow
    strgrab = Space$(255)
    GetWindowText hwnd, strgrab, 254
    
    
    strgrab = Trim$(strgrab)
    Print #fnum, strgrab
    
End Sub
