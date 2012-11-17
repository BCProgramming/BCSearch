VERSION 5.00
Begin VB.Form frmFolderMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BASeCamp Folder Monitor"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFolderMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mtray As frmSysTray
Attribute mtray.VB_VarHelpID = -1
Private mMonitor As CFolderMonitor
Private Sub Form_Load()
Set mtray = New frmSysTray
Load mtray
With mtray

        
        .ToolTip = "SysTray Sample!"
        .IconHandle = Me.Icon.Handle
       ' .ShowBalloonTip "hello", "title"

End With
    Set mMonitor = New CFolderMonitor
    mMonitor.Init "D:\testtest", "D:\testto"
End Sub
