VERSION 5.00
Begin VB.Form frmRename 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BCSearch - Renamer"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   5
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Frame FraFields 
      Caption         =   "Special Fields:"
      Height          =   2475
      Left            =   60
      TabIndex        =   3
      Top             =   1740
      Width           =   4995
      Begin VB.Label lblSpecialFields 
         Caption         =   "<<Special Fields>>"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4755
      End
   End
   Begin VB.TextBox txtrenamemask 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblrenamemask 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rename Mask:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRename.frx":0000
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Special Fields For renaming:

'%originalname% : Original filename.
'%originalext% : Original Extension.
'%filesize%    : filesize, in bytes. unformatted.
'%filesizeformatted%  : filesize, formatted.
'%column name%  : name of any other column as it appears in the column header.
Private Type RenameStruct
    FieldName As String
    FieldValue As String
End Type
Dim mlvwuse As vbalListViewCtl
Dim mRenameFilter As CSearchFilter
Private Sub DoRename(lvwitem As cListItem, NewnameMask As String)
    Dim filerename As CFile
    Dim substMask As String, renameFields() As RenameStruct
    Set filerename = GetFile(lvwitem.Tag)
    'substMask = Replace$(NewnameMask, "%originalname%", filerename.basename)
    'substMask = Replace$(substMask, "%originalext%", filerename.Extension)
    
    
    Set filerename = Nothing
    
    'filerename.Rename substMask, 0
    'Name lvwitem.Tag As substMask
    
    lvwitem.Text = substMask
    lvwitem.BackColor = &HFF00
    
    
    
    

    
    




End Sub
Public Sub RenameLvwContents()

    Dim currli As Long
    Dim curritem As cListItem
    Dim filterpass As Boolean

    For currli = 1 To mlvwuse.ListItems.Count
        Set curritem = mlvwuse.ListItems(currli)
        If Not mRenameFilter Is Nothing Then
            filterpass = mRenameFilter.FilterResult(GetFile(curritem.Tag))
        Else
            filterpass = True

        End If
        
        If filterpass Then
            'filter passed or no filter set.
            DoRename curritem, txtrenamemask.Text
        
        
        
        
        End If
    Next
    
    
End Sub


Public Sub Init(lvwuse As vbalListViewCtl)
    Set mlvwuse = lvwuse
End Sub

Private Sub cmdOK_Click()
RenameLvwContents
End Sub

