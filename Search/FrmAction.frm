VERSION 5.00
Object = "{AFFDD50D-733B-4E1C-8F98-E88F1ED6980D}#1.0#0"; "vbaListView6BC.ocx"
Begin VB.Form FrmAction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Action Filters"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
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
   ScaleHeight     =   4440
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   3180
      TabIndex        =   13
      ToolTipText     =   "Execute the set of Action filters."
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      ToolTipText     =   "Closes the action filters dialog."
      Top             =   3960
      Width           =   1095
   End
   Begin VB.PictureBox PicFilterEdit 
      Height          =   4395
      Left            =   5520
      ScaleHeight     =   4335
      ScaleWidth      =   6000
      TabIndex        =   6
      Top             =   0
      Width           =   6060
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   660
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelChange 
         Caption         =   "&Undo"
         Height          =   375
         Left            =   60
         TabIndex        =   11
         Top             =   3840
         Width           =   1035
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   3840
         Width           =   1035
      End
      Begin VB.ComboBox cboActionType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "nopersist"
         Top             =   120
         Width           =   1815
      End
      Begin VB.PictureBox PicActions 
         Height          =   2775
         Index           =   1000
         Left            =   180
         ScaleHeight     =   2715
         ScaleWidth      =   5475
         TabIndex        =   21
         Tag             =   "Scripted"
         Top             =   1020
         Width           =   5535
         Begin VB.CommandButton cmdeditScript 
            Caption         =   "&Edit..."
            Height          =   375
            Left            =   780
            TabIndex        =   31
            Top             =   780
            Width           =   1035
         End
         Begin VB.TextBox txtScriptLanguage 
            Height          =   285
            Left            =   960
            TabIndex        =   29
            Top             =   180
            Width           =   2895
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   $"FrmAction.frx":0000
            Height          =   585
            Left            =   120
            TabIndex        =   32
            Top             =   1320
            Width           =   5100
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Script:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Language:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   180
            Width           =   765
         End
      End
      Begin VB.PictureBox PicActions 
         Height          =   2775
         Index           =   1
         Left            =   240
         ScaleHeight     =   2715
         ScaleWidth      =   5355
         TabIndex        =   18
         Tag             =   "move"
         Top             =   1020
         Width           =   5415
         Begin VB.ComboBox cboMoveMask 
            Height          =   315
            Left            =   1020
            TabIndex        =   26
            Top             =   420
            Width           =   3795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Move Mask:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.PictureBox PicActions 
         Height          =   2775
         Index           =   3
         Left            =   600
         ScaleHeight     =   2715
         ScaleWidth      =   5355
         TabIndex        =   20
         Tag             =   "attributes"
         Top             =   1020
         Width           =   5415
         Begin VB.ListBox lstAttributes 
            Height          =   1860
            Left            =   1020
            Style           =   1  'Checkbox
            TabIndex        =   35
            Top             =   540
            Width           =   2595
         End
         Begin VB.ComboBox cboAttributesChangeMode 
            Height          =   315
            ItemData        =   "FrmAction.frx":00CE
            Left            =   1020
            List            =   "FrmAction.frx":00DB
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Tag             =   "nopersist"
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mode:"
            Height          =   195
            Left            =   420
            TabIndex        =   34
            Top             =   180
            Width           =   450
         End
      End
      Begin VB.PictureBox PicActions 
         Height          =   2775
         Index           =   0
         Left            =   480
         ScaleHeight     =   2715
         ScaleWidth      =   5355
         TabIndex        =   9
         Top             =   1080
         Width           =   5415
         Begin VB.ComboBox cborenamemask 
            Height          =   315
            Index           =   0
            Left            =   1260
            TabIndex        =   15
            Top             =   660
            Width           =   3735
         End
         Begin VB.Label lblRenameMask 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rename Mask:"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   720
            Width           =   1050
         End
      End
      Begin VB.PictureBox PicActions 
         Height          =   2775
         Index           =   2
         Left            =   420
         ScaleHeight     =   2715
         ScaleWidth      =   5355
         TabIndex        =   19
         Tag             =   "copy"
         Top             =   1020
         Width           =   5415
         Begin VB.ComboBox cboCopyMask 
            Height          =   315
            Left            =   1020
            TabIndex        =   27
            Top             =   660
            Width           =   3915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Copy Mask:"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   720
            Width           =   840
         End
      End
      Begin VB.Label lbldescription 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   375
         Left            =   60
         TabIndex        =   23
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Action Type:"
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.PictureBox PicCurrentFilters 
      BorderStyle     =   0  'None
      Height          =   3315
      Left            =   60
      ScaleHeight     =   3315
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   540
      Width           =   5415
      Begin VB.ComboBox cboActon 
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   120
         Width           =   1875
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Edit the currently Selected Action Filter"
         Top             =   2820
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   2340
         TabIndex        =   4
         ToolTipText     =   "Remove the currently Selected Action Filter"
         Top             =   2820
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Add a new Action Filter"
         Top             =   2820
         Width           =   1095
      End
      Begin vbaBClListViewLib6.vbalListViewCtl lvwActionFilters 
         Height          =   2355
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4154
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
      Begin VB.Label lblAct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Act on:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   525
      End
   End
   Begin VB.Label Label1 
      Caption         =   """Action"" Filters allow you to manipulate the files that were found with Search Filters."
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5415
   End
End
Attribute VB_Name = "FrmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mActionFilters As Collection 'simple ol' collection of CActionFilter objects.

Private mFilterEdit As CActionFilter
Private mItemEdit As cListItem
Private mrenamefilterAutocomplete As CAutoCompleteCombo
Private mOwner As FrmSearch

Private addingFilter As String, AddingItem As String
Public Sub ShowDialog(OwnerForm As FrmSearch)
    Set mOwner = OwnerForm
    cboActon.Clear
    cboActon.AddItem "All Results"
    cboActon.AddItem "Selected Items only"
    'If mOwner.GetSelectedItems.Count = 0 Then
    '    cboActon.enabled = False
    'End If
    Me.Show , OwnerForm
End Sub
Private Sub cboActionType_Click()
    'display the correct picturebox. for example, listindex 0 is the "rename" picturebox.
    'PicActions(cboActionType.ListIndex).ZOrder 0
    Dim Visibleframe As Variant, InvisibleFrames As Variant, inviscount As Long
    Dim I As Long, currentindex As Long, loopctl As Object
    currentindex = cboActionType.ItemData(cboActionType.ListIndex)
    
    'first, save the visible and invisible pictureboxes...
    Set Visibleframe = PicActions(cboActionType.ListIndex)
    
    
    For Each loopctl In Me.Controls
        If loopctl.Name = "PicActions" Then
            I = loopctl.Index
            If I <> currentindex Then
                inviscount = inviscount + 1
                ReDim Preserve InvisibleFrames(inviscount - 1)
                Set InvisibleFrames(inviscount - 1) = PicActions(I)
            End If
        End If
    Next
        
'    For I = PicActions.LBound To PicActions.UBound
'    Next
    
    
    
    
    With PicActions(currentindex)
    
    .Move lbldescription.Left, lbldescription.Top + lbldescription.Height
    
    .ZOrder 0
    
    
    
    
    
    End With
End Sub
Private Function FindFilterForItem(citem As cListItem) As CActionFilter

    Dim LoopFilter As CActionFilter
        For Each LoopFilter In mActionFilters
            If LoopFilter.Key = citem.Tag Then
                Set FindFilterForItem = LoopFilter
                Exit Function
            End If
        Next LoopFilter
        Set FindFilterForItem = Nothing

End Function
Private Function FindItemForFilter(FilterObj As CActionFilter) As cListItem
    Dim loopItem As Long, currli As cListItem
        For loopItem = 1 To lvwActionFilters.ListItems.Count
                Set currli = lvwActionFilters.ListItems.Item(loopItem)
                If currli.Key = FilterObj.Key Then
                    Set FindItemForFilter = currli
                    Exit Function
                End If
                    
        
        Next loopItem
    
End Function



Private Sub cmdAdd_Click()
    Dim newfilter As CActionFilter
    Set newfilter = New CActionFilter
    addfilter newfilter
End Sub
Private Sub Removeitem(itemremove As cListItem)
    mActionFilters.Remove itemremove.Tag
    lvwActionFilters.ListItems.Remove itemremove.Tag




End Sub
Private Sub EditItem(itemEdit As cListItem)
PicFilterEdit.Move 0, 0
PicFilterEdit.ZOrder 0

Dim I As Long
'cboActionType.ListIndex = mFilterEdit.Actiontype
For I = 0 To cboActionType.ListCount
    If cboActionType.ItemData(I) = mFilterEdit.Actiontype Then
        cboActionType.ListIndex = I
        Exit For
    End If

Next I

'load all the data from the item into our controls...

'rename info...
cborenamemask(0).Text = mFilterEdit.ActionRename_Mask

txtDescription.Text = mFilterEdit.Description

cboCopyMask.Text = mFilterEdit.ActionMove_Mask

cboMoveMask.Text = mFilterEdit.ActionMove_Mask


For I = 0 To lstAttributes.ListCount - 1
    If (mFilterEdit.ActionAttributes_Attributes And lstAttributes.ItemData(I)) = lstAttributes.ItemData(I) Then
        lstAttributes.selected(I) = True
    Else
        lstAttributes.selected(I) = False
    End If

Next I





End Sub
Private Sub addfilter(Filteradd As CActionFilter)
'add the item and show the editing area for it.
    Dim newitem As cListItem, usekey As String
    usekey = "KEY" & Int(Rnd * 16777216) & CStr(ObjPtr(Filteradd))
    Filteradd.Key = usekey
    Set newitem = lvwActionFilters.ListItems.Add(, usekey, Filteradd.Description)
    newitem.Tag = usekey
    Set mItemEdit = newitem
    Set mFilterEdit = Filteradd
    AddingItem = usekey
    EditItem newitem
    



End Sub

Private Sub cmdBack_Click()
    'apply the changes to the active item.
    Dim I As Long
    With mFilterEdit
        .ActionRename_Mask = cborenamemask(0).Text
        .Description = txtDescription.Text
        .Actiontype = cboActionType.ListIndex
        
        
        If .Actiontype = Action_Copy Then
            .ActionMove_Mask = cboCopyMask.Text
        ElseIf .Actiontype = Action_Move Then
            .ActionMove_Mask = cboMoveMask.Text
        End If
        mItemEdit.Text = .Description
        mItemEdit.SubItems(1).caption = .Actiontype
        If cboAttributesChangeMode.ListIndex > 0 Then
        .ActionAttributes_AttributeModifyMode = cboAttributesChangeMode.ItemData(cboAttributesChangeMode.ListIndex)
        End If
        For I = 0 To lstAttributes.ListCount - 1
            If lstAttributes.selected(I) Then .ActionAttributes_Attributes = .ActionAttributes_Attributes + lstAttributes.ItemData(I)
        Next
    
    End With
    mActionFilters.Add mFilterEdit
    
    PicFilterEdit.ZOrder vbSendToBack


End Sub

Private Sub cmdCancelChange_Click()
    lvwActionFilters.ListItems.Remove AddingItem
    PicFilterEdit.ZOrder vbSendToBack
    

End Sub

Private Sub cmdEdit_Click()

    Dim EditThis As CActionFilter
    If Not lvwActionFilters.SelectedItem Is Nothing Then
    EditItem lvwActionFilters.SelectedItem
    End If

End Sub

Private Sub cmdeditScript_Click()
    Static UseEditor As FrmTextEdit
    If UseEditor Is Nothing Then Set UseEditor = New FrmTextEdit
    mFilterEdit.ActionScripted_Code = UseEditor.EditTextFunc(mFilterEdit.ActionScripted_Code, Me)
End Sub

Private Sub cmdExecute_Click()
    Dim LoopFilter As CActionFilter
    Dim GotSel As Collection
    If cboActon.ListIndex = 1 Then
        Set GotSel = mOwner.GetSelectedItems
    End If
    For Each LoopFilter In mActionFilters
        If Not GotSel Is Nothing Then
            LoopFilter.DoAction GotSel
        Else
            LoopFilter.DoAction mOwner.lvwfiles.ListItems
        End If
    
    
    Next
    
    
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdRemove_Click()
    If Not lvwActionFilters.SelectedItem Is Nothing Then
        Removeitem lvwActionFilters.SelectedItem
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 5610
    PicFilterEdit.ZOrder vbSendToBack
    If mActionFilters Is Nothing Then Set mActionFilters = New Collection
    cboActionType.Clear
    '    Action_Scripted = 1000
    '    Action_Rename
    '    Action_Move
    '    Action_Copy
    '    Action_Attributes
    cboActionType.AddItem "Rename"
    cboActionType.ItemData(0) = 0
    cboActionType.AddItem "Move"
    cboActionType.ItemData(1) = 1
    cboActionType.AddItem "Copy"
    cboActionType.ItemData(2) = 2
    cboActionType.AddItem "Attributes"
    cboActionType.ItemData(3) = 3
    
    cboActionType.AddItem "Scripted"
    cboActionType.ItemData(4) = 1000
    
    cboActionType.AddItem "Modify Contents"
    cboActionType.ItemData(5) = 4
    With cboAttributesChangeMode
        .AddItem "Normal"
        .ItemData(0) = FILE_ATTRIBUTE_NORMAL
        .AddItem "Read-Only"
        .ItemData(1) = FILE_ATTRIBUTE_READONLY
        .AddItem "Archive"
        .ItemData(2) = FILE_ATTRIBUTE_ARCHIVE
        .AddItem "Hidden"
        .ItemData(3) = FILE_ATTRIBUTE_HIDDEN
        .AddItem "System"
        .ItemData(4) = FILE_ATTRIBUTE_SYSTEM
        .AddItem "Encrypted"
        .ItemData(5) = FILE_ATTRIBUTE_ENCRYPTED
        .AddItem "Compressed"
        .ItemData(6) = FILE_ATTRIBUTE_COMPRESSED
    
    
    End With
    On Error Resume Next
    
    With lvwActionFilters
        .columns.Add , "DESCRIPTION", "Description"
        .columns.Add , "TYPE", "Type"
        
    
    End With
    If mActionFilters.Count = 0 Then
        cmdRemove.enabled = False
        cmdEdit.enabled = False
    Else
        cmdRemove.enabled = True
        cmdEdit.enabled = True
    End If
    With cboAttributesChangeMode
    .Clear
'    Attribute_Toggle
'    Attribute_Add
'    Attribute_Remove
'    Attribute_Set
    .AddItem "Toggle"
    .AddItem "Add"
    .AddItem "Remove"
    .AddItem "Set"
    .ItemData(0) = attribute_toggle
    .ItemData(1) = attribute_add
    .ItemData(2) = attribute_Remove
    .ItemData(3) = attribute_Set
    End With
    Set mrenamefilterAutocomplete = New CAutoCompleteCombo
    mrenamefilterAutocomplete.Init cborenamemask, False
    cboActionType.ListIndex = 1

End Sub

Private Sub lvwActionFilters_Click()
    cmdEdit.enabled = Not (lvwActionFilters.SelectedItem Is Nothing)
    cmdRemove.enabled = cmdEdit.enabled
End Sub

Private Sub lvwActionFilters_ItemClick(Item As vbaBClListViewLib6.cListItem)
    cmdEdit.enabled = Not (lvwActionFilters.SelectedItem Is Nothing)
    cmdRemove.enabled = cmdEdit.enabled
End Sub
