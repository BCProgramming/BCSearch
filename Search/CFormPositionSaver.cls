VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormDataSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CFormPosition Saver-
'saves the forms position and width/height attributes when the form is loaded and unloaded.
'saves to specified cINIData Object. This should be initialized, btw....


'uses the CMonitor class and MMonitors module as well, to ensure that the loaded positions are on-screen and on the proper screen.


'1/29/2010 19:46:23 TODO: change code that saves/restores the window position to API calls to setwindowplacement and getwindowplacement.
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32.dll" (ByVal hwnd As Long, ByRef lpwndpl As WINDOWPLACEMENT) As Long


Private mINI As CINIData, mvarMonitor As CMonitor
Private WithEvents Formhandle As Form
Attribute Formhandle.VB_VarHelpID = -1

Public Sub Initialize(WithForm As Form, WithINI As CINIData)
    'MsgBox "initialize of CFormDataSaver"
    Set Formhandle = WithForm
    Set mINI = WithINI
    Set mvarMonitor = New CMonitor
    'should be called during forms "initialize" event, so we get the Load event as well.

End Sub
Private Function DoLoadControlState(FromObject) As Boolean

On Error Resume Next
DoLoadControlState = FromObject.SaveControlStates


End Function
'Private Sub SaveFormPlacement(FormSave As Form, saveti As CINIData)
'
'
'
'
'End Sub
'Private Sub LoadFormPlacement(FormLoad As Form, loadfrom As CINIData)
'    Dim readplacement As WINDOWPLACEMENT
'
'    readplacement.Length = Len(readplacement)
'    'getwindowplacement formload.hWnd,readplacement
'
'
'
'
'
'
'End Sub
Private Sub Formhandle_Load()
    Dim loopcontrol As Control
    Dim CurrIndex As Long, lcount As Long
    Dim savesection As String, HasttProperty As Boolean
    Dim LoadControlStates As Boolean
    Dim gotleft As String, gottop As String, GotWidth As String, GotHeight As String
    Dim SavedMonitor As String
    LoadControlStates = DoLoadControlState(Formhandle)
On Error GoTo BREAKOUT
    gotleft = mINI.ReadProfileSetting("WindowPositions", Formhandle.Name & ".Left")
    gottop = mINI.ReadProfileSetting("WindowPositions", Formhandle.Name & ".Top")
    GotWidth = mINI.ReadProfileSetting("WindowPositions", Formhandle.Name & ".Width")
    GotHeight = mINI.ReadProfileSetting("WindowPositions", Formhandle.Name & ".Height")
    'LoadFormPlacement Formhandle, mINI
    SavedMonitor = mINI.ReadProfileSetting("WindowPositions", Formhandle.Name & ".Monitor")
    If gotleft <> "" Or gottop <> "" Or GotWidth <> "" Or GotHeight <> "" Then
        Formhandle.Move gotleft, gottop, GotWidth, GotHeight
    End If
    Dim monitorObj As CMonitor
    'set width and height...

    
    
    If SavedMonitor <> "" Then
        MMonitors.MonitorsRefresh
        Set monitorObj = MMonitors.MonitorFromDeviceName(SavedMonitor)
        'switch monitors...
        MMonitors.WindowSwitchMonitor Formhandle.hwnd, monitorObj.Index
        Call MMonitors.WindowOffscreen(Formhandle.hwnd, False, True)
    End If
    
    
    
    'Enumerate controls- attempt to load default data from INI file for particular types of controls (textbox,Combobox, etc.
'    On Error Resume Next
'    Set tt = Formhandle.TooltipObject
'    If Err <> 0 Then
'        HasttProperty = False
'    Else
'        HasttProperty = True
'    End If
    For Each loopcontrol In Formhandle.Controls
        'ComboBox,Listbox saving...
        If Not InStr(1, loopcontrol.Tag, "nopersist", vbTextCompare) <> 0 Then
            If TypeOf loopcontrol Is ListBox Or TypeOf loopcontrol Is ComboBox Then
                'save text and previous entries.
                'Section names? - "PersistProperties.<Objectname>"
                'with value names, "Text", ListCount, and List(Index)....
                On Error Resume Next
                savesection = "PersistProperties." & Trim$(loopcontrol.Name) & "." & Trim$(loopcontrol.Index)
                If Err <> 0 Then
                    savesection = "PersistProperties." & Trim$(loopcontrol.Name)
                    Err.Clear
                End If
                'mINI.WriteProfileSetting SaveSection, "Text", LoopControl.Text
                'mINI.WriteProfileSetting SaveSection, "ListCount", LoopControl.ListCount
                lcount = Val(mINI.ReadProfileSetting(savesection, "ListCount"))
                If Err = 0 Then
                    CDebug.Post "Loading Combo/List box- section = " & savesection & "..."
                    'If InStr(savesection, "cboFilter") > 0 Then Stop
                    For CurrIndex = 0 To lcount - 1
                        Dim addme As String
                        'mINI.WriteProfileSetting SaveSection, "List(" & Trim$(CurrIndex + 1) & ")", LoopControl.List(CurrIndex)
                        addme = mINI.ReadProfileSetting(savesection, "List(" & Trim$(CurrIndex + 1) & ")")
                        If addme <> "" Then
                        loopcontrol.AddItem addme
                        End If
                    Next CurrIndex
                
                End If
                
            
            ElseIf TypeOf loopcontrol Is vbalListViewCtl Then
            
                LoadListViewColumnConfig loopcontrol, loopcontrol.Name & "_Columns"
            
            
            End If
        '//
    '       CDebug.Post  "load tooltips here. for " & loopcontrol.Name
    '        If HasttProperty Then
    '            Dim ttsection As String
    '            Dim tttext As String, Forecolor As Long, BackColor As Long, useballoon As Boolean
    '            Dim Strtitle As String, iconval As Long, mshowmode As Long, mmargin As Long
    '            Dim centered As Boolean, istransparent As Boolean, align As Long, width As Long
    '            On Error Resume Next
    '            savesection = "ToolTipProperties." & Trim$(loopcontrol.Name) & "." & Trim$(loopcontrol.Index)
    '            If Err <> 0 Then
    '                savesection = "ToolTipProperties." & Trim$(loopcontrol.Name)
    '                Err.Clear
    '            End If
    '
    '            tttext = mINI.ReadProfileSetting(savesection, "Text")
    '            Forecolor = Val(mINI.ReadProfileSetting(savesection, "ForeColor"))
    '            BackColor = Val(mINI.ReadProfileSetting(savesection, "BackColor"))
    '            useballoon = CBool(mINI.ReadProfileSetting(savesection, "Balloon"))
    '            Strtitle = mINI.ReadProfileSetting(savesection, "Title")
    '            iconval = mINI.ReadProfileSetting(savesection, "icon")
    '            mshowmode = ttsmDefault
    '            align = ttaleft
    '
    '
    '
    '            tt.AddTool loopcontrol, tttext, Forecolor, BackColor, useballoon, Strtitle, iconval, mshowmode, mmargin, centered, istransparent, align, width
    '
    '
    '        End If
        End If
    Next
    
    
    
    
    'CHANGED: now forces default fonts...
    
    'Ideally should save and Load Control fonts, but that seems a bit excessive...
    'CDebug.Post  "changing default font of form," & Formhandle.Name & " which is currently set to " & Formhandle.Font.Name
    'ChangeToDefaultFont Formhandle
    
    
BREAKOUT:
End Sub

Private Sub Formhandle_Unload(Cancel As Integer)
Dim SaveControlStates As Boolean, CurrIndex As Long
Dim savesection As String, loopcontrol As Control
    mINI.WriteProfileSetting "WindowPositions", Formhandle.Name & ".Left", Formhandle.Left
    mINI.WriteProfileSetting "WindowPositions", Formhandle.Name & ".Top", Formhandle.Top
    mINI.WriteProfileSetting "WindowPositions", Formhandle.Name & ".Width", Formhandle.Width
    mINI.WriteProfileSetting "WindowPositions", Formhandle.Name & ".Height", Formhandle.Height
    Dim mmon As CMonitor
    'save monitor device
       Set mmon = MMonitors.Monitors(MMonitors.MonitorFromWindow(Formhandle.hwnd))
        mINI.WriteProfileSetting "WindowPositions", Formhandle.Name & ".Monitor", mmon.Device
    
    
For Each loopcontrol In Formhandle.Controls
    If StrComp(loopcontrol.Tag, "nopersist", vbTextCompare) <> 0 Then
            If TypeOf loopcontrol Is ListBox Or TypeOf loopcontrol Is ComboBox Then
                'save text and previous entries.
                'Section names? - "PersistProperties.<Objectname>"
                'with value names, "Text", ListCount, and List(Index)....
                On Error Resume Next
                savesection = "PersistProperties." & Trim$(loopcontrol.Name) & "." & Trim$(loopcontrol.Index)
                If Err <> 0 Then
                    savesection = "PersistProperties." & Trim$(loopcontrol.Name)
                End If
                'If InStr(savesection, "cboFilter") > 0 Then Stop
                mINI.WriteProfileSetting savesection, "Text", loopcontrol.Text
                mINI.WriteProfileSetting savesection, "ListCount", loopcontrol.ListCount
                CDebug.Post "Persisting Combo/List box- section = " & savesection & "..."
                For CurrIndex = 0 To loopcontrol.ListCount - 1
                    mINI.WriteProfileSetting savesection, "List(" & Trim$(CurrIndex + 1) & ")", loopcontrol.List(CurrIndex)
                    
                
                Next CurrIndex
                
                
                
            ElseIf TypeOf loopcontrol Is vbalListViewCtl Then
                    
                    SaveListViewColumnConfig loopcontrol, loopcontrol.Name & "_Columns"
            
            
            
            End If
        End If
    
    Next
    
    
    
    

End Sub
Private Sub SaveListViewColumnConfig(LvwSave As Object, ByVal Section As String)

Dim castit As vbalListViewCtl
'save column positions and widths.
Dim loopcolumn As cColumn
Dim loopint As Long
Set castit = LvwSave
For loopint = 1 To castit.columns.Count
    Set loopcolumn = castit.columns.Item(loopint)
    'save the column information. <Columnkey>_Width= and <columnkey>_position=
    mINI.WriteProfileSetting Section, loopcolumn.Key & "_Width", loopcolumn.Width
    mINI.WriteProfileSetting Section, loopcolumn.Key & "_Position", loopcolumn.position

Next loopint






End Sub
Public Sub LoadListViewColumnConfig(LvwLoad As Object, ByVal Section As String)
    Dim castit As vbalListViewCtl
    Dim loopcolumn As cColumn
    Dim loopint As Long
    Set castit = LvwLoad
    For loopint = 1 To castit.columns.Count
        Set loopcolumn = castit.columns.Item(loopint)
        Dim newwidth As Long
        newwidth = mINI.ReadProfileSetting(Section, loopcolumn.Key & "_Width", 40)
        
        loopcolumn.Width = newwidth
        
        Dim gotposition As Long
        gotposition = Val(mINI.ReadProfileSetting(Section, loopcolumn.Key & "_Position"))
        If gotposition <> 0 Then
            loopcolumn.position = mINI.ReadProfileSetting(Section, loopcolumn.Key & "_Position")
        End If


    Next
End Sub
