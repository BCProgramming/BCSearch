VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFolderMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements BCFile.IChangeNotification
Private mDirNotify As BCFile.CFileChangeNotify
Private mCopyTo As String
Private mDirMonitor As String
Public Property Get CopyTo() As String
    CopyTo = mCopyTo
End Property

Public Property Get DirMonitor() As String
    DirMonitor = mDirMonitor
End Property
Public Sub Init(ByVal DirectoryMonitor As String, ByVal CopyTo As String)
    If Not mDirNotify Is Nothing Then mDirNotify.Cancel
    Set mDirNotify = New CFileChangeNotify
    mCopyTo = CopyTo
    mDirMonitor = DirectoryMonitor
    
    'alright, start monitoring...
    mDirNotify.Start DirectoryMonitor, True, Me


End Sub
'Public Property Let CopyTo(ByVal Vdata As String)
'    mCopyTo = Vdata
'    If Not mDirNotify Is Nothing Then
'        mDirNotify.Cancel
'        Set mDirNotify = Nothing
'    End If
'    'Set mDirNotify = New CFileChangeNotify
'
'End Property


Private Sub IChangeNotification_Change(FromNotify As BCFile.CFileChangeNotify, Changed As BCFile.NotifyChangeStruct)
'
Dim strbuildstate As String
Dim i As Long, grabfile As CFile
Load frmFolderMonitor.mtray
For i = 0 To Changed.NumChanges - 1
    With Changed.Changes(i)
        Select Case .Changetype
        
        
        Case FILE_ACTION_ADDED
            strbuildstate = strbuildstate & "File added,"
        Case FILE_ACTION_MODIFIED
        strbuildstate = strbuildstate & "File modified,"
        Case FILE_ACTION_REMOVED
        strbuildstate = strbuildstate & "File removed,"
        Case FILE_ACTION_RENAMED_NEW_NAME
        strbuildstate = strbuildstate & "File renamed(new)"
        Case FILE_ACTION_RENAMED_OLD_NAME
        strbuildstate = strbuildstate & "File renamed(old) "
        
        End Select
    strbuildstate = strbuildstate & .ChangedItem & vbCrLf
    
    
    If mCopyTo <> "" Then
        Set grabfile = GetFile(.ChangedItem)
        grabfile.Copy mCopyTo, frmFolderMonitor.hwnd
    End If
    End With

Next i


frmFolderMonitor.mtray.ShowBalloonTip strbuildstate, "Change detected.", NIIF_NONE, 30000
End Sub
