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
Attribute VB_Exposed = True
Option Explicit

'CfolderMonitor class.
Implements IChangeNotification
Private mFolderMonitor As String
Private mFolderCopyTo As String
Private mLogFile As FileStream
Private mMonitor As CFileChangeNotify



Public Sub Initialize(ByVal FolderMonitor As String, CopyTo As String, ByVal logFile As FileStream)
    mFolderMonitor = FolderMonitor
    mFolderCopyTo = CopyTo
    Set mLogFile = logFile
    Set mMonitor = New CFileChangeNotify
    mMonitor.Start mFolderMonitor, False, Me
    
    
    
    
End Sub
Public Sub StopMonitor()
    mMonitor.Cancel

End Sub

Private Sub IChangeNotification_Change(FromNotify As BCFile.CFileChangeNotify, Changed As BCFile.NotifyChangeStruct)
'

Dim I As Long
'mLogFile.WriteString "Received " & I & "Changes." & vbCrLf
For I = 1 To Changed.NumChanges
    'mLogFile.WriteString Changed.Changes(I).ChangedItem & " changed."
    
    If Changed.Changes(I).Changetype = FILE_ACTION_ADDED Or Changed.Changes(I).Changetype = FILE_ACTION_MODIFIED Then
        Debug.Print "copy file, " & Changed.Changes(I).ChangedItem
    
    
    End If

Next I



End Sub
