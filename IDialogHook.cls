VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDialogHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'these where easier to obtain then you might think.
'I recommend Resource Hacker for anybody.
'most of these are only in the Save Dialog.
Public Enum ExplorerStyleSaveControlIdentifierConstants
    IDC_lblLookIn = 1091        'The Look In: Label.
    IDC_CboDriveCombo = 1137   'The Drive Combobox.
    IDC_tbrToolBar = 1088   'I think the dialog replaces it with the toolbar control at run-time.
    IDC_LvwListView = 1120 'the listview where files are displayed.
    IDC_LblFileName = 1090
    IDC_txtFileName = 1152
    IDC_lblFileType = 1089
    IDC_cboFileType = 1136
    IDC_chkReadOnly = 1040
    IDC_cmdSave = 1
    IDC_cmdCancel = 2
    IDC_CmdHelp = 1038
End Enum
Public Enum SharingViolationOperation
    OFN_SHAREFALLTHROUGH& = 2
    OFN_SHARENOWARN& = 1
    OFN_SHAREWARN& = 0
End Enum
'Private mCached As CFileDialog


    

'The Dialog Hook Interface for the CFileDialog, and other dialogs.
Public Function SharingViolation() As SharingViolationOperation
SharingViolation = OFN_SHAREWARN
End Function
Public Sub TypeChange()
    'fired when a new File type is selected.
    CDebug.PostMessage "TypeChange"
End Sub
Public Sub ControlEvent(ByVal ControlID As ExplorerStyleSaveControlIdentifierConstants)
    'occurs when an event occurs to a control on the dialog.
    Debug.Print "ControlEvent, ID=" & ControlID
End Sub
Public Sub InitDone()
    'Initialization complete.
    CDebug.PostMessage "InitDone"
End Sub
Public Sub SelChange(ByVal newFilename As String)
    'when the selection changes.
    CDebug.PostMessage "SelChange:" & newFilename
End Sub
Public Function IncludeItem(ByVal strPath As String, ByVal pidl As Long) As Boolean
    CDebug.PostMessage "includeitem:" & strPath
    IncludeItem = True
End Function
Public Function VerifySelection(ByVal pickedFile As String) As Boolean
    'function should verify the file(s) selected, returning true or false.
    VerifySelection = True
    CDebug.PostMessage "VerifySelection:" & pickedFile
End Function
Public Sub HelpClick()
    'the Help button was clicked, if present.
    MsgBox "Wow, you're dumb."
End Sub
Public Sub FolderChange(ByVal newFolder As String)
    'return value ignored.
    CDebug.PostMessage "Folder change to " & newFolder
End Sub


Public Function DirectHook(ByVal hdlg As Long, _
                            ByVal uMsg As WindowsMessages, wParam As Long, lParam As Long) As Long
                            
                            
'The one and only. a direct interception of the Common Dialog Box (save or open) Messages.
'you'll probably want to know what you're doing...
                    CDebug.Post "DirectHook"
End Function


