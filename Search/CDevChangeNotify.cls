VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDevChangeNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CDevChangeNotify: uses SHChangeNotify to watch for and alert addition/removal of drives.

Private Const SHCNE_ALLEVENTS As Long = &H7FFFFFFF
Private Const SHCNE_ASSOCCHANGED As Long = &H8000000
Private Const SHCNE_ATTRIBUTES As Long = &H800&
Private Const SHCNE_CREATE As Long = &H2&
Private Const SHCNE_DELETE As Long = &H4&
Private Const SHCNE_DISKEVENTS As Long = &H2381F
Private Const SHCNE_DRIVEADD As Long = &H100&
Private Const SHCNE_DRIVEADDGUI As Long = &H10000
Private Const SHCNE_DRIVEREMOVED As Long = &H80&
Private Const SHCNE_EXTENDED_EVENT As Long = &H4000000
Private Const SHCNE_FREESPACE As Long = &H40000
Private Const SHCNE_GLOBALEVENTS As Long = &HC0581E0
Private Const SHCNE_INTERRUPT As Long = &H80000000
Private Const SHCNE_MEDIAINSERTED As Long = &H20&
Private Const SHCNE_MEDIAREMOVED As Long = &H40&
Private Const SHCNE_MKDIR As Long = &H8&
Private Const SHCNE_NETSHARE As Long = &H200&
Private Const SHCNE_NETUNSHARE As Long = &H400&
Private Const SHCNE_RENAMEFOLDER As Long = &H20000
Private Const SHCNE_RENAMEITEM As Long = &H1&
Private Const SHCNE_RMDIR As Long = &H10&
Private Const SHCNE_SERVERDISCONNECT As Long = &H4000&
Private Const SHCNE_UPDATEDIR As Long = &H1000&
Private Const SHCNE_UPDATEIMAGE As Long = &H8000&
Private Const SHCNE_UPDATEITEM As Long = &H2000&
Private Const WM_USER As Long = &H400

Private Const WM_USER_MEDIACHANGED = WM_USER + 88

Implements iSuperClass
Dim mclass As cSuperClass

Private Declare Sub SHChangeNotify Lib "Shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByRef dwItem1 As Any, ByRef dwItem2 As Any)
Private Type PIDLSTRUCT
  'Fully qualified pidl (relative to
  'the desktop folder) of the folder
  'to monitor changes in. 0 can also
  'be specified for the desktop folder.
   pidl As Long
   
  'Value specifying whether changes in
  'the folder's subfolders trigger a
  'change notification event.
   bWatchSubFolders As Long
End Type

Private Declare Function SHChangeNotifyRegister Lib "shell32" Alias "#2" _
   (ByVal hWnd As Long, _
    ByVal uFlags As Long, _
    ByVal dwEventID As Long, _
    ByVal uMsg As Long, _
    ByVal cItems As Long, _
    lpps As PIDLSTRUCT) As Long
    
    Private Declare Function SHChangeNotifyDeregister Lib "Shell32.dll" (ByVal UUID As Long) As Long
    
    
Private mSubclassedForm As FrmDevNotify
Private mRegisterLng As Long
Public Sub Start()
Dim pidstruct As PIDLSTRUCT
    If Not mclass Is Nothing Then
        mclass.UnSubclass
        Set mclass = Nothing
    
    
    End If
        Set mSubclassedForm = New FrmDevNotify
        Set mclass = New cSuperClass
        
        
        'get the desktop folder pidl...
        Dim deskpidl As Long
        deskpidl = SHGetSpecialFolderLocation(mSubclassedForm.hWnd, CSIDL_DESKTOP)
        
        
        pidstruct.pidl = deskpidl
        pidstruct.bWatchSubFolders = -1
       mRegisterLng = SHChangeNotifyRegister(mSubclassedForm.hWnd, SHCNE_DRIVEADD, &H2, WM_USER_MEDIACHANGED, 1, pidstruct)
        
        
        
        
        mclass.Subclass mSubclassedForm.hWnd, Me, True
        
End Sub
Public Sub Stop_()
On Error Resume Next
    SHChangeNotifyDeregister mRegisterLng
    
    If Err = 0 Then
    mRegisterLng = 0
    Else
        Debug.Print "Error#" & Err.Number & " Desc:" & Err.Description
    End If
    
End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
'
If uMsg = WM_USER_MEDIACHANGED Then
    'handle this message...
    Debug.Print "WM_USER_MEDIACHANGED:" & lParam & " " & wParam

ElseIf uMsg = WM_DESTROY Then
    mclass.UnSubclass
    Set mclass = Nothing


End If
Debug.Print "uMSG=" & uMsg
End Sub

Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

End Sub
