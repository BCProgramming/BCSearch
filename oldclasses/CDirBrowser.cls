VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDirBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Implements iSuperClass

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Public Enum BrowseInfoFlags
BIF_BROWSEFORCOMPUTER = &H1000
BIF_BROWSEFORPRINTER = &H2000
BIF_BROWSEINCLUDEFILES = &H4000
BIF_BROWSEINCLUDEURLS = &H80
BIF_DONTGOBELOWDOMAIN = &H2
BIF_EDITBOX = &H10
BIF_NEWDIALOGSTYLE = &H40
BIF_RETURNFSANCESTORS = &H8
BIF_RETURNONLYFSDIRS = &H1
BIF_SHAREABLE = &H8000
BIF_STATUSTEXT = &H4
BIF_USENEWUI = &H40
BIF_VALIDATE = &H20
End Enum

#If BCFILE = 1 Then
Public Function BrowseForDirectory(ByVal hwnd As Long, ByVal Title As String, ByVal Flags As BrowseInfoFlags) As Directory
#Else



#End If
    Dim DirGet As Directory
    Dim BIF As BROWSEINFO, ret As Long
    Dim ppath As String
    
    BIF.hOwner = hwnd
    BIF.lpszTitle = "TITLE"
    BIF.ulFlags = BIF_EDITBOX
    BIF.lpfn = RetParam(AddressOf BrowseCallbackProc)
    ret = SHBrowseForFolder(BIF)
    ppath = Space$(2048)
    SHGetPathFromIDList ret, ppath
    ppath = Trim$(Replace$(ppath, vbNullChar, " "))
    If ppath = "" Then
        Set BrowseForDirectory = Nothing
    Else
        Set DirGet = FileSystem.GetDirectory(ppath)
    End If
    Set BrowseForDirectory = DirGet
End Function

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub
