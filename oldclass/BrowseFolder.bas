Attribute VB_Name = "BrowseFolder"
Option Explicit

'BrowseFolder Callback routine.

'
'int CALLBACK BrowseCallbackProc(
'    HWND hwnd,
'    UINT uMsg,
'    LPARAM lParam,
'    lParam lpData
');
Public Function BrowseCallbackProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal Lparam As Long, ByVal lpdata As Long) As Long
'
Debug.Print "browsecallback ", Hwnd, uMsg, Lparam, lpdata
End Function
Public Function RetParam(retme As Long) As Long
RetParam = retme
End Function
