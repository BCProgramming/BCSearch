VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CContextSubClasser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements iSuperClass
'Private mCtxCached As Object
Private mctx3 As IContextMenu3
Private mctx2 As olelib.IContextMenu2
'Private mCtxCached As olelib.IContextMenu2  'cached object reference, used during subclassing, for it's HandleMenuMsg() method.
Private mClasser As cSuperClass
Private mCallback As IContextCallback
Public Sub Init(ByVal hWnd As Long, ICtx As IContextMenu2, callbackobj As IContextCallback)
    'Dim Ictx2 As olelib.IContextMenu2
    'Dim Ictx3 As IContextMenu3
On Error GoTo ReportInitError
Set mCallback = callbackobj
Debug.Print "type of Ictx is " & TypeName(ICtx)
500    If Not mClasser Is Nothing Then
501        mClasser.UnSubclass
502        Set mClasser = Nothing
503    End If
504    Set mClasser = New cSuperClass
    
505    On Error Resume Next
506    If TypeOf ICtx Is IContextMenu3 Then
507        Debug.Print "IContextMenu3"
508        Set mctx3 = ICtx
509    Else
510        Set mctx2 = ICtx
511        Debug.Print "IContextMenu2"
512    End If
    'Set mCtxCached = ICtx2
    'subclass all messages....
513    mClasser.AddMessages True, CWM_MEASUREITEM, CWM_MENUSELECT, CWM_DRAWITEM, CWM_INITMENU, CWM_INITMENUPOPUP, CWM_MENUCHAR, CWM_MENUSELECT
    'mClasser.Subclass hwnd, Me, True
514    mClasser.Subclass hWnd, Me, False
    Exit Sub
ReportInitError:
    Debug.Print "error in CContextSubClasser.Init- on line " & Erl & " " & Err.Description
End Sub


Private Sub Class_Terminate()
    If Not mClasser Is Nothing Then
    mClasser.UnSubclass
    End If
    Set mClasser = Nothing
    Set mctx3 = Nothing
    Set mctx2 = Nothing
End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
'   Debug.Print "AFTER:"
End Sub

Private Sub iSuperClass_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'

9 Dim resultof As Long
10  On Error GoTo ProcError
'11    Debug.Print "BEFORE!" & uMsg
12    If Not mctx3 Is Nothing Then
'        Debug.Print "calling handleMenuMsg2..."
13            Call mctx3.HandleMenuMsg2(uMsg, wParam, lParam, resultof)
14            lHandled = 1
15    ElseIf Not mctx2 Is Nothing Then
        
        
        
        
'        Debug.Print "calling HandleMenuMsg: "
16        mctx2.HandleMenuMsg uMsg, wParam, lParam
17        lHandled = 1
18        Exit Sub
19    End If
Exit Sub
ProcError:
    Debug.Print "error occured in iSuperClass_Before- " & Err.Description & " Line:" & Erl
    
End Sub
