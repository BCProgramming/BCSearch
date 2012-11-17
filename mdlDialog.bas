Attribute VB_Name = "mdlDialog"

Option Explicit

'Module for
Private Const WM_COMMAND = &H111
Public Const WM_NOTIFY As Long = &H4E&
Public Const WM_INITDIALOG As Long = &H110
Public Const CDN_FIRST As Long = -601
Public Const CDN_INITDONE As Long = (CDN_FIRST - &H0&)

Public Const SHVIEW_ICON As Long = &H7029
Public Const SHVIEW_LIST As Long = &H702B
Public Const SHVIEW_REPORT As Long = &H702C
Public Const SHVIEW_THUMBNAIL As Long = &H702D
Public Const SHVIEW_TILE As Long = &H702E
Private Const WM_SETICON As Long = &H80
Private Const DWL_MSGRESULT As Long = 0


Private Const sizeuse = 32767

Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "User32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long


Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private Declare Function GetParent Lib "User32.dll" (ByVal hwnd As Long) As Long
'typedef struct _OFNOTIFY {
'  NMHDR          hdr;
'  LPOPENFILENAME lpOFN;
'  LPTSTR         pszFile;
'} OFNOTIFY, *LPOFNOTIFY;
Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type




Private Type OFNOTIFY
    hdr As NMHDR
    lpOFN As Long   'long pointer to the OFN struct. we won't need this member, since we have the firing object.
    pszFile As String
End Type
'typedef struct _OFNOTIFYEX {
'  NMHDR           hdr;
'  LPOPENFILENAME  lpOFN;
'  LPVOID          psf;
'  LPVOID          pidl;
'} OFNOTIFYEX, *LPOFNOTIFYEX;
Private Type OFNOTIFYEX
    hdr As NMHDR
    lpOFN As Long
    psf As Long     'LPVOID(void *)
    lpidl As Long   'LPVOID(void *)
End Type
Public mFileDialog As CFileDialog   'dialog being Hooked.

'UINT_PTR CALLBACK OFNHookProc(
'    HWND hdlg,
'    UINT uiMsg,
'    WPARAM wParam,
'    lParam lParam
');

'BOOL CALLBACK ComDlg32HkProc(HWND hDlg,
'                                UINT uMsg,
'                                WPARAM wParam,
'                                lParam lPar
Public Function ComDlgHook(ByVal hdlg As Long, ByVal uMsg As WindowsMessages, wParam As Long, lParam As Long) As Long
'Is anybody surprised that I haven't written a application in  or
'C++ that was useful (well, except my association viewer...)



'perform stuff in the Hook procedure.
Static BrushMake As Long
Dim OFEX As OFNOTIFYEX, tempval As Long
Dim OFSTRUCT As OFNOTIFY, Strbuffer As String
Dim nhdr As NMHDR
Dim tstream As FileStream  ', FSO As FileSystemObject
On Error GoTo reportComHookerror
'Set FSO = New FileSystemObject
'Set tstream = FSO.OpenTextFile("C:\LOGHOOK.TXT", ForAppending, False)
Dim fn As Long
fn = FreeFile

Debug.Print "HOWDY FROM THE HOOK"
'Debug.Print "ComDlgHook Invoked. hDlg="; hDlg; " uMsg="; uMsg; " wparam="; wParam; " lparam="; lParam
Debug.Print uMsg

'tstream.WriteLine "Hook message:" & uMsg
'we want our class to be the MOST flexible one ever made.
'so what do we do? Well, first off, we need to
Select Case uMsg
    Case CWM_INITDIALOG
        'ha. ad the icon.
    '    tstream.WriteLine "CDN_INITDIALOG"
    Debug.Print "CDN_INITDIALOG"
        If Not mFileDialog.Icon Is Nothing Then
            SendMessage GetParent(hdlg), WM_SETICON, 0, ByVal mFileDialog.Icon.Handle
            
        End If
        'also, give them a HDlg.
        mFileDialog.hdlg = hdlg
    'Case WM_CTLColorDLG
        'return the background color of the dialog.
        'CDebug.PostMessage "CTLCOLORDLG"
        'tstream.WriteLine "WM_CTLColorDLG"
            'ComDlgHook = mFileDialog.BGBrush.BrushHandle
            
    
    
    
    
    
'////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////
    
    Case WM_NOTIFY
        Debug.Print "WM_NOTIFY"
        'a notification from one of the child controls.
        'first, get the OFStruct member.
        'this will be pointed to by lparam, so copy the bugger.
     '   Call tstream.WriteLine("Attempting CopyMemory...")
     Debug.Print "attempting copymemory..."
     Debug.Print lParam
        CopyMemory nhdr, lParam, Len(nhdr)
'        Call tstream.WriteLine("Code=" & OFSTRUCT.hdr.Code)
        If OFSTRUCT.hdr.code > 0 Then
        '    tstream.WriteLine "EMERGENCY! what the fudge?"
            Debug.Print "HDR.CODE=0!"
            Exit Function
        Else
            '
        End If
              'we send it to the parent of this dialog.
            Select Case OFSTRUCT.hdr.code
            Case CDN_FILEOK
                Debug.Print "CDN_OK"
                    Strbuffer = Space$(sizeuse)
                    SendMessage GetParent(hdlg), CDM_GETFILEPATH, Len(Strbuffer), ByVal Strbuffer
                    Strbuffer = Trim$(Strbuffer)
                    If mFileDialog.EventCallback.VerifySelection(Strbuffer) Then
                        'verified. they are allowing it.
                    Else
                        'REJECTION!
                        ComDlgHook = -1
                        SetWindowLong hdlg, DWL_MSGRESULT, -1
                    End If
               Case CDN_FOLDERCHANGE
                'fired when they change folders.
             '   tstream.WriteLine "CDN_FOLDERCHANGE"
             Debug.Print "FOLDERCHANGE"
                Strbuffer = Space$(sizeuse)
                
                SendMessage GetParent(hdlg), CDM_GETFOLDERPATH, Len(Strbuffer), ByVal Strbuffer
                Strbuffer = Trim$(Strbuffer)
                mFileDialog.EventCallback.FolderChange Strbuffer
                Case CDN_HELP
                    'this is the easiest one.
              '      tstream.WriteLine "CDN_HELP"
              Debug.Print "CDN_HELP"
                    mFileDialog.EventCallback.HelpClick
                    'done.
'                Case CDN_INCLUDEITEM
'                    'Dim Strbuffer As String
'                    tstream.WriteLine "CDN_INCLUDEITEM"
'
'                    Strbuffer = Space$(sizeuse) & vbNullChar
'                    CopyMemory OFEX, lParam, Len(OFEX)
'                    'I imagine our client will be spiteful if
'                    'we give them a PIDL, so pass them a file path
'                    'instead.
'                    Call SHGetPathFromIDList(OFEX.lpidl, Strbuffer)
'                    Strbuffer = Trim$(Strbuffer)
'                    Strbuffer = Replace$(Strbuffer, vbNullChar, "")
'                    If mFileDialog.EventCallback.IncludeItem(Strbuffer, OFEX.lpidl) Then
'                        'true
'                        'include it.
'                        ComDlgHook = 1
'
'                    Else
'                        'false
'                        'exclude.
'                        ComDlgHook = 0
'
'
'                    End If
            Case CDN_INITDONE
           '     tstream.WriteLine "CDN_INITDONE"
                mFileDialog.EventCallback.InitDone
            Case CDN_SELCHANGE
            '    tstream.WriteLine "CDN_SELCHANGE"
                Strbuffer = Space$(sizeuse)
                SendMessage GetParent(hdlg), CDM_GETFILEPATH, Len(Strbuffer), ByVal Strbuffer
                Strbuffer = Replace$(Trim$(Strbuffer), vbNullChar, "")
                mFileDialog.EventCallback.SelChange Strbuffer
            Case CDN_SHAREVIOLATION
            '    tstream.WriteLine "CDN_SHAREVIOLATION"
                tempval = mFileDialog.EventCallback.SharingViolation
                SetWindowLong hdlg, DWL_MSGRESULT, tempval
                ComDlgHook = tempval
                
            Case CDN_TYPECHANGE
                mFileDialog.EventCallback.TypeChange
           ' tstream.WriteLine "CDN_TYPECHANGE"
            
            End Select
        
       
       '////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////
End Select
'Call tstream.WriteLine("******END handling for message " & uMsg & "******")
'tstream.Close
Exit Function
reportComHookerror:
Debug.Print "com hook error, " & Err.Description & " #" & Err.Number
End Function

