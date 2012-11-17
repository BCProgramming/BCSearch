VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Represents the GetOpenFileName and GetSaveFileName Function in the Win32 API.
'note that I am somewhat reluctant to do this- not only has it been done before, but
'it seems that it would be better placed in a Common Dialog Box library...
'as such, for consistency, I will probably include the other common dialog boxes.

'This Class Module can only be used inside the BCFile project o' mine.
'well, at least without modification.


'I found a bug when I upgraded from Win98 to Windows XP. Seems the GetOpenFileName was returning 2, to
'indicate initialization failure. I found a KB article with a similiar problem in 2000, but that didn't
'help. I copied my old OpenFile() routine and IT worked, though. Huh.

'Added: Preliminary Unicode support. Uses GetOpen/SaveFilenameW when on unicode systems.





Public Enum CommonDialogNotificationConstants
 CDN_FIRST = (-601)
 CDN_FILEOK = CDN_FIRST - &H5
 CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
 CDN_HELP = (CDN_FIRST - &H4)
 'CDN_INCLUDEITEM = (CDN_FIRST - &H7)   'This is so retarded. won't work with files, only namespace extensions. grr.
 CDN_INITDONE = (CDN_FIRST - &H0)
 CDN_LAST = (-699)
 CDN_SELCHANGE = (CDN_FIRST - &H1)
 CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
 CDN_TYPECHANGE = (CDN_FIRST - &H6)
End Enum
Private Const WM_BASE As Long = &H400
Private Const WM_USER As Long = &H400


Public Enum CommonDialogMessages
    CDM_FIRST& = (WM_BASE + 100)
    CDM_GETFILEPATH& = (CDM_FIRST + &H1)
    CDM_GETFOLDERIDLIST& = (CDM_FIRST + &H3)
    CDM_GETFOLDERPATH& = (CDM_FIRST + &H2)
    CDM_GETSPEC& = (CDM_FIRST + &H0)
    CDM_HIDECONTROL& = (CDM_FIRST + &H5)
    CDM_LAST& = (WM_USER + 200)
    CDM_SETCONTROLTEXT& = (CDM_FIRST + &H4)
    CDM_SETDEFEXT& = (CDM_FIRST + &H6)
End Enum


Public Enum OpenFileNameFlags
 OFN_ALLOWMULTISELECT& = &H200
 OFN_CREATEPROMPT& = &H2000
 OFN_DONTADDTORECENT& = &H2000000
 OFN_ENABLEHOOK& = &H20
 'OFN_ENABLEINCLUDENOTIFY& = &H400000
 OFN_ENABLESIZING& = &H800000
 OFN_ENABLETEMPLATE& = &H40
 OFN_ENABLETEMPLATEHANDLE& = &H80
 OFN_EX_NOPLACESBAR& = &H1
 OFN_EXPLORER& = &H80000
 OFN_EXTENSIONDIFFERENT& = &H400
 OFN_FILEMUSTEXIST& = &H1000
 OFN_FORCESHOWHIDDEN& = &H10000000
 OFN_HIDEREADONLY& = &H4
 OFN_LONGNAMES& = &H200000
 OFN_NOCHANGEDIR& = &H8
 OFN_NODEREFERENCELINKS& = &H100000
 OFN_NOLONGNAMES& = &H40000
 OFN_NONETWORKBUTTON& = &H20000
 OFN_NOREADONLYRETURN& = &H8000
 OFN_NOTESTFILECREATE& = &H10000
 OFN_NOVALIDATE& = &H100
 OFN_OVERWRITEPROMPT& = &H2
 OFN_PATHMUSTEXIST& = &H800
 OFN_READONLY& = &H1
 OFN_SHAREAWARE& = &H4000
 OFN_SHAREWARN& = 0
 OFN_SHOWHELP& = &H10
 OFN_USEMONIKERS& = &H1000000
End Enum


Type OPENFILENAME
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     Flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
     'oops. We just broke Win95,98 compatibility here.
     'Which is why we can pass in the smaller struct size.
     pvreserved As Long
     dwReserved As Long
     FlagsEx As Long
'      void *        pvReserved;
'  DWORD         dwReserved;
'  DWORD         FlagsEx;

End Type
Type OPENFILENAMEW
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     lpstrFilter As Long
     lpstrCustomFilter As Long
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As Long
     nMaxFile As Long
     lpstrFileTitle As Long
     nMaxFileTitle As Long
     lpstrInitialDir As Long
     lpstrTitle As Long
     Flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As Long
     'oops. We just broke Win95,98 compatibility here.
     'Which is why we can pass in the smaller struct size.
     pvreserved As Long
     dwReserved As Long
     FlagsEx As Long
'      void *        pvReserved;
'  DWORD         dwReserved;
'  DWORD         FlagsEx;

End Type
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long



Private Declare Function GetOpenFileNameA Lib "comdlg32.dll" (ByRef pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameA Lib "comdlg32.dll" (ByRef pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileNameW Lib "comdlg32.dll" (ByRef pOpenFilename As OPENFILENAMEW) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32.dll" (ByRef pOpenFilename As OPENFILENAMEW) As Long
Private mvarBufferSize As Long
Private mvarFilter As String
Private mvarIcon As StdPicture

Private mvarCaption  As String
Private mvarEventCallback As IDialogHook
Private mvarflags As OpenFileNameFlags
Private mvarDefExt As String
Private mvarInitialDir As String
Private mvarhDlg As Long    'cached by the Hook procedure in WM_INITDIALOG.
Private mvarSelFilter As String
Private OFN As OPENFILENAME


Private Function GetSaveFileName(ByRef pOpenFilename As OPENFILENAME) As Long
Dim OfnW As OPENFILENAMEW
If MakeWideCalls() Then
    OpenFileNameAtoW pOpenFilename, OfnW
    GetSaveFileName = GetSaveFileNameW(OfnW)
    OpenFileNameWtoA OfnW, pOpenFilename
Else

    GetSaveFileName = GetSaveFileNameA(pOpenFilename)

End If

End Function
Private Function GetOpenFileName(ByRef pOpenFilename As OPENFILENAME) As Long
Dim OfnW As OPENFILENAMEW
If MakeWideCalls() Then
    OpenFileNameAtoW pOpenFilename, OfnW
    GetOpenFileName = GetOpenFileNameW(OfnW)
    OpenFileNameWtoA OfnW, pOpenFilename
Else
    GetOpenFileName = GetOpenFileNameA(pOpenFilename)

End If
End Function
Public Sub StringFromPointer(pOLESTR As Long, strout As String)
         Dim ByteArray(32768) As Byte
         Dim intTemp As Integer
         Dim intCount As Integer
         Dim i As Integer

         intTemp = 1
    
         'Walk the string and retrieve the first byte of each WORD.
         Do While intTemp <> 0
            CopyMemory intTemp, ByVal pOLESTR + i, 2
            'ReDim Preserve ByteArray(1 To intCount)
            ByteArray(intCount) = intTemp
            
            intCount = intCount + 1
            i = i + 2
         Loop

         'Copy the byte array to our string.
         strout = Space$(255)
         CopyMemory ByVal strout, ByteArray(0), intCount
      End Sub
Private Function StrFromPtr(ByVal Pointer As Long) As String
    Dim strreturn As String
    On Error GoTo ErrOccur
    If Pointer = 0 Then Exit Function
    StringFromPointer Pointer, strreturn
    StrFromPtr = strreturn
    Exit Function
    
ErrOccur:
    StrFromPtr = ""

End Function
Private Sub OpenFileNameWtoA(OfnW As OPENFILENAMEW, OFNA As OPENFILENAME)
    With OFNA
        .dwReserved = OfnW.dwReserved
        .Flags = OfnW.Flags
        .FlagsEx = OfnW.FlagsEx
        .hInstance = OfnW.hInstance
        .hWndOwner = OfnW.hWndOwner
        .lCustData = OfnW.lCustData
        .lpfnHook = OfnW.lpfnHook
        .lpstrCustomFilter = StrFromPtr(OfnW.lpstrCustomFilter)
        .lpstrDefExt = StrFromPtr(OfnW.lpstrDefExt)
        .lpstrFile = StrFromPtr(OfnW.lpstrFile)
        .lpstrFileTitle = StrFromPtr(OfnW.lpstrFileTitle)
        .lpstrFilter = StrFromPtr(OfnW.lpstrFilter)
        .lpstrInitialDir = StrFromPtr(OfnW.lpstrInitialDir)
        .lpstrTitle = StrFromPtr(OfnW.lpstrTitle)
        .lpTemplateName = StrFromPtr(OfnW.lpTemplateName)
        .lStructSize = GetOFNlength
    End With
End Sub
Private Sub OpenFileNameAtoW(OFNA As OPENFILENAME, OfnW As OPENFILENAMEW)
'create Unicode structure.

With OfnW
    .dwReserved = OFNA.dwReserved
    .Flags = OFNA.Flags
    .FlagsEx = OFNA.FlagsEx
    .hInstance = OFNA.hInstance
    .hWndOwner = OFNA.hWndOwner
    .lCustData = OFNA.lCustData
    .lpfnHook = OFNA.lpfnHook
    .lpstrCustomFilter = StrPtr(OFNA.lpstrCustomFilter)
    .lpstrDefExt = StrPtr(OFNA.lpstrDefExt)
    .lpstrFile = StrPtr(OFNA.lpstrFile)
    .lpstrFileTitle = StrPtr(OFNA.lpstrFile)
    .lpstrFilter = StrPtr(OFNA.lpstrFilter)
    .lpstrInitialDir = StrPtr(OFNA.lpstrInitialDir)
    .lpstrTitle = StrPtr(OFNA.lpstrTitle)
    .lpTemplateName = StrPtr(OFNA.lpTemplateName)
    .nFileExtension = OFNA.nFileExtension
    .nFileOffset = OFNA.nFileOffset
    .nFilterIndex = OFNA.nFilterIndex
    .nMaxCustFilter = OFNA.nMaxCustFilter
    .nMaxFile = OFNA.nMaxFile
    .nMaxFileTitle = OFNA.nMaxFileTitle
    .pvreserved = OFNA.pvreserved

    .lStructSize = GetOFNlength
    .nFileExtension = OFNA.nFileExtension
    .nFileOffset = OFNA.nFileOffset
    .nFilterIndex = OFNA.nFilterIndex
    .nMaxCustFilter = OFNA.nMaxCustFilter
    .nMaxFile = OFNA.nMaxFile
    .nMaxFileTitle = OFNA.nMaxFileTitle
    .pvreserved = OFNA.pvreserved
End With


End Sub
Private Function GetOFNlength() As Long
    Const OLD_LENGTH = 76
    Const NEW_LENGTH = 88
    Static cached As Long
    If cached = 0 Then
        If IsWinNT Then
            If VerWinMajor >= 5 Then
                cached = NEW_LENGTH
            Else
                cached = OLD_LENGTH
            End If
        Else
            cached = OLD_LENGTH
        
        
        End If
    End If
    GetOFNlength = cached
End Function
Public Property Let BufferSize(ByVal Vdata As Long)
Attribute BufferSize.VB_Description = "Returns the Buffer Size used when allocating strings. Defaults to 32K."
    mvarBufferSize = Vdata
End Property
Public Property Get BufferSize() As Long
    BufferSize = mvarBufferSize
End Property
Public Property Let InitialDir(ByVal Vdata As String)
Attribute InitialDir.VB_Description = "Sets/Returns the Initial Directory to use."
    mvarInitialDir = Vdata
End Property
Public Property Get InitialDir() As String
    InitialDir = mvarInitialDir
End Property
Friend Property Let hdlg(ByVal Vdata As Long)
    mvarhDlg = Vdata
End Property
Public Property Get hdlg() As Long
Attribute hdlg.VB_Description = "returns the hDlg in use. Only valid from within the Callback Interface procedures."
    hdlg = mvarhDlg
End Property
Public Property Let DefExt(ByVal Vdata As String)
Attribute DefExt.VB_Description = "Sets/Returns the Default extension to use it the user specifies no other extension."
    mvarDefExt = Vdata
End Property
Public Property Get DefExt() As String
    DefExt = mvarDefExt
End Property
Public Property Get SelFilter() As String
Attribute SelFilter.VB_Description = "Valid After a dialog is displayed. Contains the text of the selected filter."
    SelFilter = mvarSelFilter
End Property
Public Property Set EventCallback(Vdata As IDialogHook)
Attribute EventCallback.VB_Description = "Sets/Returns the IDialogHook Interface that will intercept events fired by this Dialog. Only one can be specified."
    Set mvarEventCallback = Vdata
End Property
Public Property Let Caption(ByVal Vdata As String)
    mvarCaption = Vdata

End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Get EventCallback() As IDialogHook
    If mvarEventCallback Is Nothing Then
        Set mvarEventCallback = New IDialogHook
    End If
    Set EventCallback = mvarEventCallback
End Property
Public Property Set Icon(Vdata As StdPicture)
Attribute Icon.VB_Description = "Sets/Returns the Icon to use in the dialog. This Icon will be used as the Control-Box icon."
    Set mvarIcon = Vdata
End Property
Public Property Get Icon() As StdPicture
    Set Icon = mvarIcon
End Property

Public Property Let Flags(ByVal Vdata As OpenFileNameFlags)
Attribute Flags.VB_Description = "Sets/Returns the OpenFileName Flags to use when calling The Common Dialog Box Function. Depending on which one is used, the function may modify the flags to make them suitable."
    mvarflags = Vdata
End Property
Public Property Get Flags() As OpenFileNameFlags
    Flags = mvarflags
End Property
Private Function retsame(ByVal t As Long)
    retsame = t
End Function
Public Property Let Filter(ByVal Vdata As String)
Attribute Filter.VB_Description = "Sets/Returns the Filter specification. This is a Null or  Pipe-delimited list of the form Description|Extension|descript... etc."
    mvarFilter = Replace$(Vdata, "|", vbNullChar) & vbNullChar
End Property
Public Property Get Filter() As String
    Filter = mvarFilter
End Property

Public Function GetFileDirect(ByVal hWndOwner As Long, _
                            Optional ByVal StrFilter As String = "All Files|*.*||", _
                            Optional ByVal Flags As OpenFileNameFlags = OFN_EXPLORER + OFN_ENABLEHOOK, _
                            Optional ByVal DefaultExtension As String = "txt", _
                            Optional DialogIcon As StdPicture = Nothing, _
                            Optional ByVal StartDir As String = "C:\", _
                            Optional EventHandler As IDialogHook = Nothing, _
                            Optional ByVal mvarSaving As Boolean = False) As FilesCollection
Attribute GetFileDirect.VB_Description = "perform a direct call to the dialog Library, with the specified arguments."
        Dim getf As FilesCollection, grabfile As CFile
        If EventHandler Is Nothing Then
            Set EventHandler = New IDialogHook
        End If
        Set getf = New FilesCollection
        StrFilter = Replace$(StrFilter, "|", vbNullChar)
    With Me
        .Flags = Flags
        .Filter = StrFilter
        .DefExt = DefaultExtension
        .InitialDir = StartDir
        Set .EventCallback = EventHandler
    End With
    If mvarSaving = False Then
        If Flags And OFN_ALLOWMULTISELECT Then
            Set GetFileDirect = Me.SelectOpenFiles(hWndOwner)
        Else
            Set grabfile = SelectOpenFile(hWndOwner)
            If Not grabfile Is Nothing Then
                getf.Add grabfile
                Set GetFileDirect = getf
            Else
                Set GetFileDirect = Nothing
            
            End If
        End If
    Else
        
        Set grabfile = FileSystem.CreateFile(Me.SelectSaveFile(hWndOwner))
        If Not grabfile Is Nothing Then
            getf.Add grabfile
        End If
    End If
                            
                            
                            
                            
End Function




'NAME: SelectOpenFiles
'+--------------------------------+
' purpose: Selects Multiple files using the GetOpenFileName() API.
' parameters:
' hWnd the Owner of the created dialog.
'returns:
' a Collection of CFile Objects representing the opened files.
'if Cancel was pressed, the count will be zero(0)
'*NOTE*: this function makes use of the Eventhandler Object passed
'in. otherwise, it will create a default instance.





'Public Function SelectOpenFiles(ByVal OwnerWnd As Long) As Files
'    'returns a collection of files that were selected.
'    'of course, there might still be only one.
'    'this function does all the required string parsing.
'    Dim grabDir As Directory, Grabfiles As Files
'    Dim retval As Long, getfile As String
'    Dim splfiles() As String, I As Long
'    Set grabDir = New Directory
'    Set Grabfiles = New Files
'    Me.EventCallback.InitDialog Me
'    'Clear it out. (it's module-level, and might have some crap
'    'left from before. I visualize this function as the zamboni-
'    'other functions have "skated" over the memory, so we need to
'    'clean up.
'    ZeroMemory OFN, Len(OFN)
'    'initialize the values based on our properties.
'    With OFN
'
'        'make sure it has OFN_ALLOWMULTISELECT...
'        If CBool(Not mvarflags And OFN_ALLOWMULTISELECT) Then
'            mvarflags = mvarflags + OFN_ALLOWMULTISELECT
'        End If
'        mvarSelFilter = Space$(mvarBufferSize)
'
'        .Flags = mvarflags
'        .hInstance = App.hInstance
'        .hWndOwner = OwnerWnd
'        .lCustData = 0
'        .lpfnHook = retsame(AddressOf ComDlgHook)
'        .lpstrCustomFilter = mvarSelFilter
'        .nMaxCustFilter = Len(mvarSelFilter)
'        .lpstrDefExt = mvarDefExt & vbNullChar
'        .lpstrFile = Space$(mvarBufferSize - 1)
'        .nMaxFile = Len(.lpstrFile)
'          .lStructSize = Len(OFN)
'
'    End With
'    'call getopenfilename.
'
'    retval = GetOpenFileName(OFN)
'
'    If retval = 0 Then
'    MsgBox "CMNdlg err:" & CommDlgExtendedError()
'        Exit Function
'    End If
'    'now we should have the strfile.
'    getfile = Trim$(OFN.lpstrFile)
'    'remove contiguous NULL chars.
'    Do Until InStr(getfile, String(2, vbNullChar)) = 0
'        getfile = Replace$(getfile, vbNullChar & vbNullChar, vbNullChar)
'    Loop
'    If Right$(getfile, 1) = vbNullChar Then
'        getfile = Mid$(getfile, 1, Len(getfile) - 1)
'    End If
'    'split the resultant string.
'    splfiles = Split(getfile, vbNullChar)
'    If UBound(splfiles) = 0 Then
'        Grabfiles.Add FileSystem.getfile(splfiles(0))
'
'    Else
'        Set grabDir = FileSystem.GetDirectory(splfiles(0))
'        For I = 1 To UBound(splfiles)
'            If splfiles(I) <> "" Then
'                Grabfiles.Add FileSystem.getfile(grabDir.Path & "\" & splfiles(I))
'            End If
'        Next I
'    End If
'    'return the value.
'    Set SelectOpenFiles = Grabfiles
'
'
'
'
'
'
'
'End Function
'returns a single file to open.

Public Function SelectSaveFile(ByVal hWndOwner As Long) As String
Dim returnfile As String
Dim returnthis As CFile
Dim A, Dialoghwnd As Long
Dim OFN As OPENFILENAME
 'since we are only allowing the selection of a Single file, make sure we don't
 'include the Multiselect flag.
 InitStruct OFN
 OFN.hWndOwner = hWndOwner
 OFN.Flags = Flags
If OFN.Flags And OFN_ALLOWMULTISELECT Then  'if included...
    OFN.Flags = OFN.Flags - OFN_ALLOWMULTISELECT 'remove.

End If



    'assign the module level variable for the hook procedure
    'in case the OFN_ENABLEHOOK flag was specified. We could FORCE it,
    'but why bother?
    Set mdlDialog.mFileDialog = Me
     A = GetSaveFileName(OFN)

Do Until OFN.lpstrFile <> "" And OFN.lpstrFile <> Chr$(0)
     DoEvents                                                   'start an idle loop for handling of the WM_INITDIALOG message.
Loop
returnfile = Trim$(OFN.lpstrFile)
returnfile = Trim$(Replace$(returnfile, """", ""))     'we don't replace any quotes. however, we'll probably want to take them
                                                'into account for SelectOpenFiles.
'there COULD be an error retrieving a reference to the file.
'On Error Resume Next
'Set returnthis = FileSystem.GetFile(returnfile)
'If Err <> 0 Then
'    Set returnthis = FileSystem.CreateFile(returnfile)
'End If
returnfile = Replace$(returnfile, vbNullChar, "")
returnfile = Trim$(returnfile)
If InStr(returnfile, "\") = 0 And InStr(returnfile, "/") = 0 Then
    returnfile = IIf(Right$(CurDir$, 1) = "\", CurDir$ & returnfile, CurDir$ & "\" & returnfile)
End If
    
    
SelectSaveFile = returnfile



End Function
Public Function SelectOpenFile(ByVal hWndOwner As Long) As CFile
      
Dim returnfile As String
Dim returnthis As CFile
Dim A, Dialoghwnd As Long
Dim OFN As OPENFILENAME
 'since we are only allowing the selection of a Single file, make sure we don't
 'include the Multiselect flag.
 InitStruct OFN
 OFN.hWndOwner = hWndOwner
 OFN.Flags = Flags
     OFN.lpfnHook = retsame(AddressOf ComDlgHook)
If OFN.Flags And OFN_ALLOWMULTISELECT Then  'if included...
    OFN.Flags = OFN.Flags - OFN_ALLOWMULTISELECT 'remove.

End If


    'assign the module level variable for the hook procedure
    'in case the OFN_ENABLEHOOK flag was specified. We could FORCE it,
    'but why bother?
    Set mdlDialog.mFileDialog = Me
     A = GetOpenFileName(OFN)

Do Until OFN.lpstrFile <> "" And OFN.lpstrFile <> Chr$(0)
     DoEvents                                                   'start an idle loop for handling of the WM_INITDIALOG message.
Loop
returnfile = Trim$(OFN.lpstrFile)
returnfile = Replace$(returnfile, """", "")     'we don't replace any quotes. however, we'll probably want to take them
                                                'into account for SelectOpenFiles.
'there COULD be an error retrieving a reference to the file.
returnfile = Trim$(Replace$(returnfile, vbNullChar, ""))
If returnfile <> "" Then
Set returnthis = FileSystem.GetFile(returnfile)
Else
    Set SelectOpenFile = Nothing
End If

Set SelectOpenFile = returnthis
End Function
'MultiSelect Version.
Public Function SelectOpenFiles(OwnerWnd As Long) As BCFile.Files
      
Dim returnfile As String
Dim returnthis As FilesCollection, splfiles() As String
Set returnthis = New Files
Dim A, Dialoghwnd As Long
Dim OFN As OPENFILENAME
 'since we are only allowing the selection of a Single file, make sure we don't
 'include the Multiselect flag.
 InitStruct OFN
 OFN.hWndOwner = OwnerWnd
 OFN.Flags = Flags
 OFN.lpfnHook = retsame(AddressOf ComDlgHook)
If Not (OFN.Flags And OFN_ALLOWMULTISELECT) Then  'if not included...
    OFN.Flags = OFN.Flags + OFN_ALLOWMULTISELECT 'add.

End If



    'assign the module level variable for the hook procedure
    'in case the OFN_ENABLEHOOK flag was specified. We could FORCE it,
    'but why bother?
    Set mdlDialog.mFileDialog = Me
     A = GetOpenFileName(OFN)

Do Until OFN.lpstrFile <> "" And OFN.lpstrFile <> Chr$(0)
     DoEvents                                                   'start an idle loop for handling of the WM_INITDIALOG message.
Loop
Dim i As Long
'split at null characters.
OFN.lpstrFile = Replace$(OFN.lpstrFile, """", "")
Do Until InStr(OFN.lpstrFile, vbNullChar & vbNullChar) = 0
    OFN.lpstrFile = Replace$(OFN.lpstrFile, vbNullChar & vbNullChar, vbNullChar)
Loop
If Right$(OFN.lpstrFile, 1) = vbNullChar Then
    OFN.lpstrFile = Left$(OFN.lpstrFile, Len(OFN.lpstrFile) - 1)
End If
splfiles = Split(OFN.lpstrFile, vbNullChar)

If UBound(splfiles) = 1 Then
    'single item.
    returnthis.Add FileSystem.GetFile(splfiles(0))
Else
    For i = 1 To UBound(splfiles)
        splfiles(i) = Trim$(splfiles(i))
        If splfiles(i) <> "" Then
            splfiles(i) = splfiles(0) & splfiles(i)
            returnthis.Add FileSystem.GetFile(splfiles(i))
        End If
    Next i

End If




                                                'into account for SelectOpenFiles.
'there COULD be an error retrieving a reference to the file.

    

Set SelectOpenFiles = returnthis




End Function



Private Sub InitStruct(OFN As OPENFILENAME)

    'this is the version 5+ definition of
'the OPENFILENAME structure containing
'three additional members providing
'additional options on Windows 2000
'or later. The SetOSVersion routine
'will assign either OSV_LENGTH (76)
'or OSVEX_LENGTH (88) to the OSV_VERSION_LENGTH
'variable declared above. This variable, rather
'than Len(OFN) is used to assign the required
'value to the OPENFILENAME structure's nStructSize
'member which tells the OS if extended features
'- primarily the Places Bar - are supported.

    
    'OFN.lStructSize = Len(OFN)
    OFN.lStructSize = GetOFNlength
    OFN.hInstance = App.hInstance
           
    OFN.lpstrFilter = Me.Filter
    'make the buffer the buffersize. I found a little difficulty
    'with this in Windows 2000 and up: the File textbox has that
    'many characters selected.
    'I really cannot seem to find a simple way around it.
    'Of course, if I was really determined, I could hook the dialog,
    'and send the Virtual Keypress Delete to the editbox....
    OFN.lpstrFile = Space$(mvarBufferSize - 1) & vbNullChar
    
    OFN.nMaxFile = mvarBufferSize
    OFN.lpstrFileTitle = Space$(mvarBufferSize - 1)
    OFN.nMaxFileTitle = mvarBufferSize - 1
    OFN.lpstrInitialDir = mvarInitialDir
    OFN.lpstrTitle = mvarCaption
    OFN.lpstrTitle = mvarCaption & vbNullChar
    
    OFN.lpfnHook = retsame(AddressOf ComDlgHook)
    If OFN.Flags = 0 Then
        OFN.Flags = OFN_EXPLORER + OFN_FORCESHOWHIDDEN + OFN_ENABLEHOOK
    End If
End Sub

Private Sub Class_Initialize()
'    Set mvarBgBrush = New CBrush
    mvarInitialDir = CurDir$
    mvarBufferSize = 32767
End Sub
