Attribute VB_Name = "mSaveDlg"
Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Public Function ShowSave(Optional hOwner As Long, Optional sFileName As String, Optional sDialogTitle As String, _
                Optional sInitDir As String, Optional lFlags As EOpenFile = OFN_OVERWRITEPROMPT, _
                Optional sFilter As String, Optional nFilterIndex As Long, _
                Optional sDefExt As String, Optional bCancelError As Boolean = False) As String
    Dim OFN As OPENFILENAME
    Dim sTemp As String
    With OFN
       .lStructSize = Len(OFN)
       .hwndOwner = hOwner
       .flags = lFlags
       .lpstrDefExt = sDefExt
       If sInitDir = "" Then sInitDir = App.Path
       .lpstrInitialDir = sInitDir
       sTemp = sFileName
       .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
       .nMaxFile = 255
       .lpstrFileTitle = String$(255, 0)
       .nMaxFileTitle = 255
        sTemp = sFilter
        For i = 1 To Len(sTemp)
            If Mid(sTemp, i, 1) = "|" Then
               Mid(sTemp, i, 1) = vbNullChar
            End If
        Next
        sTemp = sTemp & String$(2, 0)
        .lpstrFilter = sTemp
        .nFilterIndex = nFilterIndex
        .lpstrTitle = sDialogTitle
        .hInstance = App.hInstance
    End With
    If GetSaveFileName(OFN) Then
       ShowSave = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, vbNullChar) - 1)
    End If
End Function


