Attribute VB_Name = "modFileVersion"

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long


Public Type FILEINFO
CompanyName As String
FileDescription As String
FileVersion As String
InternalName As String
LegalCopyright As String
originalFilename As String
Productname As String
ProductVersion As String
End Type
Public Enum VersionReturnValue
eOK = 1
eNoVersion = 2
End Enum
'**************************************
'Name: Get Version Number for EXE, DLL or OCX files
'Description:This function will retrieve the version number, product
'name, original program name (like if you right click on the EXE file
'and select properties, then select Version tab, it shows you all that
'information) etc
'By: Serge
'Inputs: None
'Returns: FILEINFO structure
'Assumes:Label (named Label1 and make it wide enough, also increase the
'height of the label to have size of the form), Common Dilaog Box
'(CommonDialog1) and a Command Button (Command1)
'Side Effects: None
'This code is copyrighted and has limited warranties.
'Please see
'http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.4976/lngWId.-1/qx/vb/scripts/ShowCode.htm
'for
'details.
'**************************************
Public Function GetEXEFriendlyName(ByVal EXEFile As String) As String
    Dim ffinto As FILEINFO, tmpret As String
    If GetFileVersionInformation(EXEFile, ffinto) = eOK Then
    
    tmpret = ffinto.FileDescription
    tmpret = Replace$(tmpret, vbNullChar, "")
    If tmpret = EXEFile Then
        tmpret = ffinto.InternalName
    
    End If
    Else
        'GetEXEFriendlyName = tmpret
    
    End If

    GetEXEFriendlyName = tmpret

End Function


Public Sub GetFileVerInfoDirect(ByVal pstrFileName As String, Optional ByRef CompanyName As String, Optional ByRef FileDescription As String, _
Optional FileVersion As String, Optional InternalName As String, Optional LegalCopyright As String, Optional originalFilename As String, Optional Productname As String, _
Optional ProductVersion As String)
    Dim finfo As FILEINFO
    If GetFileVersionInformation(pstrFileName, finfo) = eOK Then
        With finfo
            CompanyName = .CompanyName
            FileDescription = .FileDescription
            FileVersion = .FileVersion
            InternalName = .InternalName
            LegalCopyright = .LegalCopyright
            originalfileame = .originalFilename
            Productname = .Productname
            ProductVersion = .ProductVersion
        End With
    End If



End Sub
Public Function GetFileVersionInformation(ByRef pstrFieName As String, _
ByRef tFileInfo As FILEINFO) As VersionReturnValue
Dim lBufferLen As Long, lDummy As Long
Dim sBuffer() As Byte
Dim lVerPointer As Long
Dim lRet As Long
Dim Lang_Charset_String As String
Dim HexNumber As Long
Dim i As Integer
Dim strTemp As String 'Clear the Buffer tFileInfo
tFileInfo.CompanyName = ""
tFileInfo.FileDescription = ""
tFileInfo.FileVersion = ""
tFileInfo.InternalName = ""
tFileInfo.LegalCopyright = ""
tFileInfo.originalFilename = ""
tFileInfo.Productname = ""
tFileInfo.ProductVersion = ""
lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)
If lBufferLen < 1 Then
GetFileVersionInformation = eNoVersion
Exit Function
End If
ReDim sBuffer(lBufferLen)
lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))
If lRet = 0 Then
GetFileVersionInformation = eNoVersion
Exit Function
End If
lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", _
lVerPointer, lBufferLen)
If lRet = 0 Then
GetFileVersionInformation = eNoVersion
Exit Function
End If
Dim bytebuffer(255) As Byte
MoveMemory bytebuffer(0), lVerPointer, lBufferLen
HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * _
&H10000 + bytebuffer(1) * &H1000000
Lang_Charset_String = Hex(HexNumber) 'Pull it all apart: '04------=
'SUBLANG_ENGLISH_USA '--09----= LANG_ENGLISH
' ----04E4 = 1252 = Codepage for Windows :Multilingual
Do While Len(Lang_Charset_String) < 8
Lang_Charset_String = "0" & Lang_Charset_String
Loop
Dim strVersionInfo(7) As String
strVersionInfo(0) = "CompanyName"
strVersionInfo(1) = "FileDescription"
strVersionInfo(2) = "FileVersion"
strVersionInfo(3) = "InternalName"
strVersionInfo(4) = "LegalCopyright"
strVersionInfo(5) = "OriginalFileName"
strVersionInfo(6) = "ProductName"
strVersionInfo(7) = "ProductVersion"
Dim buffer As String
For i = 0 To 7
buffer = String(255, 0)
strTemp = "\StringFileInfo\" & Lang_Charset_String _
& "\" & strVersionInfo(i)
lRet = VerQueryValue(sBuffer(0), strTemp, _
lVerPointer, lBufferLen)
If lRet = 0 Then
'GetFileVersionInformation = eNoVersion
'Exit Function

Else
    lstrcpy buffer, lVerPointer
    buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)
End If
Select Case i
Case 0
tFileInfo.CompanyName = buffer
Case 1
tFileInfo.FileDescription = buffer
Case 2
tFileInfo.FileVersion = buffer
Case 3
tFileInfo.InternalName = buffer
Case 4
tFileInfo.LegalCopyright = buffer
Case 5
tFileInfo.originalFilename = buffer
Case 6
tFileInfo.Productname = buffer
Case 7
tFileInfo.ProductVersion = buffer
End Select
Next i
GetFileVersionInformation = eOK
End Function '-----------
'Private Sub Command1_Click()
'Dim strFile As String
'Dim udtFileInfo As FILEINFO
'On Error Resume Next
'With CommonDialog1
'.Filter = "All Files (*.*)|*.*"
'.ShowOpen
'strFile = .Filename
'If Err.Number = cdlCancel Or strFile = "" Then Exit Sub
'End With
'If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
'MsgBox "No version available for this file", vbInformation
'Exit Sub
'End If
'Label1 = "Company Name: " & udtFileInfo.CompanyName & vbCrLf
'Label1 = Label1 & "File Description:" &
'udtFileInfo.FileDescription & vbCrLf
'Label1 = Label1 & "File Version:" & udtFileInfo.FileVersion
'& vbCrLf
'Label1 = Label1 & "Internal Name: " & udtFileInfo.InternalName
'& vbCrLf
'Label1 = Label1 & "Legal Copyright: " &
'udtFileInfo.LegalCopyright & vbCrLf
'Label1 = Label1 & "Original FileName:" &
'udtFileInfo.OriginalFileName & vbCrLf
'Label1 = Label1 & "Product Name:" & udtFileInfo.ProductName
'& vbCrLf
'Label1 = Label1 & "Product Version: " &
'udtFileInfo.ProductVersion & vbCrLf

