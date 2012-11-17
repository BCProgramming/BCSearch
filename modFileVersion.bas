Attribute VB_Name = "modFileVersion"

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long


Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hmem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hmem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)


Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long ' = 0x3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type


Public Type FILE_VERSION_INFO
    FileMajor As Long
    FileMinor As Long
    FileRevision As Long
    ProductMajor As Long
    ProductMinor As Long
    ProductRevision As Long
    Fileflags As Long
    FileOS As Long
    FileType As Long
    FileSubType As Long
    
End Type
Public Type FILEINFO
CompanyName As String
Filedescription As String
FileVersion As String
InternalName As String
LegalCopyright As String
OriginalFilename As String
ProductName As String
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
    
    tmpret = ffinto.Filedescription
    tmpret = Replace$(tmpret, vbNullChar, "")
    If tmpret = EXEFile Then
        tmpret = ffinto.InternalName
    
    End If
    Else
        'GetEXEFriendlyName = tmpret
    
    End If

    GetEXEFriendlyName = tmpret

End Function


Public Sub GetFileVerInfoDirect(ByVal pstrFileName As String, Optional ByRef CompanyName As String, Optional ByRef Filedescription As String, _
Optional FileVersion As String, Optional InternalName As String, Optional LegalCopyright As String, Optional OriginalFilename As String, Optional ProductName As String, _
Optional ProductVersion As String)
    Dim finfo As FILEINFO
    If GetFileVersionInformation(pstrFileName, finfo) = eOK Then
        With finfo
            CompanyName = .CompanyName
            Filedescription = .Filedescription
            FileVersion = .FileVersion
            InternalName = .InternalName
            LegalCopyright = .LegalCopyright
            originalfileame = .OriginalFilename
            ProductName = .ProductName
            ProductVersion = .ProductVersion
        End With
    End If



End Sub
Public Sub TestReadFileVersion()
    Dim fvi As FILE_VERSION_INFO
    fvi = ReadFileVersion("C:\windows\system32\shell32.dll")
    'Call ReadGetFileVersion("C:\windows\system32\shell32.dll")
    Stop
End Sub

Public Function ReadFileVersion(sDriverFile As String) As FILE_VERSION_INFO
   
   Dim FI As VS_FIXEDFILEINFO
   Dim ret As FILE_VERSION_INFO
   Dim sBuffer() As Byte
   Dim nBufferSize As Long
   Dim lpBuffer As Long
   Dim nVerSize As Long
   Dim nUnused As Long
   Dim tmpVer As String
   
  'GetFileVersionInfoSize determines whether the operating
  'system can obtain version information about a specified
  'file. If version information is available, it returns
  'the size in bytes of that information. As with other
  'file installation functions, GetFileVersionInfoSize
  'works only with Win32 file images.
  '
  'A empty variable must be passed as the second
  'parameter, which the call returns 0 in.
   nBufferSize = GetFileVersionInfoSize(sDriverFile, nUnused)
   
   If nBufferSize > 0 Then
   
     'create a buffer to receive file-version
     '(FI) information.
      ReDim sBuffer(nBufferSize)
      Call GetFileVersionInfo(sDriverFile, 0&, nBufferSize, sBuffer(0))
      
     'VerQueryValue function returns selected version info
     'from the specified version-information resource. Grab
     'the file info and copy it into the  VS_FIXEDFILEINFO structure.
      Call VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize)
      Call CopyMemory(FI, ByVal lpBuffer, Len(FI))
     
     'extract the file version from the FI structure
     
    ret.FileMajor = (FI.dwFileVersionMS And &HFFFFFF00) / (2 ^ 16)
    ret.FileMinor = FI.dwFileVersionMS And &HFF&
    ret.FileRevision = FI.dwFileVersionLS And &HFFFF&
    
    ret.ProductMajor = (FI.dwProductVersionMS And &HFFFFFF00) / (2 ^ 16)
    ret.ProductMinor = FI.dwProductVersionMS And &HFF&
    ret.ProductRevision = FI.dwProductVersionLS And &HFFFF&
     ret.Fileflags = FI.dwFileFlags
     ret.FileOS = FI.dwFileOS
     ret.FileType = FI.dwFileType
     ret.FileSubType = FI.dwFileSubtype
     
'      tmpVer = Format$(hiWord(FI.dwFileVersionMS)) & "." & _
'               Format$(loWord(FI.dwFileVersionMS), "00") & "."
'
'      If FI.dwFileVersionLS > 0 Then
'         tmpVer = tmpVer & Format$(hiWord(FI.dwFileVersionLS), "00") & "." & _
'                           Format$(loWord(FI.dwFileVersionLS), "00")
'      Else
'         tmpVer = tmpVer & Format$(FI.dwFileVersionLS, "0000")
'      End If
'
'      End If
   
  ' ODBCGetFileVersion = tmpVer
  ReadFileVersion = ret
   End If
End Function



'Function ReadFileVersion(fName As String) As FILE_VERSION_INFO
'Dim resInfo As FILE_VERSION_INFO, retBuffer As VS_FIXEDFILEINFO
'Dim dataBlock() As Byte, blockSize As Long
'Dim BufferLen As Long, lpBuffer As Long
'Dim hMem As Long
'
'blockSize = GetFileVersionInfoSize(fName & Chr(0), 0)
'If blockSize <= 0 Then
'Exit Function
'End If
'ReDim dataBlock(0 To blockSize - 1)
'
'If GetFileVersionInfo(fName & Chr(0), 0, blockSize, dataBlock(0)) = 0 Then
''Problem
'Exit Function
'End If
'
'hMem = GlobalAlloc(0, Len(retBuffer))
'lpBuffer = GlobalLock(hMem)
'
'VerQueryValue dataBlock(0), "\" & Chr(0), lpBuffer, BufferLen
'
'CopyMemory VarPtr(retBuffer), ByVal lpBuffer, BufferLen
'
'GlobalUnlock lpBuffer
'GlobalFree hMem
'
'resInfo.FileMajor = (retBuffer.dwFileVersionMS And &HFFFFFF00) / (2 ^ 16)
'resInfo.FileMinor = retBuffer.dwFileVersionMS And &HFF&
'resInfo.FileRevision = retBuffer.dwFileVersionLS And &HFFFF&
'
'resInfo.ProductMajor = (retBuffer.dwProductVersionMS And &HFFFFFF00) / (2 ^ 16)
'resInfo.productMinor = retBuffer.dwProductVersionMS And &HFF&
'resInfo.ProductRevision = retBuffer.dwProductVersionLS And &HFFFF&
'
'ReadFileVersion = resInfo
'End Function

Public Function GetFileVersionInformation(ByRef pstrFileName As String, _
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
tFileInfo.Filedescription = ""
tFileInfo.FileVersion = ""
tFileInfo.InternalName = ""
tFileInfo.LegalCopyright = ""
tFileInfo.OriginalFilename = ""
tFileInfo.ProductName = ""
tFileInfo.ProductVersion = ""
lBufferLen = GetFileVersionInfoSize(pstrFileName, lDummy)
If lBufferLen < 1 Then
GetFileVersionInformation = eNoVersion
Exit Function
End If
ReDim sBuffer(lBufferLen)
lRet = GetFileVersionInfo(pstrFileName, 0&, lBufferLen, sBuffer(0))
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
    If InStr(buffer, vbNullChar) > 1 Then
    buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)
    Else
        buffer = Trim$(buffer)
    End If
End If
Select Case i
Case 0
tFileInfo.CompanyName = buffer
Case 1
tFileInfo.Filedescription = buffer
Case 2
tFileInfo.FileVersion = buffer
Case 3
tFileInfo.InternalName = buffer
Case 4
tFileInfo.LegalCopyright = buffer
Case 5
tFileInfo.OriginalFilename = buffer
Case 6
tFileInfo.ProductName = buffer
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

