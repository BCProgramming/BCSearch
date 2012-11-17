Attribute VB_Name = "ModShell"
Option Explicit

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PLANES = 14            '  Number of planes

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwflags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long


Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As UUID)

Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As UUID) As Long

      Private Declare Function progIDfromCLSID Lib "ole32.dll" Alias "ProgIDFromCLSID" (pCLSID As UUID, lpszProgID As Long) As Long

      Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As olelib.UUID, lpszProgID As Long) As Long

      Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As olelib.UUID) As Long
         Private Declare Sub CLSIDFromProgIDEx Lib "ole32.dll" (ByVal lpszProgID As Long, ByVal lpclsid As Long)

Private Declare Function GetSysColor Lib "User32.dll" (ByVal nIndex As Long) As Long

      Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function LoadIcon Lib "User32.dll" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function ExtractAssociatedIcon Lib "Shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, ByRef lpiIcon As Long) As Long
Private Declare Function ExtractIcon Lib "Shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function ExtractIconEx Lib "Shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Const CLSCTX_INPROC_SERVER As Long = 1
Private Const CLSCTX_INPROC_HANDLER As Long = 2
    
    Private Const CLSCTX_INPROC As Long = (CLSCTX_INPROC_SERVER Or CLSCTX_INPROC_HANDLER)
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Sub CoCreateInstance Lib "ole32.dll" (ByVal rclsid As Long, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByVal riid As Long, ByRef ppv As Any)
     

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DllGetVersion Lib "COMCTL32" (pdvi As DLLVERSIONINFO) As Long
Private Const S_OK = 0



Public Function GenerateGUID() As String
    Dim GUIDuse As UUID
    Dim retVal As String, lngret As Long
    CoCreateGuid GUIDuse
        Call StringFromCLSID(GUIDuse, lngret)
        StringFromPointer lngret, retVal
        
    GenerateGUID = retVal
End Function
Public Function GetCLSID(ByVal progID As String) As String
         Dim strProgID As String * 255
         Dim pprogid As Long
         Dim udtclsid As UUID
         Dim strCLSID As String * 255
         Dim pCLSID As Long
         Dim lngret As Long
         Dim strTemp As String
         Dim i As Integer
         strTemp = progID
         'Take a ProgID.
         

         'Get CLSID.
         lngret = CLSIDFromProgID(StrPtr(strTemp), udtclsid)

         'Display CLSID elements.
       

         'Convert CLSID to a string and get the pointer back.
         lngret = StringFromCLSID(udtclsid, pCLSID)

         'Get the CLSID string and display it.
         StringFromPointer pCLSID, strCLSID
         GetCLSID = Trim$(Replace$(strCLSID, vbNullChar, ""))

         'Reinitialize the CLSID.
         With udtclsid
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            For i = 0 To 7
               .Data4(i) = 0
            Next
         End With

         
      End Function
Public Function GetProgID(ByVal CLSID As String) As String
    'retrieve the progID of a given CLSID
    'use the OLE API.
    'it's actually a LOT more annoying then it should be.
    'algorithm:
    'first, strip out the special formatting characters, such as "{","}" and "-"
    'then,useing this string, iterate through every two characters, converting the two values into a Hexedecimal Byte and storing it int oa byte array.
    'use the MemoryCopy Function to copy this Byte Data into the GUID variable.
    '(So far, it appears that the resulting values for the first DWORD,WORD, and WORD have somehow been swapped.
    'for example, a guid of:
    '       {22BB2698-0904-4D38-B340-6E10B0D5A240}
    '   would, after turning into a string and back again:
    '       {9826BB22-0409-384D-B340-6E10B0D5A240}
    'I realize I obviously made a mistake (unless there is a bug in OLE32.dll, which is about a 0.001% chance)
    'but right now I simply call a simple little stub that swaps them back into the correct order.
    'make a call to the StringFromCLSID() function, using this created GUID
    Dim originalinput As String
    Dim strTemp As String, sRet As String
    Dim Spointer As Long
    Dim useme As UUID
    Dim TmpLng As Long, lngtest As Long
    Dim strtest As String, lngret As Long
    Dim TmpInt As Long
    Dim i As Long, Element As Long
    Dim ByteArr(16) As Byte 'used to copy GUID structure.
    originalinput = CLSID
    On Error GoTo 0
    'this is what we need to change.
    'this data is not valid. we need to get the actual GUID for the CLSID.
    With useme
            .Data1 = 0
            .Data2 = 0
            .Data3 = 0
            For i = 0 To 7
               .Data4(i) = 0
            Next
         End With
    
    strTemp = CLSID
    strTemp = Replace$(strTemp, "{", "")
    strTemp = Replace$(strTemp, "}", "")
    strTemp = Replace$(strTemp, "-", "")
    'go through every WORD, or byte, and copy the denoted hex value.
    Element = 0
    For i = 1 To Len(strTemp) Step 2
    'for some reason, all sections except the last one are frigged. they seems the bytes are reversed.
        ByteArr(Element) = Val("&H" & Mid$(strTemp, i, 2))
        'Debug.Print Mid$(strTemp, I, 2);
        Element = Element + 1
    Next i
    FixBytes ByteArr()
  
    'fix the bytes. for some reason, my code screws it up or something.
    'now that we have the bytes of the data, we can do a Copymemory into the structure.
    CopyMemory useme, ByteArr(0), Len(useme)
    'phew. hope that works...
    '{Data1-Data2-Data3-Data4}
    ' DWORD-WORD-WORD-WORD-WORD & DWORD
    lngtest = 0
    Call StringFromCLSID(useme, lngtest)
    
    Call StringFromPointer(lngtest, strtest)
    strtest = Replace(strtest, vbNullChar, "")
    strtest = Trim$(strtest)
    'Assert they are the same...
    Debug.Assert strtest = originalinput
   
   lngret = progIDfromCLSID(useme, Spointer)
   'create each succesive value, and place it into the byte array.
   
   
   'spointer is pointer to the string data.
   sRet = Space$(255)
   Call StringFromPointer(Spointer, sRet)
   GetProgID = Trim$(Replace$(sRet, vbNullChar, ""))
   'return the string.
Exit Function
   
   Err.Raise 9, "modShell.GetProgID", "Invalid CLSID string format."
  '  Resume
End Function

Private Sub FixBytes(Bytesfix() As Byte)
Dim i As Long
Dim tmpbyte As Byte
'fixes the bytes by swapping the following:
'Byte 1 withbyte 4
'byte 2 with byte 3
'byte  5 with byte 6
'byte 7 with byte 8
For i = 0 To 1
    tmpbyte = Bytesfix(i)
    Bytesfix(i) = Bytesfix(3 - i)
    Bytesfix(3 - i) = tmpbyte
Next i
'that fixes the first DWORD value.
'now, the two words....
tmpbyte = Bytesfix(4)
Bytesfix(4) = Bytesfix(5)
Bytesfix(5) = tmpbyte

tmpbyte = Bytesfix(6)
Bytesfix(6) = Bytesfix(7)
Bytesfix(7) = tmpbyte



End Sub
Public Sub StringFromPointer(pOLESTR As Long, strOut As String)
         Dim ByteArray(255) As Byte
         Dim intTemp As Integer
         Dim intCount As Integer
         Dim i As Integer

         intTemp = 1
    
         'Walk the string and retrieve the first byte of each WORD.
         While intTemp <> 0
            CopyMemory intTemp, ByVal pOLESTR + i, 2
            ByteArray(intCount) = intTemp
            
            intCount = intCount + 1
            i = i + 2
         Wend

         'Copy the byte array to our string.
         strOut = Space$(255)
         CopyMemory ByVal strOut, ByteArray(0), intCount
      End Sub











Public Function BitsPerPixel() As Long
Dim lhDCD As Long
Dim lBitsPixel As Long
Dim lPlanes As Long
   lhDCD = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lBitsPixel = GetDeviceCaps(lhDCD, BITSPIXEL)
   lPlanes = GetDeviceCaps(lhDCD, PLANES)
   BitsPerPixel = (lBitsPixel * lPlanes)
   DeleteDC lhDCD
End Function

Public Function ComCtlVersion( _
        ByRef lMajor As Long, _
        ByRef lMinor As Long, _
        Optional ByRef lBuild As Long _
    ) As Boolean
Dim hMod As Long
Dim lR As Long
Dim lptrDLLVersion As Long
Dim tDVI As DLLVERSIONINFO

    lMajor = 0: lMinor = 0: lBuild = 0

    hMod = LoadLibrary("comctl32.dll")
    If Not (hMod = 0) Then
        lR = S_OK
        ' You must get this function explicitly because earlier versions
        ' of the DLL don't implement this function. That makes the
        ' lack of implementation of the function a version
        ' marker in itself.
        lptrDLLVersion = GetProcAddress(hMod, "DllGetVersion")
        If Not (lptrDLLVersion = 0) Then
            tDVI.cbSize = Len(tDVI)
            lR = DllGetVersion(tDVI)
            If (lR = S_OK) Then
                lMajor = tDVI.dwMajor
                lMinor = tDVI.dwMinor
                lBuild = tDVI.dwBuildNumber
            End If
        Else
            'If GetProcAddress failed, then the DLL is a
            ' version previous to the one shipped with IE 3.x.
            lMajor = 4
        End If
        FreeLibrary hMod
        ComCtlVersion = True
    End If

End Function

Public Function SupportsAlphaIcons() As Boolean
Static cached As Boolean
Static bSupportsAlphaIcons As Boolean
If cached Then
    
Else
    
   If (BitsPerPixel >= 32) Then
      Dim lMajor As Long
      Dim lMinor As Long
      ComCtlVersion lMajor, lMinor
      If (lMajor >= 6) Then
         bSupportsAlphaIcons = True
      End If
   End If
   cached = True
End If
SupportsAlphaIcons = bSupportsAlphaIcons
End Function
'Public Function IID_IShellDetails() As olelib.UUID
'  Static iid As olelib.UUID
'  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214EC, 0, 0)
'  IID_IShellDetails = iid
'
'End Function
