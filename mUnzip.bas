Attribute VB_Name = "mUnzip"
Option Explicit

' ======================================================================================
' Name:     mUnzip
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2000
'
' Requires: Info-ZIP's Unvbuzip10.dll v5.40, renamed to vbuzip10.dll
'           cUnzip.cls
'
' Copyright © 2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Part of the implementation of cUnzip.cls, a class which gives a
' simple interface to Info-ZIP's excellent, free unzipping library
' (Unvbuzip10.dll).
'
' This sample uses decompression code by the Info-ZIP group.  The
' original Info-Zip sources are freely available from their website
' at
'     http://www.cdrcom.com/pubs/infozip/
'
' Please ensure you visit the site and read their free source licensing
' information and requirements before using their code in your own
' application.
'
' ======================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' argv
Private Type UNZIPnames
    s(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 32800) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

' DCL structure
Public Type DCLIST
   ExtractOnlyNewer As Long      ' 1 to extract only newer
   SpaceToUnderScore As Long     ' 1 to convert spaces to underscore
   PromptToOverwrite As Long     ' 1 if overwriting prompts required
   fQuiet As Long                ' 0 = all messages, 1 = few messages, 2 = no messages
   ncflag As Long                ' write to stdout if 1
   ntflag As Long                ' test zip file
   nvflag As Long                ' verbose listing
   nUflag As Long                ' "update" (extract only newer/new files)
   nzflag As Long                ' display zip file comment
   ndflag As Long                ' all args are files/dir to be extracted
   noflag As Long                ' 1 if always overwrite files
   naflag As Long                ' 1 to do end-of-line translation
   nZIflag As Long               ' 1 to get zip info
   C_flag As Long                ' 1 to be case insensitive
   fPrivilege As Long            ' zip file name
   lpszZipFN As String           ' directory to extract to.
   lpszExtractDir As String
End Type

Private Type USERFUNCTION
   ' Callbacks:
   lptrPrnt As Long           ' Pointer to application's print routine
   lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
   lptrReplace As Long        ' Pointer to application's replace routine.
   lptrPassword As Long       ' Pointer to application's password routine.
   lptrMessage As Long        ' Pointer to application's routine for
                              ' displaying information about specific files in the archive
                              ' used for listing the contents of the archive.
   lptrService As Long        ' callback function designed to be used for allowing the
                              ' app to process Windows messages, or cancelling the operation
                              ' as well as giving option of progress.  If this function returns
                              ' non-zero, it will terminate what it is doing.  It provides the app
                              ' with the name of the archive member it has just processed, as well
                              ' as the original size.
                              
   ' Values filled in after processing:
   lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                              ' the archive header and central directory list.
   lTotalSize As Long         ' Total size of all files in the archive
   lCompFactor As Long        ' Overall archive compression factor
   lNumMembers As Long        ' Total number of files in the archive
   cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

Public Type ZIPVERSIONTYPE
   major As Byte
   minor As Byte
   patchlevel As Byte
   not_used As Byte
End Type

Public Type UZPVER
    structlen As Long         ' Length of structure
    flag As Long              ' 0 is beta, 1 uses zlib
    betalevel As String * 10  ' e.g "g BETA"
    date As String * 20       ' e.g. "4 Sep 95" (beta) or "4 September 1995"
    zlib As String * 10       ' e.g. "1.0.5 or NULL"
    Unzip As ZIPVERSIONTYPE
    zipinfo As ZIPVERSIONTYPE
    os2dll As ZIPVERSIONTYPE
    windll As ZIPVERSIONTYPE
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "vbuzip10.dll" _
  (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
   ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
   dcll As DCLIST, Userf As USERFUNCTION) As Long
Public Declare Sub UzpVersion2 Lib "vbuzip10.dll" (uzpv As UZPVER)

' Object for callbacks:
Private m_cUnzip As cUnzip
Private m_bCancel As Boolean





'unvbuzip10.dll and vbuzip10.dll.
'-- C Style argv
Public Type UNZIPnames
  uzFiles(0 To 99) As String
End Type

'-- Callback Large "String"
Public Type UNZIPCBChar
  ch(32800) As Byte
End Type

'-- Callback Small "String"
Public Type UNZIPCBCh
  ch(256) As Byte
End Type

'-- UNvbuzip10.dll DCL Structure
Public Type DCLIST
  ExtractOnlyNewer  As Long    ' 1 = Extract Only Newer, Else 0
  SpaceToUnderScore As Long    ' 1 = Convert Space To Underscore, Else 0
  PromptToOverwrite As Long    ' 1 = Prompt To Overwrite Required, Else 0
  fQuiet            As Long    ' 2 = No Messages, 1 = Less, 0 = All
  ncflag            As Long    ' 1 = Write To Stdout, Else 0
  ntflag            As Long    ' 1 = Test Zip File, Else 0
  nvflag            As Long    ' 0 = Extract, 1 = List Zip Contents
  nUflag            As Long    ' 1 = Extract Only Newer, Else 0
  nzflag            As Long    ' 1 = Display Zip File Comment, Else 0
  ndflag            As Long    ' 1 = Honor Directories, Else 0
  noflag            As Long    ' 1 = Overwrite Files, Else 0
  naflag            As Long    ' 1 = Convert CR To CRLF, Else 0
  nZIflag           As Long    ' 1 = Zip Info Verbose, Else 0
  C_flag            As Long    ' 1 = Case Insensitivity, 0 = Case Sensitivity
  fPrivilege        As Long    ' 1 = ACL, 2 = Privileges
  Zip               As String  ' The Zip Filename To Extract Files
  ExtractDir        As String  ' The Extraction Directory, NULL If Extracting To Current Dir
End Type

'-- UNvbuzip10.dll Userfunctions Structure
Public Type USERFUNCTION
  UZDLLPrnt     As Long     ' Pointer To Apps Print Function
  UZDLLSND      As Long     ' Pointer To Apps Sound Function
  UZDLLREPLACE  As Long     ' Pointer To Apps Replace Function
  UZDLLPASSWORD As Long     ' Pointer To Apps Password Function
  UZDLLMESSAGE  As Long     ' Pointer To Apps Message Function
  UZDLLSERVICE  As Long     ' Pointer To Apps Service Function (Not Coded!)
  TotalSizeComp As Long     ' Total Size Of Zip Archive
  TotalSize     As Long     ' Total Size Of All Files In Archive
  CompFactor    As Long     ' Compression Factor
  NumMembers    As Long     ' Total Number Of All Files In The Archive
  cchComment    As Integer  ' Flag If Archive Has A Comment!
End Type

'-- UNvbuzip10.dll Version Structure
Public Type UZPVER
  structlen       As Long         ' Length Of The Structure Being Passed
  flag            As Long         ' Bit 0: is_beta  bit 1: uses_zlib
  beta            As String * 10  ' e.g., "g BETA" or ""
  date            As String * 20  ' e.g., "4 Sep 95" (beta) or "4 September 1995"
  zlib            As String * 10  ' e.g., "1.0.5" or NULL
  Unzip(1 To 4)   As Byte         ' Version Type Unzip
  zipinfo(1 To 4) As Byte         ' Version Type Zip Info
  os2dll          As Long         ' Version Type OS2 DLL
  windll(1 To 4)  As Byte         ' Version Type Windows DLL
End Type

'-- This Assumes UNvbuzip10.dll Is In Your \Windows\System Directory!
'Private Declare Function Wiz_SingleEntryUnzip Lib "unvbuzip10.dll" _
  (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, _
   ByVal xfnc As Long, ByRef xfnv As UNZIPnames, _
   dcll As DCLIST, Userf As USERFUNCTION) As Long

'Private Declare Sub UzpVersion2 Lib "unvbuzip10.dll" (uzpv As UZPVER)

'argv
Public Type ZIPnames
    s(0 To 99) As String
End Type

'ZPOPT is used to set options in the vbuzip10.dll
Private Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type

Private Type ZIPUSERFUNCTIONS
    DLLPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type

'Structure ZCL - not used by VB
'Private Type ZCL
'    argc As Long            'number of files
'    filename As String      'Name of the Zip file
'    fileArray As ZIPnames   'The array of filenames
'End Type

' Call back "string" (sic)
Private Type CBChar
    ch(4096) As Byte
End Type

'Local declares

' Dim MYZCL As ZCL


'This assumes vbuzip10.dll is on your path.
Private Declare Function ZpInit Lib "vbuzip10.dll" _
(ByRef Zipfun As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks

Private Declare Function ZpSetOptions Lib "vbuzip10.dll" _
(ByRef Opts As ZPOPT) As Long ' Set Zip options

Private Declare Function ZpGetOptions Lib "vbuzip10.dll" _
() As ZPOPT ' used to check encryption flag only

Private Declare Function ZpArchive Lib "vbuzip10.dll" _
(ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private uZipNumber As Integer
Private uZipMessage As String
Private uZipInfo As String
Private uVBSkip As Integer
Public msOutput As String


' Puts a function pointer in a structure
Function FnPtr(ByVal lp As Long) As Long
    FnPtr = lp
End Function

' Callback for vbuzip10.dll
Function DLLPrnt(ByRef fname As CBChar, ByVal x As Long) As Long
    Dim s0$, xx As Long
    Dim sVbZipInf As String
    
    ' always put this in callback routines!
    On Error Resume Next
    s0 = ""
    For xx = 0 To x
        If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 + Chr(fname.ch(xx))
    Next xx
    
    Debug.Print sVbZipInf & s0
    msOutput = msOutput & s0
    
    sVbZipInf = ""
    
    DoEvents
    DLLPrnt = 0
    
End Function

' Callback for vbuzip10.dll ?
Function DllServ(ByRef fname As CBChar, ByVal x As Long) As Long
    
    Dim s0 As String
    Dim xx As Long
    
    On Error Resume Next
    
    s0 = ""
    
    For xx = 0 To x - 1
        If fname.ch(xx) = 0 Then Exit For
        s0 = s0 & Chr$(fname.ch(xx))
    Next
    
    DllServ = 0
End Function

' Callback for vbuzip10.dll
Function DllPass(ByRef s1 As Byte, x As Long, _
    ByRef s2 As Byte, _
    ByRef s3 As Byte) As Long

    ' always put this in callback routines!
    On Error Resume Next
    ' not supported - always return 1
    DllPass = 1
End Function

' Callback for vbuzip10.dll
Function DllComm(ByRef s1 As CBChar) As CBChar
    
    ' always put this in callback routines!
    On Error Resume Next
    ' not supported always return \0
    s1.ch(0) = vbNullString
    DllComm = s1
End Function

'Main Subroutine
Public Function VBZip(argc As Integer, zipname As String, _
        mynames As ZIPnames, junk As Integer, _
        recurse As Integer, updat As Integer, _
        freshen As Integer, basename As String, _
        Optional Encrypt As Integer = 0, _
        Optional IncludeSystem As Integer = 0, _
        Optional IgnoreDirectoryEntries As Integer = 0, _
        Optional Verbose As Integer = 0, _
        Optional Quiet As Integer = 0, _
        Optional CRLFtoLF As Integer = 0, _
        Optional LFtoCRLF As Integer = 0, _
        Optional Grow As Integer = 0, _
        Optional Force As Integer = 0, _
        Optional iMove As Integer = 0, _
        Optional DeleteEntries As Integer = 0) As Long
    
    Dim hmem As Long, xx As Integer
    Dim retcode As Long
    Dim MYUSER As ZIPUSERFUNCTIONS
    Dim MYOPT As ZPOPT
    
    On Error Resume Next ' nothing will go wrong :-)
    
    msOutput = ""
    
    ' Set address of callback functions
    MYUSER.DLLPrnt = FnPtr(AddressOf DLLPrnt)
    MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
    MYUSER.DLLCOMMENT = FnPtr(AddressOf DllComm)
    MYUSER.DLLSERVICE = 0& ' not coded yet :-)
'    retcode = ZpInit(MYUSER)
    
    ' Set zip options
    MYOPT.fSuffix = 0        ' include suffixes (not yet implemented)
    MYOPT.fEncrypt = Encrypt     ' 1 if encryption wanted
    MYOPT.fSystem = IncludeSystem        ' 1 to include system/hidden files
    MYOPT.fVolume = 0        ' 1 if storing volume label
    MYOPT.fExtra = 0         ' 1 if including extra attributes
    MYOPT.fNoDirEntries = IgnoreDirectoryEntries  ' 1 if ignoring directory entries
    MYOPT.fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
    MYOPT.fIncludeDate = 0   ' 1 if including files earlier than a specified date
    MYOPT.fVerbose = Verbose       ' 1 if full messages wanted
    MYOPT.fQuiet = Quiet         ' 1 if minimum messages wanted
    MYOPT.fCRLF_LF = CRLFtoLF        ' 1 if translate CR/LF to LF
    MYOPT.fLF_CRLF = LFtoCRLF ' 1 if translate LF to CR/LF
    MYOPT.fJunkDir = junk    ' 1 if junking directory names
    MYOPT.fRecurse = recurse ' 1 if recursing into subdirectories
    MYOPT.fGrow = Grow          ' 1 if allow appending to zip file
    MYOPT.fForce = Force         ' 1 if making entries using DOS names
    MYOPT.fMove = iMove          ' 1 if deleting files added or updated
    MYOPT.fDeleteEntries = DeleteEntries ' 1 if files passed have to be deleted
    MYOPT.fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
    MYOPT.fFreshen = freshen ' 1 if freshening zip file--overwrite only
    MYOPT.fJunkSFX = 0       ' 1 if junking sfx prefix
    MYOPT.fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
    MYOPT.fComment = 0       ' 1 if putting comment in zip file
    MYOPT.fOffsets = 0       ' 1 if updating archive offsets for sfx Files
    MYOPT.fPrivilege = 0     ' 1 if not saving privelages
    MYOPT.fEncryption = 0    'Read only property!
    MYOPT.fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
    MYOPT.flevel = 0         ' compression level - should be 0!!!
    MYOPT.date = vbNullString ' "12/31/79"? US Date?
    MYOPT.szRootDir = UCase$(basename)
    
    retcode = ZpInit(MYUSER)
    ' Set options
    retcode = ZpSetOptions(MYOPT)
    
    ' ZCL not needed in VB
    ' MYZCL.argc = 2
    ' MYZCL.filename = "c:\wiz\new.zip"
    ' MYZCL.fileArray = MYNAMES
    
    ' Go for it!
    
    retcode = ZpArchive(argc, zipname, mynames)
    
    VBZip = retcode
End Function



'-- Callback For UNvbuzip10.dll - Receive Message Function
Public Sub UZReceiveDLLMessage(ByVal ucsize As Long, _
    ByVal csiz As Long, _
    ByVal cfactor As Integer, _
    ByVal mo As Integer, _
    ByVal dy As Integer, _
    ByVal yr As Integer, _
    ByVal hh As Integer, _
    ByVal mm As Integer, _
    ByVal c As Byte, ByRef fname As UNZIPCBCh, _
    ByRef meth As UNZIPCBCh, ByVal crc As Long, _
    ByVal fCrypt As Byte)

  Dim s0     As String
  Dim xx     As Long
  Dim strout As String * 80

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  '------------------------------------------------
  '-- This Is Where The Received Messages Are
  '-- Printed Out And Displayed.
  '-- You Can Modify Below!
  '------------------------------------------------

  strout = Space(80)

  '-- For Zip Message Printing
  If uZipNumber = 0 Then
    Mid(strout, 1, 50) = "Filename:"
    Mid(strout, 53, 4) = "Size"
    Mid(strout, 62, 4) = "Date"
    Mid(strout, 71, 4) = "Time"
    uZipMessage = strout & vbNewLine
    strout = Space(80)
  End If

  s0 = ""

  '-- Do Not Change This For Next!!!
  For xx = 0 To 255
    If fname.ch(xx) = 0 Then Exit For
    s0 = s0 & Chr(fname.ch(xx))
  Next

  '-- Assign Zip Information For Printing
  Mid(strout, 1, 50) = Mid(s0, 1, 50)
  Mid(strout, 51, 7) = Right("        " & str(ucsize), 7)
  Mid(strout, 60, 3) = Right("0" & Trim(str(mo)), 2) & "/"
  Mid(strout, 63, 3) = Right("0" & Trim(str(dy)), 2) & "/"
  Mid(strout, 66, 2) = Right("0" & Trim(str(yr)), 2)
  Mid(strout, 70, 3) = Right(str(hh), 2) & ":"
  Mid(strout, 73, 2) = Right("0" & Trim(str(mm)), 2)

  ' Mid(strout, 75, 2) = Right(" " & Str(cfactor), 2)
  ' Mid(strout, 78, 8) = Right("        " & Str(csiz), 8)
  ' s0 = ""
  ' For xx = 0 To 255
  '     If meth.ch(xx) = 0 Then exit for
  '     s0 = s0 & Chr(meth.ch(xx))
  ' Next xx

  '-- Do Not Modify Below!!!
  uZipMessage = uZipMessage & strout & vbNewLine
  uZipNumber = uZipNumber + 1

End Sub

'-- Callback For UNvbuzip10.dll - Print Message Function
Public Function UZDLLPrnt(ByRef fname As UNZIPCBChar, ByVal x As Long) As Long

  Dim s0 As String
  Dim xx As Long

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  s0 = ""

  '-- Gets The UNvbuzip10.dll Message For Displaying.
  For xx = 0 To x - 1
    If fname.ch(xx) = 0 Then Exit For
    s0 = s0 & Chr(fname.ch(xx))
  Next

  '-- Assign Zip Information
  If Mid$(s0, 1, 1) = vbLf Then s0 = vbNewLine ' Damn UNIX :-)
  uZipInfo = uZipInfo & s0

msOutput = uZipInfo
    
  UZDLLPrnt = 0

End Function

'-- Callback For UNvbuzip10.dll - DLL Service Function
Public Function UZDLLServ(ByRef mname As UNZIPCBChar, ByVal x As Long) As Long

    Dim s0 As String
    Dim xx As Long
    
    '-- Always Put This In Callback Routines!
    On Error Resume Next
    
    s0 = ""
    '-- Get vbuzip10.dll Message For processing
    For xx = 0 To x - 1
        If mname.ch(xx) = 0 Then Exit For
        s0 = s0 + Chr(mname.ch(xx))
    Next
    ' At this point, s0 contains the message passed from the DLL
    ' It is up to the developer to code something useful here :)
    UZDLLServ = 0 ' Setting this to 1 will abort the zip!

End Function

'-- Callback For UNvbuzip10.dll - Password Function
Public Function UZDLLPass(ByRef p As UNZIPCBCh, _
  ByVal n As Long, ByRef m As UNZIPCBCh, _
  ByRef Name As UNZIPCBCh) As Integer

  Dim prompt     As String
  Dim xx         As Integer
  Dim szpassword As String

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  UZDLLPass = 1

  If uVBSkip = 1 Then Exit Function

  '-- Get The Zip File Password
  szpassword = InputBox("Please Enter The Password!")

  '-- No Password So Exit The Function
  If szpassword = "" Then
    uVBSkip = 1
    Exit Function
  End If

  '-- Zip File Password So Process It
  For xx = 0 To 255
    If m.ch(xx) = 0 Then
      Exit For
    Else
      prompt = prompt & Chr(m.ch(xx))
    End If
  Next

  For xx = 0 To n - 1
    p.ch(xx) = 0
  Next

  For xx = 0 To Len(szpassword) - 1
    p.ch(xx) = Asc(Mid(szpassword, xx + 1, 1))
  Next

  p.ch(xx) = Chr(0) ' Put Null Terminator For C

  UZDLLPass = 0

End Function

'-- Callback For UNvbuzip10.dll - Report Function To Overwrite Files.
'-- This Function Will Display A MsgBox Asking The User
'-- If They Would Like To Overwrite The Files.
Public Function UZDLLRep(ByRef fname As UNZIPCBChar) As Long

  Dim s0 As String
  Dim xx As Long

  '-- Always Put This In Callback Routines!
  On Error Resume Next

  UZDLLRep = 100 ' 100 = Do Not Overwrite - Keep Asking User
  s0 = ""

  For xx = 0 To 255
    If fname.ch(xx) = 0 Then xx = 99999 Else s0 = s0 & Chr(fname.ch(xx))
  Next

  '-- This Is The MsgBox Code
  xx = MsgBox("Overwrite " & s0 & "?", vbExclamation & vbYesNoCancel, _
              "VBUnZip32 - File Already Exists!")

  If xx = vbNo Then Exit Function

  If xx = vbCancel Then
    UZDLLRep = 104       ' 104 = Overwrite None
    Exit Function
  End If

  UZDLLRep = 102         ' 102 = Overwrite 103 = Overwrite All

End Function

'-- ASCIIZ To String Function
Public Function szTrim(szString As String) As String
    
    Dim pos As Integer
    Dim ln  As Integer
    
    pos = InStr(szString, Chr(0))
    ln = Len(szString)
    
    Select Case pos
        Case Is > 1
            szTrim = Trim(Left(szString, pos - 1))
        Case 1
            szTrim = ""
        Case Else
            szTrim = Trim(szString)
    End Select

End Function


Public Function VBUnzip(ByRef sZipFileName, ByRef sUnzipDirectory As String, _
    ByRef iExtractNewer As Integer, _
    ByRef iSpaceUnderScore As Integer, _
    ByRef iPromptOverwrite As Integer, _
    ByRef iQuiet As Integer, _
    ByRef iWriteStdOut As Integer, _
    ByRef iTestZip As Integer, _
    ByRef iExtractList As Integer, _
    ByRef iExtractOnlyNewer As Integer, _
    ByRef iDisplayComment As Integer, _
    ByRef iHonorDirectories As Integer, _
    ByRef iOverwriteFiles As Integer, _
    ByRef iConvertCR_CRLF As Integer, _
    ByRef iVerbose As Integer, _
    ByRef iCaseSensitivty As Integer, _
    ByRef iPrivilege As Integer) As Long


On Error GoTo vbErrorHandler

    
    Dim lRet As Long
    
    Dim UZDCL As DCLIST
    Dim UZUSER As USERFUNCTION
    Dim UZVER As UZPVER
    Dim uExcludeNames As UNZIPnames
    Dim uZipNames     As UNZIPnames
    
    msOutput = ""
    
    uExcludeNames.uzFiles(0) = vbNullString
    uZipNames.uzFiles(0) = vbNullString
    
    uZipNumber = 0
    uZipMessage = vbNullString
    uZipInfo = vbNullString
    uVBSkip = 0
    
    With UZDCL
        .ExtractOnlyNewer = iExtractOnlyNewer
        .SpaceToUnderScore = iSpaceUnderScore
        .PromptToOverwrite = iPromptOverwrite
        .fQuiet = iQuiet
        .ncflag = iWriteStdOut
        .ntflag = iTestZip
        .nvflag = iExtractList
        .nUflag = iExtractNewer
        .nzflag = iDisplayComment
        .ndflag = iHonorDirectories
        .noflag = iOverwriteFiles
        .naflag = iConvertCR_CRLF
        .nZIflag = iVerbose
        .C_flag = iCaseSensitivty
        .fPrivilege = iPrivilege
        .Zip = sZipFileName
        .ExtractDir = sUnzipDirectory
    End With
    
    With UZUSER
        .UZDLLPrnt = FnPtr(AddressOf UZDLLPrnt)
        .UZDLLSND = 0&
        .UZDLLREPLACE = FnPtr(AddressOf UZDLLRep)
        .UZDLLPASSWORD = FnPtr(AddressOf UZDLLPass)
        .UZDLLMESSAGE = FnPtr(AddressOf UZReceiveDLLMessage)
        .UZDLLSERVICE = FnPtr(AddressOf UZDLLServ)
    End With
    
    With UZVER
        .structlen = Len(UZVER)
        .beta = Space$(9) & vbNullChar
        .date = Space$(19) & vbNullChar
        .zlib = Space$(9) & vbNullChar
    End With
    
    UzpVersion2 UZVER
    
    lRet = Wiz_SingleEntryUnzip(0, uZipNames, 0, uExcludeNames, UZDCL, UZUSER)
    VBUnzip = lRet
    

    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "CodeModule::VBUnzip", Err.Description

End Function







Private Function plAddressOf(ByVal lPtr As Long) As Long
   ' VB Bug workaround fn
   plAddressOf = lPtr
End Function

Private Sub UnzipMessageCallBack( _
      ByVal ucsize As Long, _
      ByVal csiz As Long, _
      ByVal cfactor As Integer, _
      ByVal mo As Integer, _
      ByVal dy As Integer, _
      ByVal yr As Integer, _
      ByVal hh As Integer, _
      ByVal mm As Integer, _
      ByVal c As Byte, _
      ByRef fname As CBCh, _
      ByRef meth As CBCh, _
      ByVal crc As Long, _
      ByVal fCrypt As Byte _
   )
Dim sFileName As String
Dim sFolder As String
Dim dDate As Date
Dim sMethod As String
Dim iPos As Long

   On Error Resume Next
    
   ' Add to unzip class:
   With m_cUnzip
      ' Parse:
      sFileName = StrConv(fname.ch, vbUnicode)
      ParseFileFolder sFileName, sFolder
      dDate = DateSerial(yr, mo, hh)
      dDate = dDate + TimeSerial(hh, mm, 0)
      sMethod = StrConv(meth.ch, vbUnicode)
      iPos = InStr(sMethod, vbNullChar)
      If (iPos > 1) Then
         sMethod = Left$(sMethod, iPos - 1)
      End If
    
      Debug.Print fCrypt
      .DirectoryListAddFile sFileName, sFolder, dDate, csiz, crc, ((fCrypt And 64) = 64), cfactor, sMethod
   End With
   
End Sub

Private Function UnzipPrintCallback( _
      ByRef fname As CBChar, _
      ByVal x As Long _
   ) As Long
Dim iPos As Long
Dim sFIle As String
   On Error Resume Next
   
   ' Check we've got a message:
   If x > 1 And x < 1024 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), fname, x
      ' Convert to VB string:
      sFIle = StrConv(b, vbUnicode)
      
      ' Fix up backslashes:
      ReplaceSection sFIle, "/", "\"
      
      ' Tell the caller about it
      m_cUnzip.ProgressReport sFIle
   End If
   UnzipPrintCallback = 0
End Function

Private Function UnzipPasswordCallBack( _
      ByRef pwd As CBCh, _
      ByVal x As Long, _
      ByRef s2 As CBCh, _
      ByRef Name As CBCh _
   ) As Long

Dim bCancel As Boolean
Dim sPassword As String
Dim b() As Byte
Dim lSize As Long

On Error Resume Next

   ' The default:
   UnzipPasswordCallBack = 1
    
   If m_bCancel Then
      Exit Function
   End If
   
   ' Ask for password:
   m_cUnzip.PasswordRequest sPassword, bCancel
      
   sPassword = Trim$(sPassword)
   
   ' Cancel out if no useful password:
   If bCancel Or Len(sPassword) = 0 Then
      m_bCancel = True
      Exit Function
   End If
   
   ' Put password into return parameter:
   lSize = Len(sPassword)
   If lSize > 254 Then
      lSize = 254
   End If
   b = StrConv(sPassword, vbFromUnicode)
   CopyMemory pwd.ch(0), b(0), lSize
   
   ' Ask UnZip to process it:
   UnzipPasswordCallBack = 0
       
End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long
Dim eResponse As EUZOverWriteResponse
Dim iPos As Long
Dim sFIle As String

   On Error Resume Next
   eResponse = euzDoNotOverwrite
   
   ' Extract the filename:
   sFIle = StrConv(fname.ch, vbUnicode)
   iPos = InStr(sFIle, vbNullChar)
   If (iPos > 1) Then
      sFIle = Left$(sFIle, iPos - 1)
   End If
   
   ' No backslashes:
   ReplaceSection sFIle, "/", "\"
   
   ' Request the overwrite request:
   m_cUnzip.OverwriteRequest sFIle, eResponse
   
   ' Return it to the zipping lib
   UnzipReplaceCallback = eResponse
   
End Function
Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
Dim iPos As Long
Dim sInfo As String
Dim bCancel As Boolean
    
'-- Always Put This In Callback Routines!
On Error Resume Next
    
   ' Check we've got a message:
   If x > 1 And x < 1024 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), mname, x
      ' Convert to VB string:
      sInfo = StrConv(b, vbUnicode)
      iPos = InStr(sInfo, vbNullChar)
      If iPos > 0 Then
         sInfo = Left$(sInfo, iPos - 1)
      End If
      ReplaceSection sInfo, "\", "/"
      m_cUnzip.Service sInfo, bCancel
      If bCancel Then
         UnZipServiceCallback = 1
      Else
         UnZipServiceCallback = 0
      End If
   End If
   
End Function



Private Sub ParseFileFolder( _
      ByRef sFileName As String, _
      ByRef sFolder As String _
   )
Dim iPos As Long
Dim iLastPos As Long

   iPos = InStr(sFileName, vbNullChar)
   If (iPos <> 0) Then
      sFileName = Left$(sFileName, iPos - 1)
   End If
   
   iLastPos = ReplaceSection(sFileName, "/", "\")
   
   If (iLastPos > 1) Then
      sFolder = Left$(sFileName, iLastPos - 2)
      sFileName = Mid$(sFileName, iLastPos)
   End If
   
End Sub
Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Long
Dim iPos As Long
Dim iLastPos As Long
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, sString, "/")
      If (iPos > 1) Then
         Mid$(sString, iPos, 1) = "\"
         iLastPos = iPos + 1
      End If
   Loop While Not (iPos = 0)
   ReplaceSection = iLastPos

End Function

' Main subroutine
Public Function VBUnzip( _
      cUnzipObject As cUnzip, _
      tDCL As DCLIST, _
      iIncCount As Long, _
      sInc() As String, _
      iExCount As Long, _
      sExc() As String _
   ) As Long
Dim tUser As USERFUNCTION
Dim lR As Long
Dim tInc As UNZIPnames
Dim tExc As UNZIPnames
Dim i As Long

On Error GoTo ErrorHandler

   Set m_cUnzip = cUnzipObject
   ' Set Callback addresses
   tUser.lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
   tUser.lptrSound = 0& ' not supported
   tUser.lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
   tUser.lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
   tUser.lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
   tUser.lptrService = plAddressOf(AddressOf UnZipServiceCallback)
        
   ' Set files to include/exclude:
   If (iIncCount > 0) Then
      For i = 1 To iIncCount
         tInc.s(i - 1) = sInc(i)
      Next i
      tInc.s(iIncCount) = vbNullChar
   Else
      tInc.s(0) = vbNullChar
   End If
   If (iExCount > 0) Then
      For i = 1 To iExCount
         tExc.s(i - 1) = sExc(i)
      Next i
      tExc.s(iExCount) = vbNullChar
   Else
      tExc.s(0) = vbNullChar
   End If
   m_bCancel = False
   VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)
    
    'Debug.Print "--------------"
    'Debug.Print MYUSER.cchComment
    'Debug.Print MYUSER.TotalSizeComp
    'Debug.Print MYUSER.TotalSize
    'Debug.Print MYUSER.CompFactor
    'Debug.Print MYUSER.NumMembers
    'Debug.Print "--------------"

   Exit Function
   
ErrorHandler:
Dim lErr As Long, sErr As Long
   lErr = Err.Number: sErr = Err.Description
   VBUnzip = -1
   Set m_cUnzip = Nothing
   Err.Raise lErr, App.EXEName & ".VBUnzip", sErr
   Exit Function

End Function
