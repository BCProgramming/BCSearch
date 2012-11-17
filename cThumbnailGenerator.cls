VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cThumbnailGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SHGetMalloc Lib "shell32.dll" (ppMalloc As IVBMalloc) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Const S_OK = 0           ' indicates success
Private Const S_FALSE = 1&   ' special HRESULT value

Private Const MAX_PATH = 260
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function SHGetDesktopFolder Lib "shell32.dll" _
   (ppshf As IVBShellFolder) As Long

' GetItemID item ID retrieval constants
Private Const GIID_FIRST = 1
Private Const GIID_LAST = -1

Private alloc As IMalloc

Private m_lDesiredWidth As Long
Private m_lDesiredHeight As Long
Private m_sFileName As String
Private m_eOptions As EIEIFLAG

Public Property Get Options() As EIEIFLAG
   Options = m_eOptions
End Property
Public Property Let Options(ByVal eOptions As EIEIFLAG)
   m_eOptions = eOptions
End Property

Public Property Get DesiredWidth() As Long
   DesiredWidth = m_lDesiredWidth
End Property
Public Property Let DesiredWidth(ByVal lWidth As Long)
   m_lDesiredWidth = lWidth
End Property
Public Property Get DesiredHeight() As Long
   DesiredHeight = m_lDesiredHeight
End Property
Public Property Let DesiredHeight(ByVal lHeight As Long)
   m_lDesiredHeight = lHeight
End Property
Public Property Get Filename() As String
   Filename = m_sFileName
End Property
Public Property Let Filename(ByVal sFileName As String)
   m_sFileName = sFileName
End Property

Private Property Get Allocator() As IVBMalloc
    If alloc Is Nothing Then SHGetMalloc alloc
    Set Allocator = alloc
End Property

Private Function GetDirectoryName(ByVal sFileName As String) As String
Dim i As Long
Dim sDir As String
   For i = Len(sFileName) To 1 Step -1
      If (Mid(sFileName, i, 1) = "\") Then
         sDir = Left(sFileName, i - 1)
         If (Right(sDir, 1) = ":") Then
            sDir = sDir & "\"
         End If
         GetDirectoryName = sDir
         Exit Function
      End If
   Next i
End Function
Private Function GetFileName(ByVal sFileName As String) As String
Dim i As Long
   For i = Len(sFileName) To 1 Step -1
      If (Mid(sFileName, i, 1) = "\") Then
         GetFileName = Mid(sFileName, i + 1)
         Exit Function
      End If
   Next i
End Function
Public Function GetThumbNailhBmp() As Long
Dim folder As IShellFolderEx_TLB.IShellFolder
Dim sRet As String
Dim lR As Long
Dim sPath As String
Dim sFileName As String
Dim lFilePos As Long
Dim cParsed As Long
Dim afItem As Long
Dim pidlMain As Long
Dim item As IShellFolderEx_TLB.IShellFolder
Dim iidShellFolder As IShellFolderEx_TLB.GUID
Dim idenum As IShellFolderEx_TLB.IEnumIDList
Dim pidl As Long
Dim cFetched As Long
Dim afAttrib As Long
   
   sPath = GetDirectoryName(m_sFileName) '"C:\SteveMac" '"C:\Documents and Settings\Steve McMahon\My Documents"
   sRet = String$(MAX_PATH, 0)
   lR = GetFullPathName(sPath, MAX_PATH, sRet, lFilePos)
   If lR = 0 Then
      Err.Raise 45001, App.EXEName & ".cThumbnailImage", WinApiError(Err.LastDllError)
   Else
      
      Set folder = GetDesktopFolder
      afItem = 0
      On Error Resume Next
      folder.ParseDisplayName 0&, 0&, sPath, cParsed, pidlMain, 0&
      If Not (Err.Number = 0) Then
         On Error GoTo 0
         Err.Raise 45002, App.EXEName & ".cThumbnailGenerator", "Unable to locate the folder '" & sPath & "'"
         Exit Function
      End If
      
      ' IShellFolder:
      IIDFromString "{000214E6-0000-0000-C000-000000000046}", iidShellFolder
      folder.BindToObject pidlMain, 0&, iidShellFolder, item
      If Not (Err.Number = 0) Then
         On Error GoTo 0
         Err.Raise 45003, App.EXEName & ".cThumbnailGenerator", "Unable to bind to the folder '" & sPath & "'"
         Exit Function
      End If
      
      item.EnumObjects 0&, SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS, idenum
      If Not (Err.Number = 0) Then
         Allocator.Free pidlMain
         On Error GoTo 0
         Err.Raise 45004, App.EXEName & ".cThumbnailGenerator", "Unable to read the contents of the folder '" & sPath & "'"
         Exit Function
      End If
      
      On Error GoTo 0
      Dim hRes As Long
      Do
         pidl = 0
         hRes = idenum.Next(1, pidl, cFetched)
         
         If hRes Then Exit Do ' no more items left
                  
         sPath = PathFromPidl(pidl)
         If (GetFileName(sPath) = GetFileName(m_sFileName)) Then
            
            ' Let's get an IVBContextMenu object from it:
            Dim iidExtractImage As IShellFolderEx_TLB.GUID
            Dim extractImage As IExtractImage
            
            On Error Resume Next
            IIDFromString "{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", iidExtractImage
            item.GetUIObjectOf 0&, 1&, pidl, iidExtractImage, 0&, extractImage
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45005, App.EXEName & ".cThumbnailGenerator", "The object " & m_sFileName & " does not support thumbnails"
               Exit Function
            End If
            
            sRet = String$(MAX_PATH, 0)
            Dim tSize As SIZE
            Dim pdwFlags As EIEIFLAG
            tSize.cX = m_lDesiredHeight
            tSize.cY = m_lDesiredWidth
            pdwFlags = m_eOptions
            extractImage.GetLocation sRet, 260, 0&, tSize, 32, pdwFlags
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45006, App.EXEName & ".cThumbnailGenerator", "ExtractImage on " & m_sFileName & " failed."
               Exit Function
            End If
            
            Dim hBmp As Long
            hBmp = extractImage.Extract
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45007, App.EXEName & ".cThumbnailGenerator", "ExtractImage on " & m_sFileName & " failed."
               Exit Function
            End If

            Allocator.Free pidl
            pidl = 0
                                    
            If Not (hBmp = 0) Then
               On Error GoTo 0
               'Dim c As New pcMemDC
               'c.CreateFromHBitmap hBmp
               
                GetThumbNailhBmp = hBmp
            Else
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45008, App.EXEName & ".cThumbnailGenerator", "No Thumbnail was provided."
               Exit Function
            End If
            Exit Do
            
         End If
      
         ' Free the pidl from Next
         Allocator.Free pidl
         pidl = 0
      Loop
      
      ' Free the pidl:
      Allocator.Free pidlMain
      pidlMain = 0
         
   End If
   



End Function
Public Function GetThumbnail() As pcMemDC
Dim folder As IVBShellFolder
Dim sRet As String
Dim lR As Long
Dim sPath As String
Dim sFileName As String
Dim lFilePos As Long
Dim cParsed As Long
Dim afItem As Long
Dim pidlMain As Long
Dim item As IVBShellFolder
Dim iidShellFolder As UUID
Dim idenum As IVBEnumIDList
Dim pidl As Long
Dim cFetched As Long
Dim afAttrib As Long
   
   sPath = GetDirectoryName(m_sFileName) '"C:\SteveMac" '"C:\Documents and Settings\Steve McMahon\My Documents"
   sRet = String$(MAX_PATH, 0)
   lR = GetFullPathName(sPath, MAX_PATH, sRet, lFilePos)
   If lR = 0 Then
      Err.Raise 45001, App.EXEName & ".cThumbnailImage", WinApiError(Err.LastDllError)
   Else
      
      Set folder = GetDesktopFolder
      afItem = 0
      On Error Resume Next
      folder.ParseDisplayName 0&, 0&, sPath, cParsed, pidlMain, 0&
      If Not (Err.Number = 0) Then
         On Error GoTo 0
         Err.Raise 45002, App.EXEName & ".cThumbnailGenerator", "Unable to locate the folder '" & sPath & "'"
         Exit Function
      End If
      
      ' IShellFolder:
      IIDFromString "{000214E6-0000-0000-C000-000000000046}", iidShellFolder
      folder.BindToObject pidlMain, 0&, iidShellFolder, item
      If Not (Err.Number = 0) Then
         On Error GoTo 0
         Err.Raise 45003, App.EXEName & ".cThumbnailGenerator", "Unable to bind to the folder '" & sPath & "'"
         Exit Function
      End If
      
      item.EnumObjects 0&, SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS, idenum
      If Not (Err.Number = 0) Then
         Allocator.Free pidlMain
         On Error GoTo 0
         Err.Raise 45004, App.EXEName & ".cThumbnailGenerator", "Unable to read the contents of the folder '" & sPath & "'"
         Exit Function
      End If
      
      On Error GoTo 0
      Dim hRes As Long
      Do
         pidl = 0
         hRes = idenum.Next(1, pidl, cFetched)
         
         If hRes Then Exit Do ' no more items left
                  
         sPath = PathFromPidl(pidl)
         If (GetFileName(sPath) = GetFileName(m_sFileName)) Then
            
            ' Let's get an IVBContextMenu object from it:
            Dim iidExtractImage As UUID
            Dim extractImage As IExtractImage
            
            On Error Resume Next
            IIDFromString "{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", iidExtractImage
            item.GetUIObjectOf 0&, 1&, pidl, iidExtractImage, 0&, extractImage
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45005, App.EXEName & ".cThumbnailGenerator", "The object " & m_sFileName & " does not support thumbnails"
               Exit Function
            End If
            
            sRet = String$(MAX_PATH, 0)
            Dim tSize As SIZE
            Dim pdwFlags As EIEIFLAG
            tSize.cX = m_lDesiredHeight
            tSize.cY = m_lDesiredWidth
            pdwFlags = m_eOptions
            extractImage.GetLocation sRet, 260, 0&, tSize, 32, pdwFlags
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45006, App.EXEName & ".cThumbnailGenerator", "ExtractImage on " & m_sFileName & " failed."
               Exit Function
            End If
            
            Dim hBmp As Long
            extractImage.Extract hBmp
            If Not (Err.Number = 0) Then
               Allocator.Free pidl
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45007, App.EXEName & ".cThumbnailGenerator", "ExtractImage on " & m_sFileName & " failed."
               Exit Function
            End If

            Allocator.Free pidl
            pidl = 0
                                    
            If Not (hBmp = 0) Then
               On Error GoTo 0
               Dim c As New pcMemDC
               c.CreateFromHBitmap hBmp
               
               Set GetThumbnail = c
            Else
               Allocator.Free pidlMain
               On Error GoTo 0
               Err.Raise 45008, App.EXEName & ".cThumbnailGenerator", "No Thumbnail was provided."
               Exit Function
            End If
            Exit Do
            
         End If
      
         ' Free the pidl from Next
         Allocator.Free pidl
         pidl = 0
      Loop
      
      ' Free the pidl:
      Allocator.Free pidlMain
      pidlMain = 0
         
   End If
   

End Function
Public Function PathFromPidl(ByVal pidl As Long) As String
Dim sPath As String
Dim lR As Long
   sPath = String$(MAX_PATH, 0)
   lR = SHGetPathFromIDList(pidl, sPath)
   If lR <> 0 Then
      PathFromPidl = Left$(sPath, lstrlen(sPath))
   End If
End Function
Private Function GetDesktopFolder() As IVBShellFolder
Dim lR As Long
    lR = SHGetDesktopFolder(GetDesktopFolder)
End Function
Private Function WinApiError(ByVal e As Long) As String
Dim s As String, c As Long
   s = String(256, 0)
   c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                     FORMAT_MESSAGE_IGNORE_INSERTS, _
                     0&, e, 0&, s, Len(s), ByVal 0&)
   If c Then WinApiError = Left$(s, c)
End Function


Private Sub Class_Initialize()
   m_lDesiredWidth = 100
   m_lDesiredHeight = 100
   m_eOptions = IEIFLAG_NOBORDER Or IEIFLAG_SCREEN Or IEIFLAG_OFFLINE 'Or IEIFLAG_ORIGSIZE
End Sub
