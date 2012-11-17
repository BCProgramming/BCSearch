VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileChangeNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CFileChangeNotify: wraps change notification...


'typedef struct _FILE_NOTIFY_INFORMATION {
'  DWORD NextEntryOffset;
'  DWORD Action;
'  DWORD FileNameLength;
'  WCHAR FileName[1];
'} FILE_NOTIFY_INFORMATION, *PFILE_NOTIFY_INFORMATION;
'BOOL WINAPI ReadDirectoryChangesW(
'  __in         HANDLE hDirectory,
'  __out        LPVOID lpBuffer,
'  __in         DWORD nBufferLength,
'  __in         BOOL bWatchSubtree,
'  __in         DWORD dwNotifyFilter,
'  __out_opt    LPDWORD lpBytesReturned,
'  __inout_opt  LPOVERLAPPED lpOverlapped,
'  __in_opt     LPOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine
');


Private Const TIME_OUT = &H102
Private Const FILE_SHARE_DELETE = &H4
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ALL_ACCESS = &H1FF

Private Const FILE_LIST_DIRECTORY = &H1
Private Const OPEN_EXISTING = &H3
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Enum FILE_NOTIFY_CHANGE
FILE_NOTIFY_CHANGE_FILE_NAME = &H1
FILE_NOTIFY_CHANGE_DIR_NAME = &H2
FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
FILE_NOTIFY_CHANGE_SIZE = &H8
FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20
FILE_NOTIFY_CHANGE_CREATION = &H40
FILE_NOTIFY_CHANGE_SECURITY = &H100
End Enum

Public Enum FILE_ACTION
FILE_ACTION_ADDED = &H1
FILE_ACTION_REMOVED = &H2
FILE_ACTION_MODIFIED = &H3
FILE_ACTION_RENAMED_OLD_NAME = &H4
FILE_ACTION_RENAMED_NEW_NAME = &H5
End Enum



Private Type OVERLAPPED
    internal As Long
    internalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type FILE_NOTIFY_INFORMATION
dwNextEntryOffset As Long
dwAction As FILE_ACTION
dwFileNameLength As Long
wcFileName(1024 - 1) As Byte 'buffer of 1024 bytes
End Type


Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long

Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long


Private Declare Function ReadDirectoryChangesW Lib "kernel32.dll" (ByVal hDirectory As Long, ByVal lpBuffer As Long, ByVal nBufferLength As Long, ByVal bWatchSubtree As Boolean, ByVal dwNotifyFilter As FILE_NOTIFY_CHANGE, lpBytesReturned As Long, ByVal lpOverlapped As Long, ByVal lpCompletionRoutine As Long) As Long



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)


Public Type ChangedItemStruct
    Changetype As FILE_ACTION
    ChangedItem As String

End Type
Public Type NotifyChangeStruct
    NumChanges As Long
    Changes() As ChangedItemStruct


End Type


'Private Declare Function ReadDirectoryChangesW Lib "kernel32.dll" (ByVal hDirectory As Long, ByVal lpBuffer As Long, ByVal nBufferLength As Long, ByVal bWatchSubtree As Long, dwNotifyFilter As Long, ByVal lpBytesReturned As Long, lpOverlapped As Any, CompletionRoutine As Long) As Long



Private Declare Function WaitForMultipleObjects Lib "kernel32.dll" (ByVal nCount As Long, ByRef lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
'Private Declare Function ResetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
'Private Declare Function GetOverlappedResult Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpOverlapped As OVERLAPPED, ByRef lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Declare Function FindFirstChangeNotification Lib "kernel32.dll" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindNextChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long
'Private Declare Function CreateThread Lib "kernel32.dll" (ByRef lpThreadAttributes As Any, ByVal dwStackSize As Long, ByRef lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long

'Private Declare Function FindCloseChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long

Private Declare Function WaitForSingleObjectEx Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private mCancelled As Boolean
Private mhevent As Long
Private mhchange As Long, ExecAsync As CAsyncExec
Private mvarpath As String, mvarWatchTree As Boolean, dwFilter As Long
Private mwatched As IChangeNotification
Private mThreadSignalled As Boolean

Implements IAsyncProcess
'Private masync As CAsyncExec
Friend Property Get ThreadhEvent() As Long
    ThreadhEvent = mhevent
End Property
Friend Sub ThreadSignal()
    mThreadSignalled = True
End Sub
Public Sub Cancel()
If mhevent > 0 Then
    mCancelled = True
    SetEvent mhevent
End If
End Sub
Public Sub Start(ByVal strPath As String, ByVal WatchSubTree As Boolean, watcher As IChangeNotification)
    
'    If mhchange <> 0 Then
'        FindCloseChangeNotification mhchange
'        mhchange = 0
'    End If
    'set to async...
    Set ExecAsync = CreateObject("BCFile.CAsyncExec")
    dwFilter = FILE_NOTIFY_CHANGE_CREATION + FILE_NOTIFY_CHANGE_ATTRIBUTES
    mvarpath = strPath
    mvarWatchTree = WatchSubTree
    Set mwatched = watcher
    ExecAsync.Start Me, 0


End Sub

Private Sub IAsyncProcess_ExecAsync(ByVal vdata As Variant)

    Dim hDir As Long 'Directory Handler
    Dim hEvent As Long 'Event Handler
    
    
    mCancelled = False
    
    'Create the handle to mointor "c:\test"
    hDir = CreateFile(mvarpath, _
    FILE_LIST_DIRECTORY, _
    FILE_SHARE_READ Or _
    FILE_SHARE_DELETE Or FILE_SHARE_WRITE, _
    0&, _
    OPEN_EXISTING, _
    FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, _
    0&)
    
    'Documenting CreateEvent
    'Create an Event for Async mode in ReadDirectoryChangesW
    'hEvent = CreateEvent( _
    ' 0&, _ 'No Security Attribute
    ' True, _ 'Manual Reset
    ' True, _ 'Initial State is Set
    ' "DirEvent") 'Unique name of Event
    
    hEvent = CreateEvent( _
    0&, _
    True, _
    True, _
    "DirEvent")
    
    'Create an OverLapped structure for Async mode
    Dim oLap As OVERLAPPED
    'Assign the event created to the Overlapped structure
    With oLap
        .hEvent = hEvent
    End With
    
    
    'Buffer for ReadDirectoryChangesW
    Dim buf(0 To 1024 * 5 - 1) As Byte
    
    'structure to retreive file information
    Dim dirBuf As FILE_NOTIFY_INFORMATION
    
    Dim nUsed As Long 'used only for sync mode
    
    'Documenting ReadDirectoryChangesW call
    'ReadDirectoryChangesW in Asyn mode
    'Call ReadDirectoryChangesW(hDir, _ ' Directory Handler
    ' dirBuf, _ ' Buffer
    ' LenB(dirBuf), _ ' Byte Length of Buffer
    ' True, _ ' Watch sub tree
    ' FILE_NOTIFY_CHANGE_FILE_NAME, _ ' Notification Filter
    ' nUsed, _
    ' VarPtr(oLap), _ ' Pointer to OverLapped
    ' 0&)
    
    Call ReadDirectoryChangesW(hDir, _
    VarPtr(buf(0)), _
    UBound(buf) + 1, _
    True, _
    FILE_NOTIFY_CHANGE_FILE_NAME, _
    nUsed, _
    VarPtr(oLap), _
    0&)
    
    'Just for completeness, if you want to use Sync mode
    'Sync mode is called this way
    'Call ReadDirectoryChangesW(hDir, _
    ' VarPtr(buf(0)), _
    ' UBound(buf) + 1, _
    ' False, _
    ' FILE_NOTIFY_CHANGE_FILE_NAME, _
    ' nUsed, _
    ' 0&, _
    ' 0&)
    
    
    Dim bstr As String 'string for display of filename
    Dim Pos As Long 'location of file entry in buffer buf
    Dim ret As Long 'return value of wait operation
    Dim structmake As NotifyChangeStruct
    mCancelled = False
    Do
        Pos = 0
        'wait for hEvent to be unset
        ret = WaitForSingleObject(hEvent, 100)
        If ret <> TIME_OUT Then
            'get the first entry in buf
            CopyMemory VarPtr(dirBuf), VarPtr(buf(Pos)), Len(dirBuf)
            'uncomment if you want to see details of dirBuf
            Debug.Print dirBuf.dwFileNameLength & " " & _
             dirBuf.dwNextEntryOffset
            Select Case dirBuf.dwAction
                Case FILE_ACTION_ADDED
                    Debug.Print "ADDED "
               
                Case FILE_ACTION_MODIFIED
                    Debug.Print "MODIFIED ";
                    
                Case FILE_ACTION_REMOVED
                    Debug.Print "REMOVED ";
                    
                Case FILE_ACTION_RENAMED_NEW_NAME
                    Debug.Print "RENAME NEW ";
                    
                Case FILE_ACTION.FILE_ACTION_RENAMED_OLD_NAME:
                    Debug.Print "RENAME OLD ";
                    
                End Select
            
            'locate the filename returned
            'assigned a unicode byte array to a string
            bstr = dirBuf.wcFileName
            bstr = Left(bstr, dirBuf.dwFileNameLength / 2)
            structmake.NumChanges = 1
            ReDim structmake.Changes(1)
            structmake.Changes(0).ChangedItem = bstr
            structmake.Changes(0).Changetype = dirBuf.dwAction
'
'
            Debug.Print bstr
            Debug.Print structmake.NumChanges
            
            Do While dirBuf.dwNextEntryOffset <> 0 ' 0 this indicate last entry
                'get next entry in buf
                Pos = Pos + dirBuf.dwNextEntryOffset
                CopyMemory VarPtr(dirBuf), VarPtr(buf(Pos)), Len(dirBuf)
                'uncomment if you want to see details of dirBuf
                'Debug.Print dirBuf.dwFileNameLength & " " & _
                ' dirBuf.dwNextEntryOffset
                
                Select Case dirBuf.dwAction
                    Case FILE_ACTION_ADDED: Debug.Print "2ADDED ";
                    'Not mointoring File modification
                    'Case FILE_ACTION_MODIFIED: Debug.Print "MODIFIED ";
                    Case FILE_ACTION_REMOVED: Debug.Print "2REMOVED ";
                    Case FILE_ACTION_RENAMED_NEW_NAME: Debug.Print "2RENAME NEW ";
                    Case FILE_ACTION.FILE_ACTION_RENAMED_OLD_NAME:
                    Debug.Print "2RENAME OLD ";
                End Select
                bstr = dirBuf.wcFileName
                bstr = Left(bstr, dirBuf.dwFileNameLength / 2)
                Debug.Print bstr
                'CHANGE: create and populate proper structure that will be passed in event...
                ReDim Preserve structmake.Changes(structmake.NumChanges)
                With structmake.Changes(structmake.NumChanges)
                    .ChangedItem = bstr
                    .Changetype = dirBuf.dwAction
                End With
                structmake.NumChanges = structmake.NumChanges + 1
                
            Loop
        mwatched.Change Me, structmake
            
        
            
            
            'reset the event mointering and repeat the whole process
            ResetEvent hEvent
            Call ReadDirectoryChangesW(hDir, _
            VarPtr(buf(0)), _
            UBound(buf) + 1, _
            True, _
            FILE_NOTIFY_CHANGE_FILE_NAME, _
            nUsed, VarPtr(oLap), 0&)
            
        
        End If
        DoEvents
        
    Loop While Not mCancelled ' if fstop is true exit loop
    
    'close all the handles
    CloseHandle hEvent
    CloseHandle hDir

  

End Sub

'Private Sub IAsyncProcess_ExecAsync(ByVal vdata As Variant)
''
'
''BOOL bContinue = TRUE;
''While (bContinue)
''{
''    if (WaitForMultipleObjects(2, aHandles, FALSE, INFINITE) - WAIT_OBJECT_0 == 0)
''    {
''        // perform some sort of notification here
''
''        FindNextChangeNotification(hChange);
''    }
''    else // exit loop, as m_event has became signalled
''        bContinue = FALSE;
''}
'Dim bcontinue As Boolean
'Dim ahandles(1 To 2) As Long
'Dim ret As Long
'mhchange = FindFirstChangeNotification(mvarpath, Abs(mvarWatchTree), dwFilter)
'bcontinue = True
'mhEvent = CreateEvent(ByVal &H0&, False, 0, "ChangeNotify")
'
'Do While bcontinue
'    ahandles(1) = mhchange
'    ahandles(2) = mhEvent
'    'If (WaitForMultipleObjects(1, ahandles(1), False, INFINITE) - WAIT_OBJECT_0 = 0) Then
'    ret = WaitForSingleObject(mhchange, 10)
'    DoEvents
'        If mCancelled Then
'            'flagged-
'            bcontinue = False
'        End If
'
'
'
'    If ret = WAIT_TIMEOUT Then
'        'timeout value was hit.
'
'
'
'    ElseIf ret = WAIT_OBJECT_0 Then
'        'hit signalled state.
'
'        'perform notification...
'        mwatched.Change Me
'        FindNextChangeNotification mhchange
'
'
'    End If
'
'Loop

'End Sub
