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
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Private Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Private Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Private Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Private Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Private Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Private Const FILE_NOTIFY_CHANGE_ANY = FILE_NOTIFY_CHANGE_ATTRIBUTES + FILE_NOTIFY_CHANGE_CREATION + FILE_NOTIFY_CHANGE_DIR_NAME + FILE_NOTIFY_CHANGE_FILE_NAME + FILE_NOTIFY_CHANGE_SIZE
Private Const STATUS_WAIT_0 As Long = &H0
Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const WAIT_TIMEOUT As Long = 258&
Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hevent As Long
End Type


Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)

Private Const INFINITE As Long = &HFFFFFFFF
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


Public Enum FileChangeType
FILE_ACTION_ADDED = &H1
FILE_ACTION_REMOVED = &H2
'0 x00000002
FILE_ACTION_MODIFIED = &H3
'0 x00000003
FILE_ACTION_RENAMED_OLD_NAME = &H4
'0 x00000004
FILE_ACTION_RENAMED_NEW_NAME = &H5
'0 x00000005
End Enum

Public Type FILE_NOTIFY_INFORMATION
    NextEntryOffset As Long
    Action As Long
    FileNameLength As Long
    'Filename(1024 - 1) As Byte     'Wide character...
End Type

Public Type ChangedItemStruct
    ChangeType As FileChangeType
    ChangedItem As String

End Type
Public Type NotifyChangeStruct
    NumChanges As Long
    Changes() As ChangedItemStruct


End Type



Private Declare Function ReadDirectoryChangesW Lib "kernel32.dll" (ByVal hDirectory As Long, ByVal lpBuffer As Any, ByVal nBufferLength As Long, ByVal bWatchSubtree As Boolean, ByVal dwNotifyFilter As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long, ByVal lpCompletionRoutine As Long) As Long


'Private Declare Function ReadDirectoryChangesW Lib "kernel32.dll" (ByVal hDirectory As Long, ByVal lpBuffer As Long, ByVal nBufferLength As Long, ByVal bWatchSubtree As Long, dwNotifyFilter As Long, ByVal lpBytesReturned As Long, lpOverlapped As Any, CompletionRoutine As Long) As Long



Private Declare Function WaitForMultipleObjects Lib "kernel32.dll" (ByVal nCount As Long, ByRef lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function SetEvent Lib "kernel32.dll" (ByVal hevent As Long) As Long
Private Declare Function ResetEvent Lib "kernel32.dll" (ByVal hevent As Long) As Long
Private Declare Function GetOverlappedResult Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpOverlapped As OVERLAPPED, ByRef lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long

Private Declare Function FindFirstChangeNotification Lib "kernel32.dll" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindNextChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long
Private Declare Function CreateThread Lib "kernel32.dll" (ByRef lpThreadAttributes As Any, ByVal dwStackSize As Long, ByRef lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long

Private Declare Function FindCloseChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long

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
Public Sub Start(ByVal StrPath As String, ByVal WatchSubTree As Boolean, watcher As IChangeNotification)
    
    If mhchange <> 0 Then
        FindCloseChangeNotification mhchange
        mhchange = 0
    End If
    'set to async...
    Set ExecAsync = CreateObject("BCFile.CAsyncExec")
    dwFilter = FILE_NOTIFY_CHANGE_ANY
    mvarpath = StrPath
    mvarWatchTree = WatchSubTree
    Set mwatched = watcher
    ExecAsync.Start Me, 0


End Sub

Private Sub IAsyncProcess_ExecAsync(ByVal vdata As Variant)
    'other version.... for W2k+
    Dim BufferBytes(32768) As Byte
    Dim buffer(32768) As Byte
    Dim dirhandle As Long, olap As OVERLAPPED, mvarfname As String
    Dim retval As Long, retbytes As Long, hevent As Long
    Dim Structure As FILE_NOTIFY_INFORMATION, pos As Long
    
    Dim eventstruct As NotifyChangeStruct
    'open our directory...
    dirhandle = CreateFile(mvarpath, &H1, FILE_SHARE_READ + FILE_SHARE_WRITE + FILE_SHARE_DELETE, 0, OPEN_EXISTING, _
    FILE_FLAG_BACKUP_SEMANTICS + FILE_FLAG_OVERLAPPED, 0)
    'this will block....
    'hEvent = CreateEvent( _

    ' 0&, _ 'No Security Attribute
  
    ' True, _ 'Manual Reset
  
    ' True, _ 'Initial State is Set
  
    ' "DirEvent") 'Unique name of Event
    mhevent = CreateEvent(ByVal &H0&, 1, True, "BCFileDirChange" & Trim$(str$(SumAscii(mvarpath))))
    
    olap.hevent = mhevent
    
    retval = ReadDirectoryChangesW(dirhandle, buffer(0), 32768, mvarWatchTree, dwFilter, retbytes, VarPtr(olap), ByVal &H0)
    
    
   
    'mThreadSignalled = False
    'retval = CreateThread(ByVal &H0, 0, AddressOf ThreadProcNotify, ObjPtr(Me), 0, 0)
    'TESTING: use call to ensure procedure works before putting it on a separate thread.
    'Call ThreadProcNotify(ObjPtr(Me))
  
    mCancelled = False
    
    Dim currpos As Long
    Do
        retval = WaitForSingleObject(mhevent, 100)
        If retval <> WAIT_TIMEOUT Then
       ' Stop
        Debug.Print "here we raise event after copy..."
        currpos = 0
        
            CopyMemory BufferBytes(0), buffer(0), UBound(buffer) + 1
        
            'reset the event first...
             CloseHandle olap.hevent
        
            olap.hevent = CreateEvent(ByVal &H0&, 1, True, "BCFileDirChange" & Trim$(str$(SumAscii(mvarpath))))
            mhevent = olap.hevent
        
        
        
        Do
        
            'copy buff to working buffer...
            
 
            ZeroMemory buffer(0), UBound(buffer) + 1
            'while we handle this change, allow the API to continue monitoring...
            retval = ReadDirectoryChangesW(dirhandle, buffer(0), 32768, mvarWatchTree, dwFilter, retbytes, VarPtr(olap), ByVal &H0)
            
            
            'CopyMemory VarPtr(Structure), VarPtr(BufferBytes(currpos)), Len(Structure)

            ZeroMemory Structure, Len(Structure)
            CopyMemory Structure, BufferBytes(currpos), Len(Structure)
            
            
            If Structure.FileNameLength > 0 Then
                'copy to string...
                mvarfname = Space$(Structure.FileNameLength)
                
                CopyMemory mvarfname, VarPtr(BufferBytes(currpos + Len(Structure))), Structure.FileNameLength
                'mvarfname is now the unicode string...
                mvarfname = StrConv(mvarfname, vbFromUnicode)
                
                'add a new item to the structure...
                
                ReDim Preserve eventstruct.Changes(eventstruct.NumChanges)
                With eventstruct.Changes(eventstruct.NumChanges)
                    .ChangeType = Structure.Action
                    .ChangedItem = mvarfname
                End With
                'increment the count....
                eventstruct.NumChanges = eventstruct.NumChanges + 1
                Debug.Print mvarfname
                'structure.Action has what happened to this file.
                
                
                currpos = currpos + Structure.NextEntryOffset + Structure.FileNameLength
            End If
            
            
            
        Loop Until Structure.NextEntryOffset = 0
        'Raise the event here, with the current eventstruct.
        If eventstruct.NumChanges > 0 Then
            mwatched.Change Me, eventstruct
        
        End If
        
        'wasn't working proper- so I cheat and recreate the event. oh well.
       
        
        'ResetEvent olap.hevent
        
        
        
        
        
        End If
    
    DoEvents
    Loop While Not mCancelled
    
    'Stop
    'ReadDirectoryChangesW
    FindCloseChangeNotification mhchange
    'Close the Event handle, as well...
    CloseHandle olap.hevent
    mhevent = 0
    

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
