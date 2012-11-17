Attribute VB_Name = "ModMain"
Option Explicit
'Private mMonitor As CFolderMonitor
Dim cINIFile As CINIData
Public fCancel As Boolean
Private Declare Function FindFirstChangeNotification Lib "kernel32.dll" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Private Declare Function FindNextChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long
Private Declare Function FindCloseChangeNotification Lib "kernel32.dll" (ByVal hChangeHandle As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const INFINITE As Long = &HFFFFFFFF

Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Private Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Private Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Private Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Private Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Private Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Private Const WAIT_TIMEOUT As Long = 258&

Private Function WaitForDoevents(ByVal handle As Long)

Dim retval As Long
Do

    retval = WaitForSingleObject(handle, 100)
    If retval <> WAIT_TIMEOUT Then
        Exit Do
    End If


Loop

'Stop




End Function
Sub Main()
'Dim X As BCFSObject
'Set X = New BCFSObject
'Set X = BCFile.GetDirectory("C:\windows")
'MsgBox X.GetDirectory("C:\").Path
Dim inifilename As String
Dim SourceFile As String, sourcefiles() As String
'Dim streamread As FileStream
Dim filecontents As String
Dim foldermon As String, copyto As String, logfile As String
Dim hFChange As Long, filelog As Long

'load the form.
Load frmMonitor
frmMonitor.Hide
inifilename = App.Path
Set cINIFile = New CINIData
If Right$(inifilename, 1) <> "\" Then inifilename = inifilename & "\"


inifilename = inifilename & "Monitor.ini"
cINIFile.LoadINI inifilename

foldermon = cINIFile.ReadProfileSetting("BCMONITOR", "FolderMonitor")
copyto = cINIFile.ReadProfileSetting("BCMONITOR", "CopyTo")
logfile = cINIFile.ReadProfileSetting("BCMONITOR", "LogFile")
If Right$(foldermon, 1) <> "\" Then foldermon = foldermon & "\"
If Right$(copyto, 1) <> "\" Then copyto = copyto & "\"
filelog = FreeFile
On Error Resume Next
MkDir copyto
Open logfile For Append As filelog
Print #filelog, "Beginning log, " & Now
hFChange = FindFirstChangeNotification(foldermon, False, FILE_NOTIFY_CHANGE_CREATION + FILE_NOTIFY_CHANGE_LAST_WRITE + FILE_NOTIFY_CHANGE_SIZE + FILE_NOTIFY_CHANGE_ATTRIBUTES + FILE_NOTIFY_CHANGE_FILE_NAME + FILE_NOTIFY_CHANGE_LAST_ACCESS + FILE_NOTIFY_CHANGE_SECURITY)
WaitForDoevents hFChange
Dim docopy As Boolean
Do
    'a change has occured. as such, examine all the files in the folder we are monitoring, and copy those that have changed to the destination folder.
    
    Dim I As Long
    SourceFile = Dir$(foldermon)
    I = 1
    Do
        ReDim Preserve sourcefiles(1 To I)
        sourcefiles(I) = SourceFile
        SourceFile = Dir$
        I = I + 1
    Loop Until SourceFile = ""
    
    
    docopy = False
    
    'Do Until SourceFile = ""
    For I = 1 To UBound(sourcefiles)
        SourceFile = sourcefiles(I)
        If Dir$(copyto & SourceFile) = "" Then
        'the file is not in the target folder.
            docopy = True
        
        Else
            If FileDateTime(copyto & SourceFile) < FileDateTime(foldermon & SourceFile) Then
                docopy = True
            End If
        End If
        If docopy Then
        FileCopy foldermon & SourceFile, copyto & SourceFile
        Print #filelog, " copied file, """ & foldermon & SourceFile & """ to destination, """ & copyto & SourceFile & """, on " & FormatDateTime(Now, vbShortDate) & " At " & FormatDateTime(Now, vbShortTime) & "."
        End If
    
    'Loop
    Next
    



    FindNextChangeNotification hFChange
    WaitForDoevents hFChange
Loop Until fCancel
FindCloseChangeNotification hFChange
Print #filelog, "log End, Proper termination. " & Now
Close filelog

End Sub
