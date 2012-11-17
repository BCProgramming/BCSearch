VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mSettings As CINIData
Private mSearchForm As FrmSearch
Private MSettingsObject As CINIData
Private mColumnPlugins() As IColumnPlugin
Private mColPlugCount As Long
Public Enum SizeIMLEnum
    Size_Small = 0
    Size_Large
    End Enum
    'Public Const STATISTIC_FILESFOUND="FilesFound"
Private mRegistry As cRegistry
Private mSmallImages As cVBALSysImageList
Private mLargeImages As cVBALSysImageList
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathW" (ByVal hwnd As Long, ByVal pszPath As Long, ByVal csidl As Long, ByVal fCreate As Long) As Long

Private mStatistics As PropertyBag




Private Sub LoadStatistics()
    'first, see if the INI file gives a location for the .stats file...
    Dim Statsfile As String
    Debug.Print "LoadStatistics"
    Dim readbytes() As Byte, bytelength As Long, readstream As FileStream
    On Error GoTo LoadStatError
     Set mStatistics = New PropertyBag
    Statsfile = Settings.ReadProfileSetting("BCSearch", "StatsFile", "")
    If Statsfile = "" Then
        Statsfile = GetDataFolder & "bcsearch.stats"
    End If
    Set readstream = OpenStream(Statsfile)
    readbytes = readstream.readbytes(readstream.size)
    readstream.CloseStream
   
    mStatistics.Contents = readbytes
    Exit Sub
LoadStatError:
    

End Sub
Public Sub SaveStatistics()
        Dim Statsfile As String
        Debug.Print "SaveStatistics"
    Dim writebytes() As Byte, bytelength As Long, Writestream As FileStream
    On Error GoTo SaveStatError
     
    Statsfile = Settings.ReadProfileSetting("BCSearch", "StatsFile", "")
    If Statsfile = "" Then
        Statsfile = GetDataFolder & "bcsearch.stats"
    End If
    Set Writestream = CreateFile(Statsfile, True)
    writebytes = mStatistics.Contents
    Writestream.writebytes writebytes()
    
    Writestream.CloseStream
   
    
    Exit Sub
SaveStatError:
End Sub
Public Function GetStatistic(ByVal StatisticName As String) As Variant
    If Not mStatistics Is Nothing Then
    
    
        GetStatistic = mStatistics.ReadProperty(StatisticName, "0")
    
    
    

    End If
End Function
Public Sub SetStatistic(ByVal StatisticName As String, ByVal newvalue As Variant)

    If Not mStatistics Is Nothing Then
        mStatistics.WriteProperty StatisticName, newvalue, "0"

    End If

End Sub
Public Sub IncrementStatistic(ByVal StatisticName As String)
    SetStatistic StatisticName, GetStatistic(StatisticName) + 1
End Sub

Public Function GetDataFolder() As String
    Dim sFolder As String
    sFolder = Space$(32768)
    Call SHGetSpecialFolderPath(0, StrPtr(sFolder), CSIDL_APPDATA, 1)
    sFolder = Trim$(Replace$(sFolder, vbNullChar, ""))
    GetDataFolder = sFolder
    
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    sFolder = sFolder & App.ProductName & "\"
    GetDataFolder = sFolder
    
End Function


Friend Property Get registry() As cRegistry
    If mRegistry Is Nothing Then Set mRegistry = New cRegistry
    Set registry = mRegistry
End Property
Public Property Get SystemIML(ByVal SizeGet As SizeIMLEnum) As Object
        If SizeGet = Size_Small Then
            Set SystemIML = mSmallImages
        Else
            Set SystemIML = mLargeImages
        End If
        
End Property
Friend Property Get ColumnPlugins() As IColumnPlugin()
    ColumnPlugins = mColumnPlugins
End Property
Friend Property Get ColumnCount() As Long
    ColumnCount = mColPlugCount
End Property


Public Property Get Settings() As CINIData
    If MSettingsObject Is Nothing Then
        Set MSettingsObject = New CINIData
        MSettingsObject.LoadINI GetDataFolder & "BSearch.ini"
        
    
    End If

    Set Settings = MSettingsObject

End Property
Private Function AscSum(StrSum As String)
    Dim I As Long, sumrunner As Long
    For I = 1 To Len(StrSum)
        sumrunner = sumrunner + (I * Asc(Mid$(StrSum, I, 1)))
    Next I
    AscSum = sumrunner
End Function
Private Sub Class_Initialize()
    Dim vcount As Long, Values() As String
    Dim CurrProgID As Long
    Dim gotval As String, tmpdisabled As String
    Dim currcreated As Long, currobj As Object
    Dim testcast As IColumnPlugin, CurrHash As Long
    Set CurrApp = Me
    Set mSettings = New CINIData
    Set mSearchForm = New FrmSearch
    'ReDim mColumnPlugins(1 To 1)
    
    If PerformUpdateCheck(0) Then
        Exit Sub
    End If
    
    On Error Resume Next
    Call Me.Settings.EnumerateValues(Setting_System, "Column Plugins", Values(), vcount)
    If Err = 0 Then
        'enum each, they are progIDs.
        For CurrProgID = 1 To vcount
        
        
            If CBool(Me.Settings.ReadProfileSetting("Column Plugins", Values(CurrProgID))) Then
                Set currobj = CreateObject(Values(CurrProgID))
                On Error Resume Next
                Set testcast = currobj
                If Err = 0 Then
                    'success...
                    currcreated = currcreated + 1
                    ReDim Preserve mColumnPlugins(1 To currcreated)
                    Set mColumnPlugins(currcreated) = currobj
                    mColPlugCount = currcreated
                    CurrHash = CurrHash + AscSum(Values(CurrProgID))
                Else
                'error.
                
                End If
            End If
        
        Next CurrProgID
    End If
    'BUGFIX: whenever a column plugin was added, the Column position data would become desynchronized resulting in subitems no longer
    'matching the column position.
    'the fix? detect when a column plugin is added or removed, and if so delete the positional data in [lvwfiles_columns]
    Dim oldhash As Long
    oldhash = Me.Settings.ReadProfileSetting("BCSearch", "ColumnPluginsHash", "0")
    If oldhash <> CurrHash Then
        'delete the section...
        Me.Settings.DeleteSection "lvwfiles_Columns"
    
    End If
    
    
    
    'Initialize System Imagelists...
    If mSmallImages Is Nothing Then
    Set mSmallImages = New cVBALSysImageList
    mSmallImages.IconSizeX = 16
    mSmallImages.IconSizeY = 16
    mSmallImages.Create
    End If
    'mSmallImages.IconSizeX = 16
    'mSmallImages.IconSizeY = 16
    'mSmallImages.ColourDepth = ILC_COLOR32
    'mSmallImages.Create
    
    If mLargeImages Is Nothing Then
        Set mLargeImages = New cVBALSysImageList
        mLargeImages.IconSizeX = 32
        mLargeImages.IconSizeY = 32
        mLargeImages.Create
    End If
    LoadStatistics
    
    'Set mColumnPlugins(1) = New CMP3Columns
    ' mColPlugCount = 1
End Sub
Public Property Get MainForm() As Object
    Set MainForm = mSearchForm
End Property
Friend Property Get SearchForm() As FrmSearch
Set SearchForm = mSearchForm
End Property
Public Sub Show()
    mSearchForm.Show
End Sub
Public Sub SendCmdLine(ByVal CmdLine As String)
Attribute SendCmdLine.VB_MemberFlags = "40"
    'hidden procedure used by instances after the first to send command line arguments to the first instance.
    'mSearchForm.ParseCmdLine
    CDebug.Post "Instance sent command line:" & CmdLine
End Sub
Private Sub Class_Terminate()
    MSettingsObject.SaveINI
    SaveStatistics
End Sub

