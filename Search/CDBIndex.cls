VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDBIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDBConString As String
Private mCon As ADODB.Connection
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%APPDATA%\BCSearch\Index.mdb;Persist Security Info=False"


Private Sub CreateDB()
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table
Dim con As ADODB.Connection

  
  

    ' Create the new database.
    Set cat = New ADOX.Catalog
    cat.Create mDBConString

    ' Create a new table.
    'Set tbl = New ADOX.Table
    'tbl.name = "TestTable"
    'tbl.Columns.Append "FirstName", adVarWChar, 40
    'tbl.Columns.Append "LastName", adVarWChar, 40
    'cat.Tables.Append tbl

    ' Connect to the database.
    Set mCon = cat.ActiveConnection

    ' Insert a record.
    'con.Execute "INSERT INTO TestTable VALUES ('Rod', " & _
        "'Stephens')"
    mCon.Close
    

    

End Sub
Public Sub AddEntry(ForFile As CFile)





End Sub
Private Sub Class_Initialize()
    Dim mErrorCount As Long
On Error GoTo OpenDBError
    mDBConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & bcfile.GetSpecialFolder(CSIDL_APPDATA).Path & "BCSearch\Index.mdb"
    Set mCon = New Connection
    mCon.ConnectionString = mDBConString
    Exit Sub
OpenDBError:
    mErrorCount = mErrorCount + 1
    If mErrorCount >= 10 Then Exit Sub
    CreateDB
    Resume
End Sub
