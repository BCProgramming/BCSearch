VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearchIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'"Provider=Microsoft.Jet.OLEDB.4.0;User ID=username;Data Source=Database.mdb;Persist Security Info=False"

Private mConnection As ADODB.Connection


Public Sub CreateIndex(ByVal IndexFile As String)
    Set mConnection = New ADODB.Connection
    Dim ADOXTable As ADOX.Catalog, makeTable As ADOX.Table
    Dim genstring As String
    Set ADOXTable = New ADOX.Catalog
    genstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & IndexFile & """"
    'mConnection.ConnectionString =
    'On Error Resume Next
    ADOXTable.Create genstring
    
    Set makeTable = New ADOX.Table
    makeTable.name = "SearchIndex"
    
    Dim newColumn As ADOX.Column
    
    
    
    'fields, Filename,Size,Attributes
    
    Set newColumn = New ADOX.Column
    newColumn.name = "Filename"
    newColumn.Type = adBSTR
    makeTable.Columns.Append newColumn, adBSTR
    
    Set newColumn = New ADOX.Column
    newColumn.name = "Size"
    newColumn.Type = adBigInt
    makeTable.Columns.Append newColumn, adBigInt
    
    Set newColumn = New ADOX.Column
    newColumn.name = "Attributes"
    newColumn.Type = adInteger
    makeTable.Columns.Append newColumn, adInteger
    ADOXTable.Tables.Append makeTable
    
    
    'ADOXTable.ActiveConnection.Close
    



End Sub
