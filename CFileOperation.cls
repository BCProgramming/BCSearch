VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'IFileOperation GUID...
'Guid("04b0f1a7-9490-44bc-96e1-4296a31252e2")
'Private Declare Sub CLSIDFromString Lib "ole32.dll" (ByVal lpsz As String, ByVal pclsid As olelib.UUID)
'
'
Public Enum FileOperationTypeConstants
    FileOperation_Move
    FileOperation_Copy
End Enum


Public Type FileOperationData
    OperationType As FileOperationTypeConstants
    sFrom As String
    sTo As String
End Type


Private mFileOperations() As FileOperationData
Private mFileOpCount As Long


Public Sub AddOperation(ByVal OperationType As FileOperationTypeConstants, ByVal sFrom As String, ByVal sTo As String)
    mFileOpCount = mFileOpCount + 1
    ReDim Preserve mFileOperations(1 To mFileOpCount)
    With mFileOperations(mFileOpCount)
        .OperationType = OperationType
        .sFrom = Replace$(sFrom, "|", vbNullChar)
        .sTo = Replace$(sTo, "|", vbNullChar)
    
    
    End With




End Sub
Public Sub PerformOperations(ProgressObj As IProgressCallback)


End Sub
