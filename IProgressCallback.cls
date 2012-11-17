VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IProgressCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IProgressCallback Interface-
'used in routines that perform File operations with the shell
'Public Enum ProgressOperationConstants
'    Progress_Copy
'    Progress_Move
'    Progress_Delete
'
'End Enum
Public Enum ProgressReturnConstants
    PROGRESS_CANCEL = 1
    PROGRESS_CONTINUE = 0
    PROGRESS_QUIET = 3
    PROGRESS_STOP = 2
End Enum
'Public Sub Update(Byval

Public Function UpdateProgress(Source As Variant, Destination As Variant, ByVal callbackreason As CopyProgressRoutineStateChangeConstants, _
    Optional FileSize As Double = -1, Optional FileProgress As Double = -1, Optional StreamSize As Double = -1, Optional StreamProgress As Double = -1) As CopyProgressRoutineReturnConstants
'
   'FileSize = LI2DBL(TotalFileSizeHigh, TotalFileSizeLow)
            'BytesTransferred = LI2DBL(TotalBytesTransferredHigh, TotalBytesTransferredLow)
            'StreamSize = LI2DBL(StreamSizeHigh, StreamSizeLow)
            'streambytes = LI2DBL(StreamBytesTransferredHigh, StreamBytesTransferredLo)
'return true to cancel operation.
Debug.Print "Default IProgressCallback::UpdateProgress called"
If TypeOf Source Is CFile Then Debug.Print "Source is " & Source.Fullpath,
If TypeOf Destination Is CFile Then Debug.Print "destination is " & Destination.Fullpath

End Function
