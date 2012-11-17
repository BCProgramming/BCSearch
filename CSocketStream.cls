VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketOutputStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CSocketStream
Implements IOutputStream
Private WithEvents mSocketObj As CSocketMaster
Attribute mSocketObj.VB_VarHelpID = -1
Private mestablished As Boolean
Private mCurrpos As Long
'Friend Sub
Friend Sub Connect(RemoteHost, RemotePort)
    Set mSocketObj = New CSocketMaster
    mSocketObj.Protocol = sckTCPProtocol
    mSocketObj.RemoteHost = RemoteHost
    mSocketObj.RemotePort = RemotePort
    mSocketObj.Connect
    
End Sub

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    mSocketObj.Close
    Set mSocketObj = Nothing
End Sub

Private Sub IOutputStream_CloseStream()
'
End Sub

Private Function IOutputStream_GetPos() As Double
'
IOutputStream_GetPos = Currpos
End Function

Private Function IOutputStream_SeekTo(ToPos As Double, Optional ByVal SeekMode As FileSeekBase = 0&) As Variant
'
'seeks fail....
End Function

Private Function IOutputStream_Valid() As Boolean
'
End Function

Private Function IOutputStream_WriteBytes(Bytes() As Byte) As Boolean
'
Dim mbytes As Variant
mbytes = Bytes
mCurrpos = mCurrpos + UBound(Bytes) - LBound(Bytes) + 1
mSocketObj.SendData mbytes
End Function

Private Sub IOutputStream_WriteString(Stringwrite As String, Optional ByVal Stringmode As StringReadMode = -1&)
'
ModStreamFunc.WriteString Me, Stringwrite, Stringmode
End Sub




Private Sub mSocketObj_CloseSck()
mestablished = False
End Sub

Private Sub mSocketObj_Close()
mestablished = False
End Sub

Private Sub mSocketObj_Connect()
    mestablished = True
    Debug.Print "connection established"
End Sub

Private Sub mSocketObj_ConnectionRequest(ByVal requestID As Long)
'
Debug.Print "connectionrequest"
End Sub

Private Sub mSocketObj_DataArrival(ByVal bytesTotal As Long)
Debug.Print "data arrived"
End Sub

Public Sub CloseStream()
    mSocketObj.Close
End Sub
Public Sub WriteString(ByVal StrWrite As String, Optional ByVal Stringmode As StringReadMode = -1)

    IOutputStream_WriteString StrWrite, Stringmode
End Sub

Private Sub mSocketObj_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'
Debug.Print "error:" & Number & " description:" & Description
End Sub

Private Sub mSocketObj_SendComplete()
'
Debug.Print "sendcomplete"
End Sub

Private Sub mSocketObj_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'
Debug.Print "sendprogress, "; bytesSent, bytesRemaining
End Sub
