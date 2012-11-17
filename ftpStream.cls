VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ftpStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByRef hConnect As Long, ByVal lpszFileName As String, ByVal dwAccess As Long, ByVal dwFlags As Long, ByRef dwContext As Long)

Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Integer
Private Declare Sub InternetWriteFile Lib "wininet.dll" (ByRef hFile As Long, ByRef lpBuffer As Any, ByVal dwNumberOfBytesToWrite As Long, ByRef lpdwNumberOfBytesWritten As Long)

