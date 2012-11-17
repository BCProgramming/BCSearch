Attribute VB_Name = "WinINet"
Option Explicit
Public Enum INTERNET_PORT
 INTERNET_DEFAULT_FTP_PORT& = 21
 INTERNET_DEFAULT_GOPHER_PORT& = 70
 INTERNET_DEFAULT_HTTP_PORT& = 80
 INTERNET_DEFAULT_HTTPS_PORT& = 443
 INTERNET_DEFAULT_SOCKS_PORT& = 1080


End Enum
Private Declare Function InternetOpenA Lib "wininet.dll" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszproxy As String, ByVal lpszProxyBypass As String, ByVal dwflags As Long) As Long
Private Declare Function InternetOpenW Lib "wininet.dll" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszproxy As Long, ByVal lpszProxyBypass As Long, ByVal dwflags As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByRef hinternet As Long) As Boolean

Private Declare Function InternetConnectA Lib "wininet.dll" (ByRef hinternet As Long, ByVal lpszServerName As String, ByRef nServerPort As Long, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwflags As Long, ByRef dwContext As Long) As Long
Private Declare Function InternetConnectW Lib "wininet.dll" (ByRef hinternet As Long, ByVal lpszServerName As Long, ByRef nServerPort As Long, ByVal lpszUserName As Long, ByVal lpszPassword As Long, ByVal dwService As Long, ByVal dwflags As Long, ByRef dwContext As Long) As Long

Public Function InternetOpen(ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszproxy As String, ByVal lpszProxyBypass As String, ByVal dwflags As Long) As Long
    If MakeWideCalls Then
        InternetOpen = InternetOpenW(StrPtr(lpszAgent), dwAccessType, StrPtr(lpszproxy), StrPtr(lpszProxyBypass), dwflags)

    Else
        InternetOpen = InternetOpenA(lpszAgent, dwAccessType, lpszproxy, lpszProxyBypass, dwflags)
    End If


End Function

Public Function InternetConnect(ByRef hinternet As Long, ByVal lpszServerName As String, ByRef nServerPort As INTERNET_PORT, ByVal lpszUserName As String, ByVal lpszPassword As String, ByVal dwService As Long, ByVal dwflags As Long, ByRef dwContext As Long) As Long
    If MakeWideCalls Then
        InternetConnect = InternetConnectW(hinternet, StrPtr(lpszServerName), nServerPort, StrPtr(lpszUserName), StrPtr(lpszPassword), dwService, dwflags, dwContext)


    Else
        InternetConnect = InternetConnectA(hinternet, lpszServerName, nServerPort, lpszUserName, lpszPassword, dwService, dwflags, dwContext)


    End If



End Function

Public Sub inettest()
    Dim hinternet As Long, hconnection As Long
    
    hinternet = InternetOpen("Mozilla", 1, "", "", 0)
    hconnection = InternetConnect(hinternet, "207.46.192.254", INTERNET_DEFAULT_FTP_PORT, "", "", 1, 1, 0)
    
    
    
    
    InternetCloseHandle hconnection
    InternetCloseHandle hinternet



End Sub
