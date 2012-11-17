VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLZCompressedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function LZInit Lib "lz32.dll" (ByVal hfSrc As Long) As Long
Private Declare Sub LZDone Lib "lz32.dll" ()
Private Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Private Declare Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)
Private Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, ByRef lpOf As OFSTRUCT, ByVal style As Long) As Long
Private Declare Function LZRead Lib "lz32.dll" (ByVal hfFile As Long, ByVal lpvBuf As String, ByVal cbread As Long) As Long
Private Declare Function LZSeek Lib "lz32.dll" (ByVal hfFile As Long, ByVal lOffset As Long, ByVal nOrigin As Long) As Long
Private Declare Function LZStart Lib "lz32.dll" () As Long


Private LZHandle As Long


Friend Sub Init(ByVal Filehandle As Long)
    If LZHandle <> 0 Then
        LZClose LZHandle
        LZHandle = 0
    End If
    LZHandle = LZInit(Filehandle)
End Sub
