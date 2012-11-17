VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Volumes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function FindFirstVolume Lib "kernel32.dll" Alias "FindFirstVolumeA" (ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
Private Declare Function FindNextVolume Lib "kernel32.dll" Alias "FindNextVolumeA" (ByVal hFindVolume As Long, ByVal lpszVolumeName As String, ByVal cchBufferLength As Long) As Long
Private Declare Function FindVolumeClose Lib "kernel32.dll" (ByVal hFindVolume As Long) As Long
Private Declare Function GetVolumePathName Lib "kernel32.dll" Alias "GetVolumePathNameA" (ByVal lpszFileName As String, ByVal lpszVolumePathName As String, ByVal cchBufferLength As Long) As Long

Private Declare Function GetVolumePathNamesForVolumeName Lib "kernel32.dll" Alias "GetVolumePathNamesForVolumeNameA" (ByVal lpszVolumeName As String, ByVal lpszVolumePathNames As String, ByVal Bufferlength As Long, ByVal lpcchreturnlength As Long) As Long
'BOOL WINAPI GetVolumePathNamesForVolumeName(
'  __in   LPCTSTR lpszVolumeName,
'  __out  LPTSTR lpszVolumePathNames,
'  __in   DWORD cchBufferLength,
'  __out  PDWORD lpcchReturnLength
');

'Implements IVariantEnumerable
Private mhfind As Long
'Private mEnumerator As VBFCustomCollection.IEnumVARIANT


Public Function Count()
    'return count of volumes.
    Dim lpvolname As String
    Dim currcount As Long, findh As Long
    Dim ret As Long
    lpvolname = Space$(256)
    findh = FindFirstVolume(lpvolname, Len(lpvolname))
    currcount = 1
    Do
        lpvolname = Space$(256)
        ret = FindNextVolume(findh, lpvolname, Len(lpvolname))
        currcount = currcount + 1
        
    Loop Until ret = 0
    
    'if err.LastDllError=ERROR_NO_MORE_FILES
    
    
    Count = currcount
End Function

'Public Function NewEnum() As IUnknown
'
'    Dim mEnumerator As EnumVARIANT
'    Set mEnumerator = New EnumVARIANT
'    mEnumerator.Init Me
'    Set NewEnum = mEnumerator
'
'
'End Function

Private Sub Class_Initialize()
mTotalObjectCount = mTotalObjectCount + 1
End Sub

Private Sub Class_Terminate()
mTotalObjectCount = mTotalObjectCount - 1
End Sub

'Private Function IVariantEnumerable_Clone() As stdole.IEnumVARIANT
''
'End Function
Public Function GetNext(Optional ByRef ReturnVal As Object) As CVolume
    Dim ret As Variant
    Dim cast As CVolume
    If Not IVariantEnumerable_GetNext(ret) Then
        Set GetNext = Nothing
    
    Else
    Set cast = ret
    Set GetNext = cast
    Set ReturnVal = cast
    End If



End Function

Private Function IVariantEnumerable_GetNext(GetVar As Variant) As Long
'
    Dim retVal As Long, volmake As CVolume, lpvolname As String
    lpvolname = Space$(255)
    If mhfind = 0 Then
        
        mhfind = FindFirstVolume(lpvolname, Len(lpvolname) - 1)
    Else
        
        retVal = FindNextVolume(mhfind, lpvolname, Len(lpvolname) - 1)
        'return false to indicate end of the line- return true otherwise.
        If retVal = 0 Then
            If Err.LastDllError = ERROR_NO_MORE_FILES Then
                FindVolumeClose mhfind
                mhfind = 0
                IVariantEnumerable_GetNext = False
                Exit Function
            End If
        Else
         
        
        End If
    End If
    Dim volname As String
    volname = Space$(255)
    'GetVolumePathName lpvolname, volname, Len(volname)
    Dim retlength As Long
    GetVolumePathNamesForVolumeName lpvolname, volname, Len(volname), retlength
    'here be the issue...
    'just use the first one...
    lpvolname = Split(volname, vbNullChar)(0)
    'lpvolname = Left$(lpvolname, InStr(lpvolname, vbNullChar))
     'create/populate a Volume object...
     
     Debug.Print volname
     
     Set volmake = FileSystem.GetVolume(lpvolname)
     Set GetVar = volmake
     IVariantEnumerable_GetNext = True
End Function


'Private Function IVariantEnumerable_GetNext(GetVar As Variant) As Long
''
'    Dim retVal As Long, volmake As CVolume, lpvolname As String
'    lpvolname = Space$(255)
'    If mhfind = 0 Then
'
'        mhfind = FindFirstVolume(lpvolname, Len(lpvolname) - 1)
'    Else
'
'        retVal = FindNextVolume(mhfind, lpvolname, Len(lpvolname) - 1)
'        'return false to indicate end of the line- return true otherwise.
'        If retVal = 0 Then
'            If Err.LastDllError = ERROR_NO_MORE_FILES Then
'                FindVolumeClose mhfind
'                mhfind = 0
'                IVariantEnumerable_GetNext = False
'                Exit Function
'            End If
'        Else
'
'
'        End If
'    End If
'    Dim volname As String
'    volname = Space$(255)
'    'GetVolumePathName lpvolname, volname, Len(volname)
'    Dim retlength As Long
'    GetVolumePathNamesForVolumeName lpvolname, volname, Len(volname), retlength
'    'here be the issue...
'    'just use the first one...
'    lpvolname = Split(volname, vbNullChar)(0)
'    'lpvolname = Left$(lpvolname, InStr(lpvolname, vbNullChar))
'     'create/populate a Volume object...
'
'     Debug.Print volname
'
'     Set volmake = FileSystem.GetVolume(lpvolname)
'     Set GetVar = volmake
'     IVariantEnumerable_GetNext = True
'End Function

Private Sub IVariantEnumerable_Reset()
'
If mhfind <> 0 Then FindVolumeClose mhfind
End Sub

Private Sub IVariantEnumerable_Skip(ByVal celt As Long)
'
End Sub
