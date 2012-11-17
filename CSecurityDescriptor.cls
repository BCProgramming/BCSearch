VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSecurityDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'
'
'Public Type SECURITY_DESCRIPTOR
'    Revision As Byte
'    Sbz1 As Byte
'    Control As Long
'    Owner As Long
'    Group As Long
'    Sacl As ACL
'    Dacl As ACL
'End Type
Public Enum ACLSecurityInformationConstants
    OWNER_SECURITY_INFORMATION = &H1
    GROUP_SECURITY_INFORMATION = &H2
    DACL_SECURITY_INFORMATION = &H4
    SACL_SECURITY_INFORMATION = &H8
    LABEL_SECURITY_INFORMATION = &H10
    PROTECTED_DACL_SECURITY_INFORMATION = &H80000000
    PROTECTED_SACL_SECURITY_INFORMATION = &H40000000
    UNPROTECTED_DACL_SECURITY_INFORMATION = &H20000000
    UNPROTECTED_SACL_SECURITY_INFORMATION = &H10000000
End Enum

Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Private Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Function GetSecurityDescriptorControl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByRef pControl As Integer, ByRef lpdwRevision As Long) As Long
Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByRef lpbDaclPresent As Long, ByRef pDacl As ACL, ByRef lpbDaclDefaulted As Long) As Long
Private Declare Function GetSecurityDescriptorGroup Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByRef pGroup As Any, ByVal lpbGroupDefaulted As Long) As Long
Private Declare Function GetSecurityDescriptorLength Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Function GetSecurityDescriptorOwner Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByRef pOwner As Any, ByVal lpbOwnerDefaulted As Long) As Long
Private Declare Function GetSecurityDescriptorRMControl Lib "advapi32.dll" (ByRef SecurityDescriptor As SECURITY_DESCRIPTOR, ByVal RMControl As String) As Long
Private Declare Function GetSecurityDescriptorSacl Lib "advapi32.dll" (ByRef pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal lpbSaclPresent As Long, ByRef pSacl As ACL, ByVal lpbSaclDefaulted As Long) As Long

Private mSacl As CACL
Private mDacl As CACL
Private mDescriptor As SECURITY_DESCRIPTOR











Friend Sub InitFromPath(ByVal StrFilepath As String)
Dim ret As Long, lplengthneeded As Long
ret = GetFileSecurity(StrFilepath, OWNER_SECURITY_INFORMATION, mDescriptor, Len(mDescriptor), lplengthneeded)



End Sub
Friend Sub Init(ByVal Revision As Byte, ByVal Sbz1 As Byte, ByVal Control As Long, ByVal Owner As Long, _
                ByVal Group As Long, sAcl As CACL, dacl As CACL)
'

    With mDescriptor
        .Revision = Revision
        .Sbz1 = Sbz1
        .Control = Control
        .Owner = Owner
        .Group = Group
        Set mSacl = sAcl
        Set mDacl = dacl
    
    
    End With




End Sub


Public Property Let Group(ByVal Vdata As Long)
    mDescriptor.Group = Vdata
End Property
Public Property Get Group() As Long
    Owner = mDescriptor.Group
End Property


Public Property Let Owner(ByVal Vdata As Long)
    mDescriptor.Owner = Vdata
End Property
Public Property Get Owner() As Long
    Owner = mDescriptor.Owner
End Property





Public Property Let Control(ByVal Vdata As Long)
    mDescriptor.Control = Vdata
End Property
Public Property Get Control() As Long
    Control = mDescriptor.Control
End Property



Public Property Let Revision(ByVal Vdata As Byte)
    mDescriptor.Revision = Vdata
End Property
Public Property Get Revision() As Byte
    Revision = mDescriptor.Revision
End Property

Public Property Let Sbz1(ByVal Vdata As Byte)
    mDescriptor.Sbz1 = Vdata
End Property
Public Property Get Sbz1() As Byte
    Sbz1 = mDescriptor.Sbz1
End Property

