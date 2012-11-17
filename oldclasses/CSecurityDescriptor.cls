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

Private mSacl As CACL
Private mDacl As CACL
Private mDescriptor As SECURITY_DESCRIPTOR

Friend Sub Init(ByVal Revision As Byte, ByVal sbz1 As Byte, ByVal Control As Long, ByVal Owner As Long, _
                ByVal Group As Long, sAcl As CACL, dacl As CACL)
'

    With mDescriptor
        .Revision = Revision
        .sbz1 = sbz1
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

Public Property Let sbz1(ByVal Vdata As Byte)
    mDescriptor.sbz1 = Vdata
End Property
Public Property Get sbz1() As Byte
    sbz1 = mDescriptor.sbz1
End Property

