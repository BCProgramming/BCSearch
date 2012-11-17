VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CACL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'ACL class.
'Not yet used.
'Private Type ACL
'    AclRevision As Byte
'    Sbz1 As Byte
'    AclSize As Integer
'    AceCount As Integer
'    Sbz2 As Integer
'End Type

Private mACL As ACL
Friend Property Get ACLType() As ACL
    ACLType = mACL
End Property
Public Property Let AclRevision(ByVal vdata As Byte)
    mACL.AclRevision = vdata
End Property
Public Property Get AclRevision() As Byte
    AclRevision = mACL.AclRevision
End Property

Public Property Let Sbz1(ByVal vdata As Byte)
    mACL.Sbz1 = vdata
End Property
Public Property Get Sbz1() As Byte
    Sbz1 = mACL.Sbz1
End Property

Public Property Let AclSize(ByVal vdata As Integer)
    mACL.AclSize = vdata
End Property
Public Property Get AclSize() As Integer
    AclSize = mACL.AclSize
End Property

Public Property Let AceCount(ByVal vdata As Integer)
    mACL.AceCount = vdata
End Property
Public Property Get AceCount() As Integer
    AceCount = mACL.AceCount
End Property
Public Property Let Sbz2(ByVal vdata As Integer)
    mACL.Sbz2 = vdata
End Property
Public Property Get Sbz2() As Integer
    Sbz2 = mACL.Sbz2
End Property
