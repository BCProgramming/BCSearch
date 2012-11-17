VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CItemExtraData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mREMatchCol() As Object 'array of matchcollections
Private mREMatchExpr() As String 'corresponding array of regular expressions used for each matchcollection


Public Sub SetREMatchCol(newvalue() As Object)
    mREMatchCol = newvalue
End Sub
Public Function GetREMatchCol() As Object()
    GetREMatchCol = mREMatchCol
End Function
Public Sub setREmatchExpr(newvalue() As String)
    mREMatchExpr = newvalue
End Sub
Public Function getREMatchExpr() As String()
    getREMatchExpr = mREMatchExpr
End Function
