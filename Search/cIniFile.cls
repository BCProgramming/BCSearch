VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CIniFile
'Monday February 19th, 2006
'Michael Burgwin

'Even though we are all supposed to store our application's settings in the registry, there
'are some kinds of settings you want the user to be able to change, or, you just want a way
'to export and import settings.

'Although you *could* use the windows API profile functions, it is probably best to devise your
