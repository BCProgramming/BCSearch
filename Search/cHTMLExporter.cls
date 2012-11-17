VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHTMLExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IExportPlugin

Private Function IExportPlugin_AddPluginIcon(AddTo As Object) As String
'
IExportPlugin_AddPluginIcon = "EXPORT_HTML"
End Function

Private Sub IExportPlugin_Export(ItemsExport As Object)
'



End Sub


