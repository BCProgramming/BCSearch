VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExportPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    



'IExportPlugin: Interface supported by... err- export plugins.
Public Function AddPluginIcon(AddTo As Object) As String
    'add to the passed cVBAImageList object, and return the added key.
    'the built in plugins skip that and just return a key as added in the XML file, haha.
End Function



Public Sub Export(ItemsExport As Object)
    'Perform the export.
    'ItemsExport is a cListItems Object.
    
End Sub
