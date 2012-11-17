VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFileContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'IFileCOntainer: Implemented... by, well, file containers.
'If I add zip support, for example- that will likely have a IFileContainer Implementation.


'a File container will only need to support two methods: one that retrieves a stream enumerator (for directories, this would enumerate files) and one that returns a enumerator
'for more IFileContainers (subdirectories, for example).


Public Function GetStreamEnumerator() As IFileContainerEnumerator
'
End Function
Public Function GetContainerEnumerator() As IFileContainerEnumerator
'
End Function
