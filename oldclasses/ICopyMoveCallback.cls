VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICopyMoveCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'A Callback invoked by various "custom" copy/Move routines.

'These exist in various objects, Cfile, Cdirectory, etc. So no assumption should be made pertaining to what  the Source and destination objects are.



Public Function InitCopy(StreamSource As IInputStream, StreamDest As IOutputStream, ByRef chunksize As Long)

'StreamSource is about to be copied to StreamDest...

'be aware that these may be Alternate Data Streams opened via the CAlternateStreams OpenAsBinaryStream Method- in most cases this is irrelevant.
Debug.Print "initcopy, chunksize=" & chunksize

End Function

Public Function StreamProgress(StreamSource As IInputStream, StreamDest As IOutputStream, ByVal chunksize As Long, ByVal Chunkstransferred As Long, TotalSize As Long) As Boolean
    'return TRUE to cancel the copy of this stream. it will continue with the next file/stream...
    Debug.Print "stream progress, transferred " & Chunkstransferred & " chunks of " & chunksize & " bytes each."


End Function

