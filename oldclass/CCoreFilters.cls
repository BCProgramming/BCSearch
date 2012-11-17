VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCoreFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CompressionConstants
    LZW_Compress
    LZW_Expand
    Huffman_Compress
    Huffman_Expand
End Enum
'Implements the IStreamFilter interface for filtering of stream data.

Implements IStreamFilter
Private mProgressObj As IProgressCallback
Private mvarFilterType As CompressionConstants
Public Property Let FilterType(ByVal vdata As CompressionConstants)
mvarFilterType = vdata
End Property
Public Property Get FilterType() As CompressionConstants
    FilterType = mvarFilterType
End Property
Private Sub LZW_CompressData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim BytesRead() As Byte
    Dim WriteBytes() As Byte
    
    BytesRead = inputstream.ReadBytes(inputstream.size)
    
 WriteBytes = mdlCompress.LZW_CompressData(BytesRead())
    
    outputstream.WriteBytes WriteBytes

End Sub
Private Sub LZW_ExpandData(inputstream As IInputStream, outputstream As IOutputStream)
'
Dim ReadBytes() As Byte
Dim WriteBytes() As Byte

ReadBytes() = inputstream.ReadBytes(inputstream.size)
WriteBytes = LZW_DecompressData(ReadBytes)
outputstream.WriteBytes WriteBytes()


End Sub
Private Sub Huff_ExpandData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim ReadBytes() As Byte
    Dim WriteBytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    ReadBytes() = inputstream.ReadBytes(inputstream.size)
    chuff.DecodeByte ReadBytes, UBound(ReadBytes) + 1
    
    outputstream.WriteBytes ReadBytes()



End Sub
Private Sub Huff_CompressData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim ReadBytes() As Byte
    Dim WriteBytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    ReadBytes() = inputstream.ReadBytes(inputstream.size)
    chuff.EncodeByte ReadBytes, UBound(ReadBytes) + 1
    
    outputstream.WriteBytes ReadBytes()



End Sub
Private Sub IStreamFilter_FilterStream(inputstream As IInputStream, outputstream As IOutputStream)
    
    Select Case mvarFilterType
    'Set mProgressObj = mProgress
    Case LZW_Compress
        LZW_CompressData inputstream, outputstream
    Case LZW_Expand
        LZW_ExpandData inputstream, outputstream
    Case Huffman_Compress
        Huff_CompressData inputstream, outputstream
    Case Huffman_Expand
        Huff_ExpandData inputstream, outputstream
    End Select
End Sub
