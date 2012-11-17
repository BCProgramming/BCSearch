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

'LZW compressor class.
Public Enum CompressionConstants
    LZW_Compress
    LZW_Expand
    Huffman_Compress
    Huffman_Expand
End Enum
'Implements the IStreamFilter interface for filtering of stream data.

Implements IStreamFilter
Private mvarFilterType As CompressionConstants
Public Property Let FilterType(ByVal Vdata As CompressionConstants)
mvarFilterType = Vdata
End Property
Public Property Get FilterType() As CompressionConstants
    FilterType = mvarFilterType
End Property
Private Sub LZWCompress(inputstream As IInputStream, Outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim writebytes() As Byte
    
    readbytes = inputstream.readbytes(inputstream.size)
    
    writebytes = LZW_Compress(readbytes)
    
    Outputstream.writebytes writebytes

End Sub
Private Sub LZWExpand(inputstream As IInputStream, Outputstream As IOutputStream)
'
Dim readbytes() As Byte
Dim writebytes() As Byte

readbytes() = inputstream.readbytes(inputstream.size)
writebytes = LZW_Decompress(readbytes)
Outputstream.writebytes writebytes()


End Sub
Private Sub Huff_Expand(inputstream As IInputStream, Outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim writebytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    readbytes() = inputstream.readbytes(inputstream.size)
    chuff.DecodeByte readbytes, UBound(readbytes) + 1
    
    Outputstream.writebytes readbytes()



End Sub
Private Sub Huff_Compress(inputstream As IInputStream, Outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim writebytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    readbytes() = inputstream.readbytes(inputstream.size)
    chuff.EncodeByte readbytes, UBound(readbytes) + 1
    
    Outputstream.writebytes readbytes()



End Sub
Private Sub IStreamFilter_FilterStream(inputstream As IInputStream, Outputstream As IOutputStream)
    Select Case mvarFilterType
    Case LZW_Compress
        LZWCompress inputstream, Outputstream
    Case LZW_Expand
        LZWExpand inputstream, Outputstream
    Case Huffman_Compress
        Huff_Compress inputstream, Outputstream
    Case Huffman_Expand
        Huff_Expand inputstream, Outputstream
    End Select
End Sub
