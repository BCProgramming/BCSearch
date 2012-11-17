VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CZipFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Overall .ZIP file format:
'
'    [local file header 1]
'    [file data 1]
'    [data descriptor 1]
'    .
'    .
'    .
'    [local file header n]
'    [file data n]
'    [data descriptor n]
'    [archive decryption header]
'    [archive extra data record]
'    [central directory]
'    [zip64 end of central directory record]
'    [zip64 end of central directory locator]
'    [end of central directory record]

' A.  Local file header:
'
'        local file header signature     4 bytes  (0x04034b50)
'        version needed to extract       2 bytes
'        general purpose bit flag        2 bytes
'        compression method              2 bytes
'        last mod file time              2 bytes
'        last mod file date              2 bytes
'        crc-32                          4 bytes
'        compressed size                 4 bytes
'        uncompressed size               4 bytes
'        file name length                2 bytes
'        extra field length              2 bytes
'
'        file name (variable size)
'        extra field (variable size)
Private Type LOCAL_FILE_HEADER
    Signature As Long
    VersionRequired As Integer
    GPBitFlag As Integer
    CompressionMethod As Integer
    LastModifiedTime As Integer
    LastModifiedDate As Integer
    crc_32 As Long
    CompressedSize As Long
    UncompressedSize As Long
    FilenameLength As Integer
    ExtraFieldLength As Integer
    Filename As String
    ExtraField() As Byte
End Type
'data descriptor
'    crc-32                          4 bytes
'        compressed size                 4 bytes
'        uncompressed size               4 bytes
'
'      This descriptor exists only if bit 3 of the general
'      purpose bit flag is set (see below).  It is byte aligned
'      and immediately follows the last byte of compressed data.
'      This descriptor is used only when it was not possible to
'      seek in the output .ZIP file, e.g., when the output .ZIP file
'      was standard output or a non-seekable device.  For ZIP64(tm) format
'      archives, the compressed and uncompressed sizes are 8 bytes each.

Private Type Data_Descriptor_32
  crc32 As Long
  CompressedSize As Long
  UncompressedSize As Long
End Type

'64-bit for zip64:
Private Type Data_descriptor_64
   crc32 As Long
   CompressedSize As Currency
   UncompressedSize As Currency
End Type
