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
    Stream_Reverse = -1
    lzw_compress
    LZW_Expand
    LZSS_Compress
    LZSS_Expand
    Huffman_Compress
    Huffman_Expand
    encodedecode_xor
    Encrypt_CryptoAPI
    Decrypt_CryptoAPI
    
End Enum
'Implements the IStreamFilter interface for filtering of stream data.
Private Const PROV_RSA_FULL = 1

Private Const CRYPT_NEWKEYSET = &H8

Private Const ALG_CLASS_HASH = 32768
Private Const ALG_CLASS_DATA_ENCRYPT = 24576&

Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536&
Private Const ALG_TYPE_STREAM = 2048&

Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA1 = 4

Private Const ALG_SID_DES = 1
Private Const ALG_SID_3DES = 3
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1

Public Enum HashAlgorithmEnum
   MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
   MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
   MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
   SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum

Public Enum EncAlgorithmEnum
   DES = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES
   [3DES] = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES
   RC2 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2
   RC4 = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4
End Enum
Implements IStreamFilter
Private mProgressObj As IProgressCallback
Private mvarFilterType As CompressionConstants
Private mPassword As String
Private mhashAlgorithm As HashAlgorithmEnum
Private mEncAlgorithm As EncAlgorithmEnum
'password used by a number of encryption routines.
Public Property Let hashAlgorithm(ByVal vdata As HashAlgorithmEnum)
    mhashAlgorithm = vdata
End Property
Public Property Get hashAlgorithm() As HashAlgorithmEnum
    hashAlgorithm = mhashAlgorithm
End Property
Public Property Let encAlgorithm(ByVal vdata As EncAlgorithmEnum)
    mEncAlgorithm = vdata
End Property
Public Property Get encAlgorithm() As EncAlgorithmEnum
    encAlgorithm = mEncAlgorithm
End Property

Public Property Let password(ByVal vdata As String)
    mPassword = vdata
End Property
Public Property Get password() As String
    password = mPassword
End Property
Public Property Let FilterType(ByVal vdata As CompressionConstants)
mvarFilterType = vdata
End Property
Public Property Get FilterType() As CompressionConstants
    FilterType = mvarFilterType
End Property
'Private Function GetMangledPassword(ByVal pwdMangle As String) As String
''only requirement: return the same output for the same input. Ideally, mix things up as much as possible.
'
''my method? increase the entropy by looping forward through the string and adding up the ASCII values at each character; and looping back doing the same. then use MOD to return that to the domain of characters (1 to 255).
'
'
'Dim CharCounts() As Long
'Dim Currpos As Long, runningtotal As Long
'Dim retstr As String
'ReDim CharCounts(1 To Len(pwdMangle))
''reset the randomizer and make sure we get the same sequence for the same strings.
'
'Dim useseed As Long
'useseed = 12
'For Currpos = 1 To Len(pwdMangle)
'    useseed = useseed * Asc(Mid$(pwdMangle, Currpos, 1))
'    useseed = useseed Mod 32768
'Next Currpos
'
'
'Randomize useseed: Rnd -1
'For Currpos = 1 To Len(pwdMangle)
'
'    runningtotal = runningtotal + Asc(Mid$(pwdMangle, Currpos, 1)) + Int(Rnd * 255)
'
'    CharCounts(Currpos) = runningtotal
'
'Next Currpos
'runningtotal = 0
'
'For Currpos = Len(pwdMangle) To 1 Step -1
'    runningtotal = runningtotal + Asc(Mid$(pwdMangle, Currpos, 1)) * Int(Rnd * 255)
'    CharCounts(Currpos) = runningtotal
'Next Currpos
'Dim randA As Long, RandB As Long
'For Currpos = 1 To UBound(CharCounts)
'    'pick two "random" slots...
'    randA = Int(Rnd * (UBound(CharCounts) - 1)) + 1
'    RandB = Int(Rnd * (UBound(CharCounts) - 1)) + 1
'    CharCounts(randA) = CharCounts(randA) \ (CharCounts(RandB) + 1)
'    CharCounts(RandB) = CharCounts(RandB) + Int(Rnd * CharCounts(randA))
'
'Next Currpos
'
'For Currpos = 1 To UBound(CharCounts)
'    CharCounts(Currpos) = CharCounts(Currpos) Mod 255
'    retstr = retstr & Chr$(CharCounts(Currpos))
'Next
''GetMangledPassword = StrConv(CharCounts, vbUnicode)
'
'GetMangledPassword = retstr
'
'
'
'
'
'
'
'End Function


Private Sub XorEncodeDecode(inputstream As IInputStream, outputstream As IOutputStream)
    Dim pwduse As String
    pwduse = mPassword
    If pwduse = "" Then pwduse = " "
    'Debug.Print "password:" & pwduse;
    'pwduse = GetMangledPassword(pwduse)
    'Debug.Print "mangled:" & pwduse
    Dim CurrBlock() As Byte, Chunksize As Long, readchunk As Long, WriteChunk() As Byte
    Dim i As Long
    Chunksize = 32768
    readchunk = 32768
    
    
    Do Until inputstream.EOF
    
        CurrBlock = inputstream.readbytes(Chunksize, readchunk)
        ReDim WriteChunk(1 To readchunk)
        For i = 1 To readchunk
        
        
        WriteChunk(i) = CurrBlock(i) Xor CByte(Asc(Mid$(pwduse, (i Mod (Len(pwduse) - 1)) + 1, 1)))
        
        
        Next i
        
        
        outputstream.WriteBytes WriteChunk
    
    Loop
    
    
    



End Sub
Private Sub LZW_CompressData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim bytesread() As Byte
    Dim WriteBytes() As Byte
    
    bytesread = inputstream.readbytes(inputstream.Size)
    
 WriteBytes = mdlCompress.LZW_CompressData(bytesread())
    
    outputstream.WriteBytes WriteBytes

End Sub
Private Sub LZSS_CompressData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim WriteBytes() As Byte
    readbytes() = inputstream.readbytes(inputstream.Size)
    ModLZSS.InitializeLZSS
    WriteBytes() = ModLZSS.Compact(readbytes())
    outputstream.WriteBytes WriteBytes

End Sub
Private Sub LZSS_ExpandData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim WriteBytes() As Byte
    readbytes() = inputstream.readbytes(inputstream.Size)
    ModLZSS.InitializeLZSS
    WriteBytes() = ModLZSS.UnCompact(readbytes())
    outputstream.WriteBytes WriteBytes

End Sub
Private Sub LZW_ExpandData(inputstream As IInputStream, outputstream As IOutputStream)
'
Dim readbytes() As Byte
Dim WriteBytes() As Byte

readbytes() = inputstream.readbytes(inputstream.Size)
WriteBytes = LZW_DecompressData(readbytes)
outputstream.WriteBytes WriteBytes()


End Sub
Private Sub Huff_ExpandData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim WriteBytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    readbytes() = inputstream.readbytes(inputstream.Size)
    chuff.DecodeByte readbytes, UBound(readbytes) + 1
    
    outputstream.WriteBytes readbytes()



End Sub
Private Sub Huff_CompressData(inputstream As IInputStream, outputstream As IOutputStream)
    Dim readbytes() As Byte
    Dim WriteBytes() As Byte
    Static chuff As clsHuffman
    If chuff Is Nothing Then Set chuff = New clsHuffman
    readbytes() = inputstream.readbytes(inputstream.Size)
    chuff.EncodeByte readbytes, UBound(readbytes) + 1
    
    outputstream.WriteBytes readbytes()



End Sub
Private Sub EncryptCrypto(inputstream As IInputStream, outputstream As IOutputStream)
    '
    Dim readbytes() As Byte, Encrypted() As Byte
    Dim ReadString As String
    Dim StrBytes As String
    
    'ReadBytes = inputstream.ReadBytes(inputstream.Size)
    readbytes = inputstream.readbytes(inputstream.Size)
    
    Encrypted = modCrypt.EncryptData(readbytes, mPassword, mhashAlgorithm, mEncAlgorithm)
    

    outputstream.WriteBytes Encrypted()
    
    
    
    
    


End Sub
Private Sub DecryptCrypto(inputstream As IInputStream, outputstream As IOutputStream)
    '
    Dim readbytes() As Byte, decrypted() As Byte
    Dim StrBytes As String, StrDecrypted As String
    
    'ReadBytes = inputstream.ReadBytes(inputstream.Size)
    'StrBytes = StrConv(ReadBytes, vbUnicode)
    readbytes = inputstream.readbytes(inputstream.Size)
    decrypted = modCrypt.DecryptData(readbytes, mPassword, mhashAlgorithm, mEncAlgorithm)
    
    'cheat...just write it out as an ASCII string.
    'outputstream.WriteString StrDecrypted, StrRead_ANSI
    outputstream.WriteBytes decrypted()
    
    
    
    
    

End Sub

Private Sub Class_Initialize()
mEncAlgorithm = RC4
mhashAlgorithm = MD5
End Sub

Private Sub ReverseStream(inputstream As IInputStream, outputstream As IOutputStream)

Dim bytesread() As Byte
Dim reversed() As Byte
Dim i As Long
bytesread = inputstream.readbytes(inputstream.Size)
ReDim reversed(UBound(bytesread))
For i = 0 To UBound(bytesread)
    reversed(i) = bytesread(UBound(bytesread) - i)


Next i
outputstream.WriteBytes reversed()

End Sub
Public Sub FilterStream(inputstream, outputstream, Optional ByVal flReverse)
Dim castinput As IInputStream
Dim castoutput As IOutputStream
Set castinput = inputstream
Set castoutput = outputstream
    IStreamFilter_FilterStream castinput, castoutput, flReverse
End Sub
Private Sub IStreamFilter_FilterStream(inputstream As IInputStream, outputstream As IOutputStream, Optional ByVal flReverse As Boolean = False)
Dim mfiltertype As CompressionConstants
mfiltertype = mvarFilterType
If flReverse Then
    Select Case mvarFilterType
        Case lzw_compress
            mfiltertype = LZW_Expand
        Case LZW_Expand
            mfiltertype = lzw_compress
        Case LZSS_Compress
            mfiltertype = LZSS_Expand
        Case LZSS_Expand
            mfiltertype = LZSS_Compress
        Case Huffman_Compress
            mfiltertype = Huffman_Expand
        Case Huffman_Expand
            mfiltertype = Huffman_Compress
        Case encodedecode_xor
            mfiltertype = encodedecode_xor
        Case Encrypt_CryptoAPI
            mfiltertype = Decrypt_CryptoAPI
        Case Decrypt_CryptoAPI
            mfiltertype = Encrypt_CryptoAPI
    End Select
    
End If
    Select Case mfiltertype
    'Set mProgressObj = mProgress
    Case lzw_compress
        LZW_CompressData inputstream, outputstream
    Case LZW_Expand
        LZW_ExpandData inputstream, outputstream
    Case LZSS_Compress
        LZSS_CompressData inputstream, outputstream
    Case LZSS_Expand
        LZSS_ExpandData inputstream, outputstream
    Case Huffman_Compress
        Huff_CompressData inputstream, outputstream
    Case Huffman_Expand
        Huff_ExpandData inputstream, outputstream
    Case encodedecode_xor
        XorEncodeDecode inputstream, outputstream
    Case Encrypt_CryptoAPI
        EncryptCrypto inputstream, outputstream
    Case Decrypt_CryptoAPI
        DecryptCrypto inputstream, outputstream
    Case Stream_Reverse
        ReverseStream inputstream, outputstream
    End Select


End Sub
