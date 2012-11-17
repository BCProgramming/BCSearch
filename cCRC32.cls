VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



' This code is taken from the VB.NET CRC32 algorithm
' provided by Paul (wpsjr1@succeed.net) - Excellent work!

Private crc32Table() As Long
Private crc16Table() As Long
Private Const BUFFER_SIZE As Long = 8192
Private Function Crc16(cp() As Byte, Optional ByVal Size As Long = -1) As Long
  Dim i As Long
  Dim fcs As Long

  
  Const pppinitfcs16 = &HFFFF& 'Initial FCS value
  
  If Size < 0 Then Size = UBound(cp) - LBound(cp) + 1

 
  
  
  ' The initial FCS value
  fcs = pppinitfcs16
  
  ' evaluate the FCS
  For i = LBound(cp) To LBound(cp) + Size - 1
    fcs = (fcs \ &H100&) Xor crc16Table((fcs Xor cp(i)) And &HFF&)
  Next i
  
  ' return the result
  Crc16 = fcs
End Function

Public Function GetByteArrayCrc16(ByRef buffer() As Byte) As Long
    GetByteArrayCrc16 = Crc16(buffer())
End Function
Public Function GetByteArrayCrc32(ByRef buffer() As Byte) As Long
   
   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF
      
   Dim i As Integer
   Dim iLookup As Integer
   
   For i = LBound(buffer) To UBound(buffer)
      iLookup = (crc32Result And &HFF) Xor buffer(i)
      crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215 ' nasty       shr 8 with vb :/
      crc32Result = crc32Result Xor crc32Table(iLookup)
   Next i
   
   GetByteArrayCrc32 = Not (crc32Result)

End Function
Public Function GetStringCRC16(strbuffer As String) As Long
    Dim convbuffer() As Byte
    convbuffer = strbuffer
    GetStringCRC16 = GetByteArrayCrc16(convbuffer)
End Function
Public Function GetStringCRC32(strbuffer As String) As Long
    Dim convbuffer() As Byte
    convbuffer = strbuffer
    GetStringCRC32 = GetByteArrayCrc32(convbuffer)
End Function
Public Function GetStreamCRC16(ByRef iStream As IInputStream) As Long
Dim strread As String
strread = iStream.ReadAllStr(StrRead_unicode)
GetStreamCRC16 = GetStringCRC16(strread)

End Function
Public Function GetStreamCRC32(ByRef iStream As IInputStream) As Long

   Dim crc32Result As Long
   crc32Result = &HFFFFFFFF

   Dim buffer() As Byte
   Dim readSize As Long
   readSize = BUFFER_SIZE

   Dim count As Long
   iStream.SeekTo 0, STREAM_BEGIN
   buffer = iStream.readbytes(readSize, count)
   
   Dim i As Integer
   Dim iLookup As Integer
   Dim tot As Integer
   
   Do While (count > 0)
      For i = 0 To count - 1
         iLookup = (crc32Result And &HFF) Xor buffer(i)
         crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215 '
          'nasty shr 8 with vb :/
         crc32Result = crc32Result Xor crc32Table(iLookup)
      Next i
      buffer = iStream.readbytes(readSize, count)
   Loop

   GetStreamCRC32 = Not (crc32Result)

End Function

Private Sub Class_Initialize()

    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    Dim i As Integer, J As Integer

    ReDim crc32Table(256)
    'ReDim crc16Table(256)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For J = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next J
        crc32Table(i) = dwCrc
        'crc16Table(i) = dwCrc
    Next i
    'for crc16...
    ReDim crc16Table(0 To 255)
    crc16Table(0) = &H0&
    crc16Table(1) = &H1189&
    crc16Table(2) = &H2312&
    crc16Table(3) = &H329B&
    crc16Table(4) = &H4624&
    crc16Table(5) = &H57AD&
    crc16Table(6) = &H6536&
    crc16Table(7) = &H74BF&
    crc16Table(8) = &H8C48&
    crc16Table(9) = &H9DC1&
    crc16Table(10) = &HAF5A&
    crc16Table(11) = &HBED3&
    crc16Table(12) = &HCA6C&
    crc16Table(13) = &HDBE5&
    crc16Table(14) = &HE97E&
    crc16Table(15) = &HF8F7&
    crc16Table(16) = &H1081&
    crc16Table(17) = &H108&
    crc16Table(18) = &H3393&
    crc16Table(19) = &H221A&
    crc16Table(20) = &H56A5&
    crc16Table(21) = &H472C&
    crc16Table(22) = &H75B7&
    crc16Table(23) = &H643E&
    crc16Table(24) = &H9CC9&
    crc16Table(25) = &H8D40&
    crc16Table(26) = &HBFDB&
    crc16Table(27) = &HAE52&
    crc16Table(28) = &HDAED&
    crc16Table(29) = &HCB64&
    crc16Table(30) = &HF9FF&
    crc16Table(31) = &HE876&
    crc16Table(32) = &H2102&
    crc16Table(33) = &H308B&
    crc16Table(34) = &H210&
    crc16Table(35) = &H1399&
    crc16Table(36) = &H6726&
    crc16Table(37) = &H76AF&
    crc16Table(38) = &H4434&
    crc16Table(39) = &H55BD&
    crc16Table(40) = &HAD4A&
    crc16Table(41) = &HBCC3&
    crc16Table(42) = &H8E58&
    crc16Table(43) = &H9FD1&
    crc16Table(44) = &HEB6E&
    crc16Table(45) = &HFAE7&
    crc16Table(46) = &HC87C&
    crc16Table(47) = &HD9F5&
    crc16Table(48) = &H3183&
    crc16Table(49) = &H200A&
    crc16Table(50) = &H1291&
    crc16Table(51) = &H318&
    crc16Table(52) = &H77A7&
    crc16Table(53) = &H662E&
    crc16Table(54) = &H54B5&
    crc16Table(55) = &H453C&
    crc16Table(56) = &HBDCB&
    crc16Table(57) = &HAC42&
    crc16Table(58) = &H9ED9&
    crc16Table(59) = &H8F50&
    crc16Table(60) = &HFBEF&
    crc16Table(61) = &HEA66&
    crc16Table(62) = &HD8FD&
    crc16Table(63) = &HC974&
    crc16Table(64) = &H4204&
    crc16Table(65) = &H538D&
    crc16Table(66) = &H6116&
    crc16Table(67) = &H709F&
    crc16Table(68) = &H420&
    crc16Table(69) = &H15A9&
    crc16Table(70) = &H2732&
    crc16Table(71) = &H36BB&
    crc16Table(72) = &HCE4C&
    crc16Table(73) = &HDFC5&
    crc16Table(74) = &HED5E&
    crc16Table(75) = &HFCD7&
    crc16Table(76) = &H8868&
    crc16Table(77) = &H99E1&
    crc16Table(78) = &HAB7A&
    crc16Table(79) = &HBAF3&
    crc16Table(80) = &H5285&
    crc16Table(81) = &H430C&
    crc16Table(82) = &H7197&
    crc16Table(83) = &H601E&
    crc16Table(84) = &H14A1&
    crc16Table(85) = &H528&
    crc16Table(86) = &H37B3&
    crc16Table(87) = &H263A&
    crc16Table(88) = &HDECD&
    crc16Table(89) = &HCF44&
    crc16Table(90) = &HFDDF&
    crc16Table(91) = &HEC56&
    crc16Table(92) = &H98E9&
    crc16Table(93) = &H8960&
    crc16Table(94) = &HBBFB&
    crc16Table(95) = &HAA72&
    crc16Table(96) = &H6306&
    crc16Table(97) = &H728F&
    crc16Table(98) = &H4014&
    crc16Table(99) = &H519D&
    crc16Table(100) = &H2522&
    crc16Table(101) = &H34AB&
    crc16Table(102) = &H630&
    crc16Table(103) = &H17B9&
    crc16Table(104) = &HEF4E&
    crc16Table(105) = &HFEC7&
    crc16Table(106) = &HCC5C&
    crc16Table(107) = &HDDD5&
    crc16Table(108) = &HA96A&
    crc16Table(109) = &HB8E3&
    crc16Table(110) = &H8A78&
    crc16Table(111) = &H9BF1&
    crc16Table(112) = &H7387&
    crc16Table(113) = &H620E&
    crc16Table(114) = &H5095&
    crc16Table(115) = &H411C&
    crc16Table(116) = &H35A3&
    crc16Table(117) = &H242A&
    crc16Table(118) = &H16B1&
    crc16Table(119) = &H738&
    crc16Table(120) = &HFFCF&
    crc16Table(121) = &HEE46&
    crc16Table(122) = &HDCDD&
    crc16Table(123) = &HCD54&
    crc16Table(124) = &HB9EB&
    crc16Table(125) = &HA862&
    crc16Table(126) = &H9AF9&
    crc16Table(127) = &H8B70&
    crc16Table(128) = &H8408&
    crc16Table(129) = &H9581&
    crc16Table(130) = &HA71A&
    crc16Table(131) = &HB693&
    crc16Table(132) = &HC22C&
    crc16Table(133) = &HD3A5&
    crc16Table(134) = &HE13E&
    crc16Table(135) = &HF0B7&
    crc16Table(136) = &H840&
    crc16Table(137) = &H19C9&
    crc16Table(138) = &H2B52&
    crc16Table(139) = &H3ADB&
    crc16Table(140) = &H4E64&
    crc16Table(141) = &H5FED&
    crc16Table(142) = &H6D76&
    crc16Table(143) = &H7CFF&
    crc16Table(144) = &H9489&
    crc16Table(145) = &H8500&
    crc16Table(146) = &HB79B&
    crc16Table(147) = &HA612&
    crc16Table(148) = &HD2AD&
    crc16Table(149) = &HC324&
    crc16Table(150) = &HF1BF&
    crc16Table(151) = &HE036&
    crc16Table(152) = &H18C1&
    crc16Table(153) = &H948&
    crc16Table(154) = &H3BD3&
    crc16Table(155) = &H2A5A&
    crc16Table(156) = &H5EE5&
    crc16Table(157) = &H4F6C&
    crc16Table(158) = &H7DF7&
    crc16Table(159) = &H6C7E&
    crc16Table(160) = &HA50A&
    crc16Table(161) = &HB483&
    crc16Table(162) = &H8618&
    crc16Table(163) = &H9791&
    crc16Table(164) = &HE32E&
    crc16Table(165) = &HF2A7&
    crc16Table(166) = &HC03C&
    crc16Table(167) = &HD1B5&
    crc16Table(168) = &H2942&
    crc16Table(169) = &H38CB&
    crc16Table(170) = &HA50&
    crc16Table(171) = &H1BD9&
    crc16Table(172) = &H6F66&
    crc16Table(173) = &H7EEF&
    crc16Table(174) = &H4C74&
    crc16Table(175) = &H5DFD&
    crc16Table(176) = &HB58B&
    crc16Table(177) = &HA402&
    crc16Table(178) = &H9699&
    crc16Table(179) = &H8710&
    crc16Table(180) = &HF3AF&
    crc16Table(181) = &HE226&
    crc16Table(182) = &HD0BD&
    crc16Table(183) = &HC134&
    crc16Table(184) = &H39C3&
    crc16Table(185) = &H284A&
    crc16Table(186) = &H1AD1&
    crc16Table(187) = &HB58&
    crc16Table(188) = &H7FE7&
    crc16Table(189) = &H6E6E&
    crc16Table(190) = &H5CF5&
    crc16Table(191) = &H4D7C&
    crc16Table(192) = &HC60C&
    crc16Table(193) = &HD785&
    crc16Table(194) = &HE51E&
    crc16Table(195) = &HF497&
    crc16Table(196) = &H8028&
    crc16Table(197) = &H91A1&
    crc16Table(198) = &HA33A&
    crc16Table(199) = &HB2B3&
    crc16Table(200) = &H4A44&
    crc16Table(201) = &H5BCD&
    crc16Table(202) = &H6956&
    crc16Table(203) = &H78DF&
    crc16Table(204) = &HC60&
    crc16Table(205) = &H1DE9&
    crc16Table(206) = &H2F72&
    crc16Table(207) = &H3EFB&
    crc16Table(208) = &HD68D&
    crc16Table(209) = &HC704&
    crc16Table(210) = &HF59F&
    crc16Table(211) = &HE416&
    crc16Table(212) = &H90A9&
    crc16Table(213) = &H8120&
    crc16Table(214) = &HB3BB&
    crc16Table(215) = &HA232&
    crc16Table(216) = &H5AC5&
    crc16Table(217) = &H4B4C&
    crc16Table(218) = &H79D7&
    crc16Table(219) = &H685E&
    crc16Table(220) = &H1CE1&
    crc16Table(221) = &HD68&
    crc16Table(222) = &H3FF3&
    crc16Table(223) = &H2E7A&
    crc16Table(224) = &HE70E&
    crc16Table(225) = &HF687&
    crc16Table(226) = &HC41C&
    crc16Table(227) = &HD595&
    crc16Table(228) = &HA12A&
    crc16Table(229) = &HB0A3&
    crc16Table(230) = &H8238&
    crc16Table(231) = &H93B1&
    crc16Table(232) = &H6B46&
    crc16Table(233) = &H7ACF&
    crc16Table(234) = &H4854&
    crc16Table(235) = &H59DD&
    crc16Table(236) = &H2D62&
    crc16Table(237) = &H3CEB&
    crc16Table(238) = &HE70&
    crc16Table(239) = &H1FF9&
    crc16Table(240) = &HF78F&
    crc16Table(241) = &HE606&
    crc16Table(242) = &HD49D&
    crc16Table(243) = &HC514&
    crc16Table(244) = &HB1AB&
    crc16Table(245) = &HA022&
    crc16Table(246) = &H92B9&
    crc16Table(247) = &H8330&
    crc16Table(248) = &H7BC7&
    crc16Table(249) = &H6A4E&
    crc16Table(250) = &H58D5&
    crc16Table(251) = &H495C&
    crc16Table(252) = &H3DE3&
    crc16Table(253) = &H2C6A&
    crc16Table(254) = &H1EF1&
    crc16Table(255) = &HF78&
  
End Sub

