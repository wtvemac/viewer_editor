Attribute VB_Name = "eMacBitWorkarounds"
'''''''''''''''''''''''''''''''''
' WebTV IPE (In-place Edit) 4.0 '
'                               '
' By: Eric MacDonald            '
' Date: April 24, 2005          '
'                               '
' This is a patcher tool        '
' for any SuperViewer template  '
'''''''''''''''''''''''''''''''''

Option Explicit

Private OnBits(0 To 15) As Integer
Dim CRCTable(256) As Long

Public Function ComputeFCS(serial_number As String)
    Dim theFCS As Long
    Dim theFCSA As Long
    Dim snLen As Integer
    Dim i As Integer
    Dim CRCIndex As Integer

  If CRCTable(1) = 0 Then
    CRCTable(0) = &H0
    CRCTable(1) = &H1189
    CRCTable(2) = &H2312
    CRCTable(3) = &H329B
    CRCTable(4) = &H4624
    CRCTable(5) = &H57AD
    CRCTable(6) = &H6536
    CRCTable(7) = &H74BF
    CRCTable(8) = &H8C48
    CRCTable(9) = &H9DC1
    CRCTable(10) = &HAF5A
    CRCTable(11) = &HBED3
    CRCTable(12) = &HCA6C
    CRCTable(13) = &HDBE5
    CRCTable(14) = &HE97E
    CRCTable(15) = &HF8F7
    CRCTable(16) = &H1081
    CRCTable(17) = &H108
    CRCTable(18) = &H3393
    CRCTable(19) = &H221A
    CRCTable(20) = &H56A5
    CRCTable(21) = &H472C
    CRCTable(22) = &H75B7
    CRCTable(23) = &H643E
    CRCTable(24) = &H9CC9
    CRCTable(25) = &H8D40
    CRCTable(26) = &HBFDB
    CRCTable(27) = &HAE52
    CRCTable(28) = &HDAED
    CRCTable(29) = &HCB64
    CRCTable(30) = &HF9FF
    CRCTable(31) = &HE876
    CRCTable(32) = &H2102
    CRCTable(33) = &H308B
    CRCTable(34) = &H210
    CRCTable(35) = &H1399
    CRCTable(36) = &H6726
    CRCTable(37) = &H76AF
    CRCTable(38) = &H4434
    CRCTable(39) = &H55BD
    CRCTable(40) = &HAD4A
    CRCTable(41) = &HBCC3
    CRCTable(42) = &H8E58
    CRCTable(43) = &H9FD1
    CRCTable(44) = &HEB6E
    CRCTable(45) = &HFAE7
    CRCTable(46) = &HC87C
    CRCTable(47) = &HD9F5
    CRCTable(48) = &H3E04
    CRCTable(49) = &H2F8D
    CRCTable(50) = &H1291
    CRCTable(51) = &H318
    CRCTable(52) = &H77A7
    CRCTable(53) = &H662E
    CRCTable(54) = &H54B5
    CRCTable(55) = &H4ABB
    CRCTable(56) = &HBDCB
    CRCTable(57) = &HA3C5
    CRCTable(58) = &H9ED9
    CRCTable(59) = &H8F50
    CRCTable(60) = &HFBEF
    CRCTable(61) = &HEA66
    CRCTable(62) = &HD8FD
    CRCTable(63) = &HC974
    CRCTable(64) = &H4204
    CRCTable(65) = &H538D
    CRCTable(66) = &H6116
    CRCTable(67) = &H709F
    CRCTable(68) = &H420
    CRCTable(69) = &H15A9
    CRCTable(70) = &H2732
    CRCTable(71) = &H36BB
    CRCTable(72) = &HCE4C
    CRCTable(73) = &HDFC5
    CRCTable(74) = &HED5E
    CRCTable(75) = &HFCD7
    CRCTable(76) = &H8868
    CRCTable(77) = &H99E1
    CRCTable(78) = &HAB7A
    CRCTable(79) = &HBAF3
    CRCTable(80) = &H5285
    CRCTable(81) = &H430C
    CRCTable(82) = &H7197
    CRCTable(83) = &H601E
    CRCTable(84) = &H14A1
    CRCTable(85) = &H528
    CRCTable(86) = &H37B3
    CRCTable(87) = &H263A
    CRCTable(88) = &HDECD
    CRCTable(89) = &HCF44
    CRCTable(90) = &HFDDF
    CRCTable(91) = &HEC56
    CRCTable(92) = &H98E9
    CRCTable(93) = &H8960
    CRCTable(94) = &HBBFB
    CRCTable(95) = &HAA72
    CRCTable(96) = &H6306
    CRCTable(97) = &H728F
    CRCTable(98) = &H4014
    CRCTable(99) = &H519D
    CRCTable(100) = &H2522
    CRCTable(101) = &H34AB
    CRCTable(102) = &H630
    CRCTable(103) = &H17B9
    CRCTable(104) = &HEF4E
    CRCTable(105) = &HFEC7
    CRCTable(106) = &HCC5C
    CRCTable(107) = &HDDD5
    CRCTable(108) = &HA96A
    CRCTable(109) = &HB8E3
    CRCTable(110) = &H8A78
    CRCTable(111) = &H9BF1
    CRCTable(112) = &H7387
    CRCTable(113) = &H620E
    CRCTable(114) = &H5095
    CRCTable(115) = &H411C
    CRCTable(116) = &H35A3
    CRCTable(117) = &H242A
    CRCTable(118) = &H16B1
    CRCTable(119) = &H738
    CRCTable(120) = &HFFCF
    CRCTable(121) = &HEE46
    CRCTable(122) = &HDCDD
    CRCTable(123) = &HCD54
    CRCTable(124) = &HB9EB
    CRCTable(125) = &HA862
    CRCTable(126) = &H9AF9
    CRCTable(127) = &H8B70
    CRCTable(128) = &H8408
    CRCTable(129) = &H9581
    CRCTable(130) = &HA71A
    CRCTable(131) = &HB693
    CRCTable(132) = &HC22C
    CRCTable(133) = &HD3A5
    CRCTable(134) = &HE13E
    CRCTable(135) = &HF0B7
    CRCTable(136) = &H840
    CRCTable(137) = &H19C9
    CRCTable(138) = &H2B52
    CRCTable(139) = &H3ADB
    CRCTable(140) = &H4E64
    CRCTable(141) = &H5FED
    CRCTable(142) = &H6D76
    CRCTable(143) = &H7CFF
    CRCTable(144) = &H9489
    CRCTable(145) = &H8500
    CRCTable(146) = &HB79B
    CRCTable(147) = &HA612
    CRCTable(148) = &HD2AD
    CRCTable(149) = &HC324
    CRCTable(150) = &HF1BF
    CRCTable(151) = &HE036
    CRCTable(152) = &H18C1
    CRCTable(153) = &H948
    CRCTable(154) = &H3BD3
    CRCTable(155) = &H2A5A
    CRCTable(156) = &H5EE5
    CRCTable(157) = &H4F6C
    CRCTable(158) = &H7DF7
    CRCTable(159) = &H6C7E
    CRCTable(160) = &HA50A
    CRCTable(161) = &HB483
    CRCTable(162) = &H8618
    CRCTable(163) = &H9791
    CRCTable(164) = &HE32E
    CRCTable(165) = &HF2A7
    CRCTable(166) = &HC03C
    CRCTable(167) = &HD1B5
    CRCTable(168) = &H2942
    CRCTable(169) = &H38CB
    CRCTable(170) = &HA50
    CRCTable(171) = &H1BD9
    CRCTable(172) = &H6F66
    CRCTable(173) = &H7EEF
    CRCTable(174) = &H4C74
    CRCTable(175) = &H5DFD
    CRCTable(176) = &HB58B
    CRCTable(177) = &HA402
    CRCTable(178) = &H9699
    CRCTable(179) = &H8710
    CRCTable(180) = &HF3AF
    CRCTable(181) = &HE226
    CRCTable(182) = &HD0BD
    CRCTable(183) = &HC134
    CRCTable(184) = &H39C3
    CRCTable(185) = &H284A
    CRCTable(186) = &H1AD1
    CRCTable(187) = &HB58
    CRCTable(188) = &H7FE7
    CRCTable(189) = &H6E6E
    CRCTable(190) = &H5CF5
    CRCTable(191) = &H4D7C
    CRCTable(192) = &HC60C
    CRCTable(193) = &HD785
    CRCTable(194) = &HE51E
    CRCTable(195) = &HF497
    CRCTable(196) = &H8028
    CRCTable(197) = &H91A1
    CRCTable(198) = &HA33A
    CRCTable(199) = &HB2B3
    CRCTable(200) = &H4A44
    CRCTable(201) = &H5BCD
    CRCTable(202) = &H6956
    CRCTable(203) = &H78DF
    CRCTable(204) = &HC60
    CRCTable(205) = &H1DE9
    CRCTable(206) = &H2F72
    CRCTable(207) = &H3EFB
    CRCTable(208) = &HD68D
    CRCTable(209) = &HC704
    CRCTable(210) = &HF59F
    CRCTable(211) = &HE416
    CRCTable(212) = &H90A9
    CRCTable(213) = &H8120
    CRCTable(214) = &HB3BB
    CRCTable(215) = &HA232
    CRCTable(216) = &H5AC5
    CRCTable(217) = &H4B4C
    CRCTable(218) = &H79D7
    CRCTable(219) = &H685E
    CRCTable(220) = &H1CE1
    CRCTable(221) = &HD68
    CRCTable(222) = &H3FF3
    CRCTable(223) = &H2E7A
    CRCTable(224) = &HE70E
    CRCTable(225) = &HF687
    CRCTable(226) = &HC41C
    CRCTable(227) = &HD595
    CRCTable(228) = &HA12A
    CRCTable(229) = &HB0A3
    CRCTable(230) = &H8238
    CRCTable(231) = &H93B1
    CRCTable(232) = &H6B46
    CRCTable(233) = &H7ACF
    CRCTable(234) = &H4854
    CRCTable(235) = &H59DD
    CRCTable(236) = &H2D62
    CRCTable(237) = &H3CEB
    CRCTable(238) = &HE70
    CRCTable(239) = &H1FF9
    CRCTable(240) = &HF78F
    CRCTable(241) = &HE606
    CRCTable(242) = &HD49D
    CRCTable(243) = &HC514
    CRCTable(244) = &HB1AB
    CRCTable(245) = &HA022
    CRCTable(246) = &H92B9
    CRCTable(247) = &H8330
    CRCTable(248) = &H7BC7
    CRCTable(249) = &H6A4E
    CRCTable(250) = &H58D5
    CRCTable(251) = &H495C
    CRCTable(252) = &H3DE3
    CRCTable(253) = &H2C6A
    CRCTable(254) = &H1EF1
    CRCTable(255) = &HF78
  End If
    

    snLen = Len(serial_number)
    
    theFCS = 0
    
    For i = 1 To snLen
        CRCIndex = (theFCS Xor Asc(Mid(serial_number, i, 1))) And &HFF
        
        theFCSA = CRCTable(CRCIndex) And &H7FFF
        
        If theFCSA <> CRCTable(CRCIndex) Then
            theFCSA = theFCSA + 32768
        End If
        
        theFCS = theFCSA Xor (RShiftLong(theFCS, 8))
        
    Next i
    
    ComputeFCS = theFCS

End Function

Public Function LShiftLong(ByVal Value As Long, ByVal Shift As Integer) As Integer
  
    MakeOnBits
  
    If (Value And (2 ^ (15 - Shift))) Then GoTo OverFlow
  
    LShiftLong = ((Value And OnBits(31 - Shift)) * (2 ^ Shift))
  
    Exit Function

OverFlow:
  
    LShiftLong = ((Value And OnBits(15 - (Shift + 1))) * _
       (2 ^ (Shift))) Or &H8000
  
End Function

Public Function RShiftLong(ByVal Value As Long, ByVal Shift As Integer) As Integer
    Dim hi As Long
    MakeOnBits
    If (Value And &H8000) Then hi = &H4000
  
    RShiftLong = (Value And &HFFFE) \ (2 ^ Shift)
    RShiftLong = (RShiftLong Or (hi \ (2 ^ (Shift - 1))))
End Function
 


Private Sub MakeOnBits()
    Dim j As Integer, _
        v As Long
  
    For j = 0 To 14
  
        v = v + (2 ^ j)
        OnBits(j) = v
  
    Next j
  
    OnBits(j) = v + &H8000

End Sub

