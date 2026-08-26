Attribute VB_Name = "modOracle"
Public bytPacket(300) As Byte ' A bit of a kludge but HEY!
Public intPacketLength As Integer

Public Sub DoOracleCommand(ByVal strCommand As String)
  Dim bytPacketLengthHigh As Byte
  Dim bytPacketLengthLow As Byte
  Dim intCommandLength As Integer
  Dim bytCommandLengthHigh As Byte
  Dim bytCommandLengthLow As Byte
  Dim intIdx As Integer
  
  intPacketLength = 58 + Len(strCommand)
  bytPacketLengthHigh = (intPacketLength \ &H100)
  bytPacketLengthLow = (intPacketLength And &HFF)
  intCommandLength = Len(strCommand)
  bytCommandLengthHigh = (intCommandLength \ &H100)
  bytCommandLengthLow = (intCommandLength And &HFF)
  
  ' Build the packet
  bytPacket(0) = bytPacketLengthHigh ' Plen Hi
  bytPacket(1) = bytPacketLengthLow ' Plen Lo
  bytPacket(2) = &H0
  bytPacket(3) = &H0
  bytPacket(4) = &H1
  bytPacket(5) = &H0
  bytPacket(6) = &H0
  bytPacket(7) = &H0
  
  bytPacket(8) = &H1
  bytPacket(9) = &H36
  bytPacket(10) = &H1
  bytPacket(11) = &H2C
  bytPacket(12) = &H0
  bytPacket(13) = &H0
  bytPacket(14) = &H8
  bytPacket(15) = &H0
  
  bytPacket(16) = &H7F
  bytPacket(17) = &HFF
  bytPacket(18) = &H7F
  bytPacket(19) = &H8
  bytPacket(20) = &H0
  bytPacket(21) = &H0
  bytPacket(22) = &H0
  bytPacket(23) = &H1
  
  bytPacket(24) = bytCommandLengthHigh ' CLen Hi
  bytPacket(25) = bytCommandLengthLow ' CLen Lo
  bytPacket(26) = &H0
  bytPacket(27) = &H3A
  bytPacket(28) = &H0
  bytPacket(29) = &H0
  bytPacket(30) = &H0
  bytPacket(31) = &H0

  bytPacket(32) = &H0
  bytPacket(33) = &H0
  bytPacket(34) = &H0
  bytPacket(35) = &H0
  bytPacket(36) = &H0
  bytPacket(37) = &H0
  bytPacket(38) = &H0
  bytPacket(39) = &H0

  bytPacket(40) = &H0
  bytPacket(41) = &H0
  bytPacket(42) = &H0
  bytPacket(43) = &H0
  bytPacket(44) = &H34
  bytPacket(45) = &HE6
  bytPacket(46) = &H0
  bytPacket(47) = &H0

  bytPacket(48) = &H0
  bytPacket(49) = &H1
  bytPacket(50) = &H0
  bytPacket(51) = &H0
  bytPacket(52) = &H0
  bytPacket(53) = &H0
  bytPacket(54) = &H0
  bytPacket(55) = &H0
  
  bytPacket(56) = &H0
  bytPacket(57) = &H0
  
  For intIdx = 1 To Len(strCommand)
    bytPacket(57 + intIdx) = Asc(Mid(strCommand, intIdx, 1))
  Next
  
  If (frmMain.tcpMain.State <> sckClosed) Then frmMain.tcpMain.Close
  frmMain.tcpMain.Connect "89.7.3.11", 50101
End Sub

