VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmLeave 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "frmLeave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdModify 
      DragIcon        =   "frmLeave.frx":030A
      DragMode        =   1  'Automatic
      Height          =   615
      Left            =   14520
      Picture         =   "frmLeave.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   199
      ToolTipText     =   "Click and drag onto name to add phone number etc, Click and drag onto mailbox to add message"
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer tmrBarcodeCount 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6120
      Top             =   6000
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6600
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DTREnable       =   -1  'True
      BaudRate        =   38400
   End
   Begin VB.Timer tmrBarcode 
      Interval        =   500
      Left            =   5640
      Top             =   6000
   End
   Begin VB.CommandButton cmdPennant 
      Caption         =   "Pennant No's"
      Height          =   855
      Left            =   8760
      Picture         =   "frmLeave.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   5160
      Top             =   6000
   End
   Begin VB.CommandButton cmdID 
      Caption         =   "New ID..."
      Height          =   855
      Left            =   7680
      Picture         =   "frmLeave.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4680
      Top             =   6000
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "Wakey"
      Height          =   855
      Index           =   4
      Left            =   6600
      Picture         =   "frmLeave.frx":12DA
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "OOD"
      Height          =   855
      Index           =   3
      Left            =   5520
      Picture         =   "frmLeave.frx":171C
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "Carry On"
      Height          =   855
      Index           =   2
      Left            =   4440
      Picture         =   "frmLeave.frx":1A26
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "Still"
      Height          =   855
      Index           =   1
      Left            =   3360
      Picture         =   "frmLeave.frx":1E68
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   10560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   7011
      _Version        =   327681
      Tabs            =   9
      Tab             =   1
      TabsPerRow      =   9
      TabHeight       =   706
      BackColor       =   -2147483637
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chief Petty Officers"
      TabPicture(0)   =   "frmLeave.frx":22AA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "imgLocation(62)"
      Tab(0).Control(1)=   "lblName(62)"
      Tab(0).Control(2)=   "imgMail(62)"
      Tab(0).Control(3)=   "imgLocation(61)"
      Tab(0).Control(4)=   "lblName(61)"
      Tab(0).Control(5)=   "imgMail(61)"
      Tab(0).Control(6)=   "imgLocation(60)"
      Tab(0).Control(7)=   "lblName(60)"
      Tab(0).Control(8)=   "imgMail(60)"
      Tab(0).Control(9)=   "imgLocation(59)"
      Tab(0).Control(10)=   "lblName(59)"
      Tab(0).Control(11)=   "imgMail(59)"
      Tab(0).Control(12)=   "imgLocation(58)"
      Tab(0).Control(13)=   "lblName(58)"
      Tab(0).Control(14)=   "imgMail(58)"
      Tab(0).Control(15)=   "imgLocation(57)"
      Tab(0).Control(16)=   "lblName(57)"
      Tab(0).Control(17)=   "imgMail(57)"
      Tab(0).Control(18)=   "imgLocation(56)"
      Tab(0).Control(19)=   "lblName(56)"
      Tab(0).Control(20)=   "imgMail(56)"
      Tab(0).Control(21)=   "imgLocation(55)"
      Tab(0).Control(22)=   "lblName(55)"
      Tab(0).Control(23)=   "imgMail(55)"
      Tab(0).Control(24)=   "imgLocation(54)"
      Tab(0).Control(25)=   "lblName(54)"
      Tab(0).Control(26)=   "imgMail(54)"
      Tab(0).Control(27)=   "imgLocation(53)"
      Tab(0).Control(28)=   "lblName(53)"
      Tab(0).Control(29)=   "imgMail(53)"
      Tab(0).Control(30)=   "imgLocation(52)"
      Tab(0).Control(31)=   "lblName(52)"
      Tab(0).Control(32)=   "imgMail(52)"
      Tab(0).Control(33)=   "imgLocation(51)"
      Tab(0).Control(34)=   "lblName(51)"
      Tab(0).Control(35)=   "imgMail(51)"
      Tab(0).Control(36)=   "imgLocation(50)"
      Tab(0).Control(37)=   "lblName(50)"
      Tab(0).Control(38)=   "imgMail(50)"
      Tab(0).Control(39)=   "imgLocation(49)"
      Tab(0).Control(40)=   "lblName(49)"
      Tab(0).Control(41)=   "imgMail(49)"
      Tab(0).Control(42)=   "imgLocation(48)"
      Tab(0).Control(43)=   "lblName(48)"
      Tab(0).Control(44)=   "imgMail(48)"
      Tab(0).Control(45)=   "imgLocation(47)"
      Tab(0).Control(46)=   "lblName(47)"
      Tab(0).Control(47)=   "imgMail(47)"
      Tab(0).Control(48)=   "imgLocation(46)"
      Tab(0).Control(49)=   "lblName(46)"
      Tab(0).Control(50)=   "imgMail(46)"
      Tab(0).Control(51)=   "imgLocation(45)"
      Tab(0).Control(52)=   "lblName(45)"
      Tab(0).Control(53)=   "imgMail(45)"
      Tab(0).Control(54)=   "imgLocation(44)"
      Tab(0).Control(55)=   "lblName(44)"
      Tab(0).Control(56)=   "imgMail(44)"
      Tab(0).Control(57)=   "imgLocation(43)"
      Tab(0).Control(58)=   "lblName(43)"
      Tab(0).Control(59)=   "imgMail(43)"
      Tab(0).Control(60)=   "imgLocation(42)"
      Tab(0).Control(61)=   "lblName(42)"
      Tab(0).Control(62)=   "imgMail(42)"
      Tab(0).Control(63)=   "imgLocation(41)"
      Tab(0).Control(64)=   "lblName(41)"
      Tab(0).Control(65)=   "imgMail(41)"
      Tab(0).Control(66)=   "imgLocation(40)"
      Tab(0).Control(67)=   "lblName(40)"
      Tab(0).Control(68)=   "imgMail(40)"
      Tab(0).Control(69)=   "imgLocation(39)"
      Tab(0).Control(70)=   "lblName(39)"
      Tab(0).Control(71)=   "imgMail(39)"
      Tab(0).Control(72)=   "imgLocation(38)"
      Tab(0).Control(73)=   "lblName(38)"
      Tab(0).Control(74)=   "imgMail(38)"
      Tab(0).Control(75)=   "imgLocation(37)"
      Tab(0).Control(76)=   "lblName(37)"
      Tab(0).Control(77)=   "imgMail(37)"
      Tab(0).Control(78)=   "imgLocation(36)"
      Tab(0).Control(79)=   "lblName(36)"
      Tab(0).Control(80)=   "imgMail(36)"
      Tab(0).Control(81)=   "imgLocation(35)"
      Tab(0).Control(82)=   "lblName(35)"
      Tab(0).Control(83)=   "imgMail(35)"
      Tab(0).ControlCount=   84
      TabCaption(1)   =   "Petty Officers"
      TabPicture(1)   =   "frmLeave.frx":22C6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "imgLocation(63)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblName(63)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "imgMail(63)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "imgLocation(64)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblName(64)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "imgMail(64)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "imgLocation(65)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblName(65)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "imgMail(65)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "imgLocation(66)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblName(66)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "imgMail(66)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "imgLocation(67)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblName(67)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "imgMail(67)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "imgLocation(68)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblName(68)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "imgMail(68)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "imgLocation(69)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lblName(69)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "imgMail(69)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "imgLocation(70)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lblName(70)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "imgMail(70)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "imgLocation(71)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "lblName(71)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "imgMail(71)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "imgLocation(72)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "lblName(72)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "imgMail(72)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "imgLocation(73)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "lblName(73)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "imgMail(73)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "imgLocation(74)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lblName(74)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "imgMail(74)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "imgLocation(75)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lblName(75)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "imgMail(75)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "imgLocation(76)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "lblName(76)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "imgMail(76)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "imgLocation(77)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "lblName(77)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "imgMail(77)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "imgLocation(78)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "lblName(78)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "imgMail(78)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "imgLocation(79)"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "lblName(79)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "imgMail(79)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "imgLocation(80)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lblName(80)"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "imgMail(80)"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "imgLocation(81)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lblName(81)"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "imgMail(81)"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "imgLocation(82)"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "lblName(82)"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "imgMail(82)"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "imgLocation(83)"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "lblName(83)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "imgMail(83)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "imgLocation(84)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "lblName(84)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "imgMail(84)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "imgLocation(85)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "lblName(85)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "imgMail(85)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "imgLocation(86)"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "lblName(86)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "imgMail(86)"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "imgLocation(87)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "lblName(87)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "imgMail(87)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "imgLocation(88)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "lblName(88)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "imgMail(88)"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "imgLocation(89)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "lblName(89)"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "imgMail(89)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "imgLocation(90)"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "lblName(90)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "imgMail(90)"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).ControlCount=   84
      TabCaption(2)   =   "Aviation"
      TabPicture(2)   =   "frmLeave.frx":22E2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "imgMail(118)"
      Tab(2).Control(1)=   "lblName(118)"
      Tab(2).Control(2)=   "imgLocation(118)"
      Tab(2).Control(3)=   "imgMail(117)"
      Tab(2).Control(4)=   "lblName(117)"
      Tab(2).Control(5)=   "imgLocation(117)"
      Tab(2).Control(6)=   "imgMail(116)"
      Tab(2).Control(7)=   "lblName(116)"
      Tab(2).Control(8)=   "imgLocation(116)"
      Tab(2).Control(9)=   "imgMail(115)"
      Tab(2).Control(10)=   "lblName(115)"
      Tab(2).Control(11)=   "imgLocation(115)"
      Tab(2).Control(12)=   "imgMail(114)"
      Tab(2).Control(13)=   "lblName(114)"
      Tab(2).Control(14)=   "imgLocation(114)"
      Tab(2).Control(15)=   "imgMail(113)"
      Tab(2).Control(16)=   "lblName(113)"
      Tab(2).Control(17)=   "imgLocation(113)"
      Tab(2).Control(18)=   "imgMail(112)"
      Tab(2).Control(19)=   "lblName(112)"
      Tab(2).Control(20)=   "imgLocation(112)"
      Tab(2).Control(21)=   "imgMail(111)"
      Tab(2).Control(22)=   "lblName(111)"
      Tab(2).Control(23)=   "imgLocation(111)"
      Tab(2).Control(24)=   "imgMail(110)"
      Tab(2).Control(25)=   "lblName(110)"
      Tab(2).Control(26)=   "imgLocation(110)"
      Tab(2).Control(27)=   "imgMail(109)"
      Tab(2).Control(28)=   "lblName(109)"
      Tab(2).Control(29)=   "imgLocation(109)"
      Tab(2).Control(30)=   "imgMail(108)"
      Tab(2).Control(31)=   "lblName(108)"
      Tab(2).Control(32)=   "imgLocation(108)"
      Tab(2).Control(33)=   "imgMail(107)"
      Tab(2).Control(34)=   "lblName(107)"
      Tab(2).Control(35)=   "imgLocation(107)"
      Tab(2).Control(36)=   "imgMail(106)"
      Tab(2).Control(37)=   "lblName(106)"
      Tab(2).Control(38)=   "imgLocation(106)"
      Tab(2).Control(39)=   "imgMail(105)"
      Tab(2).Control(40)=   "lblName(105)"
      Tab(2).Control(41)=   "imgLocation(105)"
      Tab(2).Control(42)=   "imgMail(104)"
      Tab(2).Control(43)=   "lblName(104)"
      Tab(2).Control(44)=   "imgLocation(104)"
      Tab(2).Control(45)=   "imgMail(103)"
      Tab(2).Control(46)=   "lblName(103)"
      Tab(2).Control(47)=   "imgLocation(103)"
      Tab(2).Control(48)=   "imgMail(102)"
      Tab(2).Control(49)=   "lblName(102)"
      Tab(2).Control(50)=   "imgLocation(102)"
      Tab(2).Control(51)=   "imgMail(101)"
      Tab(2).Control(52)=   "lblName(101)"
      Tab(2).Control(53)=   "imgLocation(101)"
      Tab(2).Control(54)=   "imgMail(100)"
      Tab(2).Control(55)=   "lblName(100)"
      Tab(2).Control(56)=   "imgLocation(100)"
      Tab(2).Control(57)=   "imgMail(99)"
      Tab(2).Control(58)=   "lblName(99)"
      Tab(2).Control(59)=   "imgLocation(99)"
      Tab(2).Control(60)=   "imgMail(98)"
      Tab(2).Control(61)=   "lblName(98)"
      Tab(2).Control(62)=   "imgLocation(98)"
      Tab(2).Control(63)=   "imgMail(97)"
      Tab(2).Control(64)=   "lblName(97)"
      Tab(2).Control(65)=   "imgLocation(97)"
      Tab(2).Control(66)=   "imgMail(96)"
      Tab(2).Control(67)=   "lblName(96)"
      Tab(2).Control(68)=   "imgLocation(96)"
      Tab(2).Control(69)=   "imgMail(95)"
      Tab(2).Control(70)=   "lblName(95)"
      Tab(2).Control(71)=   "imgLocation(95)"
      Tab(2).Control(72)=   "imgMail(94)"
      Tab(2).Control(73)=   "lblName(94)"
      Tab(2).Control(74)=   "imgLocation(94)"
      Tab(2).Control(75)=   "imgMail(93)"
      Tab(2).Control(76)=   "lblName(93)"
      Tab(2).Control(77)=   "imgLocation(93)"
      Tab(2).Control(78)=   "imgMail(92)"
      Tab(2).Control(79)=   "lblName(92)"
      Tab(2).Control(80)=   "imgLocation(92)"
      Tab(2).Control(81)=   "imgMail(91)"
      Tab(2).Control(82)=   "lblName(91)"
      Tab(2).Control(83)=   "imgLocation(91)"
      Tab(2).ControlCount=   84
      TabCaption(3)   =   "Marine Engineering"
      TabPicture(3)   =   "frmLeave.frx":22FE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "imgLocation(146)"
      Tab(3).Control(1)=   "lblName(146)"
      Tab(3).Control(2)=   "imgMail(146)"
      Tab(3).Control(3)=   "imgLocation(145)"
      Tab(3).Control(4)=   "lblName(145)"
      Tab(3).Control(5)=   "imgMail(145)"
      Tab(3).Control(6)=   "imgLocation(144)"
      Tab(3).Control(7)=   "lblName(144)"
      Tab(3).Control(8)=   "imgMail(144)"
      Tab(3).Control(9)=   "imgLocation(143)"
      Tab(3).Control(10)=   "lblName(143)"
      Tab(3).Control(11)=   "imgMail(143)"
      Tab(3).Control(12)=   "imgLocation(142)"
      Tab(3).Control(13)=   "lblName(142)"
      Tab(3).Control(14)=   "imgMail(142)"
      Tab(3).Control(15)=   "imgLocation(141)"
      Tab(3).Control(16)=   "lblName(141)"
      Tab(3).Control(17)=   "imgMail(141)"
      Tab(3).Control(18)=   "imgLocation(140)"
      Tab(3).Control(19)=   "lblName(140)"
      Tab(3).Control(20)=   "imgMail(140)"
      Tab(3).Control(21)=   "imgLocation(139)"
      Tab(3).Control(22)=   "lblName(139)"
      Tab(3).Control(23)=   "imgMail(139)"
      Tab(3).Control(24)=   "imgLocation(138)"
      Tab(3).Control(25)=   "lblName(138)"
      Tab(3).Control(26)=   "imgMail(138)"
      Tab(3).Control(27)=   "imgLocation(137)"
      Tab(3).Control(28)=   "lblName(137)"
      Tab(3).Control(29)=   "imgMail(137)"
      Tab(3).Control(30)=   "imgLocation(136)"
      Tab(3).Control(31)=   "lblName(136)"
      Tab(3).Control(32)=   "imgMail(136)"
      Tab(3).Control(33)=   "imgLocation(135)"
      Tab(3).Control(34)=   "lblName(135)"
      Tab(3).Control(35)=   "imgMail(135)"
      Tab(3).Control(36)=   "imgLocation(134)"
      Tab(3).Control(37)=   "lblName(134)"
      Tab(3).Control(38)=   "imgMail(134)"
      Tab(3).Control(39)=   "imgLocation(133)"
      Tab(3).Control(40)=   "lblName(133)"
      Tab(3).Control(41)=   "imgMail(133)"
      Tab(3).Control(42)=   "imgLocation(132)"
      Tab(3).Control(43)=   "lblName(132)"
      Tab(3).Control(44)=   "imgMail(132)"
      Tab(3).Control(45)=   "imgLocation(131)"
      Tab(3).Control(46)=   "lblName(131)"
      Tab(3).Control(47)=   "imgMail(131)"
      Tab(3).Control(48)=   "imgLocation(130)"
      Tab(3).Control(49)=   "lblName(130)"
      Tab(3).Control(50)=   "imgMail(130)"
      Tab(3).Control(51)=   "imgLocation(129)"
      Tab(3).Control(52)=   "lblName(129)"
      Tab(3).Control(53)=   "imgMail(129)"
      Tab(3).Control(54)=   "imgLocation(128)"
      Tab(3).Control(55)=   "lblName(128)"
      Tab(3).Control(56)=   "imgMail(128)"
      Tab(3).Control(57)=   "imgLocation(127)"
      Tab(3).Control(58)=   "lblName(127)"
      Tab(3).Control(59)=   "imgMail(127)"
      Tab(3).Control(60)=   "imgLocation(126)"
      Tab(3).Control(61)=   "lblName(126)"
      Tab(3).Control(62)=   "imgMail(126)"
      Tab(3).Control(63)=   "imgLocation(125)"
      Tab(3).Control(64)=   "lblName(125)"
      Tab(3).Control(65)=   "imgMail(125)"
      Tab(3).Control(66)=   "imgLocation(124)"
      Tab(3).Control(67)=   "lblName(124)"
      Tab(3).Control(68)=   "imgMail(124)"
      Tab(3).Control(69)=   "imgLocation(123)"
      Tab(3).Control(70)=   "lblName(123)"
      Tab(3).Control(71)=   "imgMail(123)"
      Tab(3).Control(72)=   "imgLocation(122)"
      Tab(3).Control(73)=   "lblName(122)"
      Tab(3).Control(74)=   "imgMail(122)"
      Tab(3).Control(75)=   "imgLocation(121)"
      Tab(3).Control(76)=   "lblName(121)"
      Tab(3).Control(77)=   "imgMail(121)"
      Tab(3).Control(78)=   "imgLocation(120)"
      Tab(3).Control(79)=   "lblName(120)"
      Tab(3).Control(80)=   "imgMail(120)"
      Tab(3).Control(81)=   "imgLocation(119)"
      Tab(3).Control(82)=   "lblName(119)"
      Tab(3).Control(83)=   "imgMail(119)"
      Tab(3).ControlCount=   84
      TabCaption(4)   =   "Operations"
      TabPicture(4)   =   "frmLeave.frx":231A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Support"
      TabPicture(5)   =   "frmLeave.frx":2336
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Weapons Electrical"
      TabPicture(6)   =   "frmLeave.frx":2352
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Escorted Visitors"
      TabPicture(7)   =   "frmLeave.frx":236E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Unescorted Visitors"
      TabPicture(8)   =   "frmLeave.frx":238A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   146
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   146
         Left            =   -63480
         TabIndex        =   198
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   146
         Left            =   -60360
         Picture         =   "frmLeave.frx":23A6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   145
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   145
         Left            =   -63480
         TabIndex        =   197
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   145
         Left            =   -60360
         Picture         =   "frmLeave.frx":27E8
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   144
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   144
         Left            =   -63480
         TabIndex        =   196
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   144
         Left            =   -60360
         Picture         =   "frmLeave.frx":2C2A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   143
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   143
         Left            =   -63480
         TabIndex        =   195
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   143
         Left            =   -60360
         Picture         =   "frmLeave.frx":306C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   142
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   142
         Left            =   -63480
         TabIndex        =   194
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   142
         Left            =   -60360
         Picture         =   "frmLeave.frx":34AE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   141
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   141
         Left            =   -63480
         TabIndex        =   193
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   141
         Left            =   -60360
         Picture         =   "frmLeave.frx":38F0
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   140
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   140
         Left            =   -63480
         TabIndex        =   192
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   140
         Left            =   -60360
         Picture         =   "frmLeave.frx":3D32
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   139
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   139
         Left            =   -67200
         TabIndex        =   191
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   139
         Left            =   -64080
         Picture         =   "frmLeave.frx":4174
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   138
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   138
         Left            =   -67200
         TabIndex        =   190
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   138
         Left            =   -64080
         Picture         =   "frmLeave.frx":45B6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   137
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   137
         Left            =   -67200
         TabIndex        =   189
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   137
         Left            =   -64080
         Picture         =   "frmLeave.frx":49F8
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   136
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   136
         Left            =   -67200
         TabIndex        =   188
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   136
         Left            =   -64080
         Picture         =   "frmLeave.frx":4E3A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   135
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   135
         Left            =   -67200
         TabIndex        =   187
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   135
         Left            =   -64080
         Picture         =   "frmLeave.frx":527C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   134
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   134
         Left            =   -67200
         TabIndex        =   186
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   134
         Left            =   -64080
         Picture         =   "frmLeave.frx":56BE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   133
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   133
         Left            =   -67200
         TabIndex        =   185
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   133
         Left            =   -64080
         Picture         =   "frmLeave.frx":5B00
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   132
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   132
         Left            =   -71040
         TabIndex        =   184
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   132
         Left            =   -67920
         Picture         =   "frmLeave.frx":5F42
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   131
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   131
         Left            =   -71040
         TabIndex        =   183
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   131
         Left            =   -67920
         Picture         =   "frmLeave.frx":6384
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   130
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   130
         Left            =   -71040
         TabIndex        =   182
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   130
         Left            =   -67920
         Picture         =   "frmLeave.frx":67C6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   129
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   129
         Left            =   -71040
         TabIndex        =   181
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   129
         Left            =   -67920
         Picture         =   "frmLeave.frx":6C08
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   128
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   128
         Left            =   -71040
         TabIndex        =   180
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   128
         Left            =   -67920
         Picture         =   "frmLeave.frx":704A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   127
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   127
         Left            =   -71040
         TabIndex        =   179
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   127
         Left            =   -67920
         Picture         =   "frmLeave.frx":748C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   126
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   126
         Left            =   -71040
         TabIndex        =   178
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   126
         Left            =   -67920
         Picture         =   "frmLeave.frx":78CE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   125
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   125
         Left            =   -74880
         TabIndex        =   177
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   125
         Left            =   -71760
         Picture         =   "frmLeave.frx":7D10
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   124
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   124
         Left            =   -74880
         TabIndex        =   176
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   124
         Left            =   -71760
         Picture         =   "frmLeave.frx":8152
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   123
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   123
         Left            =   -74880
         TabIndex        =   175
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   123
         Left            =   -71760
         Picture         =   "frmLeave.frx":8594
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   122
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   122
         Left            =   -74880
         TabIndex        =   174
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   122
         Left            =   -71760
         Picture         =   "frmLeave.frx":89D6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   121
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   121
         Left            =   -74880
         TabIndex        =   173
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   121
         Left            =   -71760
         Picture         =   "frmLeave.frx":8E18
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   120
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   120
         Left            =   -74880
         TabIndex        =   172
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   120
         Left            =   -71760
         Picture         =   "frmLeave.frx":925A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   119
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   119
         Left            =   -74880
         TabIndex        =   171
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   119
         Left            =   -71760
         Picture         =   "frmLeave.frx":969C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   118
         Left            =   -60360
         Picture         =   "frmLeave.frx":9ADE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   118
         Left            =   -63480
         TabIndex        =   159
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   118
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   117
         Left            =   -60360
         Picture         =   "frmLeave.frx":9F20
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   117
         Left            =   -63480
         TabIndex        =   158
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   117
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   116
         Left            =   -60360
         Picture         =   "frmLeave.frx":A362
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   116
         Left            =   -63480
         TabIndex        =   157
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   116
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   115
         Left            =   -60360
         Picture         =   "frmLeave.frx":A7A4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   115
         Left            =   -63480
         TabIndex        =   156
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   115
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   114
         Left            =   -60360
         Picture         =   "frmLeave.frx":ABE6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   114
         Left            =   -63480
         TabIndex        =   155
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   114
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   113
         Left            =   -60360
         Picture         =   "frmLeave.frx":B028
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   113
         Left            =   -63480
         TabIndex        =   154
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   113
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   112
         Left            =   -60360
         Picture         =   "frmLeave.frx":B46A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   112
         Left            =   -63480
         TabIndex        =   153
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   112
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   111
         Left            =   -64080
         Picture         =   "frmLeave.frx":B8AC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   111
         Left            =   -67200
         TabIndex        =   152
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   111
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   110
         Left            =   -64080
         Picture         =   "frmLeave.frx":BCEE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   110
         Left            =   -67200
         TabIndex        =   151
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   110
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   109
         Left            =   -64080
         Picture         =   "frmLeave.frx":C130
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   109
         Left            =   -67200
         TabIndex        =   150
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   109
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   108
         Left            =   -64080
         Picture         =   "frmLeave.frx":C572
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   108
         Left            =   -67200
         TabIndex        =   149
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   108
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   107
         Left            =   -64080
         Picture         =   "frmLeave.frx":C9B4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   107
         Left            =   -67200
         TabIndex        =   148
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   107
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   106
         Left            =   -64080
         Picture         =   "frmLeave.frx":CDF6
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   106
         Left            =   -67200
         TabIndex        =   147
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   106
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   105
         Left            =   -64080
         Picture         =   "frmLeave.frx":D238
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   105
         Left            =   -67200
         TabIndex        =   146
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   105
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   104
         Left            =   -67920
         Picture         =   "frmLeave.frx":D67A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   104
         Left            =   -71040
         TabIndex        =   145
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   104
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   103
         Left            =   -67920
         Picture         =   "frmLeave.frx":DABC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   103
         Left            =   -71040
         TabIndex        =   144
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   103
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   102
         Left            =   -67920
         Picture         =   "frmLeave.frx":DEFE
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   102
         Left            =   -71040
         TabIndex        =   143
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   102
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   101
         Left            =   -67920
         Picture         =   "frmLeave.frx":E340
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   101
         Left            =   -71040
         TabIndex        =   142
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   101
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   100
         Left            =   -67920
         Picture         =   "frmLeave.frx":E782
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   100
         Left            =   -71040
         TabIndex        =   141
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   100
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   99
         Left            =   -67920
         Picture         =   "frmLeave.frx":EBC4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   99
         Left            =   -71040
         TabIndex        =   140
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   99
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   98
         Left            =   -67920
         Picture         =   "frmLeave.frx":F006
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   98
         Left            =   -71040
         TabIndex        =   139
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   98
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   97
         Left            =   -71760
         Picture         =   "frmLeave.frx":F448
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   97
         Left            =   -74880
         TabIndex        =   138
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   97
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   96
         Left            =   -71760
         Picture         =   "frmLeave.frx":F88A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   96
         Left            =   -74880
         TabIndex        =   137
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   96
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   95
         Left            =   -71760
         Picture         =   "frmLeave.frx":FCCC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   95
         Left            =   -74880
         TabIndex        =   136
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   95
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   94
         Left            =   -71760
         Picture         =   "frmLeave.frx":1010E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   94
         Left            =   -74880
         TabIndex        =   135
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   94
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   93
         Left            =   -71760
         Picture         =   "frmLeave.frx":10550
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   93
         Left            =   -74880
         TabIndex        =   134
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   93
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   92
         Left            =   -71760
         Picture         =   "frmLeave.frx":10992
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   92
         Left            =   -74880
         TabIndex        =   133
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   92
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   91
         Left            =   -71760
         Picture         =   "frmLeave.frx":10DD4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   91
         Left            =   -74880
         TabIndex        =   132
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   91
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   90
         Left            =   14640
         Picture         =   "frmLeave.frx":11216
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   90
         Left            =   11520
         TabIndex        =   131
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   90
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   89
         Left            =   14640
         Picture         =   "frmLeave.frx":11658
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   89
         Left            =   11520
         TabIndex        =   130
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   89
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   88
         Left            =   14640
         Picture         =   "frmLeave.frx":11A9A
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   88
         Left            =   11520
         TabIndex        =   129
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   88
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   87
         Left            =   14640
         Picture         =   "frmLeave.frx":11EDC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   87
         Left            =   11520
         TabIndex        =   128
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   87
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   86
         Left            =   14640
         Picture         =   "frmLeave.frx":1231E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   86
         Left            =   11520
         TabIndex        =   127
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   86
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   85
         Left            =   14640
         Picture         =   "frmLeave.frx":12760
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   85
         Left            =   11520
         TabIndex        =   126
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   85
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   84
         Left            =   14640
         Picture         =   "frmLeave.frx":12BA2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   84
         Left            =   11520
         TabIndex        =   125
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   84
         Left            =   14160
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   83
         Left            =   10920
         Picture         =   "frmLeave.frx":12FE4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   83
         Left            =   7800
         TabIndex        =   124
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   83
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   82
         Left            =   10920
         Picture         =   "frmLeave.frx":13426
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   82
         Left            =   7800
         TabIndex        =   123
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   82
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   81
         Left            =   10920
         Picture         =   "frmLeave.frx":13868
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   81
         Left            =   7800
         TabIndex        =   122
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   81
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   80
         Left            =   10920
         Picture         =   "frmLeave.frx":13CAA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   80
         Left            =   7800
         TabIndex        =   121
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   80
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   79
         Left            =   10920
         Picture         =   "frmLeave.frx":140EC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   79
         Left            =   7800
         TabIndex        =   120
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   79
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   78
         Left            =   10920
         Picture         =   "frmLeave.frx":1452E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   78
         Left            =   7800
         TabIndex        =   119
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   78
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   77
         Left            =   10920
         Picture         =   "frmLeave.frx":14970
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   77
         Left            =   7800
         TabIndex        =   118
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   77
         Left            =   10440
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   76
         Left            =   7080
         Picture         =   "frmLeave.frx":14DB2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   76
         Left            =   3960
         TabIndex        =   117
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   76
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   75
         Left            =   7080
         Picture         =   "frmLeave.frx":151F4
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   75
         Left            =   3960
         TabIndex        =   116
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   75
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   74
         Left            =   7080
         Picture         =   "frmLeave.frx":15636
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   74
         Left            =   3960
         TabIndex        =   115
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   74
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   73
         Left            =   7080
         Picture         =   "frmLeave.frx":15A78
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   73
         Left            =   3960
         TabIndex        =   114
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   73
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   72
         Left            =   7080
         Picture         =   "frmLeave.frx":15EBA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   72
         Left            =   3960
         TabIndex        =   113
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   72
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   71
         Left            =   7080
         Picture         =   "frmLeave.frx":162FC
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   71
         Left            =   3960
         TabIndex        =   112
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   71
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   70
         Left            =   7080
         Picture         =   "frmLeave.frx":1673E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   70
         Left            =   3960
         TabIndex        =   111
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   70
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   69
         Left            =   3240
         Picture         =   "frmLeave.frx":16B80
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   69
         Left            =   120
         TabIndex        =   110
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   69
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   68
         Left            =   3240
         Picture         =   "frmLeave.frx":16FC2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   68
         Left            =   120
         TabIndex        =   109
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   68
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   67
         Left            =   3240
         Picture         =   "frmLeave.frx":17404
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   67
         Left            =   120
         TabIndex        =   108
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   67
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   66
         Left            =   3240
         Picture         =   "frmLeave.frx":17846
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   66
         Left            =   120
         TabIndex        =   107
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   66
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   65
         Left            =   3240
         Picture         =   "frmLeave.frx":17C88
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   65
         Left            =   120
         TabIndex        =   106
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   65
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   64
         Left            =   3240
         Picture         =   "frmLeave.frx":180CA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   64
         Left            =   120
         TabIndex        =   105
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   64
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   63
         Left            =   3240
         Picture         =   "frmLeave.frx":1850C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   63
         Left            =   120
         TabIndex        =   104
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   63
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   62
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   62
         Left            =   -63480
         TabIndex        =   103
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   62
         Left            =   -60360
         Picture         =   "frmLeave.frx":1894E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   61
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   61
         Left            =   -63480
         TabIndex        =   102
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   61
         Left            =   -60360
         Picture         =   "frmLeave.frx":18D90
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   60
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   60
         Left            =   -63480
         TabIndex        =   101
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   60
         Left            =   -60360
         Picture         =   "frmLeave.frx":191D2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   59
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   59
         Left            =   -63480
         TabIndex        =   100
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   59
         Left            =   -60360
         Picture         =   "frmLeave.frx":19614
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   58
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   58
         Left            =   -63480
         TabIndex        =   99
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   58
         Left            =   -60360
         Picture         =   "frmLeave.frx":19A56
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   57
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   57
         Left            =   -63480
         TabIndex        =   98
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   57
         Left            =   -60360
         Picture         =   "frmLeave.frx":19E98
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   56
         Left            =   -60840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   56
         Left            =   -63480
         TabIndex        =   97
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   56
         Left            =   -60360
         Picture         =   "frmLeave.frx":1A2DA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   55
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   55
         Left            =   -67200
         TabIndex        =   96
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   55
         Left            =   -64080
         Picture         =   "frmLeave.frx":1A71C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   54
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   54
         Left            =   -67200
         TabIndex        =   95
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   54
         Left            =   -64080
         Picture         =   "frmLeave.frx":1AB5E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   53
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   53
         Left            =   -67200
         TabIndex        =   94
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   53
         Left            =   -64080
         Picture         =   "frmLeave.frx":1AFA0
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   52
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   52
         Left            =   -67200
         TabIndex        =   93
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   52
         Left            =   -64080
         Picture         =   "frmLeave.frx":1B3E2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   51
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   51
         Left            =   -67200
         TabIndex        =   92
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   51
         Left            =   -64080
         Picture         =   "frmLeave.frx":1B824
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   50
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   50
         Left            =   -67200
         TabIndex        =   91
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   50
         Left            =   -64080
         Picture         =   "frmLeave.frx":1BC66
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   49
         Left            =   -64560
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   -67200
         TabIndex        =   90
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   49
         Left            =   -64080
         Picture         =   "frmLeave.frx":1C0A8
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   48
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   48
         Left            =   -71040
         TabIndex        =   89
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   48
         Left            =   -67920
         Picture         =   "frmLeave.frx":1C4EA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   47
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   47
         Left            =   -71040
         TabIndex        =   88
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   47
         Left            =   -67920
         Picture         =   "frmLeave.frx":1C92C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   46
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   46
         Left            =   -71040
         TabIndex        =   87
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   46
         Left            =   -67920
         Picture         =   "frmLeave.frx":1CD6E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   45
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   45
         Left            =   -71040
         TabIndex        =   86
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   45
         Left            =   -67920
         Picture         =   "frmLeave.frx":1D1B0
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   44
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   44
         Left            =   -71040
         TabIndex        =   85
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   44
         Left            =   -67920
         Picture         =   "frmLeave.frx":1D5F2
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   43
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   43
         Left            =   -71040
         TabIndex        =   84
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   43
         Left            =   -67920
         Picture         =   "frmLeave.frx":1DA34
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   42
         Left            =   -68400
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   42
         Left            =   -71040
         TabIndex        =   83
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   42
         Left            =   -67920
         Picture         =   "frmLeave.frx":1DE76
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   41
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   41
         Left            =   -74880
         TabIndex        =   82
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   41
         Left            =   -71760
         Picture         =   "frmLeave.frx":1E2B8
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   3360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   40
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   40
         Left            =   -74880
         TabIndex        =   81
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   40
         Left            =   -71760
         Picture         =   "frmLeave.frx":1E6FA
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   39
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   -74880
         TabIndex        =   80
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   39
         Left            =   -71760
         Picture         =   "frmLeave.frx":1EB3C
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   2400
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   38
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   -74880
         TabIndex        =   79
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   38
         Left            =   -71760
         Picture         =   "frmLeave.frx":1EF7E
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1920
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   37
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   -74880
         TabIndex        =   78
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   37
         Left            =   -71760
         Picture         =   "frmLeave.frx":1F3C0
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   1440
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   36
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   -74880
         TabIndex        =   77
         Top             =   960
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   36
         Left            =   -71760
         Picture         =   "frmLeave.frx":1F802
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image imgLocation 
         Height          =   360
         Index           =   35
         Left            =   -72240
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   1935
      End
      Begin VB.Image imgMail 
         Height          =   360
         Index           =   35
         Left            =   -71760
         Picture         =   "frmLeave.frx":1FC44
         Stretch         =   -1  'True
         ToolTipText     =   "Double Click to activate"
         Top             =   480
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdDone 
      Height          =   855
      Left            =   14160
      Picture         =   "frmLeave.frx":20086
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdsound 
      Caption         =   "Fire"
      Height          =   855
      Index           =   0
      Left            =   2280
      Picture         =   "frmLeave.frx":20390
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncident 
      Caption         =   "Incident Log"
      Height          =   855
      Left            =   1200
      Picture         =   "frmLeave.frx":207D2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10560
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      Picture         =   "frmLeave.frx":20C14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10560
      Width           =   1095
   End
   Begin MCI.MMControl mmcPlay 
      Height          =   330
      Left            =   12240
      TabIndex        =   164
      Top             =   6000
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      _Version        =   327681
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      Enabled         =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Image imgInfo 
      Height          =   495
      Left            =   7200
      Picture         =   "frmLeave.frx":21056
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   14400
      X2              =   14400
      Y1              =   840
      Y2              =   6360
   End
   Begin VB.Label lblLoccode 
      Height          =   255
      Left            =   12360
      TabIndex        =   170
      Top             =   10680
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   360
      Left            =   11880
      Picture         =   "frmLeave.frx":21498
      Stretch         =   -1  'True
      Top             =   10560
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblLocation 
      Height          =   255
      Left            =   10080
      TabIndex        =   169
      Top             =   11040
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   9960
      Top             =   10560
      Width           =   4095
   End
   Begin VB.Label lblBarName 
      Height          =   255
      Left            =   10800
      TabIndex        =   168
      Top             =   10680
      Width           =   975
   End
   Begin VB.Label lblBarcode 
      Height          =   255
      Left            =   10080
      TabIndex        =   167
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14040
      TabIndex        =   165
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image imgHome 
      Height          =   375
      Left            =   11760
      Picture         =   "frmLeave.frx":218DA
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLeave 
      Height          =   360
      Left            =   11280
      Picture         =   "frmLeave.frx":227A4
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   34
      Left            =   13920
      Picture         =   "frmLeave.frx":22AAE
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   10200
      TabIndex        =   73
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   11760
      TabIndex        =   72
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   34
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   360
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   15240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   33
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   32
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   31
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   30
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   29
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   28
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   27
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   26
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   25
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   24
      Left            =   13440
      Stretch         =   -1  'True
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   23
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   22
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   21
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   20
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   19
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   18
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   17
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   16
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   15
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   14
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   13
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   12
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   11
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   10
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   9
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   8
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   7
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   6
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   5
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   4
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   3
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   360
      Index           =   2
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   960
      Width           =   360
   End
   Begin VB.Image imgLocation 
      Height          =   480
      Index           =   0
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgLocation 
      Height          =   480
      Index           =   1
      Left            =   12600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIn 
      Height          =   360
      Left            =   10800
      Picture         =   "frmLeave.frx":22EF0
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOut 
      Height          =   360
      Left            =   10320
      Picture         =   "frmLeave.frx":231FA
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   11760
      TabIndex        =   71
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   10560
      TabIndex        =   70
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   33
      Left            =   13920
      Picture         =   "frmLeave.frx":23504
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   11760
      TabIndex        =   69
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   10560
      TabIndex        =   68
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   32
      Left            =   13920
      Picture         =   "frmLeave.frx":23946
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   11760
      TabIndex        =   67
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   10560
      TabIndex        =   66
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   31
      Left            =   13920
      Picture         =   "frmLeave.frx":23D88
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   11760
      TabIndex        =   65
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   10560
      TabIndex        =   64
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   30
      Left            =   13920
      Picture         =   "frmLeave.frx":241CA
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   11760
      TabIndex        =   63
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   10560
      TabIndex        =   62
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   29
      Left            =   13920
      Picture         =   "frmLeave.frx":2460C
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   11760
      TabIndex        =   61
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   10560
      TabIndex        =   60
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   28
      Left            =   13920
      Picture         =   "frmLeave.frx":24A4E
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   11760
      TabIndex        =   59
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   10560
      TabIndex        =   58
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   27
      Left            =   13920
      Picture         =   "frmLeave.frx":24E90
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   11760
      TabIndex        =   57
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   10560
      TabIndex        =   56
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   26
      Left            =   13920
      Picture         =   "frmLeave.frx":252D2
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   11760
      TabIndex        =   55
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   10560
      TabIndex        =   54
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   25
      Left            =   13920
      Picture         =   "frmLeave.frx":25714
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   11760
      TabIndex        =   53
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   10560
      TabIndex        =   52
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   24
      Left            =   13920
      Picture         =   "frmLeave.frx":25B56
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   6600
      TabIndex        =   51
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   5400
      TabIndex        =   50
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   23
      Left            =   8760
      Picture         =   "frmLeave.frx":25F98
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   6600
      TabIndex        =   49
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   5400
      TabIndex        =   48
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   22
      Left            =   8760
      Picture         =   "frmLeave.frx":263DA
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   21
      Left            =   8760
      Picture         =   "frmLeave.frx":2681C
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5400
      TabIndex        =   47
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   6600
      TabIndex        =   46
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   20
      Left            =   8760
      Picture         =   "frmLeave.frx":26C5E
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   5400
      TabIndex        =   45
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   6600
      TabIndex        =   44
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   19
      Left            =   8760
      Picture         =   "frmLeave.frx":270A0
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   5400
      TabIndex        =   43
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   6600
      TabIndex        =   42
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   18
      Left            =   8760
      Picture         =   "frmLeave.frx":274E2
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   5400
      TabIndex        =   41
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   6600
      TabIndex        =   40
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   17
      Left            =   8760
      Picture         =   "frmLeave.frx":27924
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   5400
      TabIndex        =   39
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   6600
      TabIndex        =   38
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   6600
      TabIndex        =   37
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   5400
      TabIndex        =   36
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   14
      Left            =   8760
      Picture         =   "frmLeave.frx":27D66
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   6600
      TabIndex        =   35
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   5400
      TabIndex        =   34
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   16
      Left            =   8760
      Picture         =   "frmLeave.frx":281A8
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   6600
      TabIndex        =   33
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   5400
      TabIndex        =   32
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   15
      Left            =   8760
      Picture         =   "frmLeave.frx":285EA
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line Line2 
      X1              =   9240
      X2              =   9240
      Y1              =   840
      Y2              =   6360
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   13
      Left            =   8760
      Picture         =   "frmLeave.frx":28A2C
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   5400
      TabIndex        =   31
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   6600
      TabIndex        =   30
      Top             =   960
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   12
      Left            =   3960
      Picture         =   "frmLeave.frx":28E6E
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5760
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   1320
      TabIndex        =   28
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   11
      Left            =   3960
      Picture         =   "frmLeave.frx":292B0
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   5280
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1320
      TabIndex        =   26
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   10
      Left            =   3960
      Picture         =   "frmLeave.frx":296F2
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   24
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Image imgMIDN 
      Height          =   465
      Left            =   120
      Picture         =   "frmLeave.frx":29B34
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSBLT 
      Height          =   480
      Left            =   720
      Picture         =   "frmLeave.frx":29E3E
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLEUT 
      Height          =   480
      Left            =   1320
      Picture         =   "frmLeave.frx":2A148
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLCDR 
      Height          =   465
      Left            =   1920
      Picture         =   "frmLeave.frx":2A452
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRank 
      Height          =   570
      Index           =   1
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgCMDR 
      Height          =   465
      Left            =   2520
      Picture         =   "frmLeave.frx":2A75C
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgRank 
      Height          =   570
      Index           =   0
      Left            =   840
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   9
      Left            =   3960
      Picture         =   "frmLeave.frx":2AA66
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   4320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   22
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   8
      Left            =   3960
      Picture         =   "frmLeave.frx":2AEA8
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   20
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   7
      Left            =   3960
      Picture         =   "frmLeave.frx":2B2EA
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   3360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   18
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1320
      TabIndex        =   16
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   6
      Left            =   3960
      Picture         =   "frmLeave.frx":2B72C
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   5
      Left            =   3960
      Picture         =   "frmLeave.frx":2BB6E
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   2400
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   4
      Left            =   3960
      Picture         =   "frmLeave.frx":2BFB0
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   13
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   3
      Left            =   3960
      Picture         =   "frmLeave.frx":2C3F2
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   1440
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image imgMail 
      Height          =   360
      Index           =   2
      Left            =   3960
      Picture         =   "frmLeave.frx":2C834
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   840
      Y2              =   6360
   End
   Begin VB.Label lblBillet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   1
      Left            =   13200
      Picture         =   "frmLeave.frx":2CC76
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgStamp 
      Height          =   360
      Index           =   1
      Left            =   9840
      Picture         =   "frmLeave.frx":2D0B8
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgStamp 
      Height          =   360
      Index           =   0
      Left            =   9360
      Picture         =   "frmLeave.frx":2D4FA
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgMail 
      Height          =   480
      Index           =   0
      Left            =   6360
      Picture         =   "frmLeave.frx":2D93C
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to activate"
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblBillet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblBillet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Scan()
    Static PersonNo
    a = 0
    If Len(lblBarcode) = 2 Then ' location code
        tmrBarcodeCount.Enabled = True
        Exit Sub
    End If
    If Len(lblBarcode) = 3 And Left$(lblBarcode, 1) = "U" Then ' Unescorted pass
        MsgBox "Unescorted passcode here"
        tmrBarcodeCount.Enabled = True
        Exit Sub
    End If
    If Len(lblBarcode) = 3 And Left$(lblBarcode, 1) = "E" Then ' Escorted pass
        MsgBox "Escorted pass code here"
        tmrBarcodeCount.Enabled = True
        Exit Sub
    End If
    If Len(lblBarcode) = 6 Then ' Normal Personal Number
        Do Until (gPersonalNumber(a) = lblBarcode) Or (a = 232)
            a = a + 1
        Loop
    End If
    If Len(lblBarcode) <> 3 And Len(lblBarcode) <> 6 And Len(lblBarcode) <> 2 Then
        MsgBox "Unknown barcode"
        tmrBarcodeCount.Enabled = True
        Exit Sub
    End If
    ' check if the person is already out
    If gLocation(a) <> 0 Then
        ' person out already
        imgLocation(a).Picture = imgIn.Picture
        gLocation(a) = 0
        tmrDelay.Enabled = True
        KeyPreview = False
        lblBarcode = ""
        lblBarName = ""
        If gMail(a) <> "" Then
            mmcPlay.Command = "Close"
            mmcPlay.filename = "d:\winnt\media\tada.wav"
            mmcPlay.Command = "Open"
            mmcPlay.Command = "Play"
            imgMail(a).Picture = imgInfo.Picture
        End If
        gOtherInfo(a) = ""
        Exit Sub
    End If
    If gLocation(a) = 0 And imgArrow.Visible = False Then
        ' person not out, and name scanned
        'txtLocationCode.SetFocus
        For a = 1 To 232
            If lblBarcode = gPersonalNumber(a) Then
                lblBarName = gName(a)
                PersonNo = a
            End If
        Next a
        If lblBarName = "" Then
            lblBarName = "Unknown name"
            tmrBarcodeCount.Enabled = True
        End If
        imgArrow.Visible = True
    End If
    If Left$(lblLoccode, 1) = "U" Or Left$(lblLoccode, 1) = "E" Then
        lblLocation = "Unknown destination"
        imgArrow.Visible = False
        Exit Sub
    End If
    If lblLoccode <> "" Then
        If Left$(lblLoccode, 1) = "0" Then lblLoccode = Right$(lblLoccode, Len(lblLoccode) - 1)
        If lblLoccode <> "0" And Len(lblLoccode) < 4 Then
            lblLocation = gLocationText(Val(lblLoccode))
            gLocation(PersonNo) = Val(lblLoccode)
            Select Case gLocation(PersonNo)
                Case 1
                    frmLeave.imgLocation(PersonNo).Picture = frmLeave.imgLeave.Picture
                Case 2
                    frmLeave.imgLocation(PersonNo).Picture = frmLeave.imgHome.Picture
                Case 3
                    frmLeave.imgLocation(PersonNo).Picture = frmLeave.imgOut.Picture
                Case 4
                    frmLeave.imgLocation(PersonNo).Picture = frmLeave.imgOut.Picture
            End Select
        End If
        If lblLocation = "" Then lblLocation = "Unknown destination"
        If Right$(lblLoccode, 1) = "P" Then ' Extra info required.. phone no etc.
            gOtherInfo(PersonNo) = InputBox("Enter extra info", "More detail required")
        End If
        imgArrow.Visible = False
    End If
End Sub
Private Sub cmdDone_Click()
    mmcPlay.Command = "Close"
    End
End Sub


Private Sub cmdID_Click()
    frmBarcode.Show
End Sub

Private Sub cmdIncident_Click()
    frmLog.Show vbModal
End Sub

Private Sub cmdPennant_Click()
    frmPennant.Show vbModal
End Sub

Private Sub cmdSound_Click(Index As Integer)
    mmcPlay.Command = "Close"
    Select Case Index
        Case 0
            mmcPlay.filename = "c:\wakey.wav"
            soundname = "Fire"
        Case 1
            mmcPlay.filename = "d:\program files\windows nt\pinball\sound3.wav"
            soundname = "Still"
        Case 2
            mmcPlay.filename = "c:\carry-on.wav"
            soundname = "Carry On"
        Case 3
            mmcPlay.filename = "c:\ood.wav"
            soundname = "OOD"
        Case 4
            mmcPlay.filename = "c:\wakey.wav"
            soundname = "Wakey Wakey"
    End Select
    frmPlay.cmdPlay.Caption = "Play " & soundname & " sound"
    frmPlay.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblBarcode <> "" Then
        frmOnOff.txtPersonalNumber = Chr$(KeyCode)
        frmOnOff.txtPersonalNumber.SelStart = 2
    Else
        frmOnOff.txtLocationCode = Chr$(KeyCode)
        frmOnOff.txtLocationCode.SelStart = 2
    End If
    frmOnOff.Show vbModal
End Sub

Private Sub Form_Load()
    MSComm1.PortOpen = True
    On Error GoTo done
    For a = 0 To 232
        lblName(a) = gRank(a) & " " & gName(a)
        If a < 2 Then
            If gRank(a) = "CMDR" Then imgRank(a).Picture = imgCMDR.Picture
            If gRank(a) = "LCDR" Then imgRank(a).Picture = imgLCDR.Picture
            If gRank(a) = "LEUT" Then imgRank(a).Picture = imgLEUT.Picture
            If gRank(a) = "SBLT" Then imgRank(a).Picture = imgSBLT.Picture
            If gRank(a) = "MIDN" Then imgRank(a).Picture = imgMIDN.Picture
        End If
        If a < 35 Then lblBillet(a) = gBilletText(a)
        If gName(a) <> "" Then
            imgMail(a).Visible = True
            imgLocation(a).Picture = imgIn.Picture
        End If
    Next a
done:
End Sub


Private Sub imgLocation_DblClick(Index As Integer)
    gIndex = Index
    If gLocationText(gLocation(Index)) = gLocationText(0) Then
        ' make the person 'out'
        Load frmOnOff
        frmOnOff.txtPersonalNumber = gPersonalNumber(Index)
        frmOnOff.Show vbModal
        imgLocation(Index).ToolTipText = frmOnOff.lblLocation
    Else
        ' make the person 'in'
        response = MsgBox("Log " & gRank(Index) & " " & gName(Index) & " as returned onboard ?", vbYesNo + vbQuestion, "Confirm")
        If response = vbYes Then
            gLocation(Index) = 0
            imgLocation(Index).ToolTipText = gLocationText(0)
            imgLocation(Index) = imgIn.Picture
        End If
    End If
End Sub

Private Sub imgLocation_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And gLocation(Index) <> 0 Then
        MsgBox ("(" & Date & " " & gTime(Index) & ") " & gRank(Index) & " " & gName(Index) & " is " & gLocationText(gLocation(Index)))
    End If
End Sub

Private Sub imgMail_DblClick(Index As Integer)
    ' set a label for the form
    frmMessage.Caption = "Messages for " & frmLeave.lblName(Index)
    ' assign the message to the textbox
    frmMessage.txtMessage = gMail(Index)
    ' if a message already exists, then add a CR and LF
    If frmMessage.txtMessage <> "" Then
        frmMessage.txtMessage = frmMessage.txtMessage & vbCrLf
    End If
    ' now add the date and time
    frmMessage.txtMessage = frmMessage.txtMessage & Now & " "
    ' show the message form
    frmMessage.Show vbModal
    If frmMessage.txtMessage <> "" Then
        frmLeave.imgMail(Index).Picture = frmLeave.imgStamp(1).Picture
    Else
        frmLeave.imgMail(Index).Picture = frmLeave.imgStamp(0).Picture
    End If
    gMail(Index) = frmMessage.txtMessage
    Unload frmMessage
End Sub

Private Sub imgMail_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call imgMail_DblClick(Index)
End Sub

Private Sub imgMail_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If gMail(Index) <> "" Then
            response = MsgBox("Clear this message ?:" & vbCrLf & vbCrLf & gMail(Index), vbQuestion + vbOKCancel, "Confirmation required")
            If response = vbOK Then
                gMail(Index) = ""
                frmLeave.imgMail(Index).Picture = frmLeave.imgStamp(0).Picture
            End If
        End If
    End If
End Sub

Private Sub lblLocation_Change()
    tmrBarcodeCount.Enabled = True
End Sub


Private Sub lblName_DblClick(Index As Integer)
    If gName(Index) <> "" Then
        frmPerson.lblNumber = Index
        frmPerson.Show vbModal
    End If
End Sub

Private Sub lblName_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    response = InputBox("Add info...", "More info required")
    gOtherInfo(Index) = response
End Sub

Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And gOtherInfo(Index) <> "" Then
        MsgBox gOtherInfo(Index)
    End If
End Sub

Private Sub mmcPlay_Done(NotifyCode As Integer)
    mmcPlay.Command = "Close"
End Sub

Private Sub tmrBarcode_Timer()
    Static strBuffer As String
    strBuffer = strBuffer & MSComm1.Input
    If CBool(InStr(1, strBuffer, vbCr, vbTextCompare)) Then
        strBarcode = Mid(strBuffer, 1, InStr(1, strBuffer, vbCr, vbTextCompare) - 1)
        strBuffer = Mid(strBuffer, InStr(1, strBuffer, vbCr, vbTextCompare) + 1)
    End If
    If strBarcode <> "" Then
        If lblBarcode = "" Then
            lblBarcode = strBarcode
        Else
            lblLoccode = strBarcode
        End If
        Call Scan
        strBarcode = ""
        strBuffer = ""
    End If
End Sub

Private Sub tmrBarcodeCount_Timer()
    lblBarcode = ""
    lblBarName = ""
    lblLoccode = ""
    lblLocation = ""
    imgArrow.Visible = False
    tmrBarcodeCount.Enabled = False
End Sub

Private Sub tmrDelay_Timer()
    frmLeave.KeyPreview = True
End Sub

Private Sub tmrTime_Timer()
    If Time <> lblTime Then lblTime = Time
End Sub
