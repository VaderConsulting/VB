VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPennant 
   Caption         =   "Pennant Numbers"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3704
      _Version        =   327681
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ANZAC"
      TabPicture(0)   =   "frmPennant.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblShip(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblShip(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblShip(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DDG"
      TabPicture(1)   =   "frmPennant.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblShip(5)"
      Tab(1).Control(1)=   "lblShip(4)"
      Tab(1).Control(2)=   "lblShip(3)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "FFG"
      TabPicture(2)   =   "frmPennant.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblShip(11)"
      Tab(2).Control(1)=   "lblShip(10)"
      Tab(2).Control(2)=   "lblShip(9)"
      Tab(2).Control(3)=   "lblShip(8)"
      Tab(2).Control(4)=   "lblShip(7)"
      Tab(2).Control(5)=   "lblShip(6)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Patrol Boat"
      TabPicture(3)   =   "frmPennant.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblShip(26)"
      Tab(3).Control(1)=   "lblShip(25)"
      Tab(3).Control(2)=   "lblShip(24)"
      Tab(3).Control(3)=   "lblShip(23)"
      Tab(3).Control(4)=   "lblShip(22)"
      Tab(3).Control(5)=   "lblShip(21)"
      Tab(3).Control(6)=   "lblShip(20)"
      Tab(3).Control(7)=   "lblShip(19)"
      Tab(3).Control(8)=   "lblShip(18)"
      Tab(3).Control(9)=   "lblShip(17)"
      Tab(3).Control(10)=   "lblShip(16)"
      Tab(3).Control(11)=   "lblShip(15)"
      Tab(3).Control(12)=   "lblShip(14)"
      Tab(3).Control(13)=   "lblShip(13)"
      Tab(3).Control(14)=   "lblShip(12)"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Auxiliary"
      TabPicture(4)   =   "frmPennant.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblShip(27)"
      Tab(4).Control(1)=   "lblShip(28)"
      Tab(4).Control(2)=   "lblShip(29)"
      Tab(4).ControlCount=   3
      Begin VB.Label lblShip 
         Caption         =   "241 Protector"
         Height          =   255
         Index           =   29
         Left            =   -74760
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "195 Westralia"
         Height          =   255
         Index           =   28
         Left            =   -74760
         TabIndex        =   30
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "304 Success"
         Height          =   255
         Index           =   27
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "217 Bunbury"
         Height          =   255
         Index           =   26
         Left            =   -70440
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "216 Gladstone"
         Height          =   255
         Index           =   25
         Left            =   -70440
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "215 Geelong"
         Height          =   255
         Index           =   24
         Left            =   -70440
         TabIndex        =   26
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "214 Dubbo"
         Height          =   255
         Index           =   23
         Left            =   -70440
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "213 Geraldton"
         Height          =   255
         Index           =   22
         Left            =   -70440
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "212 Gawler"
         Height          =   255
         Index           =   21
         Left            =   -72720
         TabIndex        =   23
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "211 Bendigo"
         Height          =   255
         Index           =   20
         Left            =   -72720
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "210 Cessnock"
         Height          =   255
         Index           =   19
         Left            =   -72720
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "209 Ipswich"
         Height          =   255
         Index           =   18
         Left            =   -72720
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "208 Whyalla"
         Height          =   255
         Index           =   17
         Left            =   -72720
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "207 Launceston"
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "206 Wollongong"
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "205 Townsville"
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "204 Warrnambool"
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblShip 
         Caption         =   "203 Fremantle"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblShip 
         Caption         =   "06 Newcastle"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "05 Melbourne"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "04 Darwin"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   11
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "03 Sydney"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "02 Canberra"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "01 Adelaide"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "41 Brisbane"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "39 Hobart"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "38 Perth"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "152 Warramunga"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "151 Arunta"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblShip 
         Caption         =   "150 ANZAC"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmPennant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub
