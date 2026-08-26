VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Thresholds"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeDisk"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fmePing"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame fmePing 
         Caption         =   "Ping Times"
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Frame fmeDisk 
         Caption         =   "Disk Space"
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   5775
         Begin VB.Label lblInformation 
            Caption         =   "Information"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
End Sub
