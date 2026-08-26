VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRegisterClan 
      Caption         =   "Register Clan"
      Height          =   615
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdRegisterPlayer 
      Caption         =   "Register Player"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgSoldier 
      Height          =   855
      Left            =   9720
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   495
   End
   Begin VB.Image imgClan 
      Height          =   735
      Left            =   9360
      Picture         =   "frmMain.frx":08EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgSelection 
      Height          =   5415
      Left            =   1200
      Picture         =   "frmMain.frx":1DC1
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5310
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
