VERSION 5.00
Begin VB.Form frmPrint 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5550
   Begin VB.PictureBox picBarcode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   840
      ScaleHeight     =   855
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Shape shpBox 
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin VB.Image imgRAN 
      Height          =   1815
      Left            =   3960
      Picture         =   "frmPrint.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblBillet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
   End
   Begin VB.Label lblNameRank 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HMAS ARUNTA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Image imgCrest 
      Height          =   1935
      Left            =   240
      Picture         =   "frmPrint.frx":1E8C2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
