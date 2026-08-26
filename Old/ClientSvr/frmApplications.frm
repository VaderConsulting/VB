VERSION 5.00
Begin VB.Form frmApplications 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Applications"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmApplications.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeNew 
      Caption         =   "New"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   4575
      Begin VB.TextBox txtManual 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAddP 
         Caption         =   "Add to Permanent"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddT 
         Caption         =   "Add to Temporary"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdTPALL 
      Caption         =   "<<"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdPTAll 
      Caption         =   ">>"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdTP 
      Caption         =   "<"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPT 
      Caption         =   ">"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.ListBox lstTemporary 
      Height          =   2205
      ItemData        =   "frmApplications.frx":0442
      Left            =   2760
      List            =   "frmApplications.frx":0444
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdRemoveP 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ListBox lstPermanent 
      Height          =   2205
      ItemData        =   "frmApplications.frx":0446
      Left            =   120
      List            =   "frmApplications.frx":0448
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblTemporary 
      Alignment       =   2  'Center
      Caption         =   "Temporary"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblPermanent 
      Alignment       =   2  'Center
      Caption         =   "Permanent"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    frmServer.Show
    Unload Me
End Sub

