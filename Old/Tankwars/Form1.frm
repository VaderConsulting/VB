VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tank wars"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox lstLevelName 
      Height          =   1065
      Left            =   240
      Pattern         =   "*.dat"
      System          =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GOooo...."
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "You can make your own levels if you want, I'm sure you can figuer out how they work. "
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Select the level you want to run"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":00C6
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Visible = False
Tank.Visible = True
End Sub
