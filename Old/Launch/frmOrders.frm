VERSION 5.00
Begin VB.Form frmOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orders"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fleFiles 
      Height          =   3210
      Left            =   1920
      TabIndex        =   9
      Top             =   480
      Width           =   2655
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   2790
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Text            =   "D:\Program Files\Wordview\Wordview.exe"
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtBrowse 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "J:\Bulletin"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open in Word Viewer"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblPath 
      Caption         =   "Wordview Path"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblBrowse 
      AutoSize        =   -1  'True
      Caption         =   "Browse Directory"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    frmOther.Show
    Unload Me
End Sub
