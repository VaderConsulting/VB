VERSION 5.00
Begin VB.Form frmPreferences 
   Caption         =   "Preferences"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPreferences.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtSignalsDir 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblSignalsDir 
      Caption         =   "Signals Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    gSignalsDir = txtSignalsDir
    Unload Me
End Sub

Private Sub Form_Load()
    txtSignalsDir = gSignalsDir
    txtSignalsDir.SelStart = 0
    txtSignalsDir.SelLength = Len(txtSignalsDir)
End Sub
