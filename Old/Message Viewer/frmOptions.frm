VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtLogDir 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   4680
      Y1              =   740
      Y2              =   740
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblLogDir 
      Caption         =   "Logfile Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    LogDir = txtLogDir
    SaveSetting "Message Viewer", "Options", "LogDir", LogDir
    Unload Me
End Sub

Private Sub Form_Load()
    txtLogDir = LogDir
End Sub
