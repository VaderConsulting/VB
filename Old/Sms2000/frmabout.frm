VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About SMS2000"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optPricing 
      Caption         =   "Developer (includes all source code)"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.OptionButton optPricing 
      Caption         =   "Corporate use"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton optPricing 
      Caption         =   "Personal use"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Label lblRegistered 
      Alignment       =   2  'Center
      Caption         =   "Thank you for purchasing SMS2000."
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label lblPricing 
      Caption         =   "There are three pricing models:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblInstall 
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblInstalled 
      Caption         =   "SMS2000 installation date:"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label lblPricing 
      Caption         =   "$27.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label lblPricing 
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblPricing 
      Alignment       =   2  'Center
      Caption         =   "If you wish to use this program beyond the standard 30 day evaluation period, you must purchase it."
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   7095
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "SMS2000 uses the SMS protocols as published by Telstra Corporation 01 September 1998."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "This program is shareware, and copyright (C) 2000 D. Robinson. "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' gRegistered = True ' Testing only
    If Right(App.Path, 1) <> "\" Then
        AppDate = FileDateTime(App.Path & "\" & App.EXEName & ".exe")
    Else
        AppDate = FileDateTime(App.Path & App.EXEName & ".exe")
    End If
    lblInstall = AppDate
    If gRegistered = True Then
        lblRegistered.Visible = True
        For lp = lblPricing().LBound To lblPricing().UBound
            lblPricing(lp).Visible = False
        Next lp
        For lp = optPricing().LBound To optPricing().UBound
            optPricing(lp).Visible = False
        Next lp
    End If
End Sub

Private Sub optPricing_Click(Index As Integer)
    Select Case Index
        Case 0
            lblPricing(3) = "$27.50 inc GST"
        Case 1
            lblPricing(3) = "$25 plus $10 per mobile phone plus 10% GST"
        Case 2
            lblPricing(3) = "$1100 inc GST"
    End Select
End Sub
