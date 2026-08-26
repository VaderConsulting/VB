VERSION 5.00
Begin VB.Form frmAdvanced 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Settings"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame fmeExceptions 
      Caption         =   "Exceptions"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   5175
      Begin VB.TextBox txtExceptions 
         Height          =   615
         Left            =   840
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   600
         Width           =   4215
      End
      Begin VB.Image imgExceptions 
         Height          =   615
         Left            =   120
         Picture         =   "frmAdvanced.frx":0000
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Use semicolons ( ; ) to separate entries."
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Do not use proxy server for addresses beginning with:"
         Height          =   255
         Left            =   840
         TabIndex        =   28
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame fmeServers 
      Caption         =   "Servers"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtPort 
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   25
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   24
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Index           =   2
         Left            =   4320
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   21
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   20
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   19
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkUseSame 
         Caption         =   "Use the same proxy server for all protocols"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox txtProxy 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Index           =   0
         Left            =   4320
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Image imgServers 
         Height          =   615
         Left            =   120
         Picture         =   "frmAdvanced.frx":01E6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblSocks 
         Caption         =   "Socks:"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblGopher 
         Caption         =   "Gopher:"
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblFTP 
         Caption         =   "FTP:"
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblSecure 
         Caption         =   "Secure:"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblHTTP 
         Caption         =   "HTTP:"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   9
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   8
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   7
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblColon 
         Caption         =   ":"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblPort 
         Caption         =   "Port"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblProxyAddress 
         Caption         =   "Proxy address to use"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseSame_Click()
    Dim intLoop As Integer
    If chkUseSame.Value = vbChecked Then
        For intLoop = 1 To 4
            txtPort(intLoop).Text = txtPort(0).Text
        Next intLoop
        For intLoop = 1 To 4
            txtProxy(intLoop).Text = txtProxy(0).Text
        Next intLoop
        For intLoop = 1 To 4
            txtProxy(intLoop).Enabled = False
            txtPort(intLoop).Enabled = False
            txtProxy(intLoop).BackColor = frmAdvanced.BackColor
            txtPort(intLoop).BackColor = frmAdvanced.BackColor
        Next
    Else
        For intLoop = 1 To 4
            txtProxy(intLoop).Enabled = True
            txtPort(intLoop).Enabled = True
            txtProxy(intLoop).BackColor = txtProxy(0).BackColor
            txtPort(intLoop).BackColor = txtPort(0).BackColor
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub txtPort_Change(Index As Integer)
    Dim intLoop As Integer
    If Index = 0 And chkUseSame.Value = vbChecked Then
        For intLoop = 1 To 4
            txtPort(intLoop).Text = txtPort(0).Text
        Next intLoop
    End If
End Sub

Private Sub txtProxy_Change(Index As Integer)
    Dim intLoop As Integer
    If Index = 0 And chkUseSame.Value = vbChecked Then
        For intLoop = 1 To 4
            txtProxy(intLoop).Text = txtProxy(0).Text
        Next intLoop
    End If
End Sub
