VERSION 5.00
Begin VB.Form frmAddProxy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Area Network (LAN) Settings"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame fmeProxy 
      Caption         =   "Proxy server"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   5415
      Begin VB.CommandButton cmdAdvanced 
         Caption         =   "Advanced..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Bypass proxy server for local addresses"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtPort 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtProxy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkProxy 
         Caption         =   "Use a proxy server for your LAN (These settings will not apply to dial-up or VPN connections)."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblPort 
         Caption         =   "Port:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblProxy 
         Caption         =   "Address:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame fmeAutomatic 
      Caption         =   "Automatic Configuration"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtScript 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1560
         Width           =   2775
      End
      Begin VB.CheckBox chkScript 
         Caption         =   "Use automatic configuration script"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chkDetect 
         Caption         =   "Automatically detect setttings"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lblScript 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblAutomatic 
         Caption         =   "Automatic configuration may override manual settings.  To ensure the use of manual settings, disable automatic configuration."
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmAddProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkProxy_Click()
    If chkProxy.Value = vbChecked Then
        lblProxy.Enabled = True
        txtProxy.Enabled = True
        lblPort.Enabled = True
        txtPort.Enabled = True
        cmdAdvanced.Enabled = True
        chkLocal.Enabled = True
        txtProxy.BackColor = &H80000005
        txtPort.BackColor = &H80000005
    Else
        lblProxy.Enabled = False
        txtProxy.Enabled = False
        lblPort.Enabled = False
        txtPort.Enabled = False
        cmdAdvanced.Enabled = False
        chkLocal.Enabled = False
        txtProxy.BackColor = &H8000000F
        txtPort.BackColor = &H8000000F
    End If
End Sub

Private Sub chkScript_Click()
    If chkScript.Value = vbChecked Then
        lblScript.Enabled = True
        txtScript.Enabled = True
        txtScript.BackColor = &H80000005
    Else
        lblScript.Enabled = False
        txtScript.Enabled = False
        txtScript.BackColor = &H8000000F
    End If
End Sub

Private Sub cmdAdvanced_Click()
    frmAdvanced.txtProxy(0).Text = txtProxy.Text
    frmAdvanced.txtPort(0).Text = txtPort.Text
    frmAdvanced.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
