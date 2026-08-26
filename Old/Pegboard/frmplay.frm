VERSION 5.00
Begin VB.Form frmPlay 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1770
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   2385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlay 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmLeave.mmcPlay.Command = "Stop"
    Unload Me
End Sub

Private Sub cmdPlay_Click()
    frmLeave.mmcPlay.Command = "Open"
    frmLeave.mmcPlay.Command = "Play"
    If cmdPlay.Caption = "Play Fire sound" Then
        frmLog.Show vbModal
        Unload Me
    End If
End Sub
