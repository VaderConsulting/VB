VERSION 5.00
Begin VB.Form frmSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Send"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Net Send"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtMsg 
      Height          =   735
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblMsg 
      Caption         =   "Message (<128 chars)"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSend_Click()
    msg = "Net Send " & txtUsername & " " & txtMsg
    On Error Resume Next
        Kill "c:\temp\netsend.bat"
    On Error GoTo 0
    Open "c:\temp\netsend.bat" For Output As #1
        Print #1, msg
    Close 1
    Shell "c:\temp\netsend.bat", vbHide
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If Len(txtMsg.Text) > 127 Then
        KeyAscii = 0
    End If
End Sub
