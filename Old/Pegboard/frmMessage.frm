VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtMessage 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtMessage = ""
End Sub

Private Sub cmdDone_Click()
    frmMessage.Hide
End Sub

Private Sub Form_Activate()
    txtMessage.SetFocus
    txtMessage.SelStart = Len(txtMessage)
End Sub

