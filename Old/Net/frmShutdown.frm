VERSION 5.00
Begin VB.Form frmShutdown 
   Caption         =   "Form2"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2985
   LinkTopic       =   "Form2"
   ScaleHeight     =   1290
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "server"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abort"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

   If (MsgBox("Are you sure you want to initiate a forced, timed shutdown?", vbYesNo Or vbQuestion) = vbYes) Then

      NTForceTimedShutdown CLng(Text1.Text), "You're gonna get shutdown in " & Text1.Text & " seconds", Text2.Text, True, True
   End If

End Sub

Private Sub Command2_Click()
   NTAbortTimedShutdown Text2.Text
End Sub

Private Sub Form_Load()
   Text1.Text = 60
End Sub
