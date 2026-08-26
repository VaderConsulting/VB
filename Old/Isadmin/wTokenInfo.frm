VERSION 5.00
Begin VB.Form frmIsAdmin 
   Caption         =   "IsAdmin"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAdmin 
      Caption         =   "Are You Admin?"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmIsAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Developed for you by Elvio Serrao
' Elvio.Serrao@nrma.com.au

Private Sub btnAdmin_Click()
    If IsAdmin Then
        MsgBox "Your an Administrator", vbInformation, Caption
    Else
        MsgBox "Keep Dreaming", vbInformation, Caption
    End If
End Sub
