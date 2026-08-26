VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    d = Dir("\\mojadmin\SMSLOGON\SITES\", vbDirectory)
    Do Until d = ""
        Debug.Print d
        If d <> "." And d <> ".." Then
            Open "\\mojadmin\SMSLOGON\SITES\" & d & "\netconf.ncf" For Input As #1
                Do Until EOF(1)
                    Line Input #1, f

                    Debug.Print f
                Loop
            Close 1
        End If
        d = Dir
    Loop
End Sub
