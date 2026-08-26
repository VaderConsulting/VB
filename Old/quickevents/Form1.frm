VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Event Viewer"
   ClientHeight    =   2412
   ClientLeft      =   3840
   ClientTop       =   2880
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2412
   ScaleWidth      =   3480
   Begin VB.OptionButton Opt 
      Caption         =   "Application"
      Height          =   228
      Index           =   2
      Left            =   1296
      TabIndex        =   5
      Top             =   1404
      Width           =   1560
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Security"
      Height          =   228
      Index           =   1
      Left            =   1296
      TabIndex        =   4
      Top             =   1116
      Width           =   1560
   End
   Begin VB.OptionButton Opt 
      Caption         =   "System"
      Height          =   228
      Index           =   0
      Left            =   1296
      TabIndex        =   3
      Top             =   828
      Width           =   1560
   End
   Begin VB.TextBox txtServer 
      Height          =   288
      Left            =   1296
      TabIndex        =   1
      Text            =   "\\MYSERVER"
      Top             =   288
      Width           =   1668
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write event strings to file"
      Height          =   336
      Left            =   216
      TabIndex        =   0
      Top             =   1836
      Width           =   3036
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Server :"
      Height          =   192
      Left            =   180
      TabIndex        =   2
      Top             =   324
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private eType As String
Private Sub Command1_Click()
Dim ret As Boolean

Screen.MousePointer = vbHourglass
DoEvents

If ReadEvents(txtServer, eType) Then
    Screen.MousePointer = vbDefault
    MsgBox "Done !", vbInformation
Else
    Screen.MousePointer = vbDefault
    MsgBox "Error !", vbCritical
End If

End Sub


Private Sub Form_Load()

Opt(0).Value = True
Open App.Path & "\evt.txt" For Append As #1


End Sub

Private Sub Form_Unload(Cancel As Integer)

Close #1

End Sub


Private Sub Opt_Click(Index As Integer)

Select Case Index
    Case 0
        eType = EVNT_SYSTEM
    Case 1
        eType = EVNT_SECURITY
    Case 2
        eType = EVNT_APP
End Select

End Sub


