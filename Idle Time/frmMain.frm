VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Idle Time"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   2580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrIdle 
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin VB.Label lblIdle 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrIdle_Timer()
    Dim LInfo As LastInputInfo
    
    LInfo.cbSize = Len(LInfo)
    
    If GetLastInputInfo(LInfo) <> 0 Then
        'Show how many seconds the user has been idle
        lblIdle.Caption = (GetTickCount() - LInfo.dwTime) \ 1000
    End If
    

End Sub
