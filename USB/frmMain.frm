VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Detect"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   2055
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Hook Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook
    End
End Sub



