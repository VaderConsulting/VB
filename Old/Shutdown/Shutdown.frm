VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shut me down"
   ClientHeight    =   1005
   ClientLeft      =   2820
   ClientTop       =   2370
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1005
   ScaleWidth      =   2685
   Begin VB.CommandButton Command1 
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   390
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bState As Boolean
Private Sub Command1_Click()

    If bState Then
        Command1.Caption = "Shut Down"
        NTPrivilege.AbortShutDown ""
    Else
        Command1.Caption = "Abort"
        NTPrivilege.ShutDown "", "Buh-bye", 20, True, False
    End If
    bState = Not bState

End Sub



