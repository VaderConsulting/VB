VERSION 5.00
Object = "*\AxLogon.vbp"
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
   Begin xLogon.drLogon drLogon1 
      Left            =   120
      Top             =   120
      _ExtentX        =   450
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  drLogon1.Username = "PD71005n"
  drLogon1.Password = "password"
  drLogon1.Domain = "POLICE"
  drLogon1.Logon
  Debug.Print drLogon1.Authenticated
End Sub

