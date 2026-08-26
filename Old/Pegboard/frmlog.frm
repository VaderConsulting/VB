VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incident Log"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLog 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   9375
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   5280
      Width           =   855
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    txtLog = txtLog & Date & " " & Time & " "
    txtLog.SelStart = Len(txtLog.Text)
End Sub

