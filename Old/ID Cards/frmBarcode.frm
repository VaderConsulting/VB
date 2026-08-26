VERSION 5.00
Begin VB.Form frmBarcode 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2340
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2580
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame fmeType 
      Caption         =   "Barcode Type"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optType 
         Caption         =   "Other"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton optType 
         Caption         =   "Location"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optType 
         Caption         =   "Escorted Visitor Pass"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optType 
         Caption         =   "Unescorted Visitor Pass"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton optType 
         Caption         =   "Gangway Pass"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    End
End Sub

Private Sub Form_GotFocus()
    frmPrint.txtName.SetFocus
End Sub

Private Sub Form_Load()
    frmPrint.Show
    Left = (Screen.Width - Width - frmPrint.Width) / 2
    Top = (Screen.Height - Height) / 2
    frmPrint.Left = Width + Left
    frmPrint.Top = Top
End Sub

Private Sub optType_Click(Index As Integer)
    Select Case Index
        Case 0
            frmPrint.lblLabel = "Service Number"
            frmPrint.lblType = "Gangway Pass"
            frmPrint.lblName = "Name"
        Case 1
            frmPrint.lblLabel = "Pass Number"
            frmPrint.lblType = "Unescorted Visitor"
            frmPrint.lblName = "Name"
        Case 2
            frmPrint.lblLabel = "Pass Number"
            frmPrint.lblType = "Escorted Visitor"
            frmPrint.lblName = "Name"
        Case 3
            frmPrint.lblLabel = "Location Code"
            frmPrint.lblType = "Location"
            frmPrint.lblName = "Name"
        Case 4
            frmPrint.lblLabel = "ID Number"
            frmPrint.lblType = ""
            frmPrint.lblName = "Item"
    End Select
    frmPrint.txtName.SetFocus
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtData.SetFocus
    End If
End Sub
