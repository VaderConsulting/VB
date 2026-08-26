VERSION 5.00
Object = "{1309C0C7-C7B6-11D1-87CB-000000000000}#1.0#0"; "barpic.ocx"
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin BarPic.cbbBarPic barBarcode 
      Height          =   840
      Left            =   2760
      TabIndex        =   6
      Top             =   2160
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   1482
   End
   Begin VB.TextBox txtData 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image imgCrest 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   4920
      Picture         =   "frmPrint.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblShip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HMAS ARUNTA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gangway Pass"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    'Height = frmBarcode.Height
End Sub

Private Sub txtData_Change()
    barBarcode.Data = txtData
    barBarcode.Left = (frmPrint.Width - barBarcode.Width) / 2
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtData.SetFocus
    End If
End Sub

