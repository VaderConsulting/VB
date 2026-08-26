VERSION 5.00
Begin VB.Form frmSales 
   Caption         =   "Sales"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Default         =   -1  'True
      Height          =   975
      Left            =   2400
      Picture         =   "frmSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtProductID 
      BackColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.ListBox lstProducts 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   1440
      TabIndex        =   11
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   10
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   1440
      TabIndex        =   14
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   13
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   480
      TabIndex        =   12
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblView 
      Alignment       =   2  'Center
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblItems 
      Alignment       =   1  'Right Justify
      Caption         =   "0 items"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Dim s As String * 20
    subtotal = Val(lblSubtotal) + Val(txtPrice)
    lblSubtotal = Format(subtotal, "#######.##")
    RSet s = txtProductID
    lstProducts.AddItem s & " " & txtPrice
    txtPrice = ""
    txtProductID = ""
    If lstProducts.ListCount <> 1 Then
        lblItems = lstProducts.ListCount & " items"
    Else
        lblItems = lstProducts.ListCount & " item"
    End If
End Sub

Private Sub cmdDigit_Click(Index As Integer)
    txtPrice = txtPrice & Chr$(48 + Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 65 Or KeyCode > 90 Then
        txtProductID = txtProductID & Chr$(KeyCode)
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii > 57 Then KeyAscii = 0
    Select Case KeyAscii
    Case 43
        cmdAdd_Click
    Case Is > 47 < 58
        cmdDigit_Click (KeyAscii - 48)
    End Select
End Sub
