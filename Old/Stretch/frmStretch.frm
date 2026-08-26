VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Text            =   "16"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Here"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Here"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Here"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Here"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   3120
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3120
      Picture         =   "frmStretch.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Get Icon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "End"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
    End
End Sub

Private Sub Command5_Click()
    Command1.Picture = Picture1.Picture
End Sub

Private Sub Command6_Click()
    Command2.Picture = Picture2.Picture
End Sub

Private Sub Command8_Click()
    Command3.Picture = Picture2.Picture
End Sub

Private Sub Command9_Click()
    Dim sx&, sy&, sw&, sh&, dx&, dy&, dw&, dh&, di&, dj&
    sx = 0
    sy = 0
    sw = 32
    sh = 32
    dx = 0
    dy = 0
    dw = CInt(txtSize.Text)
    dh = CInt(txtSize.Text)
    di& = StretchBlt(Picture2.hdc, dx, dy, dw, dh, Picture1.hdc, sx, sy, sw, sh, SRCCOPY)
    Rem dj& = SelectObject(Picture2.hdc, Picture2.hdc)
End Sub

