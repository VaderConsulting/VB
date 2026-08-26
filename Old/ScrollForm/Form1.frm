VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scrolling Form"
   ClientHeight    =   1425
   ClientLeft      =   5550
   ClientTop       =   3510
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4845
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   4560
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1035
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oldPos As Integer

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim iFullFormHeigth As Integer
Dim iDisplayHeight  As Integer

iFullFormHeigth = 3765
iDisplayHeight = 1800

Me.Height = iDisplayHeight

With VScroll1
    .Height = Me.ScaleHeight
    .Min = 0
    .Max = iFullFormHeigth - iDisplayHeight
    .SmallChange = Screen.TwipsPerPixelY * 10
    .LargeChange = .SmallChange
End With
End Sub


Private Sub pScrollForm()
Dim ctl As Control

For Each ctl In Me.Controls
    If Not (TypeOf ctl Is VScrollBar) And _
        Not (TypeOf ctl Is CommandButton) Then
        ctl.Top = ctl.Top + oldPos - VScroll1.Value
    End If
Next

oldPos = VScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Call pScrollForm
End Sub


Private Sub VScroll1_Scroll()
    Call pScrollForm
End Sub


