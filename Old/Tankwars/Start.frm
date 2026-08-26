VERSION 5.00
Begin VB.Form Start 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Tag             =   "190"
      Top             =   0
      Width           =   6135
      Begin VB.Timer Timer5 
         Left            =   2640
         Top             =   360
      End
      Begin VB.Timer Timer4 
         Left            =   2040
         Top             =   360
      End
      Begin VB.Timer Timer3 
         Left            =   1440
         Top             =   360
      End
      Begin VB.Timer Timer2 
         Left            =   840
         Top             =   360
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   240
         Top             =   360
      End
   End
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TWp As Integer
Private Sub Timer1_Timer()
 View.Cls
 TWp = TWp + 10
 View.FontSize = 25
 View.PSet (120, TWp), RGB(195, 195, 195)
 View.Print "Tank Wars"
 If TWp = 100 Then Timer1.Interval = 0: Timer2.Interval = 10: TWp = 200
End Sub

Private Sub Timer2_Timer()
 View.Cls
 TWp = TWp - 10
 View.FontSize = 25
 View.PSet (120, 100), RGB(195, 195, 195)
 View.Print "Tank Wars"
 View.FontSize = 45
 View.PSet (120, TWp), RGB(195, 195, 195)
 View.Print "3D"
 If TWp = 120 Then Timer2.Interval = 0: Timer3.Interval = 2: TWp = 400
End Sub
Private Sub Timer3_Timer()
 View.Cls
 TWp = TWp - 10
 View.FontSize = 25
 View.PSet (120, 100), RGB(195, 195, 195)
 View.Print "Tank Wars"
 View.FontSize = 45
 View.PSet (120, 120), RGB(195, 195, 195)
 View.Print "3D"
 View.FontSize = 70
 View.PSet (TWp, 120), RGB(195, 195, 195)
 View.ForeColor = 255
 View.Print "2"
 View.ForeColor = 0
 If TWp = 120 Then Timer3.Interval = 0: Timer4.Interval = 1
End Sub

Private Sub Timer4_Timer()
 x = View.Tag / 2
 View.BackColor = RGB(View.Tag, 170 - x, 170 - x)
 View.Tag = View.Tag + 10
 If View.Tag > 255 Then Timer4.Interval = 0: Timer5.Interval = 1
End Sub

Private Sub Timer5_Timer()
 View.BackColor = RGB(View.Tag, 0, 0)
 View.Tag = View.Tag - 10
 If View.Tag < 10 Then
 Timer5.Interval = 0
 Form1.Visible = True
 Unload Me
 End If
End Sub

Private Sub View_Click()
 Form1.Visible = True
 Unload Me
End Sub

Private Sub View_KeyPress(KeyAscii As Integer)
 Form1.Visible = True
 Unload Me
End Sub
