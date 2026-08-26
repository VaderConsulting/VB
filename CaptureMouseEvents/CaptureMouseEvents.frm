VERSION 5.00
Begin VB.Form frmCaptureMouseEvents 
   BackColor       =   &H00FF8080&
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4080
      Top             =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmCaptureMouseEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©2003-2010 Vivek Nigam, All Rights Reserved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Purpose : See if the left or right  mouse button is clicked or relised outside and inside the program
' You are free to use this code within your own applications only,
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Const Cap1 = "Capture Mouse Events"
Private Sub Form_Load()
    frmCaptureMouseEvents.Caption = Cap1
End Sub

Private Sub Timer1_Timer()
    '// To check left mouse button pass 1,
    If GetAsyncKeyState(1) = 0 Then
        Label1.Caption = "Your Left Mouse Button Is UP"
    Else
        Label1.Caption = "Your Left Mouse Button Is Down"
    End If
    
    '// To check right mouse button pass 2,
    
'    If GetAsyncKeyState(2) = 0 Then
'        Label1.Caption = "Your Right Mouse Button Is UP"
'    Else
'        Label1.Caption = "Your Right Mouse Button Is Down"
'    End If
'
End Sub
