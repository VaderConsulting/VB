VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DTREnable       =   -1  'True
      BaudRate        =   38400
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    MSComm1.PortOpen = True
End Sub

Private Sub Timer1_Timer()
    Static strBuffer As String
    strBuffer = strBuffer & MSComm1.Input
    If CBool(InStr(1, strBuffer, vbCr, vbTextCompare)) Then
        strBarcode = Mid(strBuffer, 1, InStr(1, strBuffer, vbCr, vbTextCompare) - 1)
        strBuffer = Mid(strBuffer, InStr(1, strBuffer, vbCr, vbTextCompare) + 1)
    End If
    If strBarcode <> "" Then
        Text1 = strBarcode
    End If
End Sub
