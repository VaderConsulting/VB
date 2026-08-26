VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oNetworkSpeed As drNetwork.clsSpeed

Private Sub Form_Load()
    Set oNetworkSpeed = CreateObject("drNetwork.clsSpeed")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set oNetworkSpeed = Nothing
End Sub

Private Sub Timer1_Timer()
    Dim ec As Long
    Dim em As String
    Dim speed As Long
    
    speed = oNetworkSpeed.GetConnectionSpeed("\\server\c$", "dave", "r6qjDETfB54bxl", ec, em)
    Label1.Caption = "NetworkSpeed : " & speed / 1024 / 1024 & " Mbytes per second"
End Sub
