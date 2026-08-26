VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TOMS Monitor"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckTOMS 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "89.7.3.11"
      RemotePort      =   50101
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    sckTOMS.Connect "89.7.3.11", 50101
    Debug.Print sckTOMS.State
End Sub

Private Sub sckTOMS_Connect()
    Dim cmd As Byte
    
    Debug.Print sckTOMS.State
    cmd = ""
    cmd = cmd & Chr(0) & Chr(62) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0)
    cmd = cmd & Chr(1) & Chr(54) & Chr(1) & Chr(44) & Chr(0) & Chr(0) & Chr(8) & Chr(0)
    cmd = cmd & Chr(127) & Chr(255) & Chr(127) & Chr(8) & Chr(0) & Chr(0) & Chr(0) & Chr(1)
    cmd = cmd & Chr(0) & Chr(4) & Chr(0) & Chr(58) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
    cmd = cmd & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
    cmd = cmd & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(52) & Chr(230) & Chr(0) & Chr(0)
    cmd = cmd & Chr(0) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)
    cmd = cmd & Chr(0) & Chr(0)
    cmd = cmd & "ping"
    Debug.Print "Command: " & cmd
    sckTOMS.SendData cmd
End Sub

Private Sub sckTOMS_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    sckTOMS.GetData strData, vbString
    Debug.Print "Received data"
    Debug.Print strData
End Sub

