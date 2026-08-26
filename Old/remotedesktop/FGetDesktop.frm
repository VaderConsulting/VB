VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FGetDesktop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get desktop"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6200
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   2580
         TabIndex        =   2
         Text            =   "192.168.169.137"
         Top             =   160
         Width           =   1575
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Get desktop"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "IP Number or Computer Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   180
         Width           =   2295
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5280
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
      RemotePort      =   7001
      LocalPort       =   7000
   End
   Begin VB.Image ImgView 
      Height          =   4575
      Left            =   120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "FGetDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReceiveData As String
Dim intFileNum As Integer
Dim FileName As String
Dim iTop As Integer


Private Sub cmdSend_Click()
On Error Resume Next
If FileName <> "" Then Kill (App.Path & "\" + FileName)
Winsock1.RemoteHost = txtIP.Text
Winsock1.SendData "IP:" & Winsock1.LocalIP
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData ReceiveData
FileName = txtIP.Text & ".jpg"
If InStr(1, ReceiveData, "OpenFile") <> 0 Then
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    intFileNum = FreeFile
    If Dir(App.Path & "\" + FileName) Then
        Kill (App.Path & "\" + FileName)
    End If
    Open App.Path & "\" + FileName For Binary Access Write As #intFileNum
End If
If InStr(1, ReceiveData, "°°°°") <> 0 Then
    Put #intFileNum, , Right$(ReceiveData, Len(ReceiveData) - 4)
End If
If ReceiveData = "Close" Then
    Close #intFileNum
    ImgView.Picture = LoadPicture(App.Path & "\" & FileName)
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End If
End Sub


