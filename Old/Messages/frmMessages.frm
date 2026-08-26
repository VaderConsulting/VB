VERSION 5.00
Begin VB.Form frmMessages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messages"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11985
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox filMessages 
      Height          =   1650
      Left            =   120
      Pattern         =   "*.txt"
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next Message"
      Default         =   -1  'True
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   11775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11160
      TabIndex        =   0
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label lblIntro 
      Alignment       =   1  'Right Justify
      Caption         =   "Total number of messages:"
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdNext_Click()
    Static Number As Integer
    txtMessage = ""
    message = ""
    varPath = filMessages.Path & "\" & filMessages.List(Number)
    Open varPath For Input As #1
        Do Until EOF(1)
            Line Input #1, msg
            message = message & msg & vbCrLf
        Loop
    Close 1
    txtMessage.Text = message
    Number = Number + 1
    If Number = filMessages.ListCount Then
        cmdNext.Enabled = False
        cmdClose.Enabled = True
        cmdClose.Default = True
    End If
End Sub

Private Sub Form_Load()
    Dim Server As String
    Server = Environ$("logonserver")
    filMessages.Path = Server & "\netlogon\messages"
    lblNumber = filMessages.ListCount
    If filMessages.ListCount = 0 Then
        End
    End If
    cmdNext_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
End Sub
