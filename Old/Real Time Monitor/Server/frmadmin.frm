VERSION 5.00
Begin VB.Form frmAdmin 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Session"
   ClientHeight    =   3405
   ClientLeft      =   2085
   ClientTop       =   2460
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3405
   ScaleWidth      =   6165
   Begin VB.TextBox textSocketHandle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ListBox listTranscript 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   105
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   5940
   End
   Begin VB.TextBox textMessageEntry 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   4665
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Client Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Caption         =   "Client IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblClientPort 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblClientIP 
      Caption         =   "0.0.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Transcript:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      TabIndex        =   2
      Top             =   240
      Width           =   810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -15
      X2              =   7800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   7815
      Y1              =   2360
      Y2              =   2360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Enter Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents SocketX As SocketXObj
Attribute SocketX.VB_VarHelpID = -1
Dim boolTx As Boolean, boolRx As Boolean

Private Sub SocketX_Close(ByVal ErrorCode As Integer)
    SocketX.Close
    Unload Me
End Sub

Private Sub SocketX_Connect(ByVal ErrorCode As Integer)
    mdiMain.RXStart
    listTranscript.AddItem "Client Connected"
    mdiMain.RXEnd
End Sub

Private Sub SocketX_Receive(ByVal ErrorCode As Integer)
    Dim t As String
    '
    ' Receive data available, get it
    '
    mdiMain.RXStart
    t = SocketX.Receive
    mdiMain.RXEnd
    '
    ' Echo to client
    '
    Send "Received:" & t
    '
    ' Add to transcript
    '
    If (Len(t) = 1 And Asc(t) = 13) Then
        t = ""
    End If
    listTranscript.AddItem "Received data: " & t
    'listTranscript.ListIndex = listTranscript.ListCount - 1
End Sub

Private Sub Form_Load()
    Set SocketX = New SocketXObj
    SocketX.Blocking = False
    SocketX.EventMask = -1
End Sub

Private Sub SocketX_Send(ByVal ErrorCode As Integer)
    'listTranscript.AddItem "OnSend"
End Sub

Private Sub textMessageEntry_KeyPress(KeyAscii As Integer)
    Dim t As String

    If (KeyAscii = 13) Then
        t = textMessageEntry.Text
        'listTranscript.AddItem "send: " & t
        'listTranscript.ListIndex = listTranscript.ListCount - 1
        If (t = "") Then
            t = Chr$(KeyAscii)
        End If
        Send t
        textMessageEntry.Text = ""
        KeyAscii = 0
    End If
End Sub

Private Sub textSocketHandle_Change()
    SocketX.Socket = CLng(textSocketHandle.Text)
End Sub

Sub Send(strText As String)
    mdiMain.TXStart
    SocketX.Send strText
    mdiMain.TXEnd
End Sub

