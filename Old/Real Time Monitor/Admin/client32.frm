VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RTM Admin Client"
   ClientHeight    =   3465
   ClientLeft      =   1515
   ClientTop       =   1950
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   6720
   Begin VB.TextBox textMessageEntry 
      Height          =   900
      Left            =   1560
      TabIndex        =   0
      Top             =   2475
      Width           =   5055
   End
   Begin VB.ListBox listTranscript 
      Height          =   1815
      Left            =   240
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   330
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Message and Press Enter:"
      Height          =   570
      Left            =   240
      TabIndex        =   1
      Top             =   2445
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   6600
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   6600
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Transcript:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   75
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents SocketX As SocketXObj
Attribute SocketX.VB_VarHelpID = -1

Private Sub SocketX_Close(ByVal ErrorCode As Integer)
   ' Close the program if our server goes away
   Unload Me
End Sub

Private Sub SocketX_Receive(ByVal ErrorCode As Integer)
   Dim t As String
   
   ' Get the text received
   t = SocketX.Receive
   
   ' If all we have is a carriage return ...
   If (Len(t) = 1 And Asc(Left(t, 1)) = 13) Then
      ' ... just use a blank string.
      t = ""
   End If
      
   ' Add our new string to the transcript
   listTranscript.AddItem t
   listTranscript.ListIndex = listTranscript.ListCount - 1
   If Left(t, 9) = "Received:" Then
    'Unload Me
    'End
   End If
End Sub
   
Private Sub Form_Load()
   Me.Show
   Set SocketX = New SocketXObj
   SocketX.Blocking = False
   
   ' Get the server's IP address.
   frmEnterIP.Show 1
   'SocketX.RemoteAddress = frmEnterIP.txtIP
   
   ' If the user pressed Esc/Cancel exit.
   'If SocketX.RemoteAddress = "" Then
   '   End
   'End If
   
   ' Set up the socket control.
   SocketX.RemoteNameAddrXlate = False
   SocketX.RemotePort = 25
   SocketX.LocalAddress = "192.168.0.20"
   SocketX.LocalPort = 0
   SocketX.Create
   
   ' Connect to the server.
   On Error Resume Next
   SocketX.Connect "dellp1000"
   If (Err <> 0 And Err <> 10035) Then
      'MsgBox Error
   End If
   On Error GoTo 0
   SocketX.Send "Connected...."
End Sub
   
Private Sub Form_Unload(Cancel As Integer)
   ' Disconnect from test server
   SocketX.Close
End Sub


Private Sub textMessageEntry_KeyPress(KeyAscii As Integer)
   Dim t As String
   
   ' Did the user press Enter?
   If (KeyAscii = 13) Then
      t = textMessageEntry.Text
      If (textMessageEntry.Text = "") Then
         t = Chr$(13)
      End If
      '
      ' If you want to send a string containing binary data
      ' you must assign it to SendBuffer and then use Send
      ' without any arguments.  If you use Send strTemp VB
      ' will truncate the string at the first chr$(0) if it
      ' contains any.
      '
      SocketX.Send t
      textMessageEntry.Text = ""
      KeyAscii = 0
   End If
End Sub
