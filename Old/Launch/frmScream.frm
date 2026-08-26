VERSION 5.00
Begin VB.Form frmScream 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   975
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblAuto 
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Image imgScream 
      Height          =   495
      Left            =   120
      Picture         =   "frmScream.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblPage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Page the Administrator(s) with the following message:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmScream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmOther.Show
    Unload Me
End Sub

Private Sub cmdSend_Click()
    MsgBox "Sorry.. currently not implemented."
End Sub

Private Sub Form_Activate()
    lblAuto = Chr$(34) & "Signed " & gUserName
    lblAuto = lblAuto & " using " & gComputerName
    lblAuto = lblAuto & Chr$(34)
    lblAuto = lblAuto & " will automatically be added to the end of your message."
    txtMessage = "Sorry to interrupt you, but I require your assistance immediately.  Please come to my aid."
    txtMessage.SetFocus
    txtMessage.SelStart = 0
    txtMessage.SelLength = Len(txtMessage)
End Sub

