VERSION 5.00
Begin VB.Form FrmService 
   Caption         =   "Service"
   ClientHeight    =   2655
   ClientLeft      =   1305
   ClientTop       =   6180
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Status"
      Height          =   615
      Left            =   90
      TabIndex        =   3
      Top             =   1980
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause Schedule"
      Height          =   615
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Schedule"
      Height          =   615
      Left            =   90
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Schedule"
      Height          =   615
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2175
   End
End
Attribute VB_Name = "FrmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Here NTSUP7 is my machine and MSDTC is a Service
    vbStopService MachineName, ServiceName
End Sub


Private Sub Command2_Click()
    'Here NTSUP7 is my machine and MSDTC is a Service
    vbStartService MachineName, ServiceName
End Sub


Private Sub Command3_Click()
    'Here NTSUP7 is my machine and MSDTC is a Service
    vbPauseService MachineName, ServiceName
End Sub


Private Sub Command4_Click()
    ServiceStatus MachineName, ServiceName
End Sub


Private Sub Form_Load()

End Sub
