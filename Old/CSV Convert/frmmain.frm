VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CSV Convert"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4365
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbrDirectory 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "H:\Data\New Sites"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "H:\Data\Sites"
      Top             =   120
      Width           =   2295
   End
   Begin VB.FileListBox filCSV 
      Height          =   675
      Left            =   4200
      Pattern         =   "*.csv"
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Ready"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lblDestination 
      Caption         =   "Destination Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblSource 
      Caption         =   "Source Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    filCSV.Path = "H:\Data\Sites"
End Sub

Private Sub mnuConvert_Click()
    lblStatus.Caption = "Converting..."
    lblStatus.Refresh
    pbrDirectory.Max = filCSV.ListCount - 1
    For lp = 0 To filCSV.ListCount - 1
        FileName = filCSV.Path & "\" & filCSV.List(lp)
        newFileName = "H:\data\New Sites\" & filCSV.List(lp)
        Open FileName For Input As #1
            Open newFileName For Output As #2
                Line Input #1, stuff
                Print #2, stuff & ",Drive6(U),Drive7(V),Drive6(W)"
                While Not EOF(1)
                    Line Input #1, stuff
                    Print #2, stuff & ",na,na,na"
                Wend
            Close 2
        Close 1
        pbrDirectory.Value = lp
        pbrDirectory.Refresh
    Next lp
    lblStatus.Caption = "Ready"
End Sub

Private Sub mnuExit_Click()
    End
End Sub
