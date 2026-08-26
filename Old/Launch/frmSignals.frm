VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "richtx32.ocx"
Begin VB.Form frmSignals 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signals"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13545
   Icon            =   "frmSignals.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   13545
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox filSignals 
      Height          =   7695
      Left            =   120
      Pattern         =   "*.txt"
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin RichTextLib.RichTextBox rtbSignal 
      Height          =   7695
      Left            =   4800
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   13573
      _Version        =   327681
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmSignals.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblContents 
      Alignment       =   2  'Center
      Caption         =   "Contents"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label lblFilename 
      Alignment       =   2  'Center
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSignals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub filSignals_Click()
    rtbSignal.filename = filSignals.Path & "\" & filSignals.List(filSignals.ListIndex)
End Sub

Private Sub Form_Activate()
    filSignals.Path = gSignalsDir
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mnuExit_Click
End Sub

Private Sub mnuExit_Click()
    If gUserName <> "" Then ' check a generic variable
        frmOther.Show
        Unload Me
    Else
        End
    End If
End Sub

Private Sub mnuPreferences_Click()
    frmPreferences.Show vbModal
End Sub
