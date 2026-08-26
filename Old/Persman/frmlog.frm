VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logfile"
   ClientHeight    =   7520
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   8304
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7520
   ScaleWidth      =   8304
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtfLog 
      Height          =   6928
      Left            =   128
      TabIndex        =   1
      Top             =   128
      Width           =   8080
      _ExtentX        =   14576
      _ExtentY        =   12498
      _Version        =   327681
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmLog.frx":0442
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   7552
      TabIndex        =   0
      Top             =   7040
      Width           =   656
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    rtfLog.filename = gLogFile
End Sub

