VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFirstRun 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup for first run..."
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2070
   Icon            =   "frmFirstRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   2070
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fmeSpeechMode 
      Caption         =   "Speech Mode"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optSpeechMode 
         Caption         =   "Dummy option control"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optSpeechMode 
         Caption         =   "Microsoft Agent"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame fmeAgent 
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         Begin VB.CommandButton cmdBrowseforAgent 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtAgent 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.OptionButton optSpeechMode 
         Caption         =   "MS Speech API"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmFirstRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseforAgent_Click()
  Dim OpenSuccess As Boolean
  Dim DirName As String, strFilename() As String

  '-- Set a flag to track success
  OpenSuccess = False
          
  cdlOpen.CancelError = True
  
  On Error GoTo ErrHandler
      
  cdlOpen.Flags = cdlOFNHideReadOnly
  
  '-- Get the Windows directory name
  DirName = GetWindowsDir()
  
  '-- Append the Agent Chars subdirectory
  cdlOpen.InitDir = DirName + "msagent\chars"
  
  '-- Add the filter
  cdlOpen.Filter = "Microsoft Agent Characters (*.acs)|*.acs"
  cdlOpen.FilterIndex = 1
  
  '-- Show the Open dialog
  cdlOpen.ShowOpen
  
  strFilename() = Split(cdlOpen.Filename, "\")
  txtAgent = strFilename(UBound(strFilename()))
  If txtAgent = "" Then txtAgent = cdlOpen.Filename
  frmMain.GetAgent txtAgent
  SetValueString "", "Software\Alerter", "Character", txtAgent, REG_SZ
  Exit Sub
ErrHandler:
  If (Err.Number <> cdlCancel) Then
    If (OpenSuccess = False) Then
      MsgBox "There was an error opening the file " ' & CommonDialog1.FileName
    End If
    Set Character = Nothing
  End If
End Sub

Private Sub cmdOK_Click()
  SetValueString "", "Software\Alerter", "First Run", 0, REG_SZ
  If boolAgent Then frmMain.Character.MoveTo (Screen.Width / Screen.TwipsPerPixelX) - frmMain.Character.Width, (Screen.Height / Screen.TwipsPerPixelY) - frmMain.Character.Height
  boolFirstRun = False
  Unload Me
End Sub

Private Sub Form_Load()
  txtAgent = strSelectedAgent
  Me.Show
  Me.Refresh
  MsgBox "Welcome to Personal Alerter.", vbOKOnly + vbInformation, "Hello"
  
End Sub

Private Sub optSpeechMode_Click(Index As Integer)
  Dim CharX As Long
  Dim CharY As Long
  Dim Host As String
  Host = ""
  Select Case Index
    Case 0
      txtAgent.Enabled = True
      frmMain.GetAgent strSelectedAgent
      frmMain.Character.Show
      CharX = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - (frmMain.Character.Width / 2)
      CharY = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - (frmMain.Character.Height / 2)
      CharX = CharX + (frmMain.Width / Screen.TwipsPerPixelX)
      CharY = CharY + (frmMain.Height / Screen.TwipsPerPixelY)
      frmMain.Character.MoveTo CharX, CharY

      frmMain.mnuLoad.Enabled = True
      cmdBrowseforAgent.Enabled = True
      boolAgent = True
      SetValueString Host, "Software\Alerter", "Agent", "1", REG_SZ
    Case 1
      txtAgent.Enabled = False
      If frmMain.CharLoaded Then
        frmMain.Character.Hide
        frmMain.mnuLoad.Enabled = False
      End If
      cmdBrowseforAgent.Enabled = False
      boolAgent = False
      SetValueString Host, "Software\Alerter", "Agent", "0", REG_SZ
  End Select
  cmdOK.Enabled = True
End Sub
