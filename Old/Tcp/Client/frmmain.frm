VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "systray.ocx"
Begin VB.Form frmMain 
   Caption         =   "Personal Alerter"
   ClientHeight    =   5040
   ClientLeft      =   2955
   ClientTop       =   3180
   ClientWidth     =   6930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "PersonalAlerter"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   1  'CenterOwner
   Begin SysTrayCtl.cSysTray Tray 
      Left            =   2280
      Top             =   5160
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":030A
      TrayTip         =   "Personal Alerter"
   End
   Begin TabDlg.SSTab tabEvents 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4048
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Severity"
      TabPicture(0)   =   "frmMain.frx":0764
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstCriteria"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Group"
      TabPicture(1)   =   "frmMain.frx":0780
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstGroup"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "System"
      TabPicture(2)   =   "frmMain.frx":079C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstSystem"
      Tab(2).ControlCount=   1
      Begin VB.ListBox lstSystem 
         Height          =   1620
         ItemData        =   "frmMain.frx":07B8
         Left            =   -74880
         List            =   "frmMain.frx":07BA
         TabIndex        =   11
         Top             =   480
         Width           =   2655
      End
      Begin VB.ListBox lstGroup 
         Height          =   1620
         ItemData        =   "frmMain.frx":07BC
         Left            =   -74880
         List            =   "frmMain.frx":07BE
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.ListBox lstCriteria 
         Height          =   1620
         ItemData        =   "frmMain.frx":07C0
         Left            =   120
         List            =   "frmMain.frx":07C2
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   1680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "None"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "All"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdAck 
      Caption         =   "Acknowledge"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ListBox lstDetails 
      Height          =   1860
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.ListBox lstHistory 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   6735
   End
   Begin MSWinsockLib.Winsock tcpMain 
      Left            =   600
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   120
      Top             =   5160
   End
   Begin VB.Label lblQuiet 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quiet mode ON"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   1200
      Y1              =   1.333
      Y2              =   1.333
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   0
   End
   Begin AgentObjectsCtl.Agent myAgent 
      Left            =   1080
      Top             =   5160
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label lblDetails 
      Alignment       =   2  'Center
      Caption         =   "Event Details"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblHistory 
      Alignment       =   2  'Center
      Caption         =   "Recent Events (last 50)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   6735
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAgent 
         Caption         =   "&Agent"
         Begin VB.Menu mnuLoad 
            Caption         =   "&Choose"
            Shortcut        =   ^C
         End
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Character As IAgentCtlCharacterEx
Public NewBalloonStyleOption As Integer
Public CharLoaded As Boolean
Public IgnoreSizeEvent As Boolean
Dim CurrentIndex As Integer

Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8

Private Sub cmdAck_Click()
  Dim lp As Integer, SQL As String, EventID As String
  Dim adoRS As New ADODB.Recordset
  Set adoConn = New ADODB.Connection
  adoConn.Open DSN
  For lp = 0 To lstDetails.ListCount - 1
    If lstDetails.Selected(lp) = True Then
      If Left(lstDetails.List(lp), 1) <> "[" Then
        EventID = Left(lstDetails.List(lp), InStr(1, lstDetails.List(lp), "|") - 1)
        SQL = "UPDATE tblAlerts SET complete=1 WHERE ID = " & CLng(EventID)
        adoConn.Execute SQL
      End If
    End If
  Next lp
  adoConn.Close
  lstDetails.Clear
  tmrMain_Timer
End Sub

Private Sub cmdSelectAll_Click()
  Dim lp As Integer
  For lp = 0 To lstDetails.ListCount - 1
    lstDetails.Selected(lp) = True
  Next lp
  If lstDetails.ListCount <> 0 Then cmdAck.Enabled = True
End Sub

Private Sub cmdSelectNone_Click()
  Dim lp As Integer
  For lp = 0 To lstDetails.ListCount - 1
    lstDetails.Selected(lp) = False
  Next lp
  If lstDetails.ListCount <> 0 Then cmdAck.Enabled = True
End Sub

Private Sub Form_Load()
  Dim DoStart As Boolean
  
  'Defaults
  DoStart = False
  LogSeverity = 0
  
  Me.Hide
  Me.Refresh
  
  WriteToLog 0, "---------------------------"
  WriteToLog 0, App.EXEName & " Log Opened"
   
  If Command$ = "" Then
    DoStart = True
    LogSeverity = 4
  End If
   
  'Parse Command Line
  If (InStr(UCase(Command$), "/START")) Then DoStart = True
  If (InStr(UCase(Command$), "/LOG:1")) Then LogSeverity = 1
  If (InStr(UCase(Command$), "/LOG:2")) Then LogSeverity = 2
  If (InStr(UCase(Command$), "/LOG:3")) Then LogSeverity = 3
  If (InStr(UCase(Command$), "/LOG:4")) Then LogSeverity = 4
  
  If Not InitObjects Then
    MsgBox "Failed to connect to " & MasterHost & ".", vbCritical + vbApplicationModal + vbOKOnly
    Unload Me
    Exit Sub
  End If
  
  If DoStart Then StartDistroTasks
  tmrMain.Enabled = True
  '----------------------------------------------------------
  '-- When the form loads, set the IgnoreSizeEvent flag
  '-- (used to differentiate when the Character Animation
  '-- Previewer window is restored), set the CharLoaded flag
  '-- (used to track when a character is loaded),
  '-- and set the initial state of the status bar.
  '----------------------------------------------------------
  IgnoreSizeEvent = True
  
  Set oVoice = New VTxtAuto.VTxtAuto
  oVoice.Register App.Title, App.EXEName
  
  CharLoaded = False
  If Not boolFirstRun Then
    If boolAgent = True Then
      GetAgent strSelectedAgent
      Speak "Personal Alerter started in Agent mode"
      frmMain.Character.Hide
      frmMain.Character.MoveTo (Screen.Width / Screen.TwipsPerPixelX) - frmMain.Character.Width, (Screen.Height / Screen.TwipsPerPixelY) - frmMain.Character.Height
    Else
      Speak "Personal Alerter started in Speech API mode"
    End If
  Else
    frmFirstRun.Show
  End If
End Sub

Public Sub GetAgent(CharacterName As String)
  Dim OpenSuccess As Boolean
  '--Unload the previous character
  On Error Resume Next
  Set Character = Nothing
  myAgent.Characters.Unload "CharacterID"
      
  '-- Load the new character
  On Error GoTo ErrHandler
  myAgent.Characters.Load "CharacterID", CharacterName 'CommonDialog1.FileName
  
  OpenSuccess = True
  
  Set Character = myAgent.Characters("CharacterID")
  
  'frmMain.Caption = Character.Name + " - Microsoft Character Animation Previewer"
  '-- Set the character loaded flag
  CharLoaded = True
  
  '-- Set the character's language
  Character.LanguageID = &H409
  
  '-- Update the caption for the animation list box
  'AnimationFrame.Caption = "&Animations for " + Character.Name
  
  '-- Disable the Play button to avoid trying to play a null animation selection
  'Command1(0).Enabled = False
  
  '-- Load the character's animation into the list box
  'AnimationListBox.Clear
  'For Each AnimationName In Character.AnimationNames
    'AnimationListBox.AddItem AnimationName
  'Next
     
  '-- Move the character to starting position
  Character.Left = ((Screen.Width / Screen.TwipsPerPixelX) / 2) - (frmMain.Character.Width / 2)
  Character.Top = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - (frmMain.Character.Height / 2)
  
  '-- Show the character
  Character.Show
  
  '-- Update the X,Y position fields with the character's
  '-- current position
  'CharPosn(0).Text = CStr(Character.Left)
  'CharPosn(1).Text = CStr(Character.Top)
  
  '-- Update the state of the balloon style options
  SetBalloonStyleOptions
  
  '-- Initialize the pop-up menu commands
  InitPopupMenuCmds
     
  '-- Update the state of the controls to match the
  '-- character's settings
  'EnableControls
  'AnimationListBox.SetFocus
  Exit Sub
      
ErrHandler:
  If (Err.Number <> cdlCancel) Then
    If (OpenSuccess = False) Then
      MsgBox "There was an error opening the file " ' & CommonDialog1.FileName
    End If
    Set Character = Nothing
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  WriteToLog 0, App.EXEName & " Log Closed"
  CloseObjects
End Sub

Private Sub lstCriteria_Click()
  Dim SQL As String
  Dim adoRS As New ADODB.Recordset
  Select Case Left(lstCriteria.List(lstCriteria.ListIndex), 3)
    Case "Inf"
      SQL = "SELECT * FROM tblAlerts WHERE Complete=false AND Severity=1 ORDER BY DateTime DESC"
    Case "War"
      SQL = "SELECT * FROM tblAlerts WHERE Complete=false AND Severity=2 ORDER BY DateTime DESC"
    Case "Cri"
      SQL = "SELECT * FROM tblAlerts WHERE Complete=false AND Severity=3 ORDER BY DateTime DESC"
    Case Else
      Exit Sub
  End Select
  adoRS.Open SQL, DSN
  lstDetails.Clear
  Do Until adoRS.EOF
    lstDetails.AddItem adoRS("ID") & "|" & adoRS("DateTime") & "|" & adoRS("Severity") & "|" & adoRS("Group") & "|" & adoRS("System") & "|" & adoRS("Alert") & "|" & adoRS("Hint")
    adoRS.MoveNext
  Loop
  adoRS.Close
  cmdAck.Enabled = False
End Sub

Private Sub lstDetails_Click()
  If lstDetails.ListIndex <> -1 Then
    cmdAck.Enabled = True
  Else
    cmdAck.Enabled = False
  End If
End Sub

Private Sub lstGroup_Click()
  Dim SQL As String, intTemp As Integer, strTemp As String
  Dim adoRS As New ADODB.Recordset
  If lstGroup.List(lstGroup.ListIndex) = "" Then Exit Sub
  Select Case Left(lstGroup.List(lstGroup.ListIndex), 3)
    Case "NTS"
      intTemp = 1
    Case "Hel"
      intTemp = 2
    Case "EUC"
      intTemp = 4
    Case "WAN"
      intTemp = 8
  End Select
  SQL = "SELECT * FROM tblAlerts WHERE Complete=false ORDER BY DateTime DESC"
  Debug.Print SQL
  adoRS.Open SQL, DSN
  lstDetails.Clear
  Do Until adoRS.EOF
    If (adoRS("Group") And intTemp) <> 0 Then
      lstDetails.AddItem adoRS("ID") & "|" & adoRS("DateTime") & "|" & adoRS("Severity") & "|" & adoRS("Group") & "|" & adoRS("System") & "|" & adoRS("Alert") & "|" & adoRS("Hint")
      Debug.Print "Group: " & adoRS("Group")
    End If
    adoRS.MoveNext
  Loop
  adoRS.Close
  cmdAck.Enabled = False
End Sub

Private Sub lstSystem_Click()
  Dim SQL As String, intTemp As Integer, strTemp As String
  Dim adoRS As New ADODB.Recordset
  If lstSystem.List(lstSystem.ListIndex) = "" Then Exit Sub
  strTemp = Left(lstSystem.List(lstSystem.ListIndex), InStr(1, lstSystem.List(lstSystem.ListIndex), "(") - 2)
  SQL = "SELECT * FROM tblAlerts WHERE Complete=false AND System=" & Chr(34) & strTemp & Chr(34) & " ORDER BY DateTime DESC"
  adoRS.Open SQL, DSN
  lstDetails.Clear
  Do Until adoRS.EOF
    lstDetails.AddItem adoRS("ID") & "|" & adoRS("DateTime") & "|" & adoRS("Severity") & "|" & adoRS("Group") & "|" & adoRS("System") & "|" & adoRS("Alert") & "|" & adoRS("Hint")
    adoRS.MoveNext
  Loop
  adoRS.Close
  cmdAck.Enabled = False
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHide_Click()
  Me.Visible = False
  If frmMain.CharLoaded = True And boolAgent = True Then
    frmMain.Character.Hide
  End If
  Tray.InTray = True
End Sub

Private Sub mnuLoad_Click()
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
  strSelectedAgent = strFilename(UBound(strFilename()))
  If strSelectedAgent = "" Then strSelectedAgent = cdlOpen.Filename
  frmMain.GetAgent strSelectedAgent
  SetValueString "", "Software\Alerter", "Character", strSelectedAgent, REG_SZ
  Exit Sub
ErrHandler:
  If (Err.Number <> cdlCancel) Then
    If (OpenSuccess = False) Then
      MsgBox "There was an error opening the file " ' & CommonDialog1.FileName
    End If
    Set Character = Nothing
  End If
End Sub

Private Sub mnuOptions_Click()
  frmOptions.Show vbModal
End Sub

Private Sub tcpMain_Close()
  DoEvents
  SetDistroTasksState
  Debug.Print "Connection to " & MasterHost & " closed"
  MsgBox "Connection to " & MasterHost & " closed", vbCritical + vbOKOnly + vbApplicationModal
  Unload frmMain
End Sub

Private Sub tcpMain_DataArrival(ByVal BytesTotal As Long)
  TCPDataArrival BytesTotal
End Sub

Private Sub tcpMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  WriteToLog 0, "TCP Error (" & Number & ") " & Description
  CancelDisplay = True
End Sub

Private Sub tmrMain_Timer()
  Dim SQL As String, TotalRecords As Long
  Dim adoRS As New ADODB.Recordset, GroupText As String
  Dim GroupArray(4) As Integer
  frmMain.lblStatus = "Checking local database for events"
  frmMain.lblStatus.Refresh
  tmrMain.Enabled = False
  Set adoConn = New ADODB.Connection
  If boolQuiet = True Then
    If Time > CDate("7:00 PM") Then
      lblQuiet.Visible = True
    Else
      If Time < CDate("6:00 AM") Then
        lblQuiet.Visible = True
      Else
        lblQuiet.Visible = False
      End If
    End If
  Else
    lblQuiet.Visible = False
  End If
    Set adoRS = New ADODB.Recordset
      SQL = "SELECT * FROM tblAlerts WHERE Complete = false"
      adoConn.Open DSN
        adoRS.Open SQL, adoConn
          TotalRecords = 0
          Do Until adoRS.EOF
            TotalRecords = TotalRecords + 1
            adoRS.MoveNext
          Loop
        adoRS.Close
        lstCriteria.Clear
        SQL = "SELECT Count (*) as Total FROM tblAlerts WHERE Complete=false AND Severity=3"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          lstCriteria.AddItem "Critical Alerts - " & adoRS("Total")
          lstCriteria.Refresh
        End If
        adoRS.Close
        SQL = "SELECT Count (*) as total FROM tblAlerts WHERE Complete=false AND Severity=2"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          lstCriteria.AddItem "Warning Alerts - " & adoRS("Total")
          lstCriteria.Refresh
        End If
        adoRS.Close
        SQL = "SELECT Count (*) as total FROM tblAlerts WHERE Complete=false AND Severity=1"
        adoRS.Open SQL, DSN
        If adoRS("Total") > 0 Then
          lstCriteria.AddItem "Information Alerts - " & adoRS("Total")
          lstCriteria.Refresh
        End If
        adoRS.Close
        SQL = "SELECT System, Count(System) AS Total From tblAlerts WHERE Complete=false GROUP BY System ORDER BY System ASC"
        adoRS.Open SQL, DSN
        lstSystem.Clear
        Do Until adoRS.EOF
          lstSystem.AddItem adoRS("System") & " (" & adoRS("Total") & ")"
          lstSystem.Refresh
          adoRS.MoveNext
        Loop
        adoRS.Close
        SQL = "SELECT [Group] From tblAlerts WHERE Complete=false "
        adoRS.Open SQL, DSN
        lstGroup.Clear
        Do Until adoRS.EOF
          If IsNumeric(adoRS("Group")) Then
            If (CInt(adoRS("group")) And 1) <> 0 Then
              GroupArray(1) = GroupArray(1) + 1
            End If
            If (CInt(adoRS("group")) And 2) <> 0 Then
              GroupArray(2) = GroupArray(2) + 1
            End If
            If (CInt(adoRS("group")) And 4) <> 0 Then
              GroupArray(3) = GroupArray(3) + 1
            End If
            If (CInt(adoRS("group")) And 8) <> 0 Then
              GroupArray(4) = GroupArray(4) + 1
            End If
          Else
            Select Case adoRS("Group")
              Case "NTSS"
                GroupArray(1) = GroupArray(1) + 1
              Case "Helpdesk"
                GroupArray(2) = GroupArray(2) + 1
              Case "EUC"
                GroupArray(3) = GroupArray(3) + 1
              Case "WAN"
                GroupArray(4) = GroupArray(4) + 1
            End Select
            'lstGroup.AddItem adoRS("Group") & " (" & adoRS("Total") & ")"
          End If
          lstGroup.Refresh
          adoRS.MoveNext
        Loop
        If GroupArray(1) <> 0 Then
          lstGroup.AddItem "NTSS (" & GroupArray(1) & ")"
        End If
        If GroupArray(2) <> 0 Then
          lstGroup.AddItem "Helpdesk (" & GroupArray(2) & ")"
        End If
        If GroupArray(3) <> 0 Then
          lstGroup.AddItem "EUC (" & GroupArray(3) & ")"
        End If
        If GroupArray(4) <> 0 Then
          lstGroup.AddItem "WAN (" & GroupArray(4) & ")"
        End If
        adoRS.Close
      adoConn.Close
    Set adoRS = Nothing
  frmMain.Caption = App.ProductName & " (" & TotalRecords & " total events pending)"
  frmMain.Refresh
  tmrMain.Enabled = True
  frmMain.lblStatus = "Idle"
  frmMain.lblStatus.Refresh
End Sub

Private Sub MyAgent_AgentPropertyChange()
'-----------------------------------------------------
'-- Check to see if the user changed settings in
'-- Advanced Character Options
'-----------------------------------------------------

'-- Check to see if the user changed
'-- Play Character Sound Effects
'-- in Advanced Character Options
If Not myAgent.AudioOutput.SoundEffects Then
    'OutputStyleOption(0).Enabled = False
Else
    'OutputStyleOption(0).Enabled = True
End If


'-- Check to see if the user changed
'-- Display Spoken Output In Word Balloon option
'-- in Advanced Character Options
If Not Character.Balloon.Enabled Then
    'BalloonStyleOption(0).Enabled = False
    'BalloonStyleOption(1).Enabled = False
    'BalloonStyleOption(2).Enabled = False
    'BalloonStyleOption(3).Enabled = False
    
Else
    'BalloonStyleOption(0).Enabled = True
    'BalloonStyleOption(1).Enabled = True
    'BalloonStyleOption(2).Enabled = True
    'BalloonStyleOption(3).Enabled = True
    
End If

End Sub

Private Sub MyAgent_Command(ByVal UserInput As Object)
'-----------------------------------------------------
'-- If the user selects the Advanced Character Options
'-- command in the character's pop-up menu
'-- make the window visible
'-----------------------------------------------------

If UserInput.Name = "AdvCharOptions" Then
    myAgent.PropertySheet.Visible = True
End If

End Sub

Private Sub MyAgent_DragComplete(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
'-----------------------------------------------------
'-- If the user drags the character
'-- update the character position fields
'-----------------------------------------------------

'CharPosn(0) = Character.Left
'CharPosn(1) = Character.Top

End Sub

Sub InitPopupMenuCmds()
'-----------------------------------------------------
'-- Add a command to the character to provide access
'-- to the Advanced Character Options
'-----------------------------------------------------
Character.Commands.RemoveAll
Character.Commands.Add "AdvCharOptions", "&Advanced Character Options"

End Sub

Private Sub Form_Resize()
  '-------------------------------------------------------
  '-- This routines hides or shows the character when the
  '-- Character Animation Previewer window is mininized or
  '-- restored
  
  'If IgnoreSizeEvent Then
  '    IgnoreSizeEvent = False
  '    Exit Sub
  'End If
  '
  'If CharLoaded Then
  '    If frmMain.WindowState = vbMinimized Then
  '        Character.Hide True
  '    ElseIf frmMain.WindowState = vbNormal Then
  '        Character.Show True
  '    End If
  'End If
  If frmMain.CharLoaded = True Then
    frmMain.Character.Show
    frmMain.SetFocus
  End If
  lstHistory.Width = (frmMain.Width - (lstHistory.Left + lstHistory.Left))
  lstDetails.Width = (frmMain.Width - (lstDetails.Left + lstDetails.Left))
  If frmMain.Width < 7050 Then frmMain.Width = 7050
  If frmMain.Height <> 5730 Then frmMain.Height = 5730
End Sub

Sub SetBalloonStyleOptions()
'------------------------------------------------
'-- This subroutine sets the check boxes for the
'-- the word balloon settings
'------------------------------------------------

'-- Check to see if the balloon is on

If Character.Balloon.Style And BalloonOn Then
    'BalloonStyleOption(0).Value = 1
Else
    'BalloonStyleOption(0).Value = 0
End If

'-- Check to see if Auto-Hide is on

If Character.Balloon.Style And AutoHide Then
    'BalloonStyleOption(1).Value = 1
Else
    'BalloonStyleOption(1).Value = 0
End If

'-- Check to see if Auto-Pace is on

If Character.Balloon.Style And AutoPace Then
    'BalloonStyleOption(2).Value = 1
Else
    'BalloonStyleOption(2).Value = 0
End If

'-- Check to see if Size-To-Text is on

If Character.Balloon.Style And SizeToText Then
    'BalloonStyleOption(3).Value = 1
Else
    'BalloonStyleOption(3).Value = 0
End If


'-- Set the controls based on Advanced Character Options

If Not Character.Balloon.Enabled Then
    'BalloonStyleOption(0).Enabled = False
    'BalloonStyleOption(1).Enabled = False
    'BalloonStyleOption(2).Enabled = False
    'BalloonStyleOption(2).Enabled = False
Else
    'BalloonStyleOption(0).Enabled = True
    'BalloonStyleOption(1).Enabled = True
    'BalloonStyleOption(2).Enabled = True
    'BalloonStyleOption(2).Enabled = True
End If

End Sub

Private Sub Tray_MouseDblClick(Button As Integer, Id As Long)
  If frmMain.Visible = False Then
    frmMain.Visible = True
    If frmMain.CharLoaded = True And boolAgent = True Then
      frmMain.Character.Show
      frmMain.SetFocus
    End If
  Else
    frmMain.Visible = False
    If frmMain.CharLoaded = True And boolAgent = True Then
      frmMain.Character.Hide
      frmMain.SetFocus
    End If
  End If
End Sub

