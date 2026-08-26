VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "Ntsvc.ocx"
Begin VB.Form ServiceMain 
   Caption         =   "Sample NT Service"
   ClientHeight    =   1125
   ClientLeft      =   1395
   ClientTop       =   1620
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1125
   ScaleWidth      =   2970
   Begin VB.Timer Timer 
      Left            =   720
      Top             =   120
   End
   Begin NTService.NTService NTService1 
      Left            =   120
      Top             =   120
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "Sample Service"
      ServiceName     =   "Sample"
      StartMode       =   3
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   825
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ServiceMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

On Error GoTo Err_Load
    Dim strDisplayName As String
    Dim bStarted As Boolean
    
    strDisplayName = NTService1.DisplayName
    
    StatusBar.Panels(1).Text = "Loading"
    
    If Command = "-install" Then
        ' enable interaction with desktop
        NTService1.Interactive = True
        
        If NTService1.Install Then
            Call NTService1.SaveSetting("Parameters", "TimerInterval", "1000")
            MsgBox strDisplayName & " installed successfully"
        Else
            MsgBox strDisplayName & " failed to install"
        End If
        End
    ElseIf Command = "-uninstall" Then
        If NTService1.Uninstall Then
            MsgBox strDisplayName & " uninstalled successfully"
        Else
            MsgBox strDisplayName & " failed to uninstall"
        End If
        End
    ElseIf Command = "-debug" Then
        NTService1.Debug = True
    ElseIf Command <> "" Then
        MsgBox "Invalid command option"
        End
    End If
    
    Timer.Interval = CInt(NTService1.GetSetting("Parameters", "TimerInterval", "2000"))
    
    ' enable Pause/Continue. Must be set before StartService
    ' is called or in design mode
    NTService1.ControlsAccepted = svcCtrlPauseContinue
    
    ' connect service to Windows NT services controller
    NTService1.StartService
    
Err_Load:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


Private Sub NTService1_Continue(Success As Boolean)
On Error GoTo Err_Continue
    
    Timer.Enabled = True
    StatusBar.Panels(1).Text = "Running"
    Success = True
    Call NTService1.LogEvent(svcEventInformation, svcMessageInfo, "Service continued")
    
Err_Continue:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub

Private Sub NTService1_Control(ByVal EventNo As Long)

On Error GoTo Err_Control
    
    StatusBar.SimpleText = NTService1.DisplayName & " Control signal " & EventNo

Err_Control:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)

End Sub

Private Sub NTService1_Pause(Success As Boolean)
On Error GoTo Err_Pause
    
    Timer.Enabled = False
    StatusBar.Panels(1).Text = "Paused"
    Call NTService1.LogEvent(svcEventError, svcMessageError, "Service paused")
    Success = True
    
Err_Pause:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


Private Sub NTService1_Start(Success As Boolean)
On Error GoTo Err_Start
    
    StatusBar.Panels(1).Text = "Running"
    Success = True
    
Err_Start:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


Private Sub NTService1_Stop()
On Error GoTo Err_Stop
    
    StatusBar.Panels(1).Text = "Stopped"
    Unload Me
    
Err_Stop:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


Private Sub Timer_Timer()
On Error GoTo Err_Timer
    StatusBar.Panels(2).Text = Format(Now(), "hh:mm:ss")

Err_Timer:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub


