VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Helper"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   1680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrService 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   0
   End
   Begin NTService.NTService SCM 
      Left            =   120
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Interactive     =   -1  'True
      ServiceName     =   "Simple"
      StartMode       =   2
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  NOTE
' To install this application as a service, use srvinstw.exe from the Resource Kit.
'

' -------------------------------------------------------------------------------------------------------
' Returns a number indicating the ordinal number of the supplied date that is older
' Returns one of the following
'         -1 = Date1 not specified
'         -2 = Date2 not specified
'         -9 = Error
'          0 = Same
'          1 = Date1 is older
'          2 = Date2 is older
Private Function DateOlder(strDate1 As String, strDate2 As String, intOlder As Integer)
    Dim datDate1 As Date
    Dim datDate2 As Date
        
    If strDate1 = "" Then
        intOlder = -1
        Exit Function
    End If
    
    If strDate2 = "" Then
        intOlder = -2
        Exit Function
    End If
    
    datDate1 = Format(CDate(strDate1), "DD/MM/YYYY HH:MM:SS AM/PM")
    datDate2 = Format(CDate(strDate2), "DD/MM/YYYY HH:MM:SS AM/PM")
    
    Select Case DateDiff("n", datDate1, datDate2)
        Case 0
            intOlder = 0 ' Dates are the same
        Case Is > 0
            intOlder = 1 ' Date 1 is older
        Case Is < 0
            intOlder = 2 ' Date 2 is older
    End Select
    
End Function

' -------------------------------------------------------------------------------------------------------
Private Sub DisableTimer()
    tmrService.Enabled = False
End Sub

Private Sub EnableTimer()
    tmrService.Enabled = True
End Sub

Private Sub Form_Load()
    Set oDictionary = CreateObject("Scripting.Dictionary")
    ' Gnerate an App_Path variable that is the path to this application, similar to App.Path, though it will always end with a backslash.
    If Right(App.Path, 1) <> "\" Then
        strApp_Path = App.Path & "\"
    Else
        strApp_Path = App.Path
    End If
    
    With frmMain.SCM
        .ControlsAccepted = svcCtrlPauseContinue
        .DisplayName = "Installation Helper"
        .Interactive = False
        .ServiceName = "InstallationHelper"
        .Debug = False
    End With
    
   frmMain.SCM.StartService
    EnableTimer
    
    ' Run the standard subroutines when starting up.
    OpenConfig oDictionary
    
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub Pause(dblSeconds As Double)
    Dim d As Date
    d = Now
    Do Until Now > DateAdd("s", dblSeconds, d)
        DoEvents
    Loop
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub SCM_Continue(Success As Boolean)
    tmrService.Enabled = True
    Success = True
   frmMain.SCM.LogEvent svcEventInformation, 4, frmMain.SCM.DisplayName & " has continued after a pause."
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub SCM_Pause(Success As Boolean)
    tmrService.Enabled = False
    Success = True
   frmMain.SCM.LogEvent svcEventInformation, 3, frmMain.SCM.DisplayName & " has paused."
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub SCM_Start(Success As Boolean)
    EnableTimer
   frmMain.SCM.LogEvent svcEventInformation, 1, frmMain.SCM.DisplayName & " has started."
    Success = True
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub SCM_Stop()
   frmMain.SCM.LogEvent svcEventInformation, 2, frmMain.SCM.DisplayName & " has stopped."
    End
End Sub

' -------------------------------------------------------------------------------------------------------
Private Sub tmrService_Timer()
    DisableTimer
    
    EnableTimer
End Sub

