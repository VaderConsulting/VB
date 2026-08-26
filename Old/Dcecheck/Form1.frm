VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DCE Status"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Connect to..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE Integrated Logon - NOT CONFIGURED"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE Security Client - NOT CONFIGURED"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE DTS Client - NOT CONFIGURED"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE Directory Client - NOT CONFIGURED"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE RPC Service - NOT CONFIGURED"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5775
   End
   Begin VB.Label lblConfiguration 
      Alignment       =   2  'Center
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   5655
   End
   Begin VB.Label lblProcesses 
      Alignment       =   2  'Center
      Caption         =   "Processes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5760
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   5760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5760
      Y1              =   1460
      Y2              =   1460
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   5760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE DTS Client - State UNKNOWN"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE Directory Client -State UNKNOWN"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblDCE 
      Caption         =   "DCE RPC Service - State UNKNOWN"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 1997-1999 Microsoft Corporation

Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Dim services As SWbemServices
Dim Hostname As String

Private Sub Command1_Click()
    MousePointer = vbHourglass
    Command1.Enabled = False
    ' Perform the asynchronous enumeration of processes
    lblStatus.Caption = "Enumerating processes..."
    lblStatus.Refresh
    services.InstancesOfAsync sink, "Win32_process"
End Sub

Private Sub Command2_Click()
  Unload frmMain
  Load frmMain
  'Set services = Nothing
  'sink.Cancel
  'set sink = Nothing
 '
 ' Command2.Enabled = False
 ' ConnectToHost
End Sub

Private Sub Form_Load()
    Me.Show
    Me.Refresh
    ConnectToHost
    Command1.Enabled = True
    Command1_Click
    lblStatus = "Idle"
    lblStatus.Refresh
End Sub

Sub ConnectToHost()
  Hostname = InputBox("Enter computer name", "Input required", Environ$("computername"))
  If Left(Hostname, 2) <> "\\" Then Hostname = "\\" & UCase(Hostname)
  ' Create a sink to receive the results of the enumeration
  Set sink = New SWbemSink
  lblStatus = "Connecting to " & Hostname & "..."
  lblStatus.Refresh
  ' Connect to root\cimv2.
  'On Error Resume Next
  Set services = GetObject("winmgmts:" & Hostname)
  If Err.Number <> 0 Then
    ' 462 = The remote server machine does not exist or is unavailable
    Err.Clear
    lblStatus = "Error contacting " & Hostname
    lblStatus.Refresh
    Command1.Enabled = False
    Set sink = Nothing
    Set services = Nothing
    Command2.Enabled = True
  End If
  On Error GoTo 0
End Sub

Private Sub sink_OnCompleted(ByVal iHResult As WbemScripting.WbemErrorEnum, ByVal objWbemErrorObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
    ' This event handler is called when there are no more instances to
    ' be returned
    MousePointer = vbDefault
    Command1.Enabled = True
    
    If (iHResult <> wbemNoErr) Then
        lblStatus.Caption = "Error: " & Err.Description & " [0x" & Hex(Err.Number) & "]"
    End If
    lblStatus = "Idle"
    lblStatus.Refresh
End Sub

Private Sub sink_OnObjectReady(ByVal Process As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
    ' This event handler is called once for every process returned by the
    ' enumeration
    
    Key = "Handle:" & Process.Handle
    
    ' Determine state of services - via process list
    Select Case UCase(Process.Name)
      Case "DCED.EXE"
        lblDCE(0).Caption = "DCE RPC Service is Running"
      Case "CDSADV.EXE"
        lblDCE(1).Caption = "DCE Directory Client is Running"
      Case "DTSD.EXE"
        lblDCE(2).Caption = "DCE DTS Client is Running"
    End Select
    
    ' Determine configuration state - via text file
    If Dir("c:\temp\dcestatus.txt") <> "" Then
      Kill "c:\temp\dcestatus.txt"
    End If
    
    If Dir(Hostname & "\c$\program files\dce\dcelocal\etc\cfg.dat") <> "" Then
      FileCopy Hostname & "\c$\program files\dce\dcelocal\etc\cfg.dat", "c:\temp\dcestatus.txt"
      Open "c:\temp\dcestatus.txt" For Input As #1
        Do Until EOF(1)
          Line Input #1, DCEData
          If DCEData <> "" Then
            Status = Right(DCEData, 1)
            Select Case UCase(Left(DCEData, 4))
              Case "DCED"
                If UCase(Left(DCEData, 8)) = "DCED:RPC" Then
                  If Status = "1" Then
                    RPCConfig = "Partial"
                  End If
                  If Status = "2" Then
                    RPCConfig = "Configured"
                  End If
                  lblDCE(3) = "DCE RPC Service - " & RPCConfig
                End If
                If UCase(Left(DCEData, 8)) = "DCED:SEC" Then
                  If Status = "1" Then
                    SecConfig = "Partial"
                  End If
                  If Status = "2" Then
                    SecConfig = "Configured"
                  End If
                  lblDCE(6) = "DCE Security Client - " & RPCConfig
                End If
              Case "CDSA"
                If Status = "1" Then
                  CDSConfig = "Partial"
                End If
                If Status = "2" Then
                  CDSConfig = "Configured"
                End If
                lblDCE(4) = "DCE Directory Client - " & CDSConfig
              Case "DTSD"
                If Status = "1" Then
                  DTSConfig = "Partial"
                End If
                If Status = "2" Then
                  DTSConfig = "Configured"
                End If
                lblDCE(5) = "DCE DTS Client - " & DTSConfig
              Case "DCEC"
                If Status = "1" Then
                  DCEConfig = "Partial"
                End If
                If Status = "2" Then
                  DCEConfig = "Configured"
                End If
                lblDCE(7) = "DCE Integrated Logon - " & DCEConfig
            End Select
          End If
        Loop
      Close 1
    End If
    Command2.Enabled = True
End Sub
