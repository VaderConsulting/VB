VERSION 5.00
Begin VB.Form frmserver 
   Caption         =   "Named Pipe Server"
   ClientHeight    =   2175
   ClientLeft      =   1125
   ClientTop       =   1515
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   4230
   Begin VB.CommandButton cmdAsync2 
      Caption         =   "Send Async2"
      Height          =   435
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton sndAsync 
      Caption         =   "Send Async"
      Height          =   435
      Left            =   2520
      TabIndex        =   3
      Top             =   1020
      Width           =   1455
   End
   Begin VB.CommandButton sndBlocked 
      Caption         =   "Send Blocked"
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2940
      Top             =   180
   End
   Begin VB.Label lblStat 
      Caption         =   "Idle"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   660
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1635
   End
End
Attribute VB_Name = "frmserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 1997 by Desaware Inc. All Rights Reserved

Dim ctr&
Dim pipehnd&
Dim buffer$
Dim ol As OVERLAPPED
Dim ol2 As OVERLAPPED
Dim callbackobj As Object

Private Sub cmdAsync2_Click()
    Dim res&
    Dim written&
    Set callbackobj = CreateObject("dwWatcher.dwSyncWatch")
    ' Wait for a client to appear
    cmdAsync2.Enabled = False
    ol2.hEvent = CreateEvent(0, True, False, vbNullString)
    If ol2.hEvent = 0 Then
        MsgBox "Can't create event"
        Exit Sub
    End If
    lblStat.Caption = "Waiting for client"
    lblStat.Refresh
    res = ConnectNamedPipe(pipehnd, 0)
    
    If res <> 0 Or (res = 0 And GetLastError() = ERROR_PIPE_CONNECTED) Then
        lblStat.Caption = "Sending data"
        lblStat.Refresh
        ' Pipe is connected
        buffer$ = String$(200, 0)
        res = WriteFileAsync(pipehnd, ByVal buffer, 200, written, ol2)
        ' Turn on the callback object
        Call callbackobj.SetAppCallback(Me)
        Call callbackobj.SetObjectWatch(GetCurrentProcessId(), ol2.hEvent)
    Else
        Call CloseHandle(ol2.hEvent)
        ol2.hEvent = 0
        MsgBox "Client has disconnected"
        lblStat.Caption = "Idle"
        cmdAsync2.Enabled = True
    End If

End Sub

Private Sub Form_Load()
    pipehnd& = CreateNamedPipe("\\.\pipe\vbpgpipe1", PIPE_ACCESS_OUTBOUND Or FILE_FLAG_OVERLAPPED, PIPE_TYPE_BYTE, 1, 0, 0, 0, 0)
    If pipehnd = 0 Then
        sndBlocked.Enabled = False
        sndAsync.Enabled = False
        cmdAsync2.Enabled = False
        lblStat.Caption = "Can't create pipe"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If pipehnd& <> 0 Then
        Call CloseHandle(pipehnd)
    End If
End Sub

Private Sub sndAsync_Click()
    Dim res&
    Dim written&
    ' Wait for a client to appear
    sndAsync.Enabled = False
    ol.hEvent = CreateEvent(0, True, False, vbNullString)
    If ol.hEvent = 0 Then
        MsgBox "Can't create event"
        Exit Sub
    End If
    lblStat.Caption = "Waiting for client"
    lblStat.Refresh
    res = ConnectNamedPipe(pipehnd, 0)
    
    If res <> 0 Or (res = 0 And GetLastError() = ERROR_PIPE_CONNECTED) Then
        lblStat.Caption = "Sending data"
        lblStat.Refresh
        ' Pipe is connected
        buffer$ = String$(200, 0)
        res = WriteFileAsync(pipehnd, ByVal buffer, 200, written, ol)
    Else
        Call CloseHandle(ol2.hEvent)
        ol2.hEvent = 0
        MsgBox "Client has disconnected"
        lblStat.Caption = "Idle"
        sndAsync.Enabled = True
    End If

End Sub

Private Sub sndBlocked_Click()
    Dim res&
    Dim written&
    ' Wait for a client to appear
    sndBlocked.Enabled = False
    lblStat.Caption = "Waiting for client"
    lblStat.Refresh
    res = ConnectNamedPipe(pipehnd, 0)
    If res <> 0 Or (res = 0 And GetLastError() = ERROR_PIPE_CONNECTED) Then
        lblStat.Caption = "Sending data"
        lblStat.Refresh
        ' Pipe is connected
        buffer$ = String$(200, 0)
        res = WriteFile(pipehnd, ByVal buffer, 200, written, 0)
    Else
        MsgBox "Client has disconnected"
    End If
    lblStat.Caption = "Idle"
    sndBlocked.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim res&
    Label1.Caption = ctr&
    ctr& = ctr& + 1
    If ol.hEvent <> 0 Then
        ' Async operation is in progress
        res = WaitForSingleObject(ol.hEvent, 0)
        If res = WAIT_OBJECT_0 Then
            ' Async read is done
            Call CloseHandle(ol.hEvent)
            ol.hEvent = 0
            lblStat.Caption = "Idle"
            sndAsync.Enabled = True
        End If
    End If
End Sub

Public Sub dwSignaled(o As Object)
    Call CloseHandle(ol2.hEvent)
    ol2.hEvent = 0
    lblStat.Caption = "Idle"
    cmdAsync2.Enabled = True
End Sub
