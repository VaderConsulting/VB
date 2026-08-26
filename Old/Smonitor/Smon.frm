VERSION 5.00
Begin VB.Form SMon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Event Monitor"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4215
   Icon            =   "SMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Monitored Events "
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear &List"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close Window"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&All Events"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3240
         Top             =   3480
      End
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Oldest Record Index:"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Record Read:"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Number of Records:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   1935
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Settings"
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "Event &Viewer"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit Program"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Events"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "SMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
'***********************************************************************************
'
'                                 SMONITOR.VBP
'                           Security Event Log Monitor
'                                Richard Puckett
'                               rpuckett@snl.com
'
'***********************************************************************************
'
'  This project was designed to work through the NT Event Log APIs and construct
'  a useful tool in the process.  This tool monitors the Security Event log for
'  desired events and notifies the Admin if any occur, via Mail, Page, or Console
'  messsaging, since NT does not natively support this type of active messaging.
'  for, what I deem, mission-critical security events.
'
'  NOTE: The messenger service must be active on the system for Console Alerts to
'        function
'
'  **Feel free to use this code, but it is to be used at your own risk.**
'
'  The heart of the Read functions come from an article by Dan Appleman
'  called "The Big Event", published sometime in 1997 (?).  I would highly recommend
'  his book on programming to the Win32 API if you are at all interested in using
'  VB for Win32 progamming.  Please feel free to send on improvements, comments, etc.
'  but note that this is a work in progress and still requires a more work to make
'  the custom functions form-independent, etc.  It is an offshoot of a larger project
'  that monitors multiple servers from a single workstation.  I am releasing this
'  project early to VbAdminCode so as to make it in before the next webpage update.
'
'  This tool has been tested under NT Workstation, SP4 and SP5 only.
'
'  Caveats, Work to Do, Etc.:
'
'  E-MAIL/PAGING FUNCTIONS - BLAT
'  BLAT.EXE, for the uninitiated, is a little tool that allows mail messages to be
'  sent from the commandline, thus avoiding the evils of any MAPI functions.  It must
'  be installed on the local station in order for the mail functions to work, otherwise
'  you'll be limited to console alert only.  The BLAT tool (and info) is available from
'  http://www.interlog.com/~tcharron/blat.html.  A highly recommended Admin tool for
'  NT. NOTE: This project does NOT include BLAT, you must download it yourself.
'
'  ReadEventLog API - EVENTLOG_SEEK_READ & NotifyChangeEventLog API Options
'  A matter of choice really, if the EVENTLOG_SEEK_READ value is used, then the ReadEventLog
'  API uses the dwRecordOffset to tell it where to start reading in the Event log.  If
'  You use the EVENTLOG_SEQUENTIAL_READ, and maintain an open handle to the log, then
'  the API maintains the necessary information to begin reading from the last record read
'  as long as it is passed the same handle.  The downside of the EVENTLOG_SEEK_READ is
'  that you may receive an Error 87 if the log you are attempting to read is more than
'  2MBs in size (read - BUG, MSKB Article Q177199).  The NotifyChangeEventLog API is also
'  a choice, to poll or not to poll? If used, then the thread maintenance might involve
'  WaitForMultipleObjects and Alert resets, fun fun fun.
'
'  FormatMessage API
'  No thanks, I can parse out all of the data I would want without dealing with calling
'  this API to do it for me.  The added issue is that if you are reading an Event log on
'  a remote machine that has custom registered events that are not on the monitoring
'  station, you're going to have a problem (ie. LoadLibrary).
'
'  What's Left to Do
'  Error handling, form-independent functions, better start procedures (so as to not be
'  inundated with events EVERY time the dang thing is started.
    
    Private Sub Form_Load()
        
        Dim nServer As String
        Dim nLog As String
        Dim X As Long
        Dim StartTotal As Long
        Dim StartRecord As Long

         'Checks to see if the tool is currently active
        If App.PrevInstance = True Then
            MsgBox _
            "Check the System Tray for an active Security Event Monitor", _
            vbExclamation, "Duplicate Application Detected!"
            Unload Config
            Unload SMon
            End
        End If
        
        'Structure for the Systray Icon
        With nfIconData
            nfIconData.hwnd = Me.hwnd
            nfIconData.uID = Me.Icon
            nfIconData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
            nfIconData.uCallbackMessage = WM_MOUSEMOVE
            nfIconData.hIcon = Me.Icon.Handle
            nfIconData.szTip = "Security Event Monitor" & Chr$(0)
            nfIconData.cbSize = Len(nfIconData)
        End With
    
        'Initialize nServer with the local system name
        nServer = GetLocalSystemName & Chr$(0)
        nLog = "Security" & Chr$(0)

        'Open the Security Event log on the local station
        SecLog = OpenEventLog(nServer, nLog)
    
            'If not OK, bail...
            If SecLog = Null Then
                MsgBox "Unable to obtain a handle to the Security Event log", _
                vbCritical, "Error " & Err.LastDllError
                
                'Kills the Systray Icon
                Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
                        
                Unload Config
                Unload SMon
                End
            End If
        
            'If OK, then...
            If SecLog <> 0 Then

                'Read the total number of event entries
                X = GetNumberOfEventLogRecords(SecLog, StartTotal)

                    If X <> 0 Then
                    
                        'Read the oldest event record index number
                        X = GetOldestEventLogRecord(SecLog, StartRecord)
                               
                            If X <> 0 Then
                                
                                'Set the beginning Record number to be read later by the
                                'Polling function to determine if change has occurred
                                'to the event log during the polling interval
                                lOldTotal = StartTotal + StartRecord
                                'Debug.Print "lOldTotal, at load - " & lOldTotal

                            Else
                                MsgBox "Unable to retrieve the oldest Event record index", _
                                vbCritical, "Error " & Err.LastDllError
                            End If

                    Else
                        MsgBox "Unable to retrieve total number of Event Records", _
                        vbCritical, "Error " & Err.LastDllError
                    End If

            End If
        
   
        'Start the timer
        Timer1.Enabled = True
   
       'Adds the icon to the system tray
        Call Shell_NotifyIcon(NIM_ADD, nfIconData)
            
        'Hide the form
        SMon.Hide
    
    End Sub
        
    Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X _
    As Single, Y As Single)
    
    Dim Msg As Long
    
    If SMon.Visible And Msg = WM_RBUTTONDOWN Then SMon.Hide
    
        'the value of X will vary depending upon the scalemode setting
        If Me.ScaleMode = vbPixels Then
            Msg = X
        Else
            Msg = X / Screen.TwipsPerPixelX
        End If
           
            Select Case Msg
                'MouseMove
                Case WM_MOUSEMOVE
                'LeftMouseDown
                Case WM_LBUTTONDOWN
                'LeftMouseUp
                Case WM_LBUTTONUP
                'LeftDblClick
                Case WM_LBUTTONDBLCLK
                    
                    'Show the Monitor, disable the show menu option
                    SMon.Show
                    mnuShow.Enabled = False
                
                'RightMouseDown
                Case WM_RBUTTONDOWN
                
                    'Show the context menu
                    PopupMenu mnuPopup
                    
                'RightMouseUp
                Case WM_RBUTTONUP
                'RightDblClick
                Case WM_RBUTTONDBLCLK
                
                    'Hide the Monitor, enable the show menu option
                    SMon.Hide
                    mnuShow.Enabled = True
                    
            End Select
    
    End Sub
        
    Private Sub Form_Terminate()
    
            Dim Release As Long
        
            'For those people who like to close apps using the control box
            nfIconData.cbSize = Len(nfIconData)
            nfIconData.hwnd = Me.hwnd
            nfIconData.uID = Me.Icon
                    
            'Close the Event Log
            Release = CloseEventLog(SecLog)
            
            If Release = 0 Then
                 MsgBox "Unable to release the handle to the Security Event log", _
                vbCritical, "Error " & Err.LastDllError
            End If
            
            'Kills the Systray Icon
            Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
                    
            Unload Config
            Unload SMon
            End
    
    
    End Sub

    Private Sub mnuAbout_Click()
    
        Dim Msg As String
        
        'The obligatory glory page
        Msg = "NT Security Log Monitor" & vbLf
        MsgBox Msg & "    by Richard Puckett", , "About..."
        
    End Sub

    Private Sub mnuEvent_Click()
    
        Dim X As Long
        Dim StrType As String
               
        X = 0
        
        'Since this tool is already installed in
        'the path, call it just by the .exe name
        StrType = "EVENTVWR.EXE"
        
        'Open the Event Viewer on the local station
        X = Shell(StrType, 1)
        
        Do While X = 0
           DoEvents
        Loop
        
        If X <> 0 Then
            'Debug.Print X
        End If
    
    End Sub
    
    Private Sub Command1_Click()
    
        Dim nServer As String
        Dim nLog As String
        Dim ViewLog As Long
        Dim Release As Long
        Dim lDirection As Long
        Dim H As Long
        Dim lRecord As Long
        Dim lTRecords As Long
        
        nServer = GetLocalSystemName & Chr$(0)
        nLog = "Security" & Chr$(0)

        'Open the Security Event log on the local station
        ViewLog = OpenEventLog(nServer, nLog)
    
            If ViewLog = Null Then
                MsgBox "Unable to obtain a handle to the Security Event log", _
                vbCritical, "Error " & Err.LastDllError
                Exit Sub
            End If
        
            If ViewLog <> 0 Then

                'Read the oldest event record index number and display it
                H = GetOldestEventLogRecord(ViewLog, lRecord)
                               
                    If H <> 0 Then
                        SMon.Label6.Caption = lRecord
                    Else
                    MsgBox "Unable to retrieve the oldest record index", _
                    vbCritical, "Error " & Err.LastDllError
                    End If
                
                'Read the total number of actual event records and display it
                H = GetNumberOfEventLogRecords(ViewLog, lTRecords)
                
                    If H <> 0 Then
                        SMon.Label4.Caption = lTRecords
                    Else
                    MsgBox "Unable to retrieve total number of event records", _
                    vbCritical, "Error " & Err.LastDllError
                    End If

                'Check the configuration page for the direction the Records should
                'be read in
                If Config.Option1.Value = True Then
                    lDirection = EVENTLOG_BACKWARDS_READ Or EVENTLOG_SEQUENTIAL_READ
                End If
                
                If Config.Option2.Value = True Then
                    lDirection = EVENTLOG_FORWARDS_READ Or EVENTLOG_SEQUENTIAL_READ
                End If
        
                'Call the Event Read wrapper
                Call MonitorSecurity(ViewLog, 3, lDirection)
            
            End If
                
        'Close the handle to the event log
        Release = CloseEventLog(ViewLog)
    
        If Release = 0 Then
             MsgBox "Unable to release the handle to the Security Event log", _
            vbCritical, "Error " & Err.LastDllError
            Exit Sub
        End If
    
    End Sub
    
    Private Sub Command2_Click()
            
        List1.Clear
        Label4.Caption = "0"
        Label5.Caption = "0"
        Label6.Caption = "0"
    
        SMon.Hide
        mnuShow.Enabled = True
        
    End Sub
    
    Private Sub Command3_Click()
    
        List1.Clear
        Label4.Caption = "0"
        Label5.Caption = "0"
        Label6.Caption = "0"
    
    End Sub

    Private Sub Command4_Click()
    
        Dim nServer As String
        Dim nLog As String
        Dim ViewLog As Long
        Dim Release As Long
        Dim lDirection As Long
        Dim H As Long
        Dim lRecord As Long
        Dim lTRecords As Long
        
        'Get the local system name for nServer
        nServer = GetLocalSystemName & Chr$(0)
        nLog = "Security" & Chr$(0)

        'Open the Security Event log on the local station
        ViewLog = OpenEventLog(nServer, nLog)
    
            If ViewLog = Null Then
                MsgBox "Unable to obtain a handle to the Security Event log", _
                vbCritical, "Error " & Err.LastDllError
                Exit Sub
            End If
        
            If ViewLog <> 0 Then

                'Get the oldest record index and display it
                H = GetOldestEventLogRecord(ViewLog, lRecord)
                               
                    If H <> 0 Then
                        SMon.Label6.Caption = lRecord
                    Else
                    MsgBox "Unable to the oldest record index", _
                    vbCritical, "Error " & Err.LastDllError
                    End If
                
                'Get the total number of event records and display it
                H = GetNumberOfEventLogRecords(ViewLog, lTRecords)
                
                    If H <> 0 Then
                        SMon.Label4.Caption = lTRecords
                    Else
                    MsgBox "Unable to retrieve total number of Event Records", _
                    vbCritical, "Error " & Err.LastDllError
                    End If
        
                'Read the config page for direction of display
                If Config.Option1.Value = True Then
                    lDirection = EVENTLOG_BACKWARDS_READ Or EVENTLOG_SEQUENTIAL_READ
                End If
                
                If Config.Option2.Value = True Then
                    lDirection = EVENTLOG_FORWARDS_READ Or EVENTLOG_SEQUENTIAL_READ
                End If
            
                'Call the Event Read wrapper
                Call MonitorSecurity(ViewLog, 2, lDirection)
        
            End If
        'Close the handle
        Release = CloseEventLog(ViewLog)
        
        If Release = 0 Then
             MsgBox "Unable to release the handle to the Security Event log", _
            vbCritical, "Error " & Err.LastDllError
            Exit Sub
        End If
        
    End Sub
    
    
    Private Sub mnuConfig_Click()
    
        Config.Show
    
    End Sub

    Private Sub mnuExit_Click()
    
        Dim Release As Long
    
        nfIconData.cbSize = Len(nfIconData)
        nfIconData.hwnd = Me.hwnd
        nfIconData.uID = Me.Icon
                
        'Close the Event Log
        Release = CloseEventLog(SecLog)
        
        If Release = 0 Then
            MsgBox "Unable to release the handle to the Security Event Log", _
            vbCritical, "Error " & Err.LastDllError
        End If
        
        'Kills the Systray Icon
        Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
                
        Unload Config
        Unload SMon
        End
    
    End Sub

    Private Sub mnuShow_Click()
    
        SMon.Show
        mnuShow.Enabled = False
        
    End Sub
    
    Private Sub Timer1_Timer()
            
        Dim X As Long
        Dim lNewTotal As Long
        Dim NotifyType As Integer
        Dim lDirection As Long
        Dim NewTotal As Long
        Dim NewRecord As Long
        
        'Always read forwards with the handle to the event log that will
        'be used for subsequent reads, since the SEQUENTIAL_READ parameter
        'states that the function will maintain the record index of the
        'last record read and start from that record with every new read request
        'made from the same handle to the Event Log.
        lDirection = EVENTLOG_FORWARDS_READ Or EVENTLOG_SEQUENTIAL_READ
    
        TimeExpired = TimeExpired + Timer1.Interval
        
        'Debug.Print TimeExpired
        
        '(Conversion to Minutes)
        ExpiredMinutes = (TimeExpired / 1000) / 60
                    
            'If the polling period matches the total time elapsed, do...
            If ExpiredMinutes >= IDLEMINUTES * Val(Config.Text4) Then
                
                'Reset TimeExpired
                TimeExpired = 0

                'Read the total record set
                X = GetNumberOfEventLogRecords(SecLog, NewTotal)
                
                If X <> 0 Then
                
                    'Read the oldest event record index number
                    X = GetOldestEventLogRecord(SecLog, NewRecord)
                           
                        If X <> 0 Then
                            
                            'Total Record count + the oldest record index
                            lNewTotal = NewTotal + NewRecord
                            'Debug.Print "lOldTotal, at timer start - " & lOldTotal
                            'Debug.Print "lNewTotal, at timer start - " & lNewTotal
                            
                        Else
                            MsgBox "Unable to retrieve the oldest Event record index", _
                            vbCritical, "Error " & Err.LastDllError
                        End If

                Else
                    MsgBox "Unable to retrieve total number of Event Records", _
                    vbCritical, "Error " & Err.LastDllError
                End If

                'No new events have been written to the event log, exit sub
                If lNewTotal - lOldTotal = 0 Then Exit Sub
        
                'New events have occurred, read them
                If lNewTotal - lOldTotal <> 0 Then
                    
                    'Reset lOldTotal to the reflect the new total number of records
                    lOldTotal = lNewTotal
                    
                    'Debug.Print "lOldTotal, after calulation " & lOldTotal

                        'Call Event routines
                        Call MonitorSecurity(SecLog, 1, lDirection)
        
                 End If
                        
            End If
            
    End Sub
