Attribute VB_Name = "Secmon"
    Option Explicit
    
    '' Timer Constants
    Public TimeExpired
    Public ExpiredMinutes
    Public Const IDLEMINUTES = 1
    
    '' Handle & Buffer info for the local Event Log
    Public SecLog As Long
    Public Const BUFFER_SIZE = 32768
    Dim bBuffer(BUFFER_SIZE) As Byte
    Dim dwRead As Long
    Dim dwNeeded As Long
    Dim cRecords As Long
    Dim dwThisRecord As Long
    Dim ev As EVENTLOGRECORD
    Dim pevlr As Long
    Public lOldTotal As Long
    
    '' Definitions for the READ flags for Event logging
    Public Const EVENTLOG_SEQUENTIAL_READ = 1&
    Public Const EVENTLOG_SEEK_READ = 2&
    Public Const EVENTLOG_FORWARDS_READ = 4&
    Public Const EVENTLOG_BACKWARDS_READ = 8&
    
    '' The types of events that can be logged
    Public Const EVENTLOG_SUCCESS = 0&
    Public Const EVENTLOG_ERROR_TYPE = 1&
    Public Const EVENTLOG_WARNING_TYPE = 2&
    Public Const EVENTLOG_INFORMATION_TYPE = 4&
    Public Const EVENTLOG_AUDIT_SUCCESS = 8&
    Public Const EVENTLOG_AUDIT_FAILURE = 10&
    
    '' Definitions for the WRITE flags used by Auditing for paired events
    Public Const EVENTLOG_START_PAIRED_EVENT = 1&
    Public Const EVENTLOG_END_PAIRED_EVENT = 2&
    Public Const EVENTLOG_END_ALL_PAIRED_EVENTS = 4&
    Public Const EVENTLOG_PAIRED_EVENT_ACTIVE = 8&
    Public Const EVENTLOG_PAIRED_EVENT_INACTIVE = 16&
    
    '' Event log-specific status codes returned by GetLastError
    Public Const ERROR_EVENTLOG_FILE_CORRUPT = 1500&
    Public Const ERROR_EVENTLOG_CANT_START = 1501&
    Public Const ERROR_LOG_FILE_FULL = 1502&
    Public Const ERROR_EVENTLOG_FILE_CHANGED = 1503&
    Public Const ERROR_INSUFFICIENT_BUFFER = 122
    Public Const ERROR_HANDLE_EOF = 38&
    Public Const ERROR_INVALID_PARAMETER = 87
    
    Public Const GENERIC_READ = &H80000000
    Public Const INFINITE = &HFFFFFFFF
    
    '' Systray Icon constants
    Public Const NIF_ICON = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_TIP = &H4
    Public Const NIM_ADD = &H0
    Public Const NIM_DELETE = &H2
    Public Const MAX_TOOLTIP As Integer = 64
    Public Const NIM_MODIFY = &H1
    Public Const WM_MOUSEMOVE = &H200       'MouseMove
    Public Const WM_LBUTTONDOWN = &H201     'LeftMouseDown
    Public Const WM_LBUTTONUP = &H202       'LeftMouseUp
    Public Const WM_LBUTTONDBLCLK = &H203   'LeftDblClick
    Public Const WM_RBUTTONDOWN = &H204     'RightMouseDown
    Public Const WM_RBUTTONUP = &H205       'RightMouseUp
    Public Const WM_RBUTTONDBLCLK = &H206   'RightDblClick

    '' Constants for the Program Close Functions
    Public Const WM_CLOSE = &H10
    Public Const ERROR_NONE = 0&
    Public Const NOERROR = 0
    Public Const NO_ERROR = 0
    Public Const ERROR_SUCCESS = 0&

    '' Constants for network enumeration functions
    Public Const SV_TYPE_NT = &H1000
    Public Const SV_TYPE_DOMAIN_ENUM = &H80000000
    
    '' Systray Icon Struct Info
    Public nfIconData As NOTIFYICONDATA
    
    Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * MAX_TOOLTIP
    End Type
    
    '' NetAPI32 Struct for retrieving Workstation Info
    Public Type WKSTA_INFO_100
        wki100_platform_id As Long
        wki100_computername As Long
        wki100_langroup As Long
        wki100_ver_major As Long
        wki100_ver_minor As Long
    End Type

    '' Event Record Struct
    Public Type EVENTLOGRECORD
        Length As Long                 ' Length of full record
        Reserved As Long               ' Used by the service
        RecordNumber As Long           ' Absolute record number
        TimeGenerated As Long          ' Seconds since 1-1-1970
        TimeWritten As Long            ' Seconds since 1-1-1970
        EventID As Long                '
        EventType As Integer           '
        NumStrings As Integer          '
        EventCategory As Integer       '
        ReservedFlags As Integer       ' For use with paired events (auditing)
        ClosingRecordNumber As Long    ' For use with paired events (auditing)
        StringOffset As Long           ' Offset from beginning of record
        UserSidLength As Long          '
        UserSidOffset As Long          '
        DataLength As Long             '
        DataOffset As Long             ' Offset from beginning of record
    End Type
    
    Public Declare Sub RtlMoveMemory Lib "kernel32" _
        (dest As Any, _
        source As Any, _
        ByVal bytes As Long)
    
    Public Declare Sub lstrcpyW Lib "kernel32" _
        (vDest As Any, _
        ByVal sSrc As Any)

    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (hpvDest As Any, _
        hpvSource As Any, _
        ByVal cbCopy As Long)

    Public Declare Function lstrlenptr Lib "kernel32" _
        Alias "lstrlenA" _
        (ByVal lpString As Long) As Long

    Public Declare Function lstrcpyfromptr Lib "kernel32" _
        Alias "lstrcpyA" _
        (ByVal lpString1 As String, _
        ByVal lpString2 As Long) As Long

    Public Declare Function OpenEventLog Lib "advapi32.dll" _
        Alias "OpenEventLogA" _
        (ByVal lpUNCServerName As String, _
        ByVal lpSourceName As String) As Long
    
    Public Declare Function GetNumberOfEventLogRecords Lib "advapi32.dll" _
        (ByVal hEventLog As Long, _
        NumberOfRecords As Long) As Long

    Public Declare Function GetOldestEventLogRecord Lib "advapi32.dll" _
        (ByVal hEventLog As Long, _
        OldestRecord As Long) As Long

    Public Declare Function ReadEventLog Lib "advapi32.dll" _
        Alias "ReadEventLogA" _
        (ByVal hEventLog As Long, _
        ByVal dwReadFlags As Long, _
        ByVal dwRecordOffset As Long, _
              lpBuffer As Byte, _
        ByVal nNumberOfBytesToRead As Long, _
              pnBytesRead As Long, _
              pnMinNumberOfBytesNeeded As Long) As Long
    
    Public Declare Function CloseEventLog Lib "advapi32.dll" _
        (ByVal hEventLog As Long) As Long

    Public Declare Function NetApiBufferFree Lib "Netapi32.dll" _
        (ByVal lpBuffer As Long) As Long
    
    Public Declare Function NetWkstaGetInfo Lib "Netapi32.dll" _
        (ByVal sServerName$, _
        ByVal lLevel&, _
        vBuffer As Any) As Long
    
    Public Declare Function NetMessageBufferSend Lib "Netapi32.dll" _
        (ByVal sServerName$, _
        ByVal sMsgName$, _
        ByVal sFromName$, _
        ByVal sMessageText$, _
        ByVal lBufferLength&) As Long
           
    Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" _
        (ByVal lpString1 As String, _
        ByVal lpString2 As Long) As Long

    Public Declare Function Shell_NotifyIcon Lib "shell32.dll" _
        Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, _
        lpData As NOTIFYICONDATA) As Long

    Public Function GetLocalSystemName()
        Dim lReturnCode As Long
        Dim bBuffer(512) As Byte
        Dim I As Integer
        Dim twkstaInfo100 As WKSTA_INFO_100, lwkstaInfo100 As Long
        Dim lwkstaInfo100StructPtr As Long
        Dim sLocalName As String
        
        'Retrieve the Workstation info
        lReturnCode = NetWkstaGetInfo("", 100, lwkstaInfo100)
     
        lwkstaInfo100StructPtr = lwkstaInfo100
                     
        If lReturnCode = 0 Then
                     
            'Copy the data into the struct
            RtlMoveMemory twkstaInfo100, ByVal _
            lwkstaInfo100StructPtr, Len(twkstaInfo100)
             
            lstrcpyW bBuffer(0), twkstaInfo100.wki100_computername
    
            I = 0
                
                Do While bBuffer(I) <> 0
                    'Retrieve the workstation name, parsing out Nulls
                    sLocalName = sLocalName & Chr(bBuffer(I))
                    I = I + 2
                Loop
                
            'Set the workstation name
            GetLocalSystemName = sLocalName
             
        End If
    
    End Function
      
    ' Get a string given a long pointer
    Public Function LoadStringFromPtr(ByVal ptrval&) As String
       Dim slen&
       Dim resstring$
       If ptrval = 0 Then Exit Function
       
       slen = lstrlenptr(ptrval)
       resstring = String$(slen + 1, 0)
       
           Call lstrcpyfromptr(resstring, ptrval)
       
       LoadStringFromPtr = Left$(resstring, slen)
    
    End Function

    Public Function MonitorSecurity(ByVal lHwnd As Long, ByVal tRequest As Integer, _
    ByVal Direction As Long)
    
        Dim X As Long
        Dim H As Long
        Dim lRecord As Long
        Dim lTRecords As Long
        Dim sCustomEvt
              
        On Error GoTo ErrorHandler

        If lHwnd <> 0 Then
            
            pevlr = VarPtr(bBuffer(0))
            
            'Read the Event Log
            While (ReadEventLog(lHwnd, _
                   Direction, _
                   0, _
                   bBuffer(0), _
                   BUFFER_SIZE, _
                   dwRead, _
                   dwNeeded) <> 0)
             
                While dwRead > 0
                    
                'Copy the data into the ev structure.
                RtlMoveMemory ev, ByVal pevlr, Len(ev)
                
                'If tRequest type 3, then view entire Security Log
                If tRequest = 3 Then
                    
                    Call FormatMessage

                End If
                
                'If not 3, then pass request to EventMsg function for intrepretation
                'of the request type, re-passing tRequest to represent the call type
                If tRequest <> 3 Then
                
                sCustomEvt = Config.Text5.Text
                
                    Select Case CStr(ev.EventID)
                                
                        Case "512"
                            If Config.Chk512.Value = 1 Then _
                            X = EventMsg("512", "System Restart", tRequest)
                        Case "513"
                            If Config.Chk513.Value = 1 Then _
                            X = EventMsg("513", "System Shutdown", tRequest)
                        Case "517"
                            If Config.Chk517.Value = 1 Then _
                            X = EventMsg("517", "Audit Log Cleared", tRequest)
                        Case "528"
                            If Config.Chk528.Value = 1 Then _
                            X = EventMsg("528", "Successful Logon", tRequest)
                        Case "529"
                            If Config.Chk529.Value = 1 Then _
                            X = EventMsg("529", "Bad Username or Password", tRequest)
                        Case "530"
                            If Config.Chk530.Value = 1 Then _
                            X = EventMsg("530", "Logon Restriction Violation", tRequest)
                        Case "531"
                            If Config.Chk531.Value = 1 Then _
                            X = EventMsg("531", "Account Disabled", tRequest)
                        Case "532"
                            If Config.Chk532.Value = 1 Then _
                            X = EventMsg("532", "Account Expired", tRequest)
                        Case "533"
                            If Config.Chk533.Value = 1 Then _
                            X = EventMsg("533", "User Not Allowed to Logon", tRequest)
                        Case "534"
                            If Config.Chk534.Value = 1 Then _
                            X = EventMsg("534", "Logon Type Restricted", tRequest)
                        Case "535"
                            If Config.Chk535.Value = 1 Then _
                            X = EventMsg("535", "Password has Expired", tRequest)
                        Case "537"
                            If Config.Chk537.Value = 1 Then _
                            X = EventMsg("537", "Unsuccessful Logon", tRequest)
                        Case "538"
                            If Config.Chk538.Value = 1 Then _
                            X = EventMsg("538", "User Logoff", tRequest)
                        Case "539"
                            If Config.Chk539.Value = 1 Then _
                            X = EventMsg("539", "Account Locked Out", tRequest)
                        Case sCustomEvt
                            If Config.ChkCustom.Value = 1 Then _
                            X = EventMsg(sCustomEvt, "Custom Monitored Event", tRequest)
                        Case Else
                            
                    End Select
                
                End If
                
                dwRead = dwRead - ev.Length
                pevlr = pevlr + ev.Length
                 
                Wend
                ' Reset the pointer to the start of the buffer
                pevlr = VarPtr(bBuffer(0))

            Wend
          
        End If

        
    Exit Function
    
ErrorHandler:

    MsgBox " An error has occurred while reading the Event Log", _
    vbCritical, "Error " & Err.LastDllError


    End Function

    Public Function EventMsg(ByVal EvtID As String, ByVal MsgType As String, _
    ByVal tRequest As Integer)
    
    Dim PauseTime, Start, Finish, TotalTime
    Dim lReturnCode As Long
    Dim sUnicodeToName As String
    Dim sUnicodeFromName As String
    Dim sUnicodeMessage As String
    Dim lMessageLength As Long
    Dim txtType As String
    Dim lUser As String
    Dim lSystem As String
    Dim lEventId As String
    Dim Time_t As Long
    Dim NewTime As Date
    Dim lEvtType As String
    Dim lEvtCat As String
    Dim X As Long
    Dim Contact As String
    Dim SubjText As String

    'Read the Username from the buffer
    lUser = LoadStringFromPtr(pevlr + ev.StringOffset)
    
    'Read the Event type and translate into Record Type
    Select Case Hex$(ev.EventType)
        Case 10
            lEvtType = "Failure Audit"
        Case 8
            lEvtType = "Success Audit"
        Case 4
            lEvtType = "Information"
        Case 2
            lEvtType = "Warning"
        Case 1
            lEvtType = "Error"
        Case 0
            lEvtType = "Success"
    End Select
    
    'Read the Event Category and translate
    Select Case ev.EventCategory
        Case 1
            lEvtCat = "System Event"
        Case 2
            lEvtCat = "Logon/Logoff"
        Case Else
            lEvtCat = "None (" & ev.EventCategory & ")"
    End Select
    
    'If no recorded user, then process UNKNOWN
    If Len(lUser) = 0 Then
        lUser = "UNKNOWN"
    End If
        
    'Retrieve the system name that generated the event
    lSystem = LoadStringFromPtr(pevlr + Len(ev) + lstrlenptr(pevlr + Len(ev)) + 1)
    
    'Event ID
    lEventId = ev.EventID
        
    'Adjust for Eastern Standard Time from GMT
    '(4 hours - 14400 seconds)
    Time_t = ev.TimeGenerated - 14400
    NewTime = DateAdd("s", Time_t, #1/1/1970#)

    'If tRequest type involves notification
    If tRequest = 1 Then
    
        'Notify Administrator via Page (e-mail) - BLAT call goes here
        If Config.Check15.Value = 1 Then
            
            X = 0
            
            'Create string from event information, ERROR.TXT can contain
            'info but mine is blank, opting for just information in the
            'Subject line to warn the Admin.
            Contact = Trim("C:\WINNT\SYSTEM32\BLAT.EXE " _
                      & "C:\WINNT\SYSTEM32\ERROR.TXT -S " _
                      & Chr(34) _
                      & UCase(lUser) _
                      & " generated Event " _
                      & lEventId _
                      & " on " _
                      & lSystem _
                      & " at " _
                      & NewTime _
                      & Chr(34) _
                      & " -T " _
                      & Config.Text1.Text)
    
            Debug.Print Contact
            
            'Send the message
            X = Shell(Contact, 0)
            
                Do While X = 0
                   DoEvents
                Loop
            
            If X <> 0 Then
                'Debug.Print X
            End If
            
        End If
        
        'Notify Administrator via E-Mail - BLAT call goes here
        If Config.Check16.Value = 1 Then
        
            X = 0
            'Create string from event information, ERROR.TXT can contain
            'info but mine is blank, opting for just information in the
            'Subject line to warn the Admin.
            Contact = Trim("C:\WINNT\SYSTEM32\BLAT.EXE " _
                      & "C:\WINNT\SYSTEM32\ERROR.TXT -S " _
                      & Chr(34) _
                      & UCase(lUser) _
                      & " generated Event " _
                      & lEventId _
                      & " on " _
                      & lSystem _
                      & " at " _
                      & NewTime _
                      & Chr(34) _
                      & " -T " _
                      & Config.Text2.Text)
    
            Debug.Print Contact
            
            'Send the message
            X = Shell(Contact, 0)
            
                Do While X = 0
                   DoEvents
                Loop
            
            If X <> 0 Then
                'Debug.Print X
            End If
            
        End If
        
        'Send a console message to a system if selected
        If Config.Check17.Value = 1 Then
            
            'Event description for Console message
            txtType = "               A monitored event has occurred:" & vbLf & vbLf & _
                      "               Event ID:       " & EvtID & vbLf & _
                      "               Description:   " & Trim(MsgType) & vbLf & _
                      "               Type:             " & lEvtType & vbLf & _
                      "               Category:       " & lEvtCat & vbLf & _
                      "               Username:     " & UCase(lUser) & vbLf & _
                      "               Time:             " & NewTime & vbLf & _
                      "               System:          " & lSystem
    
            'Convert information to Unicode before sending to the NetMessageBufferSend API
            sUnicodeFromName = StrConv(GetLocalSystemName, vbUnicode)
            sUnicodeToName = StrConv(Config.Text3.Text, vbUnicode)
            sUnicodeMessage = StrConv(txtType, vbUnicode)
            lMessageLength = Len(sUnicodeMessage)
        
            'StatusBar1.Panels("msg").Text = vbNullString
    
            'Send the message
            lReturnCode = NetMessageBufferSend("", _
                sUnicodeToName, _
                sUnicodeFromName, _
                sUnicodeMessage, _
                lMessageLength)
            
                'Put a pause of 1 second in so that the message
                'function has enough time to receive a valid return
                'before clearing the buffer for the next processed
                'event entry
                
                'Set duration.
                PauseTime = 1
                
                'Set start time.
                Start = Timer
                
                Do While Timer < Start + PauseTime
                    DoEvents
                Loop
                
                'Set end time.
                Finish = Timer
    
            If lReturnCode <> 0 Then
                MsgBox "An error occurred while attempting to send an alert to " _
                & Config.Text3.Text, vbCritical, "Error " & Err.LastDllError
                Exit Function
            End If
    
        End If

    End If
    
    If tRequest = 2 Then
    
        'Read the event specified
        Call FormatMessage

    End If
    
    End Function

    Public Function FormatMessage()
        
    'This function still needs work....
    
        Dim Time_t As Long
        Dim NewTime As Date
        
        'Propogate SMon.List1 with Event Record information
        SMon.List1.AddItem "Event ID: " & ev.EventID
           
        'Select the Event Type
        Select Case Hex$(ev.EventType)
            Case 10
                SMon.List1.AddItem "Type: Failure Audit"
            Case 8
                SMon.List1.AddItem "Type: Success Audit"
            Case 4
                SMon.List1.AddItem "Type: Information"
            Case 2
                SMon.List1.AddItem "Type: Warning"
            Case 1
                SMon.List1.AddItem "Type: Error"
            Case 0
                SMon.List1.AddItem "Type: Success"
        End Select
            
        'Select the Event category
        Select Case ev.EventCategory
            Case 1
                SMon.List1.AddItem "Category: System Event"
            Case 2
                SMon.List1.AddItem "Category: Logon/Logoff"
            Case Else
                SMon.List1.AddItem "Category: None (" & ev.EventCategory & ")"
        End Select
            
        'Propogate list with additional Event Record information
        SMon.List1.AddItem "Source: " & LoadStringFromPtr(pevlr + Len(ev))
        SMon.List1.AddItem "Computer Name: " & LoadStringFromPtr(pevlr + Len(ev) + lstrlenptr(pevlr + Len(ev)) + 1)
        SMon.List1.AddItem "Username: " & LoadStringFromPtr(pevlr + ev.StringOffset)
    
            'Adjust for Eastern Standard Time from GMT
            '(4 hours - 14400 seconds)
            Time_t = ev.TimeGenerated - 14400
            NewTime = DateAdd("s", Time_t, #1/1/1970#)
            
        SMon.List1.AddItem "Time Written: " & NewTime
        'Update the last record read label
        SMon.Label5.Caption = ev.RecordNumber
        SMon.List1.AddItem "Record Number: " & ev.RecordNumber
        SMon.List1.AddItem " - - - - - "
    
    End Function

