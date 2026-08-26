Attribute VB_Name = "modEventLog"
Option Explicit
'Gianluigi Cosari - June 1999
'Italsys - Italiana Servizi Telematici S.n.c. - Galliate (NO) ITALY
'E-Mail: glcosari@isys.it

'IMPORTANT: This sofware is just an example and is provided as is.
'           I make no warrenties about this code. Use at your own risk.

Public Const EVNT_SYSTEM = "System"
Public Const EVNT_APP = "Application"
Public Const EVNT_SECURITY = "Security"

Public Const EVENTLOG_SEQUENTIAL_READ = &H1
Public Const EVENTLOG_SEEK_READ = &H2
Public Const EVENTLOG_FORWARDS_READ = &H4
Public Const EVENTLOG_BACKWARDS_READ = &H8

Type EVENTLOGRECORD
     Length As Long               'Length of full record
     Reserved As Long             'Used by the service
     RecordNumber As Long         'Absolute record number
     TimeGenerated As Long        'Seconds since 1-1-1970
     TimeWritten As Long          'Seconds since 1-1-1970
     EventID As Long
     EventType As Integer
     NumStrings As Integer
     EventCategory As Integer
     ReservedFlags As Integer     'For use with paired events (auditing)
     ClosingRecordNumber As Long  'For use with paired events (auditing)
     StringOffset As Long         'Offset from beginning of record
     UserSidLength As Long
     UserSidOffset As Long
     DataLength As Long
     DataOffset As Long           'Offset from beginning of record
End Type

Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (dst As Any, src As Any, ByVal Size As Long)
Declare Function OpenEventLog Lib "advapi32" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpEventSourceName As String) As Long
Declare Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Declare Function GetNumberOfEventLogRecords Lib "advapi32.dll" (ByVal hEventLog As Long, NumberOfRecords As Long) As Long
Declare Function ReadEventLog Lib "advapi32.dll" Alias "ReadEventLogA" (ByVal hEventLog As Long, ByVal dwReadFlags As Long, ByVal dwRecordOffset As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, pnBytesRead As Long, pnMinNumberOfBytesNeeded As Long) As Long

Public Sub WriteEventToDisk(ByVal evtString As String, ByVal evtNumber As Long)

Print #1, ""
Print #1, "Event #" & evtNumber
Print #1, evtString

End Sub


Public Function ReadEvents(ByVal ServerName As String, ByVal EventType As String) As Boolean
'Returns TRUE if successful, FALSE if failure
Dim ret As Long, EventLogHwd As Long, EvtRecNo As Long, rBytesRead As Long, rBytesNeeded As Long
Dim rBuff As EVENTLOGRECORD, EvtReadFlags As Long
Dim eBuff() As Byte, StrucLen As Long, EvtRecLen As Long
Dim strBuffer As String, strStart As Long, strStop As Long, strCount As Long, eBytePointer As Long
Dim eSourceName As String, eComputerName As String, ThisString As String

StrucLen = Len(rBuff)
ReDim eBuff(16384)
EvtReadFlags = EVENTLOG_SEQUENTIAL_READ Or EVENTLOG_FORWARDS_READ

EventLogHwd = OpenEventLog(ServerName, EventType)
If EventLogHwd = 0 Then Exit Function

ret = GetNumberOfEventLogRecords(EventLogHwd, EvtRecNo)
If ret = 0 Then Exit Function

Do While rBuff.RecordNumber < EvtRecNo
    'Reads all events in 16K chunks
    ret = ReadEventLog(EventLogHwd, EvtReadFlags, rBuff.RecordNumber + 1, eBuff(0), 16384, rBytesRead, rBytesNeeded)
    If ret = 0 Then Exit Function

    eBytePointer = 0
    Do While eBytePointer < rBytesRead
        CopyMem rBuff, eBuff(eBytePointer), StrucLen
        EvtRecLen = rBuff.Length
        'Here rBuff is already filled, then we can filter events
    
        strBuffer = Space(EvtRecLen - StrucLen)
        CopyMem ByVal strBuffer, eBuff(StrucLen + eBytePointer), (EvtRecLen - StrucLen)
        eBytePointer = eBytePointer + EvtRecLen
        
        strStart = 1
        strStop = InStr(strStart, strBuffer, Chr(0))
        eSourceName = Mid(strBuffer, strStart, strStop - strStart)
        
        strStart = strStop + 1
        strStop = InStr(strStart, strBuffer, Chr(0))
        eComputerName = Mid(strBuffer, strStart, strStop - strStart)
        
        'Put all strings together, we can parse later...
        If rBuff.NumStrings > 0 Then
            strStart = rBuff.StringOffset - StrucLen + 1
            ThisString = ""
            For strCount = 1 To rBuff.NumStrings
                strStop = InStr(strStart, strBuffer, Chr(0))
                ThisString = ThisString & Mid(strBuffer, strStart, strStop - strStart) & " "
                strStart = strStop + 1
            Next strCount
            'Here 'ThisString' contains all strings of the current event
            WriteEventToDisk ThisString, rBuff.RecordNumber
        End If
    Loop
Loop

ret = CloseEventLog(EventLogHwd)

ReadEvents = True

End Function


