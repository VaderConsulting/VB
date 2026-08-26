VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HostTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type TIME_OF_DAY
    t_elapsedt As Long
    t_msecs As Long
    t_hours As Long
    t_mins As Long
    t_secs As Long
    t_hunds As Long
    t_timezone As Long
    t_tinterval As Long
    t_day As Long
    t_month As Long
    t_year As Long
    t_weekday As Long
End Type

Private Declare Function NetRemoteTOD Lib "netapi32.dll" (ByVal server As String, buffer As Any) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal buffer As Long) As Long

Public Function GetTime(Servername As String, Optional Delta As Boolean) As String
    Dim t As TIME_OF_DAY, tPtr As Long, res As Long, szServer As String, days As Date, todays As Date
    szServer = StrConv("\\" & Servername, vbUnicode)    'Convert the server name to unicode
    res = NetRemoteTOD(szServer, tPtr)                  'You could also pass vbNullString for the server name
    If res = 0 Then
        CopyMemory t, ByVal tPtr, Len(t)                'Copy the pointer returned to a TIME_OF_DAY structure
        days = DateSerial(70, 1, 1) + (t.t_elapsedt / 60 / 60 / 24)  'Convert the elapsed time since 1/1/70 to a date
        days = days - (t.t_timezone / 60 / 24)          'Adjust for TimeZone differences
        'lblTime = days
                                                        'Get local computer information for comparison
        todays = DateSerial(70, 1, 1) + (DateDiff("s", DateSerial(70, 1, 1), Now()) / 60 / 60 / 24)
                                                        'Print DateDiff("s", DateSerial(70, 1, 1), Now()), todays, t.t_elapsedt, days
        If Delta = False Then
          GetTime = days ' & " Difference: " & DateDiff("s", days, todays) & " Seconds"
        Else
          GetTime = DateDiff("s", days, todays)
        End If
        NetApiBufferFree (tPtr)                         'Free the memory at the pointer
    Else
        Debug.Print "Error occurred call NetRemoteTOD: " & res
        'MsgBox "Error occurred call NetRemoteTOD: " & res, vbOKOnly, "NetRemoteTOD"
                                                        'Error 53: cannot find server
    End If
End Function


