VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1035
   ClientLeft      =   1275
   ClientTop       =   3945
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   8040
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6630
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function NetRemoteTOD Lib "NETAPI32.DLL" (ByVal server As String, buffer As Any) As Long
Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (ByVal buffer As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

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
Private Sub Command1_Click()
    Dim t As TIME_OF_DAY, tPtr As Long, res As Long, szServer As String, days As Date, todays As Date
    szServer = StrConv("\\Your Server!", vbUnicode)    'Convert the server name to unicode
    res = NetRemoteTOD(szServer, tPtr)  'You could also pass vbNullString for the server name
    If res = 0 Then
        CopyMemory t, ByVal tPtr, Len(t) 'Copy the pointer returned to a TIME_OF_DAY structure
        days = DateSerial(70, 1, 1) + (t.t_elapsedt / 60 / 60 / 24)  'Convert the elapsed time since 1/1/70 to a date
        days = days - (t.t_timezone / 60 / 24)  'Adjust for TimeZone differences
            'Get local computer information for comparison
        todays = DateSerial(70, 1, 1) + (DateDiff("s", DateSerial(70, 1, 1), Now()) / 60 / 60 / 24)
        Me.Cls
        Print DateDiff("s", DateSerial(70, 1, 1), Now()), todays, t.t_elapsedt, days
        NetApiBufferFree (tPtr) 'Free the memory at the pointer
    Else
        MsgBox "Error occurred call NetRemoteTOD: " & res, vbOKOnly, "NetRemoteTOD"
        'Error 53: cannot find server
    End If
End Sub

