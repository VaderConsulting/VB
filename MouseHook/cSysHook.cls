VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSysHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'events
Public Event msgout(hwnd As Long, imsg As Long, LP As Long, HP As Long, itime As Long)

'declarations
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'constants
Private Const WH_JOURNALRECORD = 0
Private Const WH_GETMESSAGE = 3

'types
Private Type EVENTMSG
    wMsg As Long
    lParamLow As Long
    lParamHigh As Long
    msgTime As Long
    hWndMsg As Long
End Type

'variables
Dim EMSG As EVENTMSG


'functions ----------------------------

Public Function SetHook() As Boolean
    If hJournalHook = 0 Then hJournalHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf JournalRecordProc, App.hInstance, 0)
    If hAppHook = 0 Then hAppHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf AppHookProc, App.hInstance, App.ThreadID)
    SetHook = True
End Function

Public Sub RemoveHook()
    UnhookWindowsHookEx hAppHook
    UnhookWindowsHookEx hJournalHook
End Sub

Private Sub Class_Initialize()
    SHptr = ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    If hJournalHook Or hAppHook Then RemoveHook
End Sub

Friend Function FireEvent(ByVal lParam As Long, ByVal hParam As Long)
    Dim s As String
    If lParam = WM_CANCELJOURNAL Then
        hJournalHook = 0
        SetHook
        Exit Function
    End If

    CopyMemory EMSG, ByVal lParam, Len(EMSG)
    
    RaiseEvent msgout(EMSG.hWndMsg, EMSG.wMsg, EMSG.lParamLow, EMSG.lParamHigh, EMSG.msgTime)

End Function

