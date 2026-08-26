Attribute VB_Name = "modHook"
Option Explicit

'declarations
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

'constants
Public Const WM_CANCELJOURNAL = &H4B

'types
Type POINTAPI
    x As Long
    y As Long
End Type

Type TMSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'veriables
Public hJournalHook As Long, hAppHook As Long
Public SHptr As Long


'functions ------------------------------

Public Function JournalRecordProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode < 0 Then
        JournalRecordProc = CallNextHookEx(hJournalHook, nCode, wParam, lParam)
        Exit Function
    End If
    ResolvePointer(SHptr).FireEvent lParam, wParam
    Call CallNextHookEx(hJournalHook, nCode, wParam, lParam)
End Function


Public Function AppHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If nCode < 0 Then
        AppHookProc = CallNextHookEx(hAppHook, nCode, wParam, lParam)
        Exit Function
    End If
    Dim msg As TMSG
    CopyMemory msg, ByVal lParam, Len(msg)
    Select Case msg.message
        Case WM_CANCELJOURNAL
            If wParam = 1 Then ResolvePointer(SHptr).FireEvent WM_CANCELJOURNAL, wParam
    End Select
    Call CallNextHookEx(hAppHook, nCode, wParam, ByVal lParam)
End Function


Private Function ResolvePointer(ByVal lpObj&) As cSysHook
    Dim oSH As cSysHook
    CopyMemory oSH, lpObj, 4&
    Set ResolvePointer = oSH
    CopyMemory oSH, 0&, 4&
End Function

