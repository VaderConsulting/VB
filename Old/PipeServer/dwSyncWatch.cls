BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "dwSyncWatch"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit

Private CallbackObject As Object

' Synchronization Object handle to watch
Private synchnd As Long
' Original object handle as known in the source
' process - for identification purposes only.
Private originalhnd As Long

' Internal state machine for dwSyncWatch object
' Ø - Idle
' 1 - Waiting on object to signal

' 2 - Object has been signalled
Private state As Integer

Public Property Get Signaled() As Boolean
    If state = 2 Then Signaled = True Else Signaled = False
End Property

' Called by the timer routine
Public Sub TimerEvent()
    Dim res&
    If state = Ø Or state = 2 Then Exit Sub
    res = WaitForSingleObject(synchnd, Ø)
    Select Case res
        Case WAIT_OBJECT_Ø
            Call CloseHandle(synchnd)
            state = 2
            On Error GoTo synccallerr

            If Not (CallbackObject Is Nothing) Then
                CallbackObject.dwSignaled Me
            End If
            On Error GoTo Ø
            ' Pull object off the timer list
            RemovingClass Me
            
        Case WAIT_FAILED
            Call CloseHandle(synchnd)
            state = Ø
    End Select
    Exit Sub
synccallerr:
    Resume Next
End Sub

' pid = Process ID of calling process
' objhnd = Handle of object to watch
Public Sub SetObjectWatch(ByVal pid&, ByVal objhnd&)

    Dim srcprocesshnd&
    Dim res&
    ' If we're already watching a process, error
    If state = 1 Then
        RaiseError 1, "dwSyncWatch", "Object already being watched"
        Exit Sub
    End If
    originalhnd = objhnd
    ' Open the process as a source for DuplicaeHandle
    srcprocesshnd = OpenProcess(PROCESS_DUP_HANDLE, True, pid)
    If srcprocesshnd = Ø Then
        RaiseError 2, "dwSyncWatch", "Unable to open the source process"
        Exit Sub

    End If
    res = DuplicateHandle(srcprocesshnd, objhnd, GetCurrentProcess(), synchnd, _
    SYNCHRONIZE, False, Ø)
    ' Close the source process handle
    Call CloseHandle(srcprocesshnd)
    If res = Ø Then
        RaiseError 4, "dwSyncWatch", "Unable to duplicate source handle"
        Exit Sub
    End If
    ' Set the new state
    state = 1
    ' Add this object to main collection
    WatchCollection.Add Me
End Sub

Public Property Get ObjectHandle() As Long

    ObjectHandle = originalhnd
End Property


' Set the callback object - make sure it is valid
Public Sub SetAppCallback(o As Object)
    If agIsValidName(o, "dwSignaled") = Ø Then
        RaiseError 3, "dwSyncWatch", "Callback object function undefined"
        Exit Sub
    End If
    
    Set CallbackObject = o
End Sub

All of the material presented here is copyrighted by either Desaware or Macmillan. No part of this material may be used or reproduced in any fashion (except in brief quotations used in critical articles and reviews) without prior consent.