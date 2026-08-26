VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private xhProcess As Long
Private xhThread As Long
Private xdwProcessID As Long
Private xdwThreadID As Long
Private xTaskID As Long

Public Sub SetProcessInfo(ByVal lProcess As Long, ByVal lThread As Long, ByVal lProcessID As Long, ByVal lThreadID As Long)
  xhProcess = lProcess
  xhThread = lThread
  xdwProcessID = lProcessID
  xdwThreadID = lThreadID
End Sub

Public Property Get TaskID() As Long
  TaskID = xTaskID
End Property

Public Property Let TaskID(Value As Long)
  xTaskID = Value
End Property

Public Property Get hProcess() As Long
  hProcess = xhProcess
End Property

Public Property Get hThread() As Long
  hThread = xhThread
End Property

Public Property Get dwProcessID() As Long
  dwProcessID = xdwProcessID
End Property

Public Property Get dwThreadID() As Long
  dwThreadID = xdwThreadID
End Property

Private Sub Class_Initialize()
 xhProcess = 0
 xhThread = 0
 xdwProcessID = 0
 xdwThreadID = 0
End Sub
