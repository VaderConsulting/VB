VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const NERR_Success As Long = 0&

Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" (yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long

Public Function BroadcastMessage(strUserOrMachine As String, strFrom As String, strMessage As String) As Boolean
    Dim byteToName() As Byte
    Dim byteFromName() As Byte
    Dim MessageToSend() As Byte
    
    If Trim(strUserOrMachine) = "" Or Trim(strFrom) = "" Or Trim(strMessage) = "" Then
        Exit Function
    End If
    
    'Put data into byte arrays
    byteToName = strUserOrMachine & vbNullChar
    byteFromName = strFrom & vbNullChar
    MessageToSend = strMessage & vbNullChar
    
    'Broadcast message via API
    If NetMessageBufferSend(ByVal 0&, byteToName(0), byteFromName(0), MessageToSend(0), UBound(MessageToSend)) = NERR_Success Then
        'Return True if it worked
        BroadcastMessage = True
    End If
    
End Function


