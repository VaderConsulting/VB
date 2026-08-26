VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassNetUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long

Const WN_Success = &H0
Const WN_Not_Supported = &H1
Const WN_Net_Error = &H2
Const WN_Bad_Pointer = &H4
Const WN_Bad_NetName = &H32
Const WN_Bad_Password = &H6
Const WN_Bad_Localname = &H33
Const WN_Access_Denied = &H7
Const WN_Out_Of_Memory = &HB
Const WN_Already_Connected = &H34
 

'-- Error number and message
Public ErrorNum     As Long
Public ErrorMsg     As String
Public rc As Long
Private Const ERROR_NO_CONNECTION = 8
Private Const ERROR_NO_DISCONNECT = 9

Public Sub Connect(sDrive, sService As String)
On Error GoTo Err_Connect
    Me.ErrorNum = 0
    Me.ErrorMsg = ""
    rc = WNetAddConnection(sService & Chr(0), "" & Chr(0), sDrive & Chr(0))
    If rc <> 0 Then GoTo Err_Connect
    Exit Sub
Err_Connect:
    Me.ErrorNum = rc
    Me.ErrorMsg = WnetError(rc)
End Sub

Public Sub DisConnect(sDrive As String)
On Error GoTo Err_DisConnect
    Me.ErrorNum = 0
    Me.ErrorMsg = ""
    rc = WNetCancelConnection(sDrive + Chr(0), 0)
    If rc <> 0 Then GoTo Err_DisConnect
    
    Exit Sub
Err_DisConnect:
    Me.ErrorNum = rc
    Me.ErrorMsg = WnetError(rc)
End Sub

Private Function WnetError(Errcode As Long) As String

Select Case Errcode
         Case WN_Not_Supported:
            WnetError = "Function is not supported."
         Case WN_Out_Of_Memory:
            WnetError = "Out of Memory."
         Case WN_Net_Error:
            WnetError = "An error occurred on the network."
         Case WN_Bad_Pointer:
            WnetError = "The Pointer was Invalid."
         Case WN_Bad_NetName:
            WnetError = "Invalid Network Resource Name."
         Case WN_Bad_Password:
            WnetError = "The Password was Invalid."
         Case WN_Bad_Localname:
            WnetError = "The local device name was invalid."
         Case WN_Access_Denied:
            WnetError = "A security violation occurred."
         Case WN_Already_Connected:
            WnetError = "The local device was connected to a remote resource."
         Case Else:
            WnetError = "Unrecognized Error " + Str(Errcode) + "."
      End Select
End Function

