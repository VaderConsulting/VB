VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DHCP Test"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkTest 
      Caption         =   "Enable"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox lstHistory 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   5175
   End
   Begin VB.Timer tmrDHCP 
      Interval        =   5000
      Left            =   4920
      Top             =   2880
   End
   Begin VB.Label lblServerName 
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblServer 
      Caption         =   "Server:"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgError 
      Height          =   360
      Left            =   5520
      Picture         =   "frmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image imgOK 
      Height          =   375
      Left            =   5520
      Picture         =   "frmMain.frx":0454
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgRenew 
      Height          =   375
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgRelease 
      Height          =   375
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblRenew 
      Caption         =   "Renew Status"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblRelease 
      Caption         =   "Release status"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Success_NoReboot = 0        'Successful completion, no reboot required.
Private Const Success_Reboot = 1          'Successful completion, reboot required.
Private Const Not_Supported = 64          'Method not supported on this platform.
Private Const Unknown_Failure = 65        'Unknown failure.
Private Const Invalid_Mask = 66           'Invalid subnet mask.
Private Const Instance_Error = 67         'An error occurred while processing an instance that was returned.
Private Const Invalid_Parameter = 68      'Invalid input parameter.
Private Const Excessive_Gateways = 69     'More than five gateways specified.
Private Const Invalid_IP = 70             'Invalid IP address.
Private Const Invalid_Gateway = 71        'Invalid gateway IP address.
Private Const Error_Registry = 72         'An error occurred while accessing the registry for the requested information.
Private Const Invalid_Domain = 73         'Invalid domain name.
Private Const Invalid_Host = 74           'Invalid host name.
Private Const No_WINS = 75                'No primary/secondary WINS server defined.
Private Const Invalid_File = 76           'Invalid file.
Private Const Invalid_path = 77           'Invalid system path.
Private Const Copy_Failed = 78            'File copy failed.
Private Const InvalidSecurity = 79        'Invalid security parameter.
Private Const Error_IP = 80               'Unable to configure TCP/IP service.
Private Const Error_DHCP = 81             'Unable to configure DHCP service.
Private Const Error_Renew = 82            'Unable to renew DHCP lease.
Private Const Error_Release = 83          'Unable to release DHCP lease.
Private Const IP_Not_Enabled = 84         'IP not enabled on adapter.
Private Const IPX_Not_Enabled = 85        'IPX not enabled on adapter.
Private Const Bounds_Error = 86           'Frame/network number bounds error.
Private Const Invalid_Frame = 87          'Invalid frame type.
Private Const Invalid_Net_Number = 88     'Invalid network number.
Private Const Duplicate_Network = 89      'Duplicate network number.
Private Const Error_Parameter = 90        'Parameter out of bounds.
Private Const Access_Denied = 91          'Access denied.
Private Const Error_Memory = 92           'Out of memory.
Private Const Already_Exists = 93         'Already exists.
Private Const Not_Found = 94              'Path, file, or object not found.
Private Const Error_Service = 95          'Unable to notify service.
Private Const Error_DNS = 96              'Unable to notify DNS service.
Private Const Error_Interface = 97        'Interface not configurable.
Private Const Error_Lease = 98            'Not all DHCP leases could be released/renewed.
Private Const Error_Not_Enabled = 100     'DHCP not enabled on adapter.

Private Function ErrorString(ErrorNumber) As String

    Select Case ErrorNumber
      Case Success_NoReboot
        ErrorString = "Successful completion, no reboot required."
      Case Success_Reboot
        ErrorString = "Successful completion, reboot required."
      Case Not_Supported
        ErrorString = "Method not supported on this platform."
      Case Unknown_Failure
        ErrorString = "Unknown failure."
      Case Invalid_Mask
        ErrorString = "Invalid subnet mask."
      Case Instance_Error
        ErrorString = "An error occurred while processing an instance that was returned."
      Case Invalid_Parameter
        ErrorString = "Invalid input parameter."
      Case Excessive_Gateways
        ErrorString = "More than five gateways specified."
      Case Invalid_IP
        ErrorString = "Invalid IP address."
      Case Invalid_Gateway
        ErrorString = "Invalid gateway IP address."
      Case Error_Registry
        ErrorString = "An error occurred while accessing the registry for the requested information."
      Case Invalid_Domain
        ErrorString = "Invalid domain name."
      Case Invalid_Host
        ErrorString = "Invalid host name."
      Case No_WINS
        ErrorString = "No primary/secondary WINS server defined."
      Case Invalid_File
        ErrorString = "Invalid file."
      Case Invalid_path
        ErrorString = "Invalid system path."
      Case Copy_Failed
        ErrorString = "File copy failed."
      Case InvalidSecurity
        ErrorString = "Invalid security parameter."
      Case Error_IP
        ErrorString = "Unable to configure TCP/IP service."
      Case Error_DHCP
        ErrorString = "Unable to configure DHCP service."
      Case Error_Renew
        ErrorString = "Unable to renew DHCP lease."
      Case Error_Release
        ErrorString = "Unable to release DHCP lease."
      Case IP_Not_Enabled
        ErrorString = "IP not enabled on adapter."
      Case IPX_Not_Enabled
        ErrorString = "IPX not enabled on adapter."
      Case Bounds_Error
        ErrorString = "Frame/network number bounds error."
      Case Invalid_Frame
        ErrorString = "Invalid frame type."
      Case Invalid_Net_Number
        ErrorString = "Invalid network number."
      Case Duplicate_Network
        ErrorString = "Duplicate network number."
      Case Error_Parameter
        ErrorString = "Parameter out of bounds."
      Case Access_Denied = 91
        ErrorString = "Access denied."
      Case Error_Memory
        ErrorString = "Out of memory."
      Case Already_Exists
        ErrorString = "Already exists."
      Case Not_Found
        ErrorString = "Path, file, or object not found."
      Case Error_Service
        ErrorString = "Unable to notify service."
      Case Error_DNS
        ErrorString = "Unable to notify DNS service."
      Case Error_Interface
        ErrorString = "Interface not configurable."
      Case Error_Lease
        ErrorString = "Not all DHCP leases could be released/renewed."
      Case Error_Not_Enabled
        ErrorString = "DHCP not enabled on adapter."
    End Select

End Function

Private Sub LogEvent(strMessage As String)

    Open App.Path & "\DHCP.log" For Append As #1
    Print #1, strMessage
    Close 1

End Sub

Private Sub Release()

  Dim AdapterConfig
  Dim ReturnValue
  Dim strReturnValue

    On Error Resume Next

        Set AdapterConfig = GetObject("winmgmts:Win32_NetworkAdapterConfiguration")

        ReturnValue = AdapterConfig.ReleaseDHCPLeaseAll
        Select Case ReturnValue
          Case Success_NoReboot, Success_Reboot
            imgRelease.Picture = imgOK.Picture
            lstHistory.AddItem Date & " " & Time & " Release OK", 0
            LogEvent Date & " " & Time & " Release OK."
            'lblServerName.Caption = "NONE"
          Case Else
            imgRelease.Picture = imgError.Picture
            lstHistory.AddItem Date & " " & Time & " Release Error (" & ErrorString(ReturnValue) & ")", 0
            LogEvent Date & " " & Time & " Release Error (" & ErrorString(ReturnValue) & ")"
        End Select

        Set AdapterConfig = Nothing

        Me.Refresh
    On Error GoTo 0
End Sub

Private Sub Renew()

  Dim AdapterConfig
  Dim ReturnValue
  Dim strReturnValue

    On Error Resume Next

        Set AdapterConfig = GetObject("winmgmts:Win32_NetworkAdapterConfiguration")

        ReturnValue = AdapterConfig.RenewDHCPleaseAll
        Select Case ReturnValue
          Case Success_NoReboot, Success_Reboot
            imgRenew.Picture = imgOK.Picture
            lstHistory.AddItem Date & " " & Time & " Renew OK.", 0
            LogEvent Date & " " & Time & " Renew OK."
            'lblServerName.Caption = AdapterConfig.DHCPServer
          Case Else
            imgRenew.Picture = imgError.Picture
            lstHistory.AddItem Date & " " & Time & " Renew Error (" & ErrorString(ReturnValue) & ")", 0
            LogEvent Date & " " & Time & " Renew Error (" & ErrorString(ReturnValue) & ")"
        End Select

        Set AdapterConfig = Nothing

        Me.Refresh
    On Error GoTo 0

End Sub

Private Sub tmrDHCP_Timer()

    If chkTest.Value = vbChecked Then
        tmrDHCP.Enabled = False
        Release
        Renew
        tmrDHCP.Enabled = True
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:46) 40 + 166 = 206 Lines
