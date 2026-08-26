VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Application Execution"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3720
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtCommand 
      Height          =   885
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -240
      X2              =   5880
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line1 
      X1              =   -240
      X2              =   5880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblCommand 
      Caption         =   "Command"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblHostname 
      Caption         =   "Hostname"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExecute_Click()
    Dim Hostname As String
    Dim oServer As SWbemServices
    Dim oProcess As Object
    Dim result As Long
    Dim procID As Long
    Dim wmiLocator As SWbemLocator
    Dim E As Long
    Hostname = txtHostname
    Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
    On Error Resume Next
    Set oServer = wmiLocator.ConnectServer(Hostname, "root/CIMV2", txtDomain & "\" & txtUsername, txtPassword)
    E = Err.Number
    On Error GoTo 0
    If E = 0 Then
    Set oProcess = oServer.Get("Win32_Process")
    result = oProcess.Create(txtCommand.Text, Null, Null, procID)
    If result = 0 Then
        txtResult = "Process created" & vbCrLf & "Create method returns: " & result & vbCrLf & "Process Id returned " & procID
    Else
        txtResult = "Process creation failed, code " & result
    End If
    Else
        txtResult = "Error " & E
    End If
    Set oServer = Nothing
    Set oProcess = Nothing
End Sub

Private Sub Form_Load()
    txtUsername = Environ("username")
    txtDomain = Environ("userdomain")
End Sub
