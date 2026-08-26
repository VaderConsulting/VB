VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logged On Users"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5040
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView tvwComputers 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Text            =   "*"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdGetUsers 
      Caption         =   "Get Users"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstUsers 
      Height          =   3180
      Left            =   6600
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdGetComputers 
      Caption         =   "Retrieve"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.ListBox lstComputers 
      Height          =   2790
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label lblFilter 
      Caption         =   "Filter:"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetComputers_Click()
    Dim iadsCont As IADsContainer
    Dim oComputer As IADsComputer
    Dim strFilter As String
    Dim n As Node
    Dim strDomain As String
    
    strDomain = Environ$("USERDOMAIN")
    
    lstComputers.Clear
    tvwComputers.Nodes.Clear
    
    tvwComputers.Nodes.Add , , strDomain, strDomain
    
    Set iadsCont = GetObject("WinNT://" & strDomain)
    iadsCont.Filter = Array("computer")
    
    For Each oComputer In iadsCont
        If oComputer.Name Like txtFilter.Text Then
            lstComputers.AddItem oComputer.Name
            tvwComputers.Nodes.Add strDomain, tvwChild, oComputer.Name, oComputer.Name
        End If
    Next

End Sub

Private Sub cmdGetUsers_Click()
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colComputer
    Dim objComputer
    Dim i As Integer
    
    'If lstComputers.ListIndex = -1 Then Exit Sub
    
    lstUsers.Clear
    
    On Error Resume Next
    
    'strComputer = lstComputers.List(lstComputers.ListIndex)
    For i = 0 To lstComputers.ListCount - 1
        strComputer = lstComputers.List(i)
        UpdateInfo "Connecting to " & strComputer
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        If Err.Number = 0 Then
            Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
            If Err.Number = 0 Then
                For Each objComputer In colComputer
                    lstUsers.AddItem objComputer.UserName
                    tvwComputers.Nodes.Add strComputer, tvwChild, objComputer.UserName, objComputer.UserName
                Next
                UpdateInfo "Idle"
            Else
                UpdateInfo "Error connecting to WMI on " & strComputer
            End If
            Set colComputer = Nothing
        Else
            UpdateInfo "Error connecting to " & strComputer
        End If
        Set objWMIService = Nothing
    Next i
    Err.Clear
End Sub

Sub UpdateInfo(strMessage As String)
    frmMain.lblInfo.Caption = strMessage
    frmMain.Refresh
End Sub

