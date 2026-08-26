VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NAT Ping Process"
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim ParentDirectory As String
    Dim s As String
    Dim DomainName As String
    Dim Computername As String
    Dim isUp As Boolean
    Dim IPAddress As String
    Dim l As Long
    
    On Error Resume Next
    
    If App.PrevInstance Then End
    
    Me.Show
    Me.Refresh
    
    ParentDirectory = "C:\NAT"
    DomainName = UCase(Command$)
    
    s = Dir(ParentDirectory & "\" & DomainName & "\*.na1", vbNormal)
    
    Do Until s = ""
        If LCase(Right(s, 4)) <> ".nat" Then
            Computername = Left(s, InStr(1, s, ".") - 1)
            lblInfo = "Pinging " & Computername
            Me.Refresh
            
            IPAddress = ""
            IPAddress = GetIPAddress(Computername)
            
            isUp = Ping(IPAddress)
            Debug.Print Computername, IPAddress, isUp
            
            If isUp Then
                lblInfo.Caption = Computername & " (" & IPAddress & ") found"
                Me.Refresh
                Name ParentDirectory & "\" & DomainName & "\" & s As ParentDirectory & "\" & DomainName & "\" & Computername & ".na2"
            Else
                lblInfo.Caption = Computername & " NOT found"
                Me.Refresh
                Name ParentDirectory & "\" & DomainName & "\" & s As ParentDirectory & "\" & DomainName & "\" & Computername & ".na1"
            End If
        End If
        Me.Refresh
        
        s = Dir
    Loop
    End
End Sub
