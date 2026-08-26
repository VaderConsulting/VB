VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Group membership"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   255
      Left            =   10320
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   8280
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.ListBox lstGroups 
      Height          =   3375
      Left            =   6360
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
   Begin VB.ListBox lstDomainGroups 
      Height          =   5910
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label lblDomainGroups 
      Alignment       =   2  'Center
      Caption         =   "Domain Groups"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label lblUsername 
      Caption         =   "User or Computer name"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLoggedonUsername 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public strUsername As String

Private Sub cmdQuery_Click()
    GetUserInfo txtUsername.Text
End Sub

Private Sub Form_Load()
    Dim oInfo As ADSystemInfo
    Dim oDomain As IADsContainer
    Dim Members As Variant
    Dim oGroup As IADsGroup
    
    Me.Show
    Me.Refresh
    
    
    lblLoggedonUsername.Caption = "Logged on user: " & Environ$("USERNAME")
    txtUsername = Environ$("USERNAME")
    GetUserInfo txtUsername
    
    Set oInfo = CreateObject("ADSystemInfo")
    
    Set oDomain = GetObject("WinNT://" & oInfo.DomainDNSName)
    
    oDomain.Filter = Array("group")
    
    For Each oGroup In oDomain
        If LCase(oGroup.Name) Like "right-usr-pf*" Or LCase(oGroup.Name) Like "right-usr-gf*" Or LCase(oGroup.Name) Like "????????_*" Then
            lstDomainGroups.AddItem oGroup.Name
        End If
        DoEvents
    Next
    
    Set oDomain = Nothing
    Set oGroup = Nothing
    Set oInfo = Nothing
End Sub

Private Sub GetUserInfo(strUsername As String)
    Dim oUser As IADsUser
    Dim oGroup As IADsGroup
    
    lstGroups.Clear
    
    DoEvents
    
    On Error Resume Next
    Set oUser = GetObject("WinNT://" & Environ$("USERDOMAIN") & "/" & strUsername)
    If Err.Number <> 0 Then
        Err.Clear
        Set oUser = GetObject("WinNT://" & Environ$("USERDOMAIN") & "/" & strUsername & "$")
    End If
    
    If Err.Number = 0 Then
        For Each oGroup In oUser.Groups
            lstGroups.AddItem oGroup.Name
            DoEvents
        Next
    End If
    
    Set oGroup = Nothing
    Set oUser = Nothing
End Sub
