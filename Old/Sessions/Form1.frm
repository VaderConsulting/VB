VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sessions"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView tvwSessions 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   12938
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim fso As IADsFileServiceOperations
  Dim ss As IADsCollection
  Dim n As Node
  
  Me.Show
  Me.Refresh
  
  Set fso = GetObject("WinNT://sperthxawa/LanmanServer")
  Set ss = fso.Sessions
  
  For Each resource In fso.Resources
    On Error Resume Next
    Set n = tvwSessions.Nodes.Add(, tvwparent, "K" & resource.Path, resource.Path)
    Debug.Print "Resource path: " & resource.Path
    For Each r In fso.Sessions
      If r.User = resource.User Then
        Debug.Print "  " & r.User
        Debug.Print "  " & r.Computer
        Debug.Print "  " & r.ConnectTime ' (minutes)
        Set n = tvwSessions.Nodes.Add("K" & resource.Path, tvwChild, "S" & r.Computer, r.Computer)
        Set n = tvwSessions.Nodes.Add("S" & r.Computer, tvwChild, "F" & r.User, r.User)
      End If
      tvwSessions.Refresh
    Next
    tvwSessions.Refresh
    DoEvents
  Next resource
End Sub
