VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form ShowUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show Users"
   ClientHeight    =   2955
   ClientLeft      =   2835
   ClientTop       =   3255
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2955
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lvwUsers 
      Height          =   2115
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3731
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   327682
   End
End
Attribute VB_Name = "ShowUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim iShowingType    As Integer
Dim psServer        As String
Dim foServer        As NOblets.CServer
Property Let Server(sServer As String)
    psServer = sServer
    Set foServer = TreeNet.fcolNodes(sServer)
End Property

Property Get ShowingType() As Integer
    ShowingType = iShowingType
End Property

Property Let ShowingType(ShowType As Integer)
    iShowingType = ShowType
    
    ' Create an object variable for the ColumnHeader object.
    Dim clmX    As ColumnHeader
    
    ' Add ColumnHeaders.  The width of the columns is the width
    ' of the control divided by the number of ColumnHeader objects.
    Select Case iShowingType
        Case SHOWING_USERS
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "User", lvwUsers.Width * 0.25)
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Comment", lvwUsers.Width * 0.65)
        Case SHOWING_GROUPS
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Group", lvwUsers.Width * 0.25)
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Comment", lvwUsers.Width * 0.65)
        Case SHOWING_SERVICES
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Name", lvwUsers.Width * 0.25)
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Display Name", lvwUsers.Width * 0.365)
            Set clmX = lvwUsers.ColumnHeaders.Add(, , "Status", lvwUsers.Width * 0.2)
    End Select
    
End Property


Private Sub Form_Load()
    lvwUsers.View = lvwReport ' Set View property to Report.
End Sub


Private Sub lvwUsers_DblClick()
    If iShowingType = SHOWING_SERVICES Then
        Call ShowServiceInfo
    ElseIf iShowingType = SHOWING_USERS Then
        'Call ShowUserinfo
    ElseIf iShowingType = SHOWING_GROUPS Then
        'Call ShowGroupinfo
    End If
End Sub


Private Sub ShowServiceInfo()
    Dim oLI         As ListItem
    Dim msg         As String
    Dim oService    As NOblets.CService
    
    Set oLI = lvwUsers.SelectedItem
    Set oService = foServer.Service(oLI.Key)
    With oService
        msg = "Executable Name: " & .EXEName & vbCrLf _
            & "Startup Type: " & .StartTypeString & vbCrLf _
            & "Start as User: " & .StartUser
    End With
    MsgBox msg, vbOKOnly, "Service " & oService.Name
End Sub


