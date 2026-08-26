VERSION 5.00
Begin VB.Form frmGroupInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Info"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblGroupname 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmGroupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim Group As IADsGroup
    Dim User As IADsUser
    
    Screen.MousePointer = vbHourglass
    Me.Refresh
    ' Set tab positions
    Call SetListTabStops(lstUsers.hWnd, 18)
    
    Set Group = GetObject("WinNT://" & gstrDomain & "/" & lblGroupname & ",group")
    
    On Error Resume Next
    For Each User In Group.Members
        lstUsers.AddItem User.Name & vbTab & User.Description
        DoEvents
    Next
    Set Group = Nothing
    Screen.MousePointer = vbDefault
End Sub

