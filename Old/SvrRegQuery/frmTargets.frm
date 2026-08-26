VERSION 5.00
Begin VB.Form frmTargets 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Targets"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtTarget 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.ListBox lstTargets 
      Height          =   2595
      ItemData        =   "frmTargets.frx":0000
      Left            =   120
      List            =   "frmTargets.frx":0002
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblTarget 
      Caption         =   "Target:"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblTargets 
      Alignment       =   2  'Center
      Caption         =   "Targets"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   5620
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   5620
      Y1              =   3135
      Y2              =   3135
   End
End
Attribute VB_Name = "frmTargets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  If Trim(txtTarget) = "" Then Exit Sub
  lstTargets.AddItem Trim(txtTarget)
  txtTarget = ""
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  frmMain.lstTargets.Clear
  For lp = 0 To lstTargets.ListCount
    frmMain.lstTargets.AddItem lstTargets.List(lp)
  Next lp
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  Dim colRemove As New Collection
  If Trim(txtTarget) = "" Then Exit Sub
  For lp = 0 To lstTargets.ListCount
    If lstTargets.List(lp) = Trim(txtTarget) Then
      colRemove.Add lp, Trim(txtTarget)
    End If
  Next lp
  For lp = 1 To colRemove.Count
    lstTargets.RemoveItem colRemove.Item(lp)
  Next lp
End Sub

Private Sub Form_Load()
  lstTargets.Clear
  For lp = 0 To frmMain.lstTargets.ListCount
    If Trim(frmMain.lstTargets.List(lp)) <> "" Then
      lstTargets.AddItem frmMain.lstTargets.List(lp)
    End If
  Next lp
End Sub

Private Sub lstTargets_Click()
  If lstTargets.ListCount > 0 Then
    txtTarget = lstTargets.List(lstTargets.ListIndex)
  End If
End Sub
