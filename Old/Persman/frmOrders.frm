VERSION 5.00
Begin VB.Form frmOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Orders"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3735
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   400
      Left            =   1920
      TabIndex        =   2
      Top             =   5120
      Width           =   912
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   400
      Left            =   896
      TabIndex        =   1
      Top             =   5120
      Width           =   912
   End
   Begin VB.FileListBox filOrders 
      Height          =   4770
      Left            =   128
      Pattern         =   "*.DOC"
      TabIndex        =   0
      Top             =   128
      Width           =   3472
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    If filOrders.ListIndex = -1 Then Exit Sub
    file = gWordViewPath & " " & Chr$(34) & filOrders.Path & "\" & filOrders.List(filOrders.ListIndex) & Chr$(34)
    Shell file, vbNormalFocus
End Sub

Private Sub filOrders_DblClick()
    file = gWordViewPath & " " & Chr$(34) & filOrders.Path & "\" & filOrders.List(filOrders.ListIndex) & Chr$(34)
    Shell file, vbNormalFocus
End Sub

Private Sub Form_Load()
    filOrders.Path = gOrderPath
End Sub
