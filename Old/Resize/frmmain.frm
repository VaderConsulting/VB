VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Resize Example..."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwList 
      Height          =   3015
      Left            =   2265
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwTree 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5318
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Image imgDrag 
      DragMode        =   1  'Automatic
      Height          =   3015
      Left            =   2160
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const BORDER = 120

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "imgDrag" Then
        PlaceControls x, y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    ' Control Width (Listview control only)
    If (frmMain.Width - lvwList.Left - BORDER < frmMain.Width) And (frmMain.Width - lvwList.Left - BORDER > 0) Then
        lvwList.Width = frmMain.Width - lvwList.Left - BORDER
    End If
    'Control Height
    ControlHeight = frmMain.Height - (BORDER * 2)
    If ControlHeight < (BORDER * 2) Then ControlHeight = (BORDER * 2)
    If frmMain.Height < ControlHeight Then frmMain.Height = ControlHeight
    tvwTree.Height = ControlHeight
    lvwList.Height = ControlHeight
    imgDrag.Height = ControlHeight
End Sub

Private Sub imgDrag_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 9
End Sub

Private Sub lvwList_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "imgDrag" Then
        PlaceControls lvwList.Left + x, y
    End If
End Sub

Private Sub lvwList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = vbDefault
End Sub

Private Sub tvwTree_DragDrop(Source As Control, x As Single, y As Single)
    If Source.Name = "imgDrag" Then
        PlaceControls x, y
    End If
End Sub

Private Sub tvwTree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = vbDefault
End Sub

Sub PlaceControls(x As Single, y As Single)
    If (x - tvwTree.Left > BORDER) And (frmMain.Width - lvwList.Left - BORDER < frmMain.Width) And (frmMain.Width - lvwList.Left - BORDER > 0) Then
        imgDrag.Left = x
        tvwTree.Width = imgDrag.Left - BORDER
        lvwList.Left = imgDrag.Left + imgDrag.Width
        lvwList.Width = frmMain.Width - lvwList.Left - BORDER
    End If
End Sub
