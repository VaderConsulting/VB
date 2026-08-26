VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   3
      Left            =   6000
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Items"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":01DA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const LB_GETITEMRECT = &H198
Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMRECT = LVM_FIRST + 14
Private Const LVIR_BOUNDS = 0
Private Const LVIR_ICON = 1
Private Const LVIR_LABEL = 2
Private Const LVIR_SELECTBOUNDS = 3

Dim TTStd As New cToolTip
Dim TTBlnRed As New cToolTip
Dim TTBlnBlue As New cToolTip
Dim TTBlnGreen As New cToolTip

Private Sub Command1_Click()
   If Left$(Command1.Caption, 1) = "R" Then
      TTBlnRed.DelToolTip Command1.hwnd
      Command1.Caption = "Set TT on Me"
   Else
      TTBlnRed.SetToolTipObj Command1.hwnd, "Multiline and balloon" & vbCrLf & "ToolTip sample" & vbCrLf & "for Command Button", False
      Command1.Caption = "Remove TT from Me"
   End If
End Sub

Private Sub Form_Load()
  Dim rc As RECT, i As Long
  Command1.Caption = "Remove TT from Me"
  Option1(0).Caption = "Icon"
  Option1(1).Caption = "SmallIcon"
  Option1(2).Caption = "List"
  Option1(3).Caption = "Report"
  With TTBlnRed
       .BkColor = vbRed
       .TxtColor = vbBlack
       .DelayTime = 300
       .VisibleTime = 1000
       .TipWidth = 10
       .Style = ttStyleBalloon
       .SetToolTipObj Command1.hwnd, "Multiline and balloon" & vbCrLf & "ToolTip sample" & vbCrLf & "for Command Button", False
  End With
  With TTBlnBlue
       .BkColor = vbBlue
       .TxtColor = vbWhite
       .Style = ttStyleBalloon
       For i = 0 To 3
          .SetToolTipObj Option1(i).hwnd, "Choose ListView Style"
       Next i
  End With
  With TTBlnGreen
       .BkColor = vbGreen
       .TxtColor = vbBlack
       .Style = ttStyleBalloon
  End With
  With TTStd
       .Style = ttStyleStandard
       .DelayTime = 100
       .VisibleTime = 1000
  End With
  For i = 0 To 20
      List1.AddItem "This is item No. " & CStr(i)
      SendMessage List1.hwnd, LB_GETITEMRECT, i, rc
      TTStd.SetToolTipItem List1.hwnd, i, rc.Left, rc.Top, rc.Right, rc.Bottom, "Multiline and Centered long Tooltip style example" & vbCrLf & "for item No. " & CStr(i), True
  Next i
  For i = 1 To 30
      With ListView1.ListItems.Add(, , "ListView Item No. " & CStr(i))
        .Icon = 1
        .SmallIcon = 1
      End With
      rc.Left = LVIR_BOUNDS
      SendMessage ListView1.hwnd, LVM_GETITEMRECT, i - 1, rc
      TTBlnGreen.SetToolTipItem ListView1.hwnd, i - 1, rc.Left, rc.Top, rc.Right, rc.Bottom, "ListView Tooltip example" & vbCrLf & "for item No. " & CStr(i)
  Next
  
  Option1(0).Value = True
  'center
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set TTBlnRed = Nothing
  Set TTBlnBlue = Nothing
  Set TTStd = Nothing
End Sub

Private Sub List1_Scroll()
  Dim rc As RECT, i As Long
  For i = 0 To List1.ListCount - 1
      SendMessage List1.hwnd, LB_GETITEMRECT, i, rc
      TTStd.AjustItemRect List1.hwnd, i, rc.Left, rc.Top, rc.Right, rc.Bottom
  Next i
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim rc As RECT, i As Long
  For i = 1 To ListView1.ListItems.Count
      rc.Left = LVIR_BOUNDS
      SendMessage ListView1.hwnd, LVM_GETITEMRECT, i - 1, rc
      TTBlnGreen.AjustItemRect ListView1.hwnd, i - 1, rc.Left, rc.Top, rc.Right, rc.Bottom
  Next i
End Sub

Private Sub Option1_Click(Index As Integer)
   ListView1.View = Index
End Sub
