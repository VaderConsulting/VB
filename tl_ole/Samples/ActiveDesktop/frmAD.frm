VERSION 5.00
Begin VB.Form frmAD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Active Desktop"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4545
      TabIndex        =   5
      Top             =   3675
      Width           =   1140
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   390
      Left            =   1290
      TabIndex        =   4
      Top             =   3675
      Width           =   1140
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
      Height          =   390
      Left            =   75
      TabIndex        =   3
      Top             =   3675
      Width           =   1140
   End
   Begin VB.ListBox lstComponents 
      Height          =   1185
      Left            =   75
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2445
      Width           =   5610
   End
   Begin VB.PictureBox picScreen 
      BackColor       =   &H80000001&
      Height          =   1950
      Left            =   1523
      ScaleHeight     =   1890
      ScaleMode       =   0  'User
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   105
      Width           =   2715
      Begin VB.PictureBox shpComp 
         DragMode        =   1  'Automatic
         Height          =   405
         Index           =   0
         Left            =   360
         ScaleHeight     =   345
         ScaleWidth      =   360
         TabIndex        =   6
         Top             =   375
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Components:"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   2205
      Width           =   930
   End
End
Attribute VB_Name = "frmAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_oAD As ActiveDesktop

Sub LoadComponents()
Dim lIdx As Long
Dim tCOMP As IE4COMPONENT
Dim lTotal As Long
Dim sName As String

   lstComponents.Clear
   
   For lIdx = 1 To shpComp.UBound
      Unload shpComp(lIdx)
   Next
   
   m_oAD.GetDesktopItemCount lTotal
   
   For lIdx = 0 To lTotal - 1
   
      tCOMP.dwSize = Len(tCOMP)
      tCOMP.cpPos.dwSize = Len(tCOMP.cpPos)
      
      m_oAD.GetDesktopItem lIdx, tCOMP
      
      sName = Left$(tCOMP.wszFriendlyName, InStr(tCOMP.wszFriendlyName, vbNullChar))
      
      If sName = vbNullChar Then sName = Left$(tCOMP.wszSource, InStr(tCOMP.wszSource, vbNullChar))
         
      lstComponents.AddItem sName
      lstComponents.ItemData(lstComponents.NewIndex) = tCOMP.dwID
   
      Load shpComp(lIdx + 1)
      
      With tCOMP.cpPos
         shpComp(lIdx + 1).Move .iLeft, .iTop, .dwWidth, .dwHeight
      End With
      
      shpComp(lIdx + 1).Visible = tCOMP.fChecked
      lstComponents.Selected(lstComponents.NewIndex) = tCOMP.fChecked
         
   Next
   
End Sub

Sub Copy2Arr(Arr() As Byte, ByVal str As String)
Dim lLen As Long

   lLen = UBound(Arr) - LBound(Arr) + 1
   
   If lLen > LenB(str) + 2 Then lLen = LenB(str) + 2
   
   MoveMemory Arr(LBound(Arr)), ByVal StrPtr(str), lLen
   
End Sub

Private Sub cmdAdd_Click()
Dim tCOMP As IE4COMPONENT
Dim sURL As String
Dim sName As String
   
   sURL = InputBox("Type the URL of the site you want to add:", "Add desktop component", "http://www.domaindlx.com/e_morcillo")
   If StrPtr(sURL) = 0 Then Exit Sub
   
   sName = InputBox("Type the name of the site:", "Add desktop component", "Edanmo's VB Page")
   If StrPtr(sName) = 0 Then Exit Sub
   
   tCOMP.dwSize = Len(tCOMP)
   tCOMP.cpPos.dwSize = Len(tCOMP.cpPos)
   tCOMP.fChecked = True
   tCOMP.iComponentType = COMP_TYPE_WEBSITE
   Copy2Arr tCOMP.wszSource, sURL
   Copy2Arr tCOMP.wszFriendlyName, sName
   
   If m_oAD.AddDesktopItem(tCOMP) = S_OK Then
      cmdApply.Enabled = True
      LoadComponents
   Else
      MsgBox "There item can't be added."
   End If
   
End Sub


Private Sub cmdApply_Click()

   m_oAD.ApplyChanges AD_APPLY_ALL
   cmdApply.Enabled = False
   
End Sub

Private Sub cmdRemove_Click()
Dim tCOMP As IE4COMPONENT

   If lstComponents.ListIndex <> -1 Then
   
      tCOMP.dwSize = Len(tCOMP)
      m_oAD.GetDesktopItemByID lstComponents.ItemData(lstComponents.ListIndex), tCOMP
      m_oAD.RemoveDesktopItem tCOMP
   
      cmdApply.Enabled = True
      LoadComponents
      
   End If
   
End Sub

Private Sub Form_Load()

   Set m_oAD = New ActiveDesktop
   
   picScreen.ScaleWidth = Screen.Width / Screen.TwipsPerPixelX
   picScreen.ScaleHeight = Screen.Height / Screen.TwipsPerPixelY
   
   LoadComponents
   
End Sub
Private Sub lstComponents_Click()
Dim lIdx As Long

   On Error Resume Next
   
   If lstComponents.ListIndex <> -1 Then
   
      For lIdx = 0 To shpComp.UBound
         shpComp(lIdx).BackColor = vbWindowBackground
      Next
      
      shpComp(lstComponents.ListIndex + 1).BackColor = vbHighlight
      
   End If
   
End Sub


Private Sub lstComponents_ItemCheck(Item As Integer)
Dim tCOMP As IE4COMPONENT

   ' Get the component data
   tCOMP.dwSize = Len(tCOMP)
   tCOMP.cpPos.dwSize = Len(tCOMP.cpPos)
   m_oAD.GetDesktopItemByID lstComponents.ItemData(Item), tCOMP
   
   ' Set the new state
   If lstComponents.Selected(Item) Then
      tCOMP.fChecked = 1
   Else
      tCOMP.fChecked = 0
   End If
   m_oAD.ModifyDesktopItem tCOMP, COMP_ELEM_CHECKED
   
   ' Show/Hide the shape in the screen picture
   shpComp(Item + 1).Visible = tCOMP.fChecked
   
   cmdApply.Enabled = True
   
End Sub
