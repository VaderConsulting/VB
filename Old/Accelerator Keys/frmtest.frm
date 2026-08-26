VERSION 5.00
Object = "{69939685-F9DE-11D2-8E21-44910FC10000}#2.0#0"; "vbalHkCt.ocx"
Begin VB.Form frmAccelTest 
   Caption         =   "Accelerator Test"
   ClientHeight    =   3705
   ClientLeft      =   4230
   ClientTop       =   3240
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5760
   Begin VB.ListBox lstAcceleratorEvents 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   5535
   End
   Begin vbalHkCt.vbalHookControl ctlAccelerator 
      Left            =   4920
      Top             =   3060
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmTest.frx":1272
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   0
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGoTOP 
      Caption         =   "&Go"
      Begin VB.Menu mnuGo 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAccelTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)
Dim sActive As String
   
   ' Show the accelerator we got and also whether the form is currently
   ' the system active form:
   If Not ctlAccelerator.IsActive Then
      sActive = "[InActive]"
   End If
   lstAcceleratorEvents.AddItem "Accelerator: " & ctlAccelerator.Key(nIndex) & " " & sActive & ""
   lstAcceleratorEvents.ListIndex = lstAcceleratorEvents.NewIndex
   
End Sub

Private Sub Form_Load()
   
   ' Create some accelerators at runtime:
   Load mnuEdit(1)
   mnuEdit(1).Visible = True
   mnuEdit(1).Caption = "&Copy" & vbTab & "Ctrl+C"
   Load mnuEdit(2)
   mnuEdit(2).Visible = True
   mnuEdit(2).Caption = "&Paste" & vbTab & "Ctrl+V"
   
   mnuGo(0).Caption = "&Back" & vbTab & "Alt+Left Arrow"
   Load mnuGo(1)
   mnuGo(1).Visible = True
   mnuGo(1).Caption = "&Next" & vbTab & "Alt+Right Arrow"
   Load mnuGo(2)
   mnuGo(2).Visible = True
   mnuGo(2).Caption = "&Up One Level"
   Load mnuGo(3)
   mnuGo(3).Visible = True
   mnuGo(3).Caption = "-"
   Load mnuGo(4)
   mnuGo(4).Visible = True
   mnuGo(4).Caption = "&Home Page" & vbTab & "Alt+Home"
   Load mnuGo(5)
   mnuGo(5).Visible = True
   mnuGo(5).Caption = "&vbAccelerator..." & vbTab & "Alt+Enter"
   
   
   
   With ctlAccelerator
      .AddAccelerator vbKeyC, vbCtrlMask, "COPY"
      .AddAccelerator vbKeyV, vbCtrlMask, "PASTE"
      .AddAccelerator vbKeyLeft, vbAltMask, "BACK"
      .AddAccelerator vbKeyRight, vbAltMask, "NEXT"
      .AddAccelerator vbKeyHome, vbAltMask, "HOME"
      .AddAccelerator vbKeyReturn, vbAltMask, "VBACCELERATOR"
      .Enabled = True
   End With
   
End Sub

Private Sub mnuEdit_Click(Index As Integer)
   ' Menu Event:
   Debug.Print "MenuEdit_Click " & Index
End Sub
