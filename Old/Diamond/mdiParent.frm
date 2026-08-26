VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{0053FB10-C707-11D1-87C9-000000000000}#6.0#0"; "Barcode.ocx"
Begin VB.MDIForm mdiParent 
   BackColor       =   &H8000000C&
   Caption         =   "Warranty"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9765
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Barcode.cbbBarcode barScan 
      Left            =   120
      Top             =   600
      _ExtentX        =   609
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PortOpen        =   -1  'True
   End
   Begin VB.Timer tmrSecond 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   17639
            MinWidth        =   17639
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "0:31"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Day"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "30/06/98"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current Date"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mdiParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    frmFlow.Left = 0
    frmFlow.Top = 0
    frmFlow.Height = mdiParent.Height - 1120
    frmFlow.Show
    frmData.Left = frmFlow.Left + frmFlow.Width
    frmData.Top = 0
    frmData.Width = mdiParent.Width - frmFlow.Width - 180
    frmData.Height = frmFlow.Height
    frmData.Show
    'sbrStatus.Left = mdiParent.Width - sbrStatus.Width - 180
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub tbrNav_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1
            reply = MsgBox("Cancel Operation?", vbQuestion + vbYesNo)
        Case 3
            gStepNo = gStepNo - 1
        Case 5
            gStepNo = gStepNo + 1
        Case 7
            gStepNo = gStepMax
    End Select
    StepChange
End Sub

Private Sub mnuSales_Click()
    frmSales.Show
End Sub

Private Sub tmrSecond_Timer()
    Static sDate
    If WeekDay(Now) <> sDate Then
        sDate = WeekDay(Now)
        sbrStatus.Panels(3).Text = gDay(WeekDay(Now))
    End If
End Sub
