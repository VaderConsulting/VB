VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHalf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gazette Translator"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10200
   Icon            =   "frmGazette.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   10200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Prev Field"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdFull 
      Caption         =   "Full Page"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next field"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskText 
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   327681
      PromptChar      =   "_"
   End
   Begin ComctlLib.Slider sldTag 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   5953
      _Version        =   327682
      Orientation     =   1
      Max             =   30
      TickStyle       =   3
   End
   Begin VB.ListBox lstTag 
      Height          =   3180
      ItemData        =   "frmGazette.frx":0442
      Left            =   360
      List            =   "frmGazette.frx":0444
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin ComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3795
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   7673
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "19:37"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtText 
      Height          =   2775
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Label lblLabel 
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuAccess 
            Caption         =   "Access"
         End
         Begin VB.Menu mnuText 
            Caption         =   "Text"
         End
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmHalf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Label(30) As String
    Dim Multi(30) As Boolean
    Dim Mask(30) As String
    Dim Length(30) As Integer
    Const vbNull = ""
    Public gModified
    
Private Sub cmdFull_Click()
    frmFull.Show
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If lstTag.ListIndex < lstTag.ListCount Then lstTag.ListIndex = lstTag.ListIndex + 1
End Sub

Private Sub cmdPrevious_Click()
    If lstTag.ListIndex > 0 Then lstTag.ListIndex = lstTag.ListIndex - 1
End Sub

Private Sub Form_Activate()
    lstTag.AddItem "Name Phone number"
    lstTag.AddItem "Portfolio"
    lstTag.AddItem "Department"
    lstTag.AddItem "Name Date"
    lstTag.AddItem "Gazette date"
    lstTag.AddItem "P&S code"
    lstTag.AddItem "Reference number"
    lstTag.AddItem "Contract date"
    lstTag.AddItem "Value"
    lstTag.AddItem "Overseas country"
    lstTag.AddItem "Supply agreement"
    lstTag.AddItem "Supplier name"
    lstTag.AddItem "Supplier locality"
    lstTag.AddItem "Supplier's state"
    lstTag.AddItem "Supplier's postcode"
    lstTag.AddItem "Supplier's postal address"
    lstTag.AddItem "Supplier A.C.N."
    lstTag.AddItem "P&S Code"
    lstTag.AddItem "Office postcode"
    lstTag.AddItem "Description"
    lstTag.AddItem "Division"
    lstTag.AddItem "Branch"
    lstTag.AddItem "JA, JX, KS, KX"
    Label(0) = "Name {space} Phone number (80 characters)"
    Length(0) = 64
    Label(1) = "Name (30 characters)"
    Length(1) = 30
    Label(2) = "Cost (xxxxxxxxx.xx)"
    Mask(2) = "#########.##"
    Label(3) = "Name (10 characters)"
    Length(3) = 10
    Label(4) = "Item (Multi-line)"
    Multi(4) = True
    sldTag.Max = lstTag.ListCount - 1
    lstTag.ListIndex = 0
End Sub

Private Sub lstTag_Click()
    sldTag.Value = lstTag.ListIndex
    lblLabel = Label(lstTag.ListIndex)
    If Multi(lstTag.ListIndex) = True Then
        mskText.Visible = False
        txtText.Visible = True
        txtText.SetFocus
    Else
        mskText.Visible = True
        txtText.Visible = False
        mskText.SetFocus
    End If
    If Mask(lstTag.ListIndex) <> "" Then
        mskText.Mask = Mask(lstTag.ListIndex)
    Else
        mskText.Mask = ""
        mskText = ""
    End If
    If Length(lstTag.ListIndex) <> 0 Then
        mskText.MaxLength = Length(lstTag.ListIndex)
    Else
        mskText.MaxLength = 64
    End If
    If lstTag.ListIndex = lstTag.ListCount - 1 Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    If lstTag.ListIndex > 0 Then
        cmdPrevious.Enabled = True
    Else
        cmdPrevious.Enabled = False
    End If
End Sub

Private Sub mnuExit_Click()
    If gModified = True Then
        response = MsgBox("You will lose all data if you exit." & vbCrLf & "Save data first?", vbCritical + vbYesNoCancel, "Exit and lose changes")
        If response = vbNo Then End
        If response = vbCancel Then Exit Sub
        If response = vbYes Then ' save data first
            Call mnuSave_Click
            End
        End If
    Else
        response = MsgBox("Are you sure", vbInformation + vbYesNo, "Exit")
        If response = vbYes Then End
    End If
End Sub

Private Sub mnuSave_Click()
    MsgBox "Save data code goes here"
End Sub

Private Sub mskText_Change()
    gModified = True
End Sub
