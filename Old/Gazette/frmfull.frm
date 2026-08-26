VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFull 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gazette Translator"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13695
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.Slider sldProgress 
      Height          =   345
      Left            =   1920
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8520
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   609
      _Version        =   327682
      Max             =   22
      SelStart        =   1
      Value           =   1
   End
   Begin VB.OptionButton optCode 
      BackColor       =   &H80000000&
      Caption         =   "JA"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   24
      Top             =   6240
      Width           =   615
   End
   Begin VB.OptionButton optCode 
      BackColor       =   &H80000000&
      Caption         =   "JX"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   25
      Top             =   6240
      Width           =   615
   End
   Begin VB.OptionButton optCode 
      BackColor       =   &H80000000&
      Caption         =   "KX"
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   27
      Top             =   6600
      Width           =   615
   End
   Begin VB.OptionButton optCode 
      BackColor       =   &H80000000&
      Caption         =   "KS"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   26
      Top             =   6600
      Width           =   615
   End
   Begin VB.Frame fmeLegend 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Legend"
      Height          =   1095
      Left            =   6360
      TabIndex        =   55
      Top             =   6720
      Width           =   2775
      Begin VB.Label lblLocked 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Locked"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOptional 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Optional"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblmandatory 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mandatory"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Data datPSCode 
      Caption         =   "P & S Codes"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\Gazette.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "P&S Codes"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "Z"
      DataSource      =   "datData"
      Height          =   405
      Index           =   22
      Left            =   3000
      TabIndex        =   28
      ToolTipText     =   "Z"
      Top             =   6360
      Width           =   375
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "G"
      DataSource      =   "datData"
      Height          =   285
      Index           =   21
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "G"
      Top             =   5880
      Width           =   12015
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "F"
      DataSource      =   "datData"
      Height          =   285
      Index           =   20
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "F"
      Top             =   5520
      Width           =   12015
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "EE"
      DataSource      =   "datData"
      Height          =   1725
      Index           =   19
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      ToolTipText     =   "EE"
      Top             =   3720
      Width           =   12015
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   435
      Left            =   12480
      TabIndex        =   29
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "E2"
      DataSource      =   "datData"
      Height          =   285
      Index           =   18
      Left            =   6720
      TabIndex        =   20
      ToolTipText     =   "E2"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "DZ"
      DataSource      =   "datData"
      Height          =   285
      Index           =   17
      Left            =   4800
      TabIndex        =   19
      ToolTipText     =   "DZ"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DY"
      DataSource      =   "datData"
      Height          =   285
      Index           =   16
      Left            =   1560
      TabIndex        =   18
      ToolTipText     =   "DY"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DX"
      DataSource      =   "datData"
      Height          =   285
      Index           =   15
      Left            =   1560
      TabIndex        =   17
      ToolTipText     =   "DX"
      Top             =   3000
      Width           =   12015
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DW"
      DataSource      =   "datData"
      Height          =   285
      Index           =   14
      Left            =   11040
      TabIndex        =   16
      ToolTipText     =   "DW"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "DV"
      DataSource      =   "datData"
      Height          =   285
      Index           =   13
      Left            =   6720
      TabIndex        =   15
      ToolTipText     =   "DV"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "DU"
      DataSource      =   "datData"
      Height          =   285
      Index           =   12
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "DU"
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "DT"
      DataSource      =   "datData"
      Height          =   285
      Index           =   11
      Left            =   6720
      TabIndex        =   13
      ToolTipText     =   "DT"
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DS"
      DataSource      =   "datData"
      Height          =   285
      Index           =   10
      Left            =   1560
      TabIndex        =   12
      ToolTipText     =   "DS"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DP"
      DataSource      =   "datData"
      Height          =   285
      Index           =   9
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "DP"
      Top             =   1920
      Width           =   10095
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "DO"
      DataSource      =   "datData"
      Height          =   285
      Index           =   8
      Left            =   9600
      TabIndex        =   10
      ToolTipText     =   "DO"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "DN"
      DataSource      =   "datData"
      Height          =   285
      Index           =   7
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "DN"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DM"
      DataSource      =   "datData"
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   8
      ToolTipText     =   "DM"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DE"
      DataSource      =   "datData"
      Height          =   285
      Index           =   5
      Left            =   13200
      TabIndex        =   7
      ToolTipText     =   "DE"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DC"
      DataSource      =   "datData"
      Height          =   285
      Index           =   4
      Left            =   9600
      TabIndex        =   6
      ToolTipText     =   "DC"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080C0FF&
      DataField       =   "DA"
      DataSource      =   "datData"
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "DA"
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "D"
      DataSource      =   "datData"
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "D"
      Top             =   840
      Width           =   12015
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      DataField       =   "C"
      DataSource      =   "datData"
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "C"
      Top             =   480
      Width           =   12015
   End
   Begin MSMask.MaskEdBox mskDate1 
      Height          =   255
      Left            =   12600
      TabIndex        =   1
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   327681
      BackColor       =   8438015
      MaxLength       =   10
      Format          =   "dd.mm.yyyy"
      Mask            =   "##.##.####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      DataField       =   "B"
      DataSource      =   "datData"
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "B"
      Top             =   120
      Width           =   10095
   End
   Begin VB.Data datData 
      Caption         =   "Gazette"
      Connect         =   "Access"
      DatabaseName    =   "C:\Temp\Gazette.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Data"
      Top             =   8520
      Width           =   1935
   End
   Begin VB.CommandButton cmdHalf 
      Caption         =   "Half Page"
      Height          =   435
      Left            =   12480
      TabIndex        =   30
      Top             =   6720
      Width           =   1095
   End
   Begin MSMask.MaskEdBox mskDate2 
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   327681
      BackColor       =   8438015
      MaxLength       =   10
      Format          =   "dd.mm.yyyy"
      Mask            =   "##.##.####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblG 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
      Height          =   255
      Left            =   120
      TabIndex        =   54
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label lblF 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label lblEE 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblE2 
      BackStyle       =   0  'Transparent
      Caption         =   "Office Postcode"
      Height          =   255
      Left            =   5400
      TabIndex        =   51
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblDZ 
      BackStyle       =   0  'Transparent
      Caption         =   "P && S Code"
      Height          =   255
      Left            =   3720
      TabIndex        =   50
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDY 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier A.C.N."
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblDX 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Address"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblDW 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Postcode"
      Height          =   255
      Left            =   9600
      TabIndex        =   47
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblDV 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier state"
      Height          =   255
      Left            =   5280
      TabIndex        =   46
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblDU 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Locality"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblDT 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
      Height          =   255
      Left            =   5280
      TabIndex        =   44
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblDS 
      BackStyle       =   0  'Transparent
      Caption         =   "Supply Agreement"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblDP 
      BackStyle       =   0  'Transparent
      Caption         =   "Overseas Country"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblDO 
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   255
      Left            =   8520
      TabIndex        =   41
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblGazettedate 
      BackStyle       =   0  'Transparent
      Caption         =   "Gazette date"
      Height          =   255
      Left            =   8520
      TabIndex        =   40
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblPSCode 
      BackStyle       =   0  'Transparent
      Caption         =   "P&&S Code"
      Height          =   255
      Left            =   12360
      TabIndex        =   39
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblDM 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref Number"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblDN 
      BackStyle       =   0  'Transparent
      Caption         =   "Contract date"
      Height          =   255
      Left            =   6240
      TabIndex        =   37
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblDate2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblDA 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "Portfolio"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDate1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   12000
      TabIndex        =   32
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmFull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHalf_Click()
    frmHalf.Show
    Unload Me
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    If Len(txtB.Text) > 69 Then KeyAscii = 0
End Sub

Private Sub txtC_KeyPress(KeyAscii As Integer)
    If Len(txtC.Text) > 99 Then KeyAscii = 0
End Sub

Private Sub cmdRefresh_Click()
    frmFull.Refresh
End Sub

Private Sub optCode_Click(Index As Integer)
    txtField(22).Text = optCode(Index).Caption
End Sub

Private Sub sldProgress_Click()
    txtField(sldProgress.Value).SetFocus
End Sub

Private Sub txtField_GotFocus(Index As Integer)
    txtField(Index).SelStart = 0
    txtField(Index).SelLength = Len(txtField(Index))
    sldProgress.Value = Index
End Sub
