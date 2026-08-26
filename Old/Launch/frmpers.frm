VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmPERS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERS Calculator"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7680
   Begin MSChart20Lib.MSChart chtData 
      Height          =   1335
      Left            =   120
      OleObjectBlob   =   "frmPERS.frx":0000
      TabIndex        =   62
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Frame fmeOptions 
      Caption         =   "Options"
      Height          =   975
      Left            =   4560
      TabIndex        =   59
      Top             =   5280
      Width           =   3015
      Begin VB.CheckBox chkChart 
         Caption         =   "Graph results"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         ToolTipText     =   "Show a Bar graph of the results"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "I am writing someone else's PERS"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "Enable help and other special options"
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdScores 
      Caption         =   "Explain scores"
      Height          =   375
      Left            =   4080
      TabIndex        =   58
      ToolTipText     =   "Show what these scores mean"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   57
      ToolTipText     =   "Save reults to H:"
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Clear all results"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Preview"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      ToolTipText     =   "Show print preview"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtFinal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   11
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   10
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   9
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   8
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   7
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox txtScore 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
   Begin ComctlLib.TabStrip tstRank 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      TabWidthStyle   =   2
      Style           =   1
      TabFixedWidth   =   1058
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SMN"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "AB"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "LS"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PO"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "CPO"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      ToolTipText     =   "Close"
      Top             =   6360
      Width           =   855
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   11
      Left            =   120
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   10
      Left            =   120
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00400000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   9
      Left            =   120
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   8
      Left            =   120
      Top             =   3480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   7
      Left            =   120
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   120
      Top             =   2760
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   120
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   120
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   120
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   120
      Top             =   1320
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape boxLegend 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblTotallbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   56
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   43
      Top             =   4560
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   42
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   41
      Top             =   3840
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   40
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   39
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   38
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   37
      Top             =   2400
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   36
      Top             =   2040
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   35
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   34
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   33
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblFinal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Final"
      Height          =   255
      Left            =   6960
      TabIndex        =   32
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      Height          =   255
      Left            =   5280
      TabIndex        =   30
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblCriteria 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   29
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label lblWeightings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Weighting"
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmPERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subLoad()
    If gRank = "SMN" Or gRank = "AB" Then
    Criteria = 6
        file = gDrive2 & "\bulletin\lan\SMN-AB.txt"
    End If
    
    If gRank = "LS" Then
        Criteria = 8
        file = gDrive2 & "\bulletin\lan\LS.txt"
    End If
    
    If gRank = "PO" Or gRank = "CPO" Then
        Criteria = 11
        file = gDrive2 & "\bulletin\lan\PO-CPO.txt"
    End If
    
    Open file For Input As #1
    For a = 0 To Criteria
        For b = 1 To 9
            Line Input #1, txtexplain
            gCriteria(a, b) = txtexplain
        Next b
    Next a
    Close 1
    If chkHelp.Value = 1 Then
        For a = 1 To 9
            frmScores.lblExplain(a) = gCriteria(0, a)
            frmScores.lblExplain(a).Refresh
        Next a
    End If
End Sub
Private Sub chkChart_Click()
    If chkChart.Value = 1 Then
        chtData.Visible = True
    Else
        chtData.Visible = False
    End If
End Sub

Private Sub chkHelp_Click()
    If chkHelp.Value = 1 Then
        Left = (Screen.Width - Width + frmScores.Width) / 2
        Top = (Screen.Height - Height) / 2
        frmScores.Show
    End If
    If chkHelp.Value <> 1 Then
        Left = (Screen.Width - Width) / 2
        Top = (Screen.Height - Height) / 2
        Unload frmScores
    End If
End Sub

Private Sub cmdClear_Click()
    For a = 0 To 11
        txtScore(a) = ""
        txtFinal(a) = ""
        lblTotal = ""
        ' only clear the column data if the column exists
        If chtData.ColumnCount > a Then
            chtData.Column = a + 1
            chtData.Data = 0
        End If
    Next a
    txtScore(0).SetFocus
    gSurname = ""
End Sub

Private Sub cmdDone_Click()
    frmOther.Show
    If chkHelp.Value = 1 Then Unload frmScores
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    erra = 0
    For a = 0 To 11
        If lblCriteria(a) <> "" And Val(txtScore(a)) = 0 Then erra = 1
    Next a
    If erra = 0 Then
        Do Until gSurname <> ""
            gSurname = InputBox("Please enter your surname")
        Loop
        frmScores.Hide
        frmPERS.Hide
        frmPERSResults.Show
    Else
        MsgBox "Please enter a score for all criteria", vbOKOnly + vbExclamation, "Error"
        txtScore(0).SetFocus
        txtScore(0).SelStart = 0
        txtScore(0).SelLength = Len(txtScore(0))
    End If
    
End Sub

Private Sub cmdSave_Click()
    MsgBox "Not implemented"
End Sub

Private Sub cmdScores_Click()
    frmScores.Hide
    frmPERS.Hide
    frmPrint.Show
End Sub

Private Sub Form_Activate()
    txtScore(0).SetFocus
    txtScore(0).SelStart = 0
    txtScore(0).SelLength = Len(txtScore(0))
    If chkHelp.Value = 1 Then
        frmScores.Left = frmPERS.Left - frmScores.Width
        frmScores.Top = frmPERS.Top
    End If
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    lblCriteria(0) = "Efficiency and Effectiveness"
    lblCriteria(1) = "Initiative"
    lblCriteria(2) = "Reliability"
    lblCriteria(3) = "Professional Knowledge"
    lblCriteria(4) = "Compliance with Orders and Acceptance of Authority"
    lblCriteria(5) = "Dress and Grooming"
    lblCriteria(6) = "Teamwork and Cooperation"
    txtWeight(0) = "3.22"
    txtWeight(1) = "2.68"
    txtWeight(2) = "3.21"
    txtWeight(3) = "5.35"
    txtWeight(4) = "3.75"
    txtWeight(5) = "1.08"
    txtWeight(6) = "2.94"
    txtWeight(7) = ""
    txtWeight(8) = ""
    txtWeight(9) = ""
    txtWeight(10) = ""
    txtWeight(11) = ""
    With chtData.Plot.Axis(VtChAxisIdY).ValueScale
        .Auto = False
        .MajorDivision = 0
        .MinorDivision = 0
        .Maximum = 9
        .Minimum = 0
    End With
    For a = 1 To chtData.ColumnCount
        chtData.Column = a
        chtData.Data = 0
    Next a
    gRank = "SMN"
    msg = "Current RAN policy does not allow us to focus on PERS scores," & vbCrLf
    msg = msg & "but rather on the individual's performance." & vbCrLf
    msg = msg & "Because of this, the author encourages the reportee" & vbCrLf
    msg = msg & "discuss the meaning of their PERS with Divisional staff."
    MsgBox msg, vbInformation + vbOKOnly, "Please note"
    subLoad
End Sub



Private Sub tstRank_Click()
    Rem clear existing surname
    gSurname = ""
    Select Case tstRank.SelectedItem.Index
    Case Is < 3
        If tstRank.SelectedItem.Index = 1 Then gRank = "SMN"
        If tstRank.SelectedItem.Index = 2 Then gRank = "AB"
        lblCriteria(0) = "Efficiency and Effectiveness"
        lblCriteria(1) = "Initiative"
        lblCriteria(2) = "Reliability"
        lblCriteria(3) = "Professional Knowledge"
        lblCriteria(4) = "Compliance with Orders and Acceptance of Authority"
        lblCriteria(5) = "Dress and Grooming"
        lblCriteria(6) = "Teamwork and Cooperation"
        lblCriteria(7) = ""
        txtScore(7).Visible = False
        txtWeight(7).Visible = False
        txtFinal(7).Visible = False
        boxLegend(7).Visible = False
        lblCriteria(8) = ""
        txtScore(8).Visible = False
        txtWeight(8).Visible = False
        txtFinal(8).Visible = False
        boxLegend(8).Visible = False
        lblCriteria(9) = ""
        txtScore(9).Visible = False
        txtWeight(9).Visible = False
        txtFinal(9).Visible = False
        boxLegend(9).Visible = False
        lblCriteria(10) = ""
        txtScore(10).Visible = False
        txtWeight(10).Visible = False
        txtFinal(10).Visible = False
        boxLegend(10).Visible = False
        lblCriteria(11) = ""
        txtScore(11).Visible = False
        txtWeight(11).Visible = False
        txtFinal(11).Visible = False
        boxLegend(11).Visible = False
        txtWeight(0) = "3.22"
        txtWeight(1) = "2.68"
        txtWeight(2) = "3.21"
        txtWeight(3) = "5.35"
        txtWeight(4) = "3.75"
        txtWeight(5) = "1.08"
        txtWeight(6) = "2.94"
        txtWeight(7) = ""
        txtWeight(8) = ""
        txtWeight(9) = ""
        txtWeight(10) = ""
        txtWeight(11) = ""
        chtData.ColumnCount = 7
    Case 3
        lblCriteria(0) = "Efficiency and Effectiveness"
        lblCriteria(1) = "Initiative"
        lblCriteria(2) = "Reliability"
        lblCriteria(3) = "Acceptance of Responsibility"
        lblCriteria(4) = "Professional Knowledge"
        lblCriteria(5) = "Compliance with Orders and Acceptance of Authority"
        lblCriteria(6) = "Dress and Grooming"
        lblCriteria(7) = "Supervising"
        txtScore(7).Visible = True
        txtWeight(7).Visible = True
        txtFinal(7).Visible = True
        boxLegend(7).Visible = True
        lblCriteria(8) = "Teamwork and Cooperation"
        txtScore(8).Visible = True
        txtWeight(8).Visible = True
        txtFinal(8).Visible = True
        boxLegend(8).Visible = True
        lblCriteria(9) = ""
        txtScore(9).Visible = False
        txtWeight(9).Visible = False
        txtFinal(9).Visible = False
        boxLegend(9).Visible = False
        lblCriteria(10) = ""
        txtScore(10).Visible = False
        txtWeight(10).Visible = False
        txtFinal(10).Visible = False
        boxLegend(10).Visible = False
        lblCriteria(11) = ""
        txtScore(11).Visible = False
        txtWeight(11).Visible = False
        txtFinal(11).Visible = False
        boxLegend(11).Visible = False
        txtWeight(0) = "2.42"
        txtWeight(1) = "2.17"
        txtWeight(2) = "2.42"
        txtWeight(3) = "2.42"
        txtWeight(4) = "5.31"
        txtWeight(5) = "2.66"
        txtWeight(6) = "0.97"
        txtWeight(7) = "1.93"
        txtWeight(8) = "1.93"
        txtWeight(9) = ""
        txtWeight(10) = ""
        txtWeight(11) = ""
        chtData.ColumnCount = 8
        gRank = "LS"
    Case 4
        lblCriteria(0) = "Efficiency and Effectiveness"
        lblCriteria(1) = "Initiative"
        lblCriteria(2) = "Acceptance of Responsibility"
        lblCriteria(3) = "Professional Knowledge"
        lblCriteria(4) = "Problem Solving"
        lblCriteria(5) = "Planning and Organising"
        lblCriteria(6) = "Divisional Responsibilities"
        lblCriteria(7) = "Compliance with Orders and Acceptance of Authority"
        txtScore(7).Visible = True
        txtWeight(7).Visible = True
        txtFinal(7).Visible = True
        boxLegend(7).Visible = True
        lblCriteria(8) = "Dress and Grooming"
        txtScore(8).Visible = True
        txtWeight(8).Visible = True
        txtFinal(8).Visible = True
        boxLegend(8).Visible = True
        lblCriteria(9) = "Supervising"
        txtScore(9).Visible = True
        txtWeight(9).Visible = True
        txtFinal(9).Visible = True
        boxLegend(9).Visible = True
        lblCriteria(10) = "Delegating"
        txtScore(10).Visible = True
        txtWeight(10).Visible = True
        txtFinal(10).Visible = True
        boxLegend(10).Visible = True
        lblCriteria(11) = "Oral and Written Expression"
        txtScore(11).Visible = True
        txtWeight(11).Visible = True
        txtFinal(11).Visible = True
        boxLegend(11).Visible = True
        txtWeight(0) = "2.05"
        txtWeight(1) = "1.71"
        txtWeight(2) = "2.74"
        txtWeight(3) = "3.42"
        txtWeight(4) = "2.05"
        txtWeight(5) = "1.36"
        txtWeight(6) = "1.37"
        txtWeight(7) = "1.37"
        txtWeight(8) = "0.68"
        txtWeight(9) = "2.73"
        txtWeight(10) = "1.38"
        txtWeight(11) = "1.37"
        chtData.ColumnCount = 12
        gRank = "PO"
    Case 5
        lblCriteria(0) = "Efficiency and Effectiveness"
        lblCriteria(1) = "Initiative"
        lblCriteria(2) = "Acceptance of Responsibility"
        lblCriteria(3) = "Professional Knowledge"
        lblCriteria(4) = "Problem Solving"
        lblCriteria(5) = "Planning and Organising"
        lblCriteria(6) = "Divisional Responsibilities"
        lblCriteria(7) = "Compliance with Orders and Acceptance of Authority"
        txtScore(7).Visible = True
        txtWeight(7).Visible = True
        txtFinal(7).Visible = True
        boxLegend(7).Visible = True
        lblCriteria(8) = "Dress and Grooming"
        txtScore(8).Visible = True
        boxLegend(8).Visible = True
        txtWeight(8).Visible = True
        txtFinal(8).Visible = True
        lblCriteria(9) = "Supervising"
        txtScore(9).Visible = True
        txtWeight(9).Visible = True
        txtFinal(9).Visible = True
        boxLegend(9).Visible = True
        lblCriteria(10) = "Delegating"
        txtScore(10).Visible = True
        txtWeight(10).Visible = True
        txtFinal(10).Visible = True
        boxLegend(10).Visible = True
        lblCriteria(11) = "Oral and Written Expression"
        txtScore(11).Visible = True
        txtWeight(11).Visible = True
        txtFinal(11).Visible = True
        boxLegend(11).Visible = True
        txtWeight(0) = "1.58"
        txtWeight(1) = "1.77"
        txtWeight(2) = "2.36"
        txtWeight(3) = "3.15"
        txtWeight(4) = "1.97"
        txtWeight(5) = "2.36"
        txtWeight(6) = "1.57"
        txtWeight(7) = "1.57"
        txtWeight(8) = "0.39"
        txtWeight(9) = "2.36"
        txtWeight(10) = "1.58"
        txtWeight(11) = "1.57"
        chtData.ColumnCount = 12
        gRank = "CPO"
    End Select
    For a = 0 To 11
        txtScore(a) = ""
        txtFinal(a) = ""
        lblTotal = ""
    Next a
    txtScore(0).SetFocus
    For a = 1 To chtData.ColumnCount
        chtData.Column = a
        chtData.Data = 0
    Next a
    subLoad
End Sub

Private Sub txtFinal_Change(Index As Integer)
    lblTotal = ""
    For a = 0 To 11
        lblTotal = Val(txtFinal(a)) + Val(lblTotal)
    Next a
End Sub

Private Sub txtScore_Change(Index As Integer)
    On Error Resume Next
    If txtScore(Index) <> "" Then
        txtFinal(Index) = Val(txtScore(Index)) * Val(txtWeight(Index))
        chtData.Column = Index + 1
        chtData.Data = Val(txtScore(Index))
    End If
    If Len(txtScore(Index)) > 1 Then
        txtScore(Index) = Left$(txtScore(Index), 1)
        ' txtScore(Index) = ""
        txtScore(Index).SelStart = 0
        txtScore(Index).SelLength = Len(txtScore(Index))
    Else
        Select Case tstRank.SelectedItem.Index
            Case 1, 2
                Criteria = 6
            Case 3
                Criteria = 8
            Case 4, 5
                Criteria = 11
        End Select
        If Index = Criteria Then
            txtScore(0).SetFocus
        Else
            txtScore(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtScore_GotFocus(Index As Integer)
    Rem show scoring help
    If chkHelp.Value = 1 Then
        For a = 1 To 9
            frmScores.lblExplain(a) = gCriteria(Index, a)
        Next a
    End If
End Sub

Private Sub txtScore_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode > 48 And KeyCode < 58 Then
        If Index < 11 Then
            If txtScore(Index + 1).Visible = True Then
                txtScore(Index + 1).SetFocus
                txtScore(Index + 1).SelStart = 0
                txtScore(Index + 1).SelLength = Len(txtScore(Index + 1))
            Else
                txtScore(0).SetFocus
                txtScore(0).SelStart = 0
                txtScore(0).SelLength = Len(txtScore(0))
            End If
        End If
    Else
        If KeyCode = 16 Then ' <> 16
            Rem [TAB] key
            txtScore(Index) = ""
            txtFinal(Index) = ""
            If chtData.ColumnCount > Index Then
                chtData.Column = Index + 1
                chtData.Data = Val(txtScore(Index))
            End If
        Else
            txtScore(Index).SetFocus
            txtScore(Index).SelStart = 0
            txtScore(Index).SelLength = Len(txtScore(Index))
        End If
    End If
    Rem Colours
    If KeyCode > 48 And KeyCode < 51 Then
        Rem score of 1 or 2
        txtScore(Index).ForeColor = &HFF&
    End If
    If KeyCode > 50 And KeyCode < 56 Then
        Rem score of 3 to 7
        txtScore(Index).ForeColor = &H0&
    End If
    If KeyCode > 55 And KeyCode < 58 Then
        Rem score of 8 or 9
        txtScore(Index).ForeColor = &HFF00&
    End If
    If KeyCode = 13 Then KeyCode = 0
    If KeyCode = 38 And Index > 0 Then txtScore(Index - 1).SetFocus ' up arrow
    If KeyCode = 40 And Index = 10 Then
        If txtScore(Index).Visible = True Then txtScore(Index).SetFocus ' down arrow
    End If
    If KeyCode = 40 And Index <> 10 Then
        If txtScore(Index + 1).Visible = True Then txtScore(Index + 1).SetFocus ' down arrow
    End If
End Sub

Private Sub txtScore_LostFocus(Index As Integer)
    ' If Len(txtScore(Index)) > 1 Then txtScore(Index) = Left$(txtScore(Index), 1)
End Sub
