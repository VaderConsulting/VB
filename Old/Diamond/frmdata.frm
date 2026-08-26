VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form frmData 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPreview 
      Caption         =   "Preview search results"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fmeFound 
      Caption         =   "Search results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   8895
      Begin VB.ListBox lstRecord 
         Height          =   1620
         ItemData        =   "frmData.frx":0000
         Left            =   7440
         List            =   "frmData.frx":0002
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox lstMatch 
         BackColor       =   &H00C0FFC0&
         Height          =   3765
         ItemData        =   "frmData.frx":0004
         Left            =   240
         List            =   "frmData.frx":0006
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin ComctlLib.ProgressBar pbrSearch 
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   4320
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Image imgDate 
         Height          =   255
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblAmount 
         Alignment       =   2  'Center
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label lblSearch 
         Caption         =   "Search Progress"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   28
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   27
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   26
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   25
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label lblField 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3120
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   23
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   22
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblNames 
         Caption         =   "Serial Number"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Product ID"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Date out"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   19
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Transaction Number"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Date in"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Supplier ID"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   16
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "System ID"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblNames 
         Caption         =   "Cost"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblNames 
         Caption         =   "Comments"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   13
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblField 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   12
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label lblTransaction 
         Alignment       =   2  'Center
         Caption         =   "Transaction No"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSerial 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Reload"
      Height          =   285
      Left            =   8280
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDb 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\data\access\diamond.mdb"
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label lblDb 
      Alignment       =   2  'Center
      Caption         =   "Database Location"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label lblDays 
      Height          =   255
      Left            =   8280
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting for Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   9720
      Width           =   8895
   End
   Begin VB.Label lblSerial 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblWarranty 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label lblStock 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9960
      Y1              =   10560
      Y2              =   10560
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLoad_Click()
    txtDb.Enabled = False
End Sub

Public Sub cmdSearch_Click()
    lstMatch.Clear
    lblFound = "0 found"
    fmeFound.Caption = "Search results"
    cmdSearch.Enabled = False
    frmData.lblStatus = "SEARCHING"
    gStepNo = 2
    StepChange
    rsStock.MoveFirst
    Screen.MousePointer = vbHourglass
    searchfield = 2
    For lp = 1 To rsStock.RecordCount
        pbrSearch.Value = lp
        a = rsStock.Fields(searchfield)
        'lblField = "Current record (" & lp & ") : " & a
        If UCase(a) = UCase(txtSerial) Then
            lstMatch.AddItem rsStock.Fields(0) ' Transaction Number
            lstRecord.AddItem lp
            fmeFound = "Search results: " & lstMatch.ListCount & " found"
            ShowResults
            lstMatch.Refresh
            fmeFound.Refresh
        End If
        If chkPreview.Value = vbChecked Then
            ShowResults
        End If
        rsStock.MoveNext
        DoEvents
    Next lp
    If lstMatch.ListCount = 0 Then ' NOT FOUND
        fmeFound.Caption = "Search results: 0 found"
        gStepNo = 7
        StepChange
        frmData.lblStatus = "NO MATCH FOUND"
        reply = MsgBox("Is there a valid receipt ?", vbYesNo + vbQuestion, "Customer Question")
        If reply = vbYes Then
            gStepNo = 8
            StepChange
            MsgBox "Perform a manual check on the item", vbInformation, "Instruction"
        End If
        If reply = vbNo Then
            gStepNo = 10
            StepChange
            MsgBox "Inform the client that you require a valid receipt, or inform the manager", vbInformation, "Instruction"
            gStepNo = 15
            StepChange
        End If
    End If
    If lstMatch.ListCount = 1 Then ' FOUND ONE
        gStepNo = 3
        StepChange
        frmData.lblStatus = "MATCH FOUND"
        rsStock.MoveFirst
        a = Int(lstRecord.List(0))
        rsStock.Move a
        If rsStock.Fields(5) <> "" Then lblDateOut = CDate(rsStock.Fields(5) & "")
        WarrantyCheck
    End If
    If lstMatch.ListCount > 1 Then ' MULTIPLE FOUND
        gStepNo = 3
        StepChange
        frmFlow.cmdOperation.Visible = False
        frmFlow.cmdCont.Visible = True
        frmFlow.cmdCont.Enabled = False
        frmData.lblStatus = "MULTIPLE MATCHES FOUND - SELECT APPROPRIATE TRANSACTION NUMBER"
    End If
    'If gStepNo = 5 Then ' Find out if it is in stock
    '    reply = MsgBox("Is it in stock?", vbQuestion + vbYesNo, "Warranty replacement")
    '    If reply = vbYes Then
    '        frmData.lblStatus = "REPLACE ITEM FROM STOCK"
    '        gStepNo = 6
    '        StepChange
    '    End If
    '    If reply = vbNo Then
    '        frmData.lblStatus = "OBTAIN A RA AND ORDER A REPLACEMENT ITEM"
    '        gStepNo = 9
    '        StepChange
    '    End If
    'End If
    Screen.MousePointer = vbNormal
End Sub

Public Sub Form_Load()
    Left = frmFlow.Left + frmFlow.Width
    Top = 0
    rsStock.MoveFirst
    rsStock.MoveLast
    gTotal = rsStock.RecordCount
    pbrSearch.Max = gTotal
    lblStock = rsStock.RecordCount & " stock items."
    rsWarranty.MoveFirst
    rsWarranty.MoveLast
    lblWarranty = rsWarranty.RecordCount & " warranty claims."
End Sub

Private Sub lblField_Change(Index As Integer)
    If Index = 5 Then ' date out
        If lstMatch.ListCount > 1 Then  'more than 1 item
            If frmData.lblField(5) <> "" Then
                imgDate.Visible = True
                days = DateDiff("d", frmData.lblField(5), Now)
                If days < 365 Then
                    imgDate.Picture = frmPic.imgStart.Picture
                Else
                    imgDate.Picture = frmPic.imgStop.Picture
                End If
            End If
        Else
            imgDate.Visible = False
        End If
    End If
End Sub

Public Sub lstMatch_Click()
    If lstMatch.ListIndex = -1 Then
        rsStock.MoveFirst
        rsStock.Move CInt(lstRecord.List(0)) - 1
        For a = 0 To 8
            lblField(a) = rsStock.Fields(a) & ""
        Next a
    End If
    If lstMatch.ListCount = 1 Then
        rsStock.MoveFirst
        'rsStock.Move CInt(lstRecord.List(lstMatch.ListIndex)) - 1
        For a = 0 To 8
            lblField(a) = rsStock.Fields(a) & ""
        Next a
    End If
    If lstMatch.ListCount > 1 Then
        frmFlow.cmdCont.Enabled = True
        rsStock.MoveFirst
        rsStock.Move CInt(lstRecord.List(lstMatch.ListIndex)) - 1
        For a = 0 To 8
            lblField(a) = rsStock.Fields(a) & ""
        Next a
        'frmData.lblStatus = "SELECT CONTINUE WHEN READY"
    End If
End Sub

Private Sub txtSerial_Change()
    If txtSerial = "" Then cmdSearch.Enabled = False
    If txtSerial <> "" Then cmdSearch.Enabled = True
End Sub

Sub ShowResults()
    For a = 0 To 8
        lblField(a) = rsStock.Fields(a) & ""
    Next a
End Sub

Private Sub txtSerial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSearch_Click
    End If
End Sub
