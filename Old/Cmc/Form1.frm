VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CheckBox chkStep 
      Caption         =   "Step"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   5280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   12
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtCMC 
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtCMC 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtCMC 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtCMC 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   6720
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   6240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtParent 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   6855
   End
   Begin VB.TextBox txtCMC 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3081
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5160
      Top             =   7800
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\WIN98\Desktop\CMC.MDB"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\WIN98\Desktop\CMC.MDB"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from lclfacility_item"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvwCMC 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8070
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Step As Boolean

Private Sub Command1_Click()
    Command1.Enabled = False
    Dim n As Node
    ProgressBar1.Max = Adodc1.Recordset.RecordCount
    ProgressBar1.Min = 0
    Adodc1.Recordset.MoveFirst
    'On Error Resume Next
    For a = 1 To Adodc1.Recordset.RecordCount
        s = Adodc1.Recordset.Fields(0).Value
        For b = 0 To 3
            c = InStr(1, s, "-", vbTextCompare)
            If c > 0 Then txtCMC(b) = Left(s, c - 1)
            'txtCMC(b).Refresh
            s = Right(s, Len(s) - c)
        Next b
        txtCMC(4) = s
        'txtCMC(4).Refresh
        
        txtDescription = Adodc1.Recordset.Fields(1).Value
        'txtDescription.Refresh
        
        t = Adodc1.Recordset.Fields(2).Value
        For b = 0 To 3
            c = InStr(1, t, "-", vbTextCompare)
            If c > 0 Then txtParent(b) = Left(t, c - 1)
            'txtParent(b).Refresh
            t = Right(t, Len(t) - c)
        Next b
        txtParent(4) = t
        'txtParent(4).Refresh
        
        ProgressBar1.Value = a
        'Form1.Refresh
        Adodc1.Recordset.MoveNext
        'Adodc1.Recordset.MoveNext
        
        'Start population of treeview control
         p = Adodc1.Recordset.Fields(0).Value
         c = Adodc1.Recordset.Fields(2).Value
         r = Adodc1.Recordset.Fields(1).Value
         
         c0 = txtCMC(0)
         c1 = txtCMC(1)
         c2 = txtCMC(2)
         c3 = txtCMC(3)
         c4 = txtCMC(4)
         
         p0 = txtParent(0)
         p1 = txtParent(1)
         p2 = txtParent(2)
         p3 = txtParent(3)
         p4 = txtParent(4)
         
        If txtParent(0) <> "" Then
            On Error Resume Next
            'Set n = tvwCMC.Nodes().Add(, , p, p & "-" & r)
            Set n = tvwCMC.Nodes().Add(, , "N0" & p0, p0)
            Set n = tvwCMC.Nodes().Add("N0" & p0, tvwChild, "N1" & p1, p1)
            Set n = tvwCMC.Nodes().Add("N1" & p1, tvwChild, "N2" & p2, p2)
            Set n = tvwCMC.Nodes().Add("N2" & p2, tvwChild, "N3" & p3, p3)
            Set n = tvwCMC.Nodes().Add("N3" & p3, tvwChild, "N4" & p4, p4)
            On Error Resume Next
            'Set n = tvwCMC.Nodes().Add(p4, tvwChild, r, r)
            'Set n = tvwCMC.Nodes().Add(p, tvwChild, c, c & "-" & r)
            'Set n = tvwCMC.Nodes().Add(c, tvwChild, r, p & "-" & r)
            
            If c0 <> p0 And p0 <> "*" Then
                Set n = tvwCMC.Nodes().Add(, , "N0" & c0, c0)
            End If
            If c1 <> p1 And p1 <> "*" Then
                Set n = tvwCMC.Nodes().Add("N0" & c0, tvwChild, "N1" & c1, c1)
            End If
            If c2 <> p2 And p2 <> "*" Then
                Set n = tvwCMC.Nodes().Add("N1" & c1, tvwChild, "N2" & c2, c2)
            End If
            If c3 <> p3 And p3 <> "*" Then
                Set n = tvwCMC.Nodes().Add("N2" & c2, tvwChild, "N3" & c3, c3)
            End If
            If c4 <> p4 And p4 <> "*" Then
                Set n = tvwCMC.Nodes().Add("N3" & c3, tvwChild, "N4" & c4, c4 & "-" & r)
            End If
            
            If c0 = p0 Or p0 = "*" Then
                If c1 = p1 Or p1 = "*" Then
                    If c2 = p2 Or p2 = "*" Then
                        If c3 = p3 Or p3 = "*" Then
                            If c4 = p4 Or p4 = "*" Then
                                Set n = tvwCMC.Nodes().Add("N3" & c3, tvwChild, "N4" & r, r)
                            End If
                        End If
                    End If
                End If
            End If
            
            'Set n = tvwCMC.Nodes().Add(c4, tvwChild, r, r)
            On Error GoTo 0
        Else
            On Error Resume Next
            Set n = tvwCMC.Nodes().Add(, , "N" & txtParent(4), txtParent(4))
            On Error GoTo 0
        End If
        tvwCMC.Refresh
        If chkStep.Value = vbChecked Then
            While Step = True
                DoEvents ' Just wait here until the next button is clicked
            Wend
        End If
        Step = True
        DoEvents
    Next a
    ProgressBar1.Value = 0
    tvwCMC.Sorted = True
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Step = False
End Sub
