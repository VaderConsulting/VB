VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOther 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other info"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSignals 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmOther.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Signals"
      Top             =   3360
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      _Version        =   327681
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "CD Server 1"
      TabPicture(0)   =   "frmOther.frx":0426
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picCD1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CD Server 2"
      TabPicture(1)   =   "frmOther.frx":0442
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picCD2"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picCD2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -74880
         ScaleHeight     =   1335
         ScaleWidth      =   5535
         TabIndex        =   9
         ToolTipText     =   "Contents of CD Server 2"
         Top             =   480
         Width           =   5535
      End
      Begin VB.PictureBox picCD1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   5535
         TabIndex        =   8
         ToolTipText     =   "Contents of CD Server 1"
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdScream 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   3600
      Picture         =   "frmOther.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Call for assistance"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   1095
      Left            =   4800
      Picture         =   "frmOther.frx":08A0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Launcher Information"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPers 
      Height          =   1095
      Left            =   2400
      Picture         =   "frmOther.frx":0CE2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "PERS calculator"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdMessages 
      Height          =   1095
      Left            =   1200
      Picture         =   "frmOther.frx":0FEC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Logon Messages"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdWords 
      Enabled         =   0   'False
      Height          =   1095
      Left            =   0
      Picture         =   "frmOther.frx":142E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Daily Orders"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Show abbreviated help"
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "frmOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
    frmOther.Hide
    frmAbout.Show
End Sub

Private Sub cmdDone_Click()
    frmStart.Show
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    msg = "Use the right mouse button anywhere on the Launcher tab" & vbCrLf
    msg = msg & "to modify your personal program shortcuts." & vbCrLf & vbCrLf
    msg = msg & "Greyed out command buttons indicate the option is not available."
    MsgBox msg
End Sub

Private Sub cmdMessages_Click()
    On Error Resume Next
    Shell "D:\Program files\Messages\Messages.exe", vbNormalFocus
    On Error GoTo 0
End Sub

Private Sub cmdPers_Click()
    frmOther.Hide
    frmPERS.Show
End Sub

Private Sub cmdScream_Click()
    frmOther.Hide
    frmScream.Show
End Sub

Private Sub cmdSignals_Click()
    frmOther.Hide
    frmSignals.Show
End Sub

Private Sub cmdWords_Click()
    frmOther.Hide
    frmOrders.Show
End Sub

Private Sub Form_Activate()
    
    Dim wrkJet As Workspace
    Dim dbsNorthwind As Database
    Number = 6
    picCD1.Cls
    On Error Resume Next
    Set dbsCD = OpenDatabase("J:\Bulletin\LAN\CD's.mdb")
    
    Set rstCD1 = dbsCD.OpenRecordset("SELECT * FROM Server1;")
    For Data = 1 To Number
        picCD1.Print rstCD1!letter & ": " & rstCD1!Contents
        rstCD1.MoveNext
    Next
    
Server2:
    picCD2.Cls
    On Error Resume Next
    Set rstCD2 = dbsCD.OpenRecordset("SELECT * FROM Server2;")
    For Data = 1 To Number
        picCD2.Print rstCD2!letter & ": " & rstCD2!Contents
        rstCD2.MoveNext
    Next
    
Done:
    dbsCD.Close
    On Error GoTo 0
End Sub
