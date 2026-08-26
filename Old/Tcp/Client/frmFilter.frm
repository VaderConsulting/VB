VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkGroup 
      Caption         =   "WAN Alerts"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox chkGroup 
      Caption         =   "EUC Alerts"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox chkGroup 
      Caption         =   "Helpdesk Alerts"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CheckBox chkGroup 
      Caption         =   "NTSS Alerts"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox chkSeverity 
      Caption         =   "Severity 2 (Warning)"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox chkSeverity 
      Caption         =   "Severity 1 (Information)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblFilter 
      Caption         =   "Don't alert me for..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   8640
      Y1              =   2660
      Y2              =   2660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   8640
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkGroup_Click(Index As Integer)
  Dim lp As Integer
  ' ensure that at least one is unchecked
  If chkGroup(0).Value = vbChecked And chkGroup(1).Value = vbChecked And chkGroup(2).Value = vbChecked And chkGroup(3).Value = vbChecked Then
    chkGroup(Index).Value = vbUnchecked
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim Host As String
  SetValueString Host, "Software\Alerter", "Ignore Information", Val(chkSeverity(0)), REG_SZ
  SetValueString Host, "Software\Alerter", "Ignore Warning", Val(chkSeverity(1)), REG_SZ
  SetValueString Host, "Software\Alerter", "Ignore NTSS", Val(chkGroup(0)), REG_SZ
  SetValueString Host, "Software\Alerter", "Ignore Helpdesk", Val(chkGroup(1)), REG_SZ
  SetValueString Host, "Software\Alerter", "Ignore EUC", Val(chkGroup(2)), REG_SZ
  SetValueString Host, "Software\Alerter", "Ignore WAN", Val(chkGroup(3)), REG_SZ
  Unload Me
End Sub

Private Sub Form_Load()
  Dim isOK As Boolean
  boolIgnoreInformation = CBool(GetValue(Host, "Software\Alerter", "Ignore Information", isOK, REG_SZ))
  boolIgnoreWarning = CBool(GetValue(Host, "Software\Alerter", "Ignore Warning", isOK, REG_SZ))
  boolIgnoreNTSS = CBool(GetValue(Host, "Software\Alerter", "Ignore NTSS", isOK, REG_SZ))
  boolIgnoreHelpdesk = CBool(GetValue(Host, "Software\Alerter", "Ignore Helpdesk", isOK, REG_SZ))
  boolIgnoreEUC = CBool(GetValue(Host, "Software\Alerter", "Ignore EUC", isOK, REG_SZ))
  boolIgnoreWAN = CBool(GetValue(Host, "Software\Alerter", "Ignore WAN", isOK, REG_SZ))
  chkSeverity(0).Value = Abs(boolIgnoreInformation)
  chkSeverity(1).Value = Abs(boolIgnoreWarning)
  chkGroup(0).Value = Abs(boolIgnoreNTSS)
  chkGroup(1).Value = Abs(boolIgnoreHelpdesk)
  chkGroup(2).Value = Abs(boolIgnoreEUC)
  chkGroup(3).Value = Abs(boolIgnoreWAN)
  Me.Refresh
End Sub
