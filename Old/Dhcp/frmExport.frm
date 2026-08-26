VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton optExtract2 
      Caption         =   "All Scopes"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton optExtract1 
      Caption         =   "Single Scope"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "c:\temp\dhcp client info.csv"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    Open txtOutput For Output As #2
        If optExtract1 = True Then
            For a = 0 To frmDHCP.lstClients.ListCount - 1
                varData = frmDHCP.lstClients.List(a)
                Print #2, varData
            Next a
        Else
            For a = 0 To frmDHCP.lstScopes.ListCount - 1
                frmDHCP.lstScopes.ListIndex = a
                frmDHCP.Refresh
                For b = 0 To frmDHCP.lstClients.ListCount - 1
                    varData = frmDHCP.lstClients.List(b)
                    Print #2, varData
                Next b
            Next a
        End If
    Close 2
End Sub

