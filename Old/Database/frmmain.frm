VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wee Department details"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Height          =   1095
      Left            =   3360
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtDarren 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Text            =   "789"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtDarren 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Text            =   "456"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtDarren 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Text            =   "123"
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Operations"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Aviation"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Support"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MEE"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "WEE"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "3"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "2"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "1"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "WEE Department"
      Connect         =   "Access"
      DatabaseName    =   "G:\WE\Regulator\WEReg.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DEPTLIST"
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Surname"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Rank"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Billet"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    For a = txtDarren.LBound To txtDarren.UBound
        txtDarren(a).Text = ""
    Next a
End Sub

Private Sub Option1_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Data1.DatabaseName = "G:\WE\Regulator\WEReg.mdb"
End Sub

Private Sub Option2_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""

End Sub

Private Sub Option3_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""

End Sub

Private Sub Option4_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""

End Sub

Private Sub Option5_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""

End Sub
