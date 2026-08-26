VERSION 5.00
Begin VB.Form frmServerLU 
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSqlList 
      Height          =   315
      ItemData        =   "frmServerLU.frx":0000
      Left            =   600
      List            =   "frmServerLU.frx":0002
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton cmdSrvList 
      Caption         =   "Get ServerList"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6720
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Servers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmServerLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    End
End Sub


Private Sub cmdSrvList_Click()
' this will call the "EnumSQLServers" function
' and add each server to the combo-box
'
Dim si As Variant       ' Variant for list enumeration
Dim vnt As Variant      ' Variant(array) with SQLServer list
   
' Get the list of SQL Servers
vnt = EnumSQLServers

' Populate the combobox
For Each si In vnt
    cboSqlList.AddItem si
Next si

' Set the ComboBox to the 1st in the list
cboSqlList.ListIndex = 0
    
End Sub

Private Sub Form_Load()
    cboSqlList.Clear
End Sub
