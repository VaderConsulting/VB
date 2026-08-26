VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Get Logical Drive Information"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive Query Ouput"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   4155
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
