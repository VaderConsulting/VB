VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL2 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   8535
   End
   Begin VB.TextBox txtURL1 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "http%3a%2f%2fsupport%2edell%2ecom%2fFileLib%2fFormat%2easpx%3fReleaseID%3dR58470%26sws%3dfw"
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtURL1_Change()
    txtURL2 = Replace(txtURL1.Text, "%3a", ":", 1, -1, vbTextCompare)
    txtURL2 = Replace(txtURL2.Text, "%2f", "/", 1, -1, vbTextCompare)
    txtURL2 = Replace(txtURL2.Text, "%2e", ".", 1, -1, vbTextCompare)
    txtURL2 = Replace(txtURL2.Text, "%3f", "?", 1, -1, vbTextCompare)
    txtURL2 = Replace(txtURL2.Text, "%3d", "=", 1, -1, vbTextCompare)
    txtURL2 = Replace(txtURL2.Text, "%26", "&", 1, -1, vbTextCompare)
End Sub
