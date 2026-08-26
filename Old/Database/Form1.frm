VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim db As Database
    Set db = Workspaces(0).OpenDatabase(App.Path & "Biblio.mdb")
    'commented out syntax followed by working example
    'db.execute "SELECT tbl.fields INTO [dbms type;DATABASE=path].[unqualified 'filename] FROM [table or tables]
    db.execute "SELECT * INTO [dBase III;DATABASE=""C:Documents].[testa] FROM [authors]"
End Sub
