VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim AE As AceArray
    Dim I As Integer
        
    If AddAccessControlElement("c:\temp\test.txt", SE_FILE_OBJECT, "UserName", COMMON_CHANGE, SET_ACCESS) Then
        MsgBox "Rights added to DACL"
    Else
        MsgBox "Unable to add rights"
    End If
    
    AE = GetAccessControlElements("c:\temp\test.txt", SE_FILE_OBJECT)
    
    For I = LBound(AE.List) To UBound(AE.List)
        Form1.Print AE.List(I).TRUSTEE.ptstrName & " - " & AccessRights(AE.List(I).grfAccessMode, AE.List(I).grfAccessPermissions)
    Next I
End Sub


