VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim adoConn As ADODB.Connection
  Dim adoRS As ADODB.Recordset
  Dim DSN As String
  Dim SQL As String
  
  Set adoConn = CreateObject("ADODB.Connection")
  Set adoRS = CreateObject("ADODB.Recordset")
  
  DSN = "Driver=" & "{Microsoft Text Driver (*.txt; *.csv)}" & ";" & _
            "DBQ=" & "C:\temp" & ";" & _
            "DefaultDir=" & "C:\temp" & ";" & _
            "Uid=Admin;Pwd=;"
  SQL = "SELECT * FROM siteinfo#csv WHERE Subnet = '10.1.30'"
  
  adoConn.Open DSN
  adoRS.Open SQL, adoConn
  Do Until adoRS.EOF Or adoRS.BOF
    Debug.Print adoRS("Subnet") & " " & adoRS("Siteloc") & " " & adoRS("Siteserver")
    adoRS.MoveNext
  Loop
  adoRS.Close
  adoConn.Close
  Set adoRS = Nothing
  Set adoConn = Nothing
  Exit Sub
  
  'If Dir("c:\temp\filename.csv") <> "" Then
  '  dtremote = FileDateTime("\\cbdxaaa\netlogon\sites\filename.csv")
  '  dtlocal = FileDateTime("c:\temp\filename.csv")
  '  If dtlocal < dtremote Then
  '      FileCopy "\\cbdxaaa\netlogon\sites\filename.csv", "c:\temp\filename.csv"
  '  End If
  'Else
  '  FileCopy "\\cbdxaaa\netlogon\sites\filename.csv", "c:\temp\filename.csv"
  'End If
End Sub
