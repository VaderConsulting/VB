Attribute VB_Name = "modMain"
Option Explicit


Public adoRS As ADODB.Recordset
Public adoConn As ADODB.Connection
Public SQL As String
Public DSN As String
Public Db As String
Public SelectedSong As String

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long



Sub Main()
    Dim ShowKey As Boolean
    
    Db = GetSetting("D. Robinson\" & App.EXEName, "Setup", "Db Path", App.Path & "\music.mdb")
    If GetSetting("D. Robinson\" & App.EXEName, "Setup", "Show Key", "True") = "True" Then
        frmDatabase.fmeKey.Visible = True
    Else
        frmDatabase.fmeKey.Visible = False
    End If
    
    Set adoRS = CreateObject("ADODB.Recordset")
    Set adoConn = CreateObject("ADODB.Connection")
    
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Db & ";Persist Security Info=False"
    frmDatabase.Show
End Sub
