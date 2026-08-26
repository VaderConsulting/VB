Attribute VB_Name = "modMain"
Public DSN As String
Public isLoading As Boolean

Public adoRS As ADODB.Recordset
Public adoConn As ADODB.Connection

Public Sub Main()
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ConnManager.mdb;Persist Security Info=False"
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    Load frmMain
End Sub
