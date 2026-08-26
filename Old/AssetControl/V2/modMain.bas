Attribute VB_Name = "modMain"
Public oDatabase As ADODB.Recordset
Public oConn As ADODB.Connection
Public DSN As String

Public Function GetRecords(strQuery As String) As Integer
    Dim iRecordCount As Integer
    Dim strLabel As String
    
    If oDatabase.State = adStateOpen Then oDatabase.Close
    oDatabase.Open strQuery, oConn, adOpenDynamic, adLockOptimistic
    
    frmMain.lblInfo = "Opened Database."
    
    iRecordCount = oDatabase.RecordCount
    
    Select Case iRecordCount
        Case 1
            strLabel = " record"
        Case Else
            strLabel = " records"
    End Select
    
    frmMain.lblInfo = iRecordCount & strLabel & " found."
    
    GetRecords = iRecordCount
End Function
