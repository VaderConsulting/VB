Attribute VB_Name = "modMain"
' This is task ID 1 in the Domain Control System.
' Any host with this task ID set will be queried by this application
'

Option Base 1

Sub Main()
    Dim adoConn As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim DSN As String, SQL As String
    Dim SQLServer As String, Database As String, Table As String
    Dim Hostname As String, Driveletter As String
    Dim UsedSpace As String, TotalSpace As String, FreeSpace As String
    Dim wmiLocator As SWbemLocator
    Dim oServer As SWbemServices
    Dim Domainname As String, Username As String, Password As String
    Dim oWMI As Object, HostID As Integer
    Dim curUsed As Currency, curFree As Currency, curTotal As Currency
    Dim E As Long, TaskID As Integer, NewTaskID As Integer
    
    Set adoConn = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
    
    If SQLServer = "" Then SQLServer = "(Local)"
    If Database = "" Then Database = "Domain Control"
    If Table = "" Then Table = "tblDiskSpace"
    
    'Build DSN from provided values
    DSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=Domain Control;Data Source=(Local)"
    
    adoConn.Open DSN
    
    SQL = "SELECT tblHosts.ID, tblHosts.TaskID, tblConnections.Domainname, tblHosts.Hostname, tblConnections.Username, tblConnections.Password FROM tblHosts LEFT JOIN tblConnections ON tblHosts.ConnectionID = tblConnections.ID WHERE (TaskID & 1) = 1"
    
    adoRS.Open SQL, adoConn
    
    Do Until adoRS.EOF
        HostID = adoRS("ID")
        Hostname = adoRS("Hostname")
        Domainname = adoRS("DomainName")
        Username = adoRS("Username")
        Password = adoRS("Password")
        TaskID = adoRS("TaskID")
        On Error Resume Next
        Set oServer = wmiLocator.ConnectServer(Hostname, "root/CIMV2", Domainname & "\" & Username, Password)
        E = Err.Number
        Select Case E
            Case 0 ' Success
                Set oWMI = oServer.ExecQuery("SELECT Name, Freespace, Size FROM Win32_LogicalDisk WHERE DriveType = 3")
                For Each disk In oWMI
                    ' Get diskspace
                    FreeSpace = disk.FreeSpace
                    TotalSpace = disk.Size
                    Driveletter = disk.Name
                    
                    ' Convert to currency datatype so we can do a calculation on the initial string value
                    curFree = CCur(FreeSpace)
                    curTotal = CCur(TotalSpace)
                    curUsed = curTotal - curFree
                    UsedSpace = CStr(curUsed)
                    
                    ' Create SQL Statement to insert data
                    SQL = ""
                    SQL = "INSERT INTO tblDiskSpace (HostID, Driveletter, UsedSpace, TotalSpace) "
                    SQL = SQL & "VALUES ("
                    SQL = SQL & HostID & ","
                    SQL = SQL & "'" & Left(Driveletter, 1) & "',"
                    SQL = SQL & "'" & UsedSpace & "',"
                    SQL = SQL & "'" & TotalSpace & "'"
                    SQL = SQL & ")"
                    
                    ' Save to SQL Database!!
                    adoConn.Execute SQL
                    
                    NewTaskID = TaskID And Not (1)
                    
                    ' Remove this task from this Host
                    SQL = "UPDATE tblHosts SET TaskID = " & NewTaskID & " WHERE ID = " & HostID
                    adoConn.Execute SQL
                Next
            Case -2147023174 ' RPC Server unavailable
                
        End Select
        adoRS.MoveNext
    Loop
    ' Done saving to SQL Database
    adoConn.Close
    Set adoConn = Nothing
    End
End Sub
