Attribute VB_Name = "gDiskSize"
Option Explicit

'-----------------------------------------------------------------------------------------------------------------------
' This project reports disk information for each volume for any number of servers in a domain.
' It uses ODBC and RDO for update of 2 tables on a SQL Server
'
' AdmDisk
' Daily Disk Information
'
' AdmDiskStat
' Monthly Disk Information (One record for each month)
'
' However you can also use a Access Database or dump it to a text file.

' You need to modify the Server() array and NumServers in Sub ScanDisks in order to make this program work for you.

' Ofc you need administrator rights and a NT workstation in order to run this program.
'
' Enjoy
' Morten
'------------------------------------------------------------------------------------------------------------------------


 Type DiskRecord
    Server As String
    Share As String
    Letter As String
    Capacity As Long
    FreeSpace As Long
    UsedSpace As Long
    VolumeLabel As String
    Filesystem As String
    Serial As String
    DateCheck As String 'yyyymm - for statistical purposes, in order to make sure that only 1 record for each month exist
    ScanDate As String
End Type

Global Const Dns = "DSN=695inf;UID=;PWD=;"
Sub Main()
    Call ScanDisks
  End
 End Sub

Sub ScanDisks()

    '-- Declarations
    Dim oDisk As New ClassDisk
    Dim I As Integer
    Dim ii As Integer
    
    Dim Disk(10000) As DiskRecord
    Dim numDisks As Integer
    ReDim Server(1000)
    Dim numServers As Integer
    
    Dim rc As Long                                              ' Return code
    Dim sConnectDrive As String                         ' Drive used for connecting to the respective share
    Dim sConnectCmd As String                          ' String which holds the server and sharename
    Dim bStatFound As Boolean                           ' Is there a month report
    Dim sDateCheck As String
                                 
    '-- Initialize
    sConnectDrive = oDisk.GetNextFreedrive
    sDateCheck = Left(Format(Date, "yyyymmdd"), 6)
    
'    If StatExists(sDateCheck) Then
'        bStatfound = True
'    Else
'        bStatfound = False
'    End If
'
'   ClearTable ("AdmDisk")
       
    '-- Get The servers from a list and put it into the Server array
    
    MsgBox "You need to modify the Server() array and NumServers in order to make this program work for you"
           
        Server(1) = "ntxx01"
        Server(2) = "ntxx02"
        Server(3) = "ntxx11"
        numServers = 3
        
    '-- Disconnect case drive is used.
    Call oDisk.DisConnect(sConnectDrive)
    '-- Get Number of Drives and names
    For I = 1 To numServers
        For ii = 99 To 107 ' Browse the shares from driveletter C$ to k$
            sConnectCmd = "\\" & Server(I) & "\" & Chr(ii) & "$"
            Call oDisk.Connect(sConnectDrive, sConnectCmd)
            If oDisk.rc = 0 Then
                numDisks = numDisks + 1
                oDisk.DriveName = sConnectDrive
                Disk(numDisks).Letter = UCase(Chr(ii)) & ":"
                Disk(numDisks).Server = Server(I)
                Disk(numDisks).Share = sConnectCmd
                Disk(numDisks).ScanDate = Date & "-" & Format(Time, "hh:mm")
                Disk(numDisks).Capacity = oDisk.CapacityMB
                Disk(numDisks).FreeSpace = oDisk.FreeSpaceMB
                Disk(numDisks).UsedSpace = oDisk.UsedSpaceMB
                Disk(numDisks).Filesystem = oDisk.Filesystem
                Disk(numDisks).Serial = oDisk.Serial
                Disk(numDisks).VolumeLabel = oDisk.VolumeLabel
                Disk(numDisks).Filesystem = oDisk.Filesystem
                Disk(numDisks).DateCheck = sDateCheck
                                
                Debug.Print "Server   : " & Disk(numDisks).Server
                Debug.Print "Disk       : " & Disk(numDisks).Letter
                Debug.Print "Capacity: " & Disk(numDisks).Capacity
                Debug.Print "Used      : " & Disk(numDisks).UsedSpace
                Debug.Print "Free       : " & Disk(numDisks).FreeSpace
                
'                If Not bStatFound Then
'                    Call UpdateTable(Disk(numDisks), "AdmDisk")
'                    Call UpdateTable(Disk(numDisks), "AdmDiskStat")
'                Else
'                    Call UpdateTable(Disk(numDisks), "AdmDisk")
'                End If
                Call oDisk.DisConnect(sConnectDrive)
            End If
        Next ii
    Next I
'-- Clean up
    Set oDisk = Nothing
End Sub
Function StatExists(sYearMonth As String) As Boolean

'-- Declarations
    Dim rdoEnv As rdoEnvironment
    Dim rdoCon As rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    Dim numRecords As Long
    Dim Found As Boolean
'-- ErrorHandling
'-- Initalize
    SQL = "Select * from ADMDISKSTAT"
    Set rdoEnv = rdoEnvironments(0)
    Set rdoCon = rdoEnv.OpenConnection("", rdDriverNoPrompt, False, Dns)
    Set rdoRs = rdoCon.OpenResultset(SQL, , rdUseServer)
    Found = False
'-- Process the resultset
    Do While Not rdoRs.EOF
        numRecords = numRecords + 1
        If sYearMonth = rdoRs!YYYYMM Then
            Found = True
        End If
        rdoRs.MoveNext
    Loop
    rdoRs.Close
    rdoCon.Close
    
    If numRecords = 0 Then
        StatExists = False
        Exit Function
    End If
        
    If Found Then
        StatExists = True
    End If

End Function

Sub ClearTable(sTable As String)
'-- Declarations
    Dim rdoEnv As rdoEnvironment
    Dim rdoCon As rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    
'-- ErrorHandling
On Error GoTo slut
'-- Initalize
    Set rdoEnv = rdoEnvironments(0)
    Set rdoCon = rdoEnv.OpenConnection("", rdDriverNoPrompt, False, Dns)
    SQL = "DELETE from " & sTable & " where capacity >" & 0
    rdoCon.Execute SQL
slut:
    rdoCon.Close
End Sub
Sub UpdateTable(DiskRec As DiskRecord, Table As String)
Dim er As rdoError
On Error GoTo CnEh

'-- Declarations
    Dim rdoEnv As rdoEnvironment
    Dim rdoCon As rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
'-- Errorhandling
'-- Initialize
    Set rdoEnv = rdoEnvironments(0)
    Set rdoCon = rdoEnv.OpenConnection("", rdDriverNoPrompt, False, Dns)
    SQL = "select * from " & Table
'--
    Set rdoRs = rdoCon.OpenResultset(SQL, , rdConcurRowVer)
    rdoRs.AddNew
    rdoRs("server") = DiskRec.Server
    rdoRs("share") = DiskRec.Share
    rdoRs("scandate") = DiskRec.ScanDate
    rdoRs("yyyymm") = DiskRec.DateCheck
    rdoRs("diskname") = DiskRec.Letter
    rdoRs("capacity") = DiskRec.Capacity
    rdoRs("free") = DiskRec.FreeSpace
    rdoRs("used") = DiskRec.UsedSpace
    rdoRs("filesystem") = DiskRec.Filesystem
    rdoRs("serial") = DiskRec.Serial
    rdoRs("volumelabel") = DiskRec.VolumeLabel
    rdoRs.Update
    rdoCon.Close
Exit Sub
CnEh:

    Debug.Print Err, Error
    For Each er In rdoErrors
        Debug.Print er.Description, er.Number
    Next er
    Resume Next
End Sub
