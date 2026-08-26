Attribute VB_Name = "Module1"
Option Base 1
Const DRIVE_LOCAL As Integer = 2
Const DRIVE_NETWORK As Integer = 3
Const DRIVE_CDROM As Integer = 4

Sub Main()
  Dim SQL As String
  Dim Hostname As String, Driveletter As String
  Dim oSys As New FileSystemObject
  Dim FreePercent As String, FreeSpace As String, UsedSpace As String, UsedPercent As String, TotalSpace As String
  
  ' General variables
  Dim intTemp As Integer
  Dim strTemp As String
  Dim lngTemp As Long
  Dim boolTemp As Boolean
  
  Hostname = UCase(Environ$("computername"))
  ' Create registry key
  boolTemp = CreateKey(Hostname, "Software", "CSC")
  ' Create registry key
  boolTemp = CreateKey(Hostname, "Software\CSC", "SOTD")
  ' Get current value of Collecting key
  strTemp = GetValue(Hostname, "Software\CSC\SOTD", "Collecting", False)
  If Trim(strTemp) <> "" Then
    OldCollection = CInt(strTemp)
  Else
    OldCollection = 0
  End If
  ' Add one to it, and save it back so we know this app is running
  SetValueString Hostname, "Software\CSC\SOTD", "Collecting", CStr(OldCollection + 1), REG_SZ
  
  ' Create directory to store information
  If Dir("C:\SOTD", vbDirectory) = "" Then
    MkDir "C:\SOTD"
  End If
  
  For Each Drive In oSys.Drives
    If Drive.DriveType = DRIVE_LOCAL Then
      Driveletter = Drive.Driveletter
      FreeSpace = Drive.FreeSpace
      TotalSpace = Drive.TotalSize
      UsedSpace = TotalSpace - FreeSpace
      UsedPercent = Int((UsedSpace / TotalSpace) * 100)
      FreePercent = Int((FreeSpace / TotalSpace) * 100)
      ' Create SQL Statement to insert data
      SQL = "INSERT INTO tblDiskSpace (HostName, Driveletter, UsedPercent, UsedSpace, FreePercent, FreeSpace, TotalSpace) "
      SQL = SQL + "VALUES ("
      SQL = SQL + "'" + Hostname + "',"
      SQL = SQL + "'" + Left(Driveletter, 1) + "',"
      SQL = SQL + "'" + UsedPercent + "',"
      SQL = SQL + "'" + UsedSpace + "',"
      SQL = SQL + "'" + FreePercent + "',"
      SQL = SQL + "'" + FreeSpace + "',"
      SQL = SQL + "'" + TotalSpace + "'"
      SQL = SQL + ")"
      
      ' Output SQL String to report
      Open "C:\SOTD\DiskSpace.txt" For Append As #1
        Print #1, SQL
      Close 1
      SQL = ""
      
    End If
  Next
  
  ' Get current value of Collecting key
  OldCollection = CInt(GetValue(Hostname, "Software\CSC\SOTD", "Collecting", False))
  If OldCollection = 0 Then OldCollection = 1
  ' Reset flag so we know this app is finished
  SetValueString Hostname, "Software\CSC\SOTD", "Collecting", CStr(OldCollection - 1), REG_SZ
  End
End Sub

Private Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim InfoTextStr As String
    
    'Init
    ECode = False
    Hostname = Trim(Hostname)
    Key = Trim(Key)
    Value = Trim(Value)
    GetValue = ""
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then Exit Function
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    InfoTextStr = RegistryQueryValue(OpenKeyVal, Value, REG_SZ)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    GetValue = InfoTextStr
    ECode = True
End Function

Private Function CreateKey(ByVal Hostname As String, ByVal Key As String, NewKey As String) As Boolean
    Dim OpenKeyVal As Long
    Dim OpenHiveVal As Long
    Dim RResult As Long
    Dim lCreateResult As Long
    
    'Init
    Hostname = Trim(Hostname)
    Key = Trim(Key)
        
    RResult = RegConnectRegistry("", HKEY_LOCAL_MACHINE, OpenHiveVal)
    If (RResult <> ERROR_SUCCESS) Then
      CreateKey = False
      Exit Function
    End If
    
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, Key)
    lCreateResult = RegistryCreateKey(OpenKeyVal, NewKey)
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    CreateKey = True
    
End Function

Private Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
    Dim OpenKeyVal As Long
    Dim RResult As Long
    Dim OpenHiveVal As Long
    Dim strValue, CMPTRName, keytogo As String, InfoTextStr As String
    Dim x As Integer
    
    'GetIPCConnection (Trim(Hostname))
    
    CMPTRName = Trim(Hostname)
    
    RResult = RegConnectRegistry(CMPTRName, HKEY_LOCAL_MACHINE, OpenHiveVal)
    keytogo = Trim(Key)
    OpenKeyVal = RegistryOpenKey(OpenHiveVal, keytogo)
    RegistryWriteValue Data, OpenKeyVal, Value, regType
    
    RegCloseKey (OpenKeyVal)
    RegCloseKey (OpenHiveVal)
    
    'DisIPCConnection (Trim(Hostname))
        
End Function

