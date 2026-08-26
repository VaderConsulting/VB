Attribute VB_Name = "modGlobal"
Option Explicit

Public Processes As New Collection

Public Function FindCompletedProcesses() As Boolean
  Dim oProc As clsProcess
  Dim i As Integer
  Dim DoAgain As Boolean
  
  FindCompletedProcesses = False

  Do
    DoAgain = False
    For i = 1 To Processes.Count
      Set oProc = Processes.Item(i)
      If ProcessCompleted(oProc) Then
        CleanupProcess oProc
        Processes.Remove i
        Set oProc = Nothing
        DoAgain = True
        Exit For
      End If
    Next
  Loop Until Not DoAgain
  
  FindCompletedProcesses = True
End Function

Public Function ExecuteTask(ByVal cmdline As String) As Boolean
  Dim oProc As clsProcess
  
  ExecuteTask = False
  Set oProc = New clsProcess
  If Not ExecCmd(cmdline, oProc) Then
    Set oProc = Nothing
    Exit Function
  End If
  
  Processes.Add oProc
  ExecuteTask = True
End Function

Public Function GetValue(ByVal Hostname As String, ByVal Key As String, ByVal Value As String, ByRef ECode As Boolean) As String
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

Public Function CreateKey(ByVal Hostname As String, ByVal Key As String, NewKey As String) As Boolean
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

Public Function SetValueString(Hostname As String, Key As String, Value As String, Data As String, regType As Long) As Long
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

