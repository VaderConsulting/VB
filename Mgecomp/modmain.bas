Attribute VB_Name = "modMain"
Option Explicit

Public strUserGroups() As String
Public strComputerGroups() As String
Public strUserOU As String
Public strComputerOU As String
Public strKey(4) As String
Public strDomainName As String
Public strApp_Path
Public intNumberOfLocations As Integer
Public strLocationNames() As String
Public strProfilePaths() As String
Public strGroups() As String

Public Enum IADSObject
    User = 0
    Computer = 1
    Group = 2
    OU = 3
End Enum

Public Enum EventType
    EventStop = 0
    EventQuestion = 1
    EventExclamation = 2
    EventInfo = 3
End Enum

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'---------------------------------------------------------------------------------------
' Procedure : AddtoAuditTrail
' DateTime  : 05-05-2003 20:36
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date            Author                 History
' 1.0   05-05-2003      Dave Robinson          Initial Version
'---------------------------------------------------------------------------------------
Public Sub AddtoAuditTrail(strMessage As String)
    Open strApp_Path & "audit.txt" For Append As #2
        Print #2, Format(Now, "YYYY-MM-DD HH:NN:SS ") & strMessage
    Close 2
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetConfig
' DateTime  : 01-05-2003 18:54
' Author    : Dave Robinson
' Purpose   : Retrieve configuration from ini file
'
'  V    Date        Author          History
' 1.0   01-05-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetConfig(strInfoType As String)
    Dim strGroup As String
    Dim intLoop As Integer
    
    intLoop = 1
    Do
        strGroup = GetIniValue(strInfoType, "Group" & CStr(intLoop), "", strApp_Path & "mgecomp.ini")
        If strGroup <> "" Then
            Select Case strInfoType
                Case "User Groups"
                    ReDim Preserve strUserGroups(intLoop - 1)
                    strUserGroups(intLoop - 1) = strGroup
                Case "Computer Groups"
                    ReDim Preserve strComputerGroups(intLoop - 1)
                    strComputerGroups(intLoop - 1) = strGroup
            End Select
            intLoop = intLoop + 1
        End If
    Loop Until strGroup = ""
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetIniValue
' DateTime  : 03-04-2003 08:21
' Author    : Dave Robinson
' Purpose   : Retrieves a value from an ini file corresponding to the section and key name passed.
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetIniValue(lpSectionName As String, lpKeyName As String, defaultValue As String, inifile As String) As String
    Dim success As Long
    Dim nSize As Long
    Dim Ret As String
    
    On Error GoTo GetIniValue_Error
    
    'call the API with the parameters passed.
    'The return value is the length of the string in ret, including the terminating null. If a default value was passed, and the section or
    'key name are not in the file, that value is returned. If no default value was passed (""), then success will = 0 if not found.
    
    'Pad a string large enough to hold the data.
    Ret = Space$(2048)
    nSize = Len(Ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, defaultValue, Ret, nSize, inifile)
    
    If success Then
        GetIniValue = Left$(Ret, success)
    End If
    
    On Error GoTo 0
    Exit Function
    
GetIniValue_Error:
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetIniValue of Module modMain", vbCritical

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetObjectPath
' DateTime  : 14-03-2003 10:30
' Author    : Dave Robinson
' Purpose   : Based on code found in the ADSI 2.5 Help files:  see "Automation - Retrieving a Property From an ADSI object"
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function GetObjectPath(strObjectName As String, ObjectType As IADSObject) As String
    Dim objConnection As ADODB.Connection
    Dim objRecordSet As ADODB.Recordset
    Dim objCommand As ADODB.Command
    Dim objGC As IADs
    Dim intMAX_DISPLAY As Integer
    Dim objRoot As IADs
    Dim strDomain As String
    Dim objDomain As IADs
    Dim strADSPath As String
    Dim strFilter As String
    Dim strAttribsToReturn As String
    Dim strDepth As String
    Dim strText As String
    Dim intNumDisplay As Integer
    Dim intCount As Integer
    Dim intTemp1 As Integer
    Dim intTemp2 As Integer
        
    On Error Resume Next
    'Maximum number of items to list.
    intMAX_DISPLAY = 1
    
    'Create ADO connection object for Active Directory
    Set objConnection = CreateObject("ADODB.Connection")
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on CreateObject in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on CreateObject(Connection)"
         Err.Clear
         Exit Function
      End If
    objConnection.Provider = "ADsDSOObject"
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on Provider in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on Provider"
         Err.Clear
         Exit Function
      End If
    objConnection.Open "Active Directory Provider"
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on Open in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on Open"
         Err.Clear
         Exit Function
      End If
     
    'Create ADO command object for the connection.
    Set objCommand = CreateObject("ADODB.Command")
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on CreateObject in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on CreateObject(Command)"
         Err.Clear
         Exit Function
      End If
    objCommand.ActiveConnection = objConnection
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on Active Connection in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on Active Connection"
         Err.Clear
         Exit Function
      End If
     
    'Get the ADsPath for the domain to search.
    Set objRoot = GetObject("LDAP://rootDSE")
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on GetObject for rootDSE in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on GetObject for rootDSE"
         Err.Clear
         Exit Function
      End If
    strDomain = objRoot.Get("defaultNamingContext")
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on Get defaultNamingContext in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on Get DefaultNamingContext"
         Err.Clear
         Exit Function
      End If
    Set objDomain = GetObject("LDAP://" & strDomain)
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on GetObject for domain in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on GetObject for Domain"
         Err.Clear
         Exit Function
      End If

    'Build the ADsPath element of the commandtext
    strADSPath = "<" & objDomain.ADsPath & ">"
    
    'Build the filter element of the commandtext
    Select Case ObjectType
        Case IADSObject.User
            strFilter = "(&(objectCategory=person)(objectClass=user)(samAccountName=" & strObjectName & "))"
        Case IADSObject.OU
            strFilter = "(&(objectClass=organizationalUnit)(name=" & strObjectName & "))"
        Case IADSObject.Group
            strFilter = "(&(objectCategory=group)(groupType:1.2.840.113556.1.4.804: = ADS_GROUP_TYPE_SECURITY_ENABLED))"
        Case IADSObject.Computer
            strFilter = "(&(objectCategory=computer)(name=" & strObjectName & "))"
     End Select
    'Build the returned attributes element of the commandtext.
    strAttribsToReturn = "adsPath"
     
    'Build the depth element of the commandtext.
    strDepth = "subTree"

    'Assemble the commandtext.
    objCommand.CommandText = strADSPath & ";" & strFilter & ";" & strAttribsToReturn & ";" & strDepth
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on CommandText in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on CommandText"
         Err.Clear
         Exit Function
      End If
    
    'Execute the query.
    Set objRecordSet = objCommand.Execute
      If (Err.Number <> 0) Then
         Logevent "**** " & App.ProductName & " produced the following error:" & Err.Number & " on Execute in GetObjectPath Function inside modMain", EventExclamation
         MsgBox "Error on Execute"
         Err.Clear
         Exit Function
      End If
     
    intNumDisplay = 0
    intCount = 0

    ' Navigate the record set
    objRecordSet.MoveFirst
    While Not objRecordSet.EOF
        intCount = intCount + 1
        For intTemp1 = 0 To objRecordSet.Fields.Count - 1
            If objRecordSet.Fields(intTemp1).Type = adVariant And Not (IsNull(objRecordSet.Fields(intTemp1).Value)) Then
              For intTemp2 = LBound(objRecordSet.Fields(intTemp1).Value) To UBound(objRecordSet.Fields(intTemp1).Value)
                 strText = objRecordSet.Fields(intTemp1).Value(intTemp2)
              Next
            Else
              
              GetObjectPath = objRecordSet.Fields(intTemp1).Value
            End If
        Next
        intNumDisplay = intNumDisplay + 1
        If intNumDisplay = intMAX_DISPLAY Then
            strText = ""
            intNumDisplay = 0
        End If
        objRecordSet.MoveNext
    Wend
    
    Set objRecordSet = Nothing
    Set objConnection = Nothing
    Set objCommand = Nothing
    Set objRoot = Nothing
    Set objDomain = Nothing
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : Logevent
' DateTime  : 03-05-2003 19:34
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author            History
' 1.0   03-05-2003  Dave Robinson     Initial Version
'---------------------------------------------------------------------------------------
Public Sub Logevent(strMessage As String, Optional iType As EventType = EventInfo)
    frmComputers.lblInfo = strMessage
    Select Case iType
        Case EventStop
            frmComputers.imgInfo.Picture = frmComputers.imlImages.ListImages.Item("Stop").Picture
        Case EventQuestion
            frmComputers.imgInfo.Picture = frmComputers.imlImages.ListImages.Item("Question").Picture
        Case EventExclamation
            frmComputers.imgInfo.Picture = frmComputers.imlImages.ListImages.Item("Exclamation").Picture
        Case EventInfo
            frmComputers.imgInfo.Picture = frmComputers.imlImages.ListImages.Item("Info").Picture
    End Select
    frmComputers.Refresh
End Sub
