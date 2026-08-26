VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUser"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Attribute VB_Description = "A network user.  Use the Init method to instantiate a CUser object.  It initializes on user sUser at server sServer.  Using a domain controller gets a network user, a regular server returns a user local at that server."
Option Explicit

Dim psNetName       As String
Dim psFullName      As String
Dim psComment       As String
Dim psHomeDir       As String
Dim psHomeDirDrive  As String
Dim plFlags         As Long
Dim psScriptPath    As String
Dim plLastLogon     As Long
Dim plLastLogoff    As Long
Dim psLogonServer   As String
Dim psProfile       As String
   
Private Type USER_INFO_3
    usri3_name              As Long     'LPWSTR in SDK
    usri3_password          As Long     'LPWSTR in SDK
    usri3_password_age      As Long     'DWORD in SDK
    usri3_priv              As Long     'DWORD in SDK
    usri3_home_dir          As Long     'LPWSTR in SDK
    usri3_comment           As Long     'LPWSTR in SDK
    usri3_flags             As Long     'DWORD in SDK
    usri3_script_path       As Long     'LPWSTR in SDK
    usri3_auth_flags        As Long     'DWORD in SDK
    usri3_full_name         As Long     'LPWSTR in SDK
    usri3_usr_comment       As Long     'LPWSTR in SDK
    usri3_parms             As Long     'LPWSTR in SDK
    usri3_workstations      As Long     'LPWSTR in SDK
    usri3_last_logon        As Long     'DWORD in SDK
    usri3_last_logoff       As Long     'DWORD in SDK
    usri3_acct_expires      As Long     'DWORD in SDK
    usri3_max_storage       As Long     'DWORD in SDK
    usri3_units_per_week    As Long     'DWORD in SDK
    usri3_logon_hours       As Long     'PBYTE in SDK
    usri3_bad_pw_count      As Long     'DWORD in SDK
    usri3_num_logons        As Long     'DWORD in SDK
    usri3_logon_server      As Long     'LPWSTR in SDK
    usri3_country_code      As Long     'DWORD in SDK
    usri3_code_page         As Long     'DWORD in SDK
    usri3_user_id           As Long     'DWORD in SDK
    usri3_primary_group_id  As Long     'DWORD in SDK
    usri3_profile           As Long     'LPWSTR in SDK
    usri3_home_dir_drive    As Long     'LPWSTR in SDK
    usri3_password_expired  As Long     'DWORD in SDK
End Type
 
Private Declare Function NetUserGetInfo Lib "netapi32.dll" _
   (strServerName As Any, _
    strUserName As Any, _
    ByVal dwLevel As Long, _
    pBuffer As Long) As Long

Private Declare Function NetGetDCName Lib "netapi32.dll" _
   (strServerName As Any, _
    strDomainName As Any, _
    bufptr As Long) As Long

' Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)



Property Get Comment() As String
Attribute Comment.VB_Description = "The Comment associated with this user."
    Comment = psComment
End Property
Property Get Flags() As Long
    Flags = plFlags
End Property
Property Get FullName() As String
Attribute FullName.VB_Description = "The full name of the user."
    FullName = psFullName
End Property

Property Get HomeDir() As String
Attribute HomeDir.VB_Description = "The users home directory, in UNC format."
    HomeDir = psHomeDir
End Property
Property Get HomeDirDrive() As String
Attribute HomeDirDrive.VB_Description = "The default drive letter to use for linking to the users homedir if none is specified."
    HomeDirDrive = psHomeDirDrive
End Property

Function Init(sNetUser As String, sServer As String) As Boolean
Attribute Init.VB_Description = "Initializes the CUser object on user sUser,  using the SAM database on server sServer.  If the server is a domain controller, the user is a domain user,  else a user local to sServer."

Dim lrc     As Long
Dim sWrk    As String

Dim pServer()   As Byte
Dim pUser()     As Byte
       
    psNetName = sNetUser
       
    pUser = sNetUser & vbNullChar
    pServer = sServer & vbNullChar
    psFullName = vbNullChar
    psComment = vbNullChar
    
    Dim dwLevel As Long
    dwLevel = 3
    
    Dim tmpBuffer    As USER_INFO_3
    Dim ptmpBuffer   As Long
    
    lrc = NetUserGetInfo(pServer(0), pUser(0), dwLevel, ptmpBuffer)
    'As last param is dimmed as long, the pointer to ptmpBuffer is
    'passed to dll, and the function returns a pointer to a pointer
    'to our UDT. Therefore ptmpBuffer on return holds a pointer
    'to our UDT.
    
     If lrc <> 0 Then
        Init = False
        Exit Function
     End If
    
    'Deference it!!!
    CopyMemory tmpBuffer, ptmpBuffer, LenB(tmpBuffer)
    
    Dim sByte() As Byte
    ReDim sByte(255)
    
    'Convert LPWSTR (Unicode string) to VB string.
    CopyMemory sByte(0), tmpBuffer.usri3_full_name, 256
    psFullName = sByte
    psFullName = Left$(psFullName, InStr(psFullName, vbNullChar) - 1)
    
    CopyMemory sByte(0), tmpBuffer.usri3_comment, 256
    psComment = sByte
    psComment = psComment & vbNullChar
    psComment = Left$(psComment, InStr(psComment, vbNullChar) - 1)
    
    CopyMemory sByte(0), tmpBuffer.usri3_home_dir, 256
    psHomeDir = sByte
    psHomeDir = psHomeDir & vbNullChar
    psHomeDir = Left$(psHomeDir, InStr(psHomeDir, vbNullChar) - 1)
    
    CopyMemory sByte(0), tmpBuffer.usri3_home_dir_drive, 256
    psHomeDirDrive = sByte
    psHomeDirDrive = psHomeDirDrive & vbNullChar
    psHomeDirDrive = Left$(psHomeDirDrive, InStr(psHomeDirDrive, vbNullChar) - 1)
    
    plFlags = tmpBuffer.usri3_flags
    plLastLogon = tmpBuffer.usri3_last_logon
    plLastLogoff = tmpBuffer.usri3_last_logoff
    
    CopyMemory sByte(0), tmpBuffer.usri3_script_path, 256
    psScriptPath = sByte
    psScriptPath = psScriptPath & vbNullChar
    psScriptPath = Left$(psScriptPath, InStr(psScriptPath, vbNullChar) - 1)
    
    CopyMemory sByte(0), tmpBuffer.usri3_logon_server, 256
    psLogonServer = sByte
    psLogonServer = psLogonServer & vbNullChar
    psLogonServer = Left$(psLogonServer, InStr(psLogonServer, vbNullChar) - 1)
    
    CopyMemory sByte(0), tmpBuffer.usri3_profile, 256
    psProfile = sByte
    psProfile = psProfile & vbNullChar
    psProfile = Left$(psProfile, InStr(psProfile, vbNullChar) - 1)
    
    Init = True

End Function
Property Get LastLogoff() As String
Attribute LastLogoff.VB_Description = "The time at which the user last logged off."
    LastLogoff = DateAdd("s", plLastLogoff, "01/01/1970 00:00:00")
End Property

Property Get LastLogon() As String
Attribute LastLogon.VB_Description = "The time at which the user last logged on to the domain."
    LastLogon = DateAdd("s", plLastLogon, "01/01/1970 00:00:00")
End Property
Property Get LogonServer() As String
Attribute LogonServer.VB_Description = "The default logon server to use when authenticating this user."
    LogonServer = psLogonServer
End Property
    
Property Get Profile() As String
Attribute Profile.VB_Description = "The UNC path to the users domain profile.  This is the users 'roaming' profile, if any."
    Profile = psProfile
End Property

Property Get ScriptPath() As String
Attribute ScriptPath.VB_Description = "The UNC path to the users logon script file, if any."
    ScriptPath = psScriptPath
End Property
Property Get UserName() As String
Attribute UserName.VB_Description = "The network user name of the user."
    UserName = psNetName
End Property
