Attribute VB_Name = "modUsers"
Public Declare Function NetUserAdd Lib "netapi32.dll" (ServerName As Byte, ByVal Level As Long, Buffer As USER_INFO_3, parm_err As Long) As Long
Public Declare Function LogonUser Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As Any, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Public Declare Function NetUserGetInfo Lib "netapi32.dll" (ServerName As Byte, Username As Byte, ByVal Level As Long, Buffer As Long) As Long
Public Declare Function NetUserEnum Lib "netapi32.dll" (ServerName As Byte, ByVal Level As Long, ByVal Filter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHwnd As Long) As Long
Public Declare Function NetUserChangePassword Lib "netapi32.dll" (ByVal domainname As String, ByVal Username As String, ByVal OldPassword As String, ByVal NewPassword As String) As Long
Public Declare Function NetUserSetInfo Lib "netapi32.dll" (ByVal ServerName As String, ByVal Username As String, ByVal Level As Long, UserInfo As Any, ParmError As Long) As Long


' ---------------------------------------------
' The USER_INFO_3 data structure
' ---------------------------------------------

Public Type USER_INFO_3
    usri3_name As Long
    usri3_password As Long
    usri3_password_age As Long
    usri3_priv As Long
    usri3_home_dir As Long
    usri3_comment As Long
    usri3_flags As Long
    usri3_script_path As Long
    usri3_auth_flags As Long
    usri3_full_name As Long
    usri3_usr_comment As Long
    usri3_parms As Long
    usri3_workstations As Long
    usri3_last_logon As Long
    usri3_last_logoff As Long
    usri3_acct_expires As Long
    usri3_max_storage As Long
    usri3_units_per_week As Long
    usri3_logon_hours As Long
    usri3_bad_pw_count As Long
    usri3_num_logons As Long
    usri3_logon_server As Long
    usri3_country_code As Long
    usri3_code_page As Long
    usri3_user_id As Long
    usri3_primary_group_id As Long
    usri3_profile As Long
    usri3_home_dir_drive As Long
    usri3_password_expired As Long
End Type

Public Type USERINFO_2_API
    usri2_name As Long
    usri2_password As Long
    usri2_password_age As Long
    usri2_priv As Long
    usri2_home_dir As Long
    usri2_comment As Long
    usri2_flags As Long
    usri2_script_path As Long
    usri2_auth_flags As Long
    usri2_full_name As Long
    usri2_usr_comment As Long
    usri2_parms As Long
    usri2_workstations As Long
    usri2_last_logon As Long
    usri2_last_logoff As Long
    usri2_acct_expires As Long
    usri2_max_storage As Long
    usri2_units_per_week As Long
    usri2_logon_hours As Long
    usri2_bad_pw_count As Long
    usri2_num_logons As Long
    usri2_logon_server As Long
    usri2_country_code As Long
    usri2_code_page As Long
End Type

Public Type USER_INFO_10_API
    Name As Long
    Comment As Long
    UsrComment As Long
    FullName As Long
End Type

Public Type USER_INFO_1003
  usri1003_password As Long
End Type


Public Function GetLoggedOnUsers(ByVal ServerName As String) As Variant
    
    Dim p_lngRtn As Long
    Dim p_lngPtrBuffer As Long
    Dim p_lngPtrUserInfoBuf As Long
    Dim p_lngEntriesRead As Long
    Dim p_lngTotalEntries As Long
    Dim p_lngResumeHwnd As Long
    Dim p_lngLoop As Long
    Dim p_lngLastLogon As Long
    Dim p_lngLastLogoff As Long
    Dim p_strUserName As String
    Dim p_abytServerName() As Byte
    Dim p_abytUserName() As Byte
    Dim p_atypUserInfo() As USER_INFO_10_API
    Dim p_typUserInfo As USERINFO_2_API
    
    ' ------------------------------------------
    ' Initialize the variable(s)
    ' ------------------------------------------
    
    If ServerName = "" Then
        p_abytServerName = Chr$(0)
    Else
        p_abytServerName = "\\" & ServerName & Chr$(0)
    End If
    
    ' ------------------------------------------
    ' Make appropriate API call and check for error
    ' ------------------------------------------
    
    p_lngRtn = NetUserEnum(ServerName:=p_abytServerName(0), _
    Level:=10, _
    Filter:=0&, _
    Buffer:=p_lngPtrBuffer, _
    PrefMaxLen:=&H4000, _
    EntriesRead:=p_lngEntriesRead, _
    TotalEntries:=p_lngTotalEntries, _
    ResumeHwnd:=p_lngResumeHwnd)
    
    If p_lngRtn <> 0 Then
        MsgBox "Had an error with NetUserEnum, " & CStr(p_lngRtn), _
        Buttons:=vbInformation, _
        Title:="GetLoggedOnUsers"
        Exit Function
    End If
    
    ' ------------------------------------------
    ' Exit if no entries found
    ' ------------------------------------------
    
    If p_lngEntriesRead < 1 Then
        Exit Function
    End If
    
    ' ------------------------------------------
    ' Redim the type array to hold this info
    ' ------------------------------------------
    
    ReDim p_atypUserInfo(0 To p_lngEntriesRead - 1)
    
    ' ------------------------------------------
    ' Copy the pointer to the buffer into the
    ' type array
    ' ------------------------------------------
    
    CopyMem p_atypUserInfo(0), _
    ByVal p_lngPtrBuffer, _
    Len(p_atypUserInfo(0)) * p_lngEntriesRead
    
    ' ------------------------------------------
    ' Fill-in the info needed to call the
    ' Add() method
    ' NOTE: We will always have +1 open pipe,
    ' since in making this call we create
    ' a pipe, "\PIPE\srvsvc"
    ' ------------------------------------------
    For p_lngLoop = 0 To p_lngEntriesRead - 1
        p_strUserName = PointerToUnicodeStr(p_atypUserInfo(p_lngLoop).Name)
        p_abytUserName = p_strUserName & Chr(0)
        p_lngRtn = NetUserGetInfo(ServerName:=p_abytServerName(0), _
        Username:=p_abytUserName(0), _
        Level:=2, _
        Buffer:=p_lngPtrUserInfoBuf)
        
        If p_lngRtn <> 0 Then
            Debug.Print "Had an error with NetUserGetInfo, " & CStr(p_lngRtn), '_
            'Buttons:=vbInformation, _
            'Title:="GetLoggedOnUsers"
            Exit Function
        End If
        
        CopyMem p_typUserInfo, _
        ByVal p_lngPtrUserInfoBuf, _
        Len(p_typUserInfo)
        
        p_lngLastLogon = p_typUserInfo.usri2_last_logon
        p_lngLastLogoff = p_typUserInfo.usri2_last_logoff
        
        If p_lngLastLogoff = 0 And p_lngLastLogon = 0 Then
            Debug.Print " **** " & p_strUserName & " has NEVER logged in"
        ElseIf (p_lngLastLogoff < p_lngLastLogon) Then
            Debug.Print p_strUserName & " is still logged in -- " & p_lngLastLogoff, p_lngLastLogon
        Else
            Debug.Print " **** " & p_strUserName & " is NOT logged in"
        End If
        
        If p_lngPtrUserInfoBuf <> 0 Then
            NetApiBufferFree p_lngPtrUserInfoBuf
        End If
        
    Next p_lngLoop
    
    ' ------------------------------------------
    ' Clean-up the buffer
    ' ------------------------------------------
    
    If p_lngPtrBuffer <> 0 Then
        NetApiBufferFree p_lngPtrBuffer
    End If

End Function

Public Function ChangePassword(strUserName As String, strDomain As String, strOldPwl As String, strNewPwl As String) As Boolean
    
    Dim sServer As String, sUser As String
    Dim sNewPass As String, sOldPass As String
    Dim UI1003 As USER_INFO_1003
    Dim dwLevel As Long
    Dim lRet As String
    Dim sNew As String
    
    ' StrConv Functions are necessary since VB will perform
    ' UNICODE/ANSI translation before passing strings to the
    ' NETAPI functions
    
    sUser = StrConv(strUserName, vbUnicode)
    sNewPass = StrConv(strNewPwl, vbUnicode)
    'See if this is Domain or Computer referenced
    
    If Left(strDomain, 2) = "\\" Then
        sServer = StrConv(strDomain, vbUnicode)
    Else
        ' Domain was referenced, get the Primary Domain Controller
        sServer = StrConv(GetPrimaryDCName(strDomain), vbUnicode)
    End If
    
    If strOldPwl = "" Then
        ' Administrative over-ride of existing password.
        ' Does not require old password
        dwLevel = 1003
        sNew = strNewPwl
        UI1003.usri1003_password = StrPtr(sNew)
        lRet = NetUserSetInfo(sServer, sUser, dwLevel, UI1003, 0&)
    Else
        ' Set the Old Password and attempt to change the user's password
        sOldPass = StrConv(strOldPwl, vbUnicode)
        lRet = NetUserChangePassword(sServer, sUser, sOldPass, sNewPass)
    End If
    
    If lRet <> 0 Then
        ChangePassword = False
    Else
        ChangePassword = True
    End If

End Function

Public Function AddUser(ByVal xi_strServerName As String, ByVal xi_strUserName As String, ByVal xi_strPassword As String, Optional ByVal xi_strUserFullName As String = vbNullString, Optional ByVal xi_strUserComment As String = vbNullString) As Boolean
    
    Dim p_strErr As String
    Dim p_lngRtn As Long
    Dim p_lngPtrUserName As Long
    Dim p_lngPtrPassword As Long
    Dim p_lngPtrUserFullName As Long
    Dim p_lngPtrUserComment As Long
    Dim p_lngParameterErr As Long
    Dim p_lngFlags As Long
    Dim p_abytServerName() As Byte
    Dim p_abytUserName() As Byte
    Dim p_abytPassword() As Byte
    Dim p_abytUserFullName() As Byte
    Dim p_abytUserComment() As Byte
    Dim p_typUserInfo3 As USER_INFO_3
    
    If xi_strUserFullName = vbNullString Then
        xi_strUserName = xi_strUserName
    End If
    
    ' ------------------------------------------
    ' Create byte arrays to avoid Unicode hassles
    ' ------------------------------------------
    
    p_abytServerName = xi_strServerName & vbNullChar
    p_abytUserName = xi_strUserName & vbNullChar
    p_abytUserFullName = xi_strUserFullName & vbNullChar
    p_abytPassword = xi_strPassword & vbNullChar
    p_abytUserComment = xi_strUserComment & vbNullChar
    
    ' ------------------------------------------
    ' Allocate buffer space
    ' ------------------------------------------
    
    p_lngRtn = NetApiBufferAllocate(UBound(p_abytUserName), p_lngPtrUserName)
    
    p_lngRtn = NetApiBufferAllocate(UBound(p_abytUserFullName), p_lngPtrUserFullName)
    
    p_lngRtn = NetApiBufferAllocate(UBound(p_abytPassword), p_lngPtrPassword)
    
    p_lngRtn = NetApiBufferAllocate(UBound(p_abytUserComment), p_lngPtrUserComment)
    
    ' ------------------------------------------
    ' Get pointers to the byte arrays
    ' ------------------------------------------
    
    p_lngPtrUserName = VarPtr(p_abytUserName(0))
    p_lngPtrUserFullName = VarPtr(p_abytUserFullName(0))
    p_lngPtrPassword = VarPtr(p_abytPassword(0))
    p_lngPtrUserComment = VarPtr(p_abytUserComment(0))
    
    ' ------------------------------------------
    ' Fill the VB structure
    ' ------------------------------------------
    
    p_lngFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT Or UF_DONT_EXPIRE_PASSWD
    
    With p_typUserInfo3
        .usri3_acct_expires = TIMEQ_FOREVER ' Never expires
        .usri3_comment = p_lngPtrUserComment ' Comment
        .usri3_flags = p_lngFlags ' There are a number of variations
        .usri3_full_name = p_lngPtrUserFullName ' User's full name
        .usri3_max_storage = USER_MAXSTORAGE_UNLIMITED ' Can use any amount
        'of disk space
        .usri3_name = p_lngPtrUserName ' Name of user account
        .usri3_password = p_lngPtrPassword ' Password for user account
        .usri3_primary_group_id = DOMAIN_GROUP_RID_USERS ' You MUST use this
        'constant for NetUserAdd
        .usri3_script_path = 0& ' Path of user's logon script
        .usri3_auth_flags = 0& ' Ignored by NetUserAdd
        .usri3_bad_pw_count = 0& ' Ignored by NetUserAdd
        .usri3_code_page = 0& ' Code page for user's language
        .usri3_country_code = 0& ' Country code for user's language
        .usri3_home_dir = 0& ' Can specify path of home directory of this
        'user
        .usri3_home_dir_drive = 0& ' Drive letter assign to user's
        'profile
        .usri3_last_logoff = 0& ' Not needed when adding a user
        .usri3_last_logon = 0& ' Ignored by NetUserAdd
        .usri3_logon_hours = 0& ' Null means no restrictions
        .usri3_logon_server = 0& ' Null means logon to domain server
        .usri3_num_logons = 0& ' Ignored by NetUserAdd
        .usri3_parms = 0& ' Used by specific applications
        .usri3_password_age = 0& ' Ignored by NetUserAdd
        .usri3_password_expired = 0& ' None-zero means user must change
        'password at next logon
        .usri3_priv = 0& ' Ignored by NetUserAdd
        .usri3_profile = 0& ' Path to a user's profile
        .usri3_units_per_week = 0& ' Ignored by NetUserAdd
        .usri3_user_id = 0& ' Ignored by NetUserAdd
        .usri3_usr_comment = 0& ' User comment
        .usri3_workstations = 0& ' Workstations a user can log onto (null
        '= all stations)
    End With
    
    ' ------------------------------------------
    ' Attempt to add the user
    ' ------------------------------------------
    
    p_lngRtn = NetUserAdd(p_abytServerName(0), _
    constUserInfoLevel3, p_typUserInfo3, p_lngParameterErr)
    
    ' ------------------------------------------
    ' Check for error
    ' ------------------------------------------
    
    If p_lngRtn <> 0 Then
        AddUser = False
        Select Case p_lngRtn
            Case ERROR_ACCESS_DENIED
                p_strErr = "User doesn't have sufficient access rights."
            Case NERR_GroupExists
                p_strErr = "The group already exists."
            Case NERR_NotPrimary
                p_strErr = "Can only do this operation on the PDC of the domain."
            Case NERR_UserExists
                p_strErr = "The user account already exists."
            Case NERR_PasswordTooShort
                p_strErr = "The password is shorter than required."
            Case NERR_InvalidComputer
                p_strErr = "The computer name is invalid."
            Case Else
                p_strErr = "Unknown error #" & CStr(p_lngRtn)
        End Select
        On Error GoTo 0
        Err.Raise Number:=p_lngRtn, _
        Description:=p_strErr & vbCrLf & _
        "Error in parameter " & p_lngParameterErr & _
        " when attempting to add the user, " & xi_strUserName, _
        Source:="Form1.AddUser"
    Else
        AddUser = True
    End If
    
    ' ------------------------------------------
    ' Be a good programmer and free the memory
    ' you've allocated
    ' ------------------------------------------
    
    p_lngRtn = NetApiBufferFree(p_lngPtrUserName)
    p_lngRtn = NetApiBufferFree(p_lngPtrPassword)
    p_lngRtn = NetApiBufferFree(p_lngPtrUserFullName)
    p_lngRtn = NetApiBufferFree(p_lngPtrUserComment)

End Function

Public Function AddUserToLocal(ByVal xi_strGroupName As String, ByVal xi_strUserName As String, ByVal xi_strServerName As String) As Boolean
 
    Dim p_lngPtrGroupName As Long
    Dim p_lngPtrUserName As Long
    Dim p_lngPtrServerName As Long
    Dim p_lngMemberCount As Long
    Dim p_lngRtn As Long
    
    ' Convert the server name to a pointer
    If Len(Trim$(xi_strServerName)) = 0 Then
        p_lngPtrServerName = 0&
    Else
        p_lngPtrServerName = StrPtr(xi_strServerName)
    End If
    
    ' Convert the group name to a pointer
    p_lngPtrGroupName = StrPtr(xi_strGroupName)
    
    ' Convert the user name to a pointer
    p_lngPtrUserName = StrPtr(xi_strUserName)
    
    ' Add the user
    p_lngMemberCount = 1
    
    p_lngRtn = NetLocalGroupAddMembers(p_lngPtrServerName, _
    p_lngPtrGroupName, _
    LocalGroupMembersInfo3, _
    p_lngPtrUserName, _
    p_lngMemberCount)
    
    If p_lngRtn = NERR_Success Then
        AddUserToLocal = True
    Else
        AddUserToLocal = False
    End If

End Function

Public Function Login(ByVal xi_strUserID As String, ByVal xi_strPassword As String) As Boolean
    
    On Error Resume Next ' Don't accept errors here
    
    Dim p_lngToken As Long
    Dim p_lngRtn As Long
    
    p_lngRtn = LogonUser(lpszUsername:=xi_strUserID, _
    lpszDomain:=0&, _
    lpszPassword:=xi_strPassword, _
    dwLogonType:=LOGON32_LOGON_NETWORK, _
    dwLogonProvider:=LOGON32_PROVIDER_DEFAULT, _
    phToken:=p_lngToken)
    
    If p_lngRtn = 0 Then
        Login = False
    Else
        Login = True
    End If
    
    On Error GoTo 0

End Function


