'Name:		NetUser.Bas
'Description:	Network User Account Functions for Windows NT
'Dependencies:	NetDecs.Bas
'		NetErr.BAS
'Included
'Functions:	NetGetUserName
'		NetResetPassword
'		NetGetUser
'		NetSetUser
'		NetAddUser
'		NetDelUser
'------------------------------------------------------------------------

Public Function NetGetUserName() As String
On Error Resume Next
    UserName$ = Space(256)
    cbusername& = Len(UserName$)
    ret& = WNetGetUser(ByVal 0&, UserName$, cbusername&)
    If ret& = 0 Then
        UserName$ = Left$(UserName$, InStr(UserName$, Chr$(0)) - 1)
    Else
        UserName$ = ""
    End If
    NetGetUserName = UserName$
End Function

Public Sub NetResetPassword(ByVal ServerName$, ByVal UserName$)
Dim USER_TYPE As Long, arrayNewPassword() As Byte
Dim UserStructPlayful As PlayfulStructType

    TripError = False
    
    NewPassword$ = UserName$
    
    Call ConvStringToByteArray(NewPassword$, arrayNewPassword)
    Call ConvArrayToPointer(arrayNewPassword, ptrNewPassword&)
    
    With UserStructPlayful
        .ptrVar = ptrNewPassword&
    End With 'UserStructPlayful
    
    USER_TYPE = 1003 'USER_INFO_TYPE_1003 (AS DESCRIBED IN WIN 32 SDK CDROM)
    Reply& = NetUserResetPassword(StrConv(ServerName$, vbUnicode), StrConv(UserName$, vbUnicode), USER_TYPE, UserStructPlayful, ParmError&)
    Call NetReturnError(Reply&, UserName$)

    Call ConvPointerToNothing(ptrNewPassword&)

End Sub

Public Sub NetSetUser(ByVal ServerName$, xUserName$, xUser As UserRecord)
'THIS FUNCTION SETS USER INFORMATION I WANTED TO CHANGE:
'---USERNAME
'---FULL NAME
'---DESCRIPTION
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

'CUSTOMIZATION NOTE.  YOU MAY WANT TO INVESTIGATE THE WIN 32 SDK CDROM TO FIND WAYS TO
'SET OTHER KINDS OF THINGS BESIDES WHAT I SET HERE

Dim arrayServerName() As Byte, arrayUserName() As Byte, arrayXUserName() As Byte, arrayFullName() As Byte, arrayDescription() As Byte
Dim UserStructPlayful As PlayfulStructType

    Dim USER_TYPE As Long, ParmError&, Reply&, UserName$, FullName$, Description$, ptrServerName&, ptrUserName&, ptrXUserName&, ptrFullName&, ptrDescription&
    
TripError = False

    UserName$ = xUser.Name
    FullName$ = xUser.FullName
    Description$ = xUser.Description
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(UserName$, arrayUserName)
    Call ConvStringToByteArray(xUserName$, arrayXUserName)
    Call ConvStringToByteArray(FullName$, arrayFullName)
    Call ConvStringToByteArray(Description$, arrayDescription)
    
    Call ConvArrayToPointer(arrayServerName, ptrServerName&)
    Call ConvArrayToPointer(arrayUserName, ptrUserName&)
    Call ConvArrayToPointer(arrayXUserName, ptrXUserName&)
    Call ConvArrayToPointer(arrayFullName, ptrFullName&)
    Call ConvArrayToPointer(arrayDescription, ptrDescription&)
    
    With UserStructPlayful
        .ptrVar = ptrFullName&
    End With 'UserStructPlayful
    
    USER_TYPE = 1011  'USER_INFO_TYPE_1011 (AS DESCRIBED IN WIN 32 SDK CDROM)
    Reply& = NetUserSetInfo(arrayServerName(0), arrayXUserName(0), USER_TYPE, UserStructPlayful, ParmError&)
    Call NetReturnError(Reply&, UserName$)
    
    With UserStructPlayful
        .ptrVar = ptrDescription&
    End With 'UserStructPlayful
    
    USER_TYPE = 1007 'USER_INFO_TYPE_1007 (AS DESCRIBED IN WIN 32 SDK CDROM)
    Reply& = NetUserSetInfo(arrayServerName(0), arrayXUserName(0), USER_TYPE, UserStructPlayful, ParmError&)
    Call NetReturnError(Reply&, UserName$)
    
    With UserStructPlayful
        .ptrVar = ptrUserName&
    End With 'UserStructPlayful
    
    USER_TYPE = 0 'USER_INFO_TYPE_0 (AS DESCRIBED IN WIN 32 SDK CDROM)
    Reply& = NetUserSetInfo(arrayServerName(0), arrayXUserName(0), USER_TYPE, UserStructPlayful, ParmError&)
    Call NetReturnError(Reply&, UserName$)

End Sub

Public Sub NetGetUser(ByVal ServerName$, ByVal UserName$, returnUserName$, returnFullName$, returnDescription$)
'THIS FUNCTION RETRIEVES USER INFORMATION I WANTED:
'---USERNAME
'---FULL NAME
'---DESCRIPTION
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

'CUSTOMIZATION NOTE.  YOU MAY WANT TO INVESTIGATE THE WIN 32 SDK CDROM TO FIND WAYS TO
'SET OTHER KINDS OF THINGS BESIDES WHAT I SET HERE

Dim arrayServerName() As Byte, arrayUserName() As Byte, arrayDescription(255) As Byte, arrayFullName(255) As Byte
Dim ptrServerName&, ptrUserName&, i&, ptrDescription&, ptrFullName&, ptrBuffer&, Reply&
Dim USER_TYPE As Integer

TripError = False

    UserName$ = Trim(UserName$)
    
    If InStr(UserName$, "$") Then Exit Sub
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(UserName$, arrayUserName)

    Call ConvArrayToPointer(arrayServerName, ptrServerName&)
    Call ConvArrayToPointer(arrayUserName, ptrUserName&)
    
    i& = 512
    Reply& = NetAPIBufferAllocate(i&, ptrBuffer&)
    USER_TYPE = 10
    Reply& = NetUserGetInfo(arrayServerName(0), arrayUserName(0), USER_TYPE, ptrBuffer&)
    Call NetReturnError(Reply&, UserName$)
    returnDescription$ = ExtractFieldFromStruct(ptrBuffer&, 2)
    ConvPointerToNothing (ptrBuffer&)
    
    'IMPORTANT NOTE:  DO YOU SEE HERE HOW I GO OUT AND CALL NetUserGetInfo ?
    'IF I DIDN'T CALL IT AGAIN FOR THE NEXT FIELD, AND I TRIED TO RETRIEVE ALL FIELDS FROM ptrBuffer&,
    'THEN THE SYSTEM WOULD CRASH.  REASON UNKNOWN, BUT I SUSPECT POINTER BECOMES "UNLOCKED"
    'AND THUS POINTS TO A DIFFERENT SECTION OF MEMORY.
    i& = 512
    Reply& = NetAPIBufferAllocate(i&, ptrBuffer&)
    USER_TYPE = 10
    Reply& = NetUserGetInfo(arrayServerName(0), arrayUserName(0), USER_TYPE, ptrBuffer&)
    Call NetReturnError(Reply&, UserName$)
    returnFullName$ = ExtractFieldFromStruct(ptrBuffer&, 4)
    ConvPointerToNothing (ptrBuffer&)

    'SEE IMPORTANT NOTE ABOVE--APPLIES HERE TOO.
    i& = 512
    Reply& = NetAPIBufferAllocate(i&, ptrBuffer&)
    USER_TYPE = 10
    Reply& = NetUserGetInfo(arrayServerName(0), arrayUserName(0), USER_TYPE, ptrBuffer&)
    Call NetReturnError(Reply&, UserName$)
    returnUserName$ = ExtractFieldFromStruct(ptrBuffer&, 1)
    ConvPointerToNothing (ptrBuffer&)
    
End Sub

Public Sub NetDelUser(ByVal ServerName$, ByVal UserName$)
'THIS FUNCTION DELETES A USER FROM AN NT DOMAIN
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayUserName() As Byte, arrayServerName() As Byte

    TripError = False
    
    Call ConvStringToByteArray(UserName$, arrayUserName)
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    
    Reply& = NetUserDel(arrayServerName(0), arrayUserName(0))
    Call NetReturnError(Reply&, UserName$)

End Sub

Public Sub NetAddUser(ByVal ServerName$, ByVal UserName$, ByVal Password$, FullName$, Comment$, HomeDrive$, HomeDir$, Script$)
'THIS FUNCTION ADDS A USER TO THE NT DOMAIN WITH THE FOLLOWING ASSUMPTIONS
'----PASSWORD WILL BE SAME AS USERNAME
'----A PROFILE FILE NAME ISN'T NECESSARY TO BE SET (YOU CAN ADJUST CODE BELOW IF YOU WANT TO ADD PROFILE NAME)
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayUserName() As Byte, arrayPassword() As Byte, arrayHomeDir() As Byte
Dim arrayHomeDrive() As Byte, arrayComment() As Byte, arrayScript() As Byte, arrayFullName() As Byte

Dim UserStructLevel2 As Level_2_User_Structure_Type
Dim UserStructHomeDrive As HomeDriveType
Dim UserStructLogonScript As LoginScriptType

Const USER_PRIV_MASK = &H3
Const USER_PRIV_GUEST = &H0
Const USER_PRIV_USER = &H1
Const USER_PRIV_ADMIN = &H2
Const UF_SCRIPT = &H1
Const UF_ACCOUNTDISABLE = &H2
Const UF_HOMEDIR_REQUIRED = &H8
Const UF_LOCKOUT = &H10
Const UF_PASSWD_NOTREQD = &H20
Const UF_PASSWD_CANT_CHANGE = &H40
Const UF_NORMAL_ACCOUNT = &H200
Const DEFAULT_FLAG = &H0
Const MAX_FLAG = &HFFFFFFFF

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(UserName$, arrayUserName)
    Call ConvStringToByteArray(Password$, arrayPassword)
    Call ConvStringToByteArray(HomeDir$, arrayHomeDir)
    Call ConvStringToByteArray(HomeDrive$, arrayHomeDrive)
    Call ConvStringToByteArray(Comment$, arrayComment)
    Call ConvStringToByteArray(Script$, arrayScript)
    Call ConvStringToByteArray(FullName$, arrayFullName)
  
    Call ConvArrayToPointer(arrayServerName, ptrServerName&)
    Call ConvArrayToPointer(arrayUserName, ptrUserName&)
    Call ConvArrayToPointer(arrayPassword, ptrPassword&)
    Call ConvArrayToPointer(arrayHomeDir, ptrHomeDir&)
    Call ConvArrayToPointer(arrayHomeDrive, ptrHomeDrive&)
    Call ConvArrayToPointer(arrayComment, ptrComment&)
    Call ConvArrayToPointer(arrayScript, ptrScript&)
    Call ConvArrayToPointer(arrayFullName, ptrFullName&)
  
    With UserStructLevel2
        .ptrName = ptrUserName&
        .ptrPassword = ptrPassword&
        .dwPasswordAge = 3
        .dwPriv = USER_PRIV_USER
        .ptrHomeDir = ptrHomeDir&
        .ptrComment = ptrComment&
        .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT
        .ptrScriptPath = DEFAULT_FLAG
        .dwAuthFlags = DEFAULT_FLAG
        .ptrFullName = ptrFullName&
        .ptrComment2 = DEFAULT_FLAG
        .ptrParms = DEFAULT_FLAG
        .ptrWorkstations = DEFAULT_FLAG
        .dwLastLogon = DEFAULT_FLAG
        .dwLastLogoff = DEFAULT_FLAG
        .dwAccountExpires = MAX_FLAG
        .dwMaxStorage = DEFAULT_FLAG
        .dwUnitsPerWeek = DEFAULT_FLAG
        .pbLogonHours = DEFAULT_FLAG
        .dwBadPasswordCount = MAX_FLAG
        .dwNumLogons = MAX_FLAG
        .ptrLogonServer = DEFAULT_FLAG
        .dwCountryCode = DEFAULT_FLAG
        .dwCodePage = DEFAULT_FLAG
    End With 'UserStructLevel2
    
    Const USER_LEVEL_2 = 2
    Reply& = NetUserAdd2(arrayServerName(0), USER_LEVEL_2, UserStructLevel2, ParmError&)
    Call NetReturnError(Reply&, UserName$)
    
    With UserStructHomeDrive
        .ptrHomeDrive = ptrHomeDrive&
    End With 'UserStructHomeDrive
    
    Const HOME_DRIVE_TYPE = 1053
    Reply& = NetUserSetInfoHomeDrive(arrayServerName(0), arrayUserName(0), HOME_DRIVE_TYPE, UserStructHomeDrive, ParmError&)
    Call NetReturnError(Reply&, UserName$)
    
    With UserStructLogonScript
        .ptrScriptPath = ptrScript&
    End With 'UserStructLogonScript
    
    Const LOGIN_SCRIPT_TYPE = 1009
    Reply& = NetUserSetInfoLogonScript(arrayServerName(0), arrayUserName(0), LOGIN_SCRIPT_TYPE, UserStructLogonScript, ParmError&)
    Call NetReturnError(Reply&, UserName$)
    
    ConvPointerToNothing (ptrServerName&)
    ConvPointerToNothing (ptrUserName&)
    ConvPointerToNothing (ptrPassword&)
    ConvPointerToNothing (ptrHomeDir&)
    ConvPointerToNothing (ptrHomeDrive&)
    ConvPointerToNothing (ptrComment&)
    ConvPointerToNothing (ptrScript&)
    ConvPointerToNothing (ptrFullName&)
    
End Sub



