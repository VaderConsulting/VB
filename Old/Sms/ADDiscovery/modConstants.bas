Attribute VB_Name = "modConstants"

' ---------------------------------------------
' Possible errors with API call
' ---------------------------------------------

Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const NERR_BASE As Long = 2100
Public Const NERR_GroupExists As Long = NERR_BASE + 123
Public Const NERR_NotPrimary As Long = NERR_BASE + 126
Public Const NERR_UserExists As Long = NERR_BASE + 124
Public Const NERR_PasswordTooShort As Long = NERR_BASE + 145
Public Const NERR_InvalidComputer As Long = NERR_BASE + 251
Public Const NERR_Success As Long = 0&

' ---------------------------------------------
' General constants used
' ---------------------------------------------

Public Const constUserInfoLevel3 As Long = 3
Public Const TIMEQ_FOREVER As Long = -1&
Public Const MAX_PATH As Long = 260&
Public Const DOMAIN_GROUP_RID_USERS As Long = &H201&
Public Const USER_MAXSTORAGE_UNLIMITED As Long = -1&
Public Const LocalGroupMembersInfo3 As Long = 3&
Public Const MAX_RESOURCES As Long = 256
Public Const NOT_A_CONTAINER As Long = -1
Public Const RESOURCE_GLOBALNET As Long = &H2&
Public Const RESOURCETYPE_ANY As Long = &H0&
Public Const RESOURCEUSAGE_ALL As Long = &H0&
Public Const NO_ERROR As Long = 0&
Public Const RESOURCE_ENUM_ALL As Long = &HFFFF

' ---------------------------------------------
' Constants used by LogonUser
' ---------------------------------------------

Public Const LOGON32_PROVIDER_DEFAULT As Long = 0&
Public Const LOGON32_PROVIDER_WINNT35 As Long = 1&
Public Const LOGON32_LOGON_INTERACTIVE As Long = 2&
Public Const LOGON32_LOGON_NETWORK As Long = 3&
Public Const LOGON32_LOGON_BATCH As Long = 4&
Public Const LOGON32_LOGON_SERVICE As Long = 5&

' ---------------------------------------------
' Used by usri3_flags element of data structure
' ---------------------------------------------

Public Const UF_SCRIPT As Long = &H1&
Public Const UF_ACCOUNTDISABLE As Long = &H2&
Public Const UF_HOMEDIR_REQUIRED As Long = &H8&
Public Const UF_LOCKOUT As Long = &H10&
Public Const UF_PASSWD_NOTREQD As Long = &H20&
Public Const UF_PASSWD_CANT_CHANGE As Long = &H40&
Public Const UF_DONT_EXPIRE_PASSWD As Long = &H10000
Public Const STILL_ACTIVE As Long = &H103&
Public Const UF_NORMAL_ACCOUNT As Long = &H200&
Public Const UF_SERVER_TRUST_ACCOUNT As Long = &H2000&
Public Const PROCESS_QUERY_INFORMATION As Long = &H400&
Public Const UF_TEMP_DUPLICATE_ACCOUNT As Long = &H100&
Public Const UF_INTERDOMAIN_TRUST_ACCOUNT As Long = &H800&
Public Const UF_WORKSTATION_TRUST_ACCOUNT As Long = &H1000&
