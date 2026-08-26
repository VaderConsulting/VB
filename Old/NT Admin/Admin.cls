VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public UserCount As Integer
Public cDomainName As String
Public NTGroup As String
Public UserList As String
Public PrimaryDomainController As String

'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.22172/lngWId.1/qx/vb/scripts/ShowCode.htm for details.

'- Module 1 taken from http://support.microsoft.com/support/kb/articles/Q159/4/98.ASP -HOWTO: Call LanMan Services from 32-bit Visual Basic Apps
  
  Option Explicit
  Option Base 0 ' Important assumption for this code

  Private Type WKSTA_INFO_101
    wki101_platform_id As Long
    wki101_computername As Long
    wki101_langroup As Long
    wki101_ver_major As Long
    wki101_ver_minor As Long
    wki101_lanroot As Long
  End Type

  Private Type WKSTA_USER_INFO_1
    wkui1_username As Long
    wkui1_logon_domain As Long
    wkui1_logon_server As Long
    wkui1_oth_domains As Long
  End Type

  Private Type MungeLong
    X As Long
    Dummy As Integer
  End Type


  Private Type MungeInt
    XLo As Integer
    XHi As Integer
    Dummy As Integer
  End Type


  Private Type TUser0 ' Level 0
    ptrName As Long
  End Type


  Private Type TUser1 ' Level 1
    ptrName As Long
    ptrPassword As Long
    dwPasswordAge As Long
    dwPriv As Long
    ptrHomeDir As Long
    ptrComment As Long
    dwFlags As Long
    ptrScriptPath As Long
  End Type
    '
    ' for dwPriv
    '
  Const USER_PRIV_MASK = &H3
  Const USER_PRIV_GUEST = &H0
  Const USER_PRIV_USER = &H1
  Const USER_PRIV_ADMIN = &H2
  '
  ' for dwFlags
  '
  Const UF_SCRIPT = &H1
  Const UF_ACCOUNTDISABLE = &H2
  Const UF_HOMEDIR_REQUIRED = &H8
  Const UF_LOCKOUT = &H10
  Const UF_PASSWD_NOTREQD = &H20
  Const UF_PASSWD_CANT_CHANGE = &H40
  Const UF_NORMAL_ACCOUNT = &H200 ' Needs to be ORed with the
  ' other flags
  '
  ' for lFilter
  '
  Const FILTER_NORMAL_ACCOUNT = &H2


  Private Declare Function NetGetDCName Lib "NETAPI32.DLL" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
  Private Declare Function NetUserDel Lib "NETAPI32.DLL" (ServerName As Byte, UserName As Byte) As Long
  Private Declare Function NetGroupAddUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
  Private Declare Function NetGroupDelUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
      ' Add using Level 1 user structure
  Private Declare Function NetUserAdd1 Lib "NETAPI32.DLL" Alias "NetUserAdd" (ServerName As Byte, ByVal Level As Long, buffer As TUser1, ParmError As Long) As Long
      ' Enumerate using Level 0 user structure
      '
  Private Declare Function NetUserEnum0 Lib "NETAPI32.DLL" Alias "NetUserEnum" (ServerName As Byte, ByVal Level As Long, ByVal lFilter As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
  Private Declare Function NetGroupEnumUsers0 Lib "NETAPI32.DLL" Alias "NetGroupGetUsers" (ServerName As Byte, GroupName As Byte, ByVal Level As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
  Private Declare Function NetGroupEnum0 Lib "NETAPI32.DLL" Alias "NetGroupEnum" (ServerName As Byte, ByVal Level As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
  Private Declare Function NetUserGetGroups0 Lib "NETAPI32.DLL" Alias "NetUserGetGroups" (ServerName As Byte, UserName As Byte, ByVal Level As Long, buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long
  Private Declare Function NetApiBufferFree Lib "NETAPI32.DLL" (ByVal Ptr As Long) As Long
  Private Declare Function NetAPIBufferAllocate Lib "NETAPI32.DLL" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
  Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
  Private Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
  Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
  Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
  Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)
  Private Declare Sub RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal size&)
  Private Declare Function NetWkstaGetInfo& Lib "Netapi32" (strServer As Any, ByVal lLevel&, pbBuffer As Any)
  Private Declare Function NetWkstaUserGetInfo& Lib "Netapi32" (reserved As Any, ByVal lLevel&, pbBuffer As Any)

Private Sub Class_Initialize()
  UserCount = 0
  DomainName = ""
  NTGroup = ""
  PrimaryDomainController = ""
End Sub

'Public Property Set DomainName(strDomainName As String)
'  cDomainName = strDomainName
'End Property

Public Property Get DomainName(sDomainName As String)
  DomainName = cDomainName
End Property

Function AddUserToGroup(ByVal SName As String, _
  ByVal GName As String, ByVal UName As String) As Long
  '
  ' This only adds users to global groups - not to local groups
  '
  Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, _
  Result As Long
  SNArray = SName & vbNullChar
  GNArray = GName & vbNullChar
  UNArray = UName & vbNullChar
  Result = NetGroupAddUser(SNArray(0), GNArray(0), UNArray(0))
  If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & GName & "'"
  AddUserToGroup = Result
End Function


Function DelUser(ByVal SName As String, ByVal UName As String) As Long
  Dim UNArray() As Byte, SNArray() As Byte
  UNArray = UName & vbNullChar
  SNArray = SName & vbNullChar
  DelUser = NetUserDel(SNArray(0), UNArray(0))
End Function


Function DelUserFromGroup(ByVal SName As String, _
  ByVal GName As String, ByVal UName As String) As Long
  '
  ' This only deletes users from global groups - not local groups
  '
  Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, _
  Result As Long
  SNArray = SName & vbNullChar
  GNArray = GName & vbNullChar
  UNArray = UName & vbNullChar
  Result = NetGroupDelUser(SNArray(0), GNArray(0), UNArray(0))
  If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & GName & "'"
  DelUserFromGroup = Result
End Function


Function EnumerateGroups(ByVal SName As String, ByVal UName As String) As Long
  '
  ' Enumerates global groups only - not local groups
  '
  ' The buffer is filled from the left with pointers to user names that
  ' are filled from the right side. For example:
  '
  ' ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
  ' ^-------------- BufPtr buffer ----------------^
  '
  ' On NT, TotalEntries is the number of entries left to be read including
  ' the currently read entries.
  '
  ' On LanMan and OS/2, it is the total number of entries, period. Code
  ' would have to be changed to reflect this if the Domain controller
  ' wasn't an NT machine.
  '
  ' BufPtr gets the address of the buffer(or ptr1 - add 4 to BufPtr for
  ' each additional pointer)
  '
  Dim Result As Long, BufPtr As Long, EntriesRead As Long, _
  TotalEntries As Long, ResumeHandle As Long, BufLen As Long, _
  SNArray() As Byte, GNArray(99) As Byte, UNArray() As Byte, _
  GName As String, i As Integer, UNPtr As Long, _
  TempPtr As MungeLong, TempStr As MungeInt
  SNArray = SName & vbNullChar ' Move to byte array
  UNArray = UName & vbNullChar ' Move to Byte array
  BufLen = 255 ' Buffer size
  ResumeHandle = 0 ' Start with the first entry


  Do
    If UName = "" Then
      Result = NetGroupEnum0(SNArray(0), 0, BufPtr, BufLen, _
      EntriesRead, TotalEntries, ResumeHandle)
    Else
      Result = NetUserGetGroups0(SNArray(0), UNArray(0), 0, BufPtr, _
      BufLen, EntriesRead, TotalEntries)
    End If

    EnumerateGroups = Result

    If Result <> 0 And Result <> 234 Then ' 234 means multiple reads required
      Debug.Print "Error " & Result & " enumerating group " & EntriesRead & " of " & TotalEntries
      Exit Function
    End If


    For i = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4 byte block of memory in 2 steps
      '
      Result = PtrToInt(TempStr.XLo, BufPtr + (i - 1) * 4, 2)
      Result = PtrToInt(TempStr.XHi, BufPtr + (i - 1) * 4 + 2, 2)
      LSet TempPtr = TempStr ' munge 2 Integers to a Long
      ' Copy string to array and convert to a string
      Result = PtrToStr(GNArray(0), TempPtr.X)
      GName = Left(GNArray, StrLen(TempPtr.X))
      Debug.Print "Group: " & GName
    Next i
    Debug.Print "-------------------------------"
  Loop Until EntriesRead = TotalEntries
  ' The above condition only valid for reading accounts on NT
  ' but not OK for OS/2 or LanMan
  Result = NetApiBufferFree(BufPtr) ' Don't leak memory
End Function


Function EnumerateUsers(ByVal SName As String, ByVal GName As String, ByRef Users As String) As Long
  '
  ' If a group name is specified, it must be a global group
  ' and not a local group.
  '
  ' The buffer is filled from the left with pointers to user names that
  ' are filled from the right side. For example:
  '
  ' ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
  ' ^-------------- BufPtr buffer ----------------^
  '
  ' On Windows NT, TotalEntries is the number of entries left to be read,
  ' including the currently read entries.
  ' On LanMan and OS/2, it is the total number of entries, period. Code
  ' would have to be changed to reflect this if the Domain controller
  ' wasn't an NT machine.
  '
  ' BufPtr gets the address of the buffer(or ptr1 - add 4 to BufPtr for
  ' each additional pointer)
  '
  ' SName should be "\\servername"
  '
  Dim Result As Long, BufPtr As Long, EntriesRead As Long, _
  TotalEntries As Long, ResumeHandle As Long, BufLen As Long, _
  SNArray() As Byte, GNArray() As Byte, UNArray(99) As Byte, _
  UName As String, i As Integer, UNPtr As Long, TempPtr As MungeLong, _
  TempStr As MungeInt
  SNArray = SName & vbNullChar ' Move to byte array
  GNArray = GName & vbNullChar ' Move to Byte array
  BufLen = 255 ' Buffer size
  ResumeHandle = 0 ' Start with the first entry


  Do
    If GName = "" Then
      Result = NetUserEnum0(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, _
      BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
    Else
      Result = NetGroupEnumUsers0(SNArray(0), GNArray(0), 0, BufPtr, _
      BufLen, EntriesRead, TotalEntries, ResumeHandle)
    End If
    EnumerateUsers = Result
    If Result <> 0 And Result <> 234 Then ' 234 means multiple reads
      ' required
      Debug.Print "Error " & Result & " enumerating user " & EntriesRead & " of " & TotalEntries
      If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & GName & "'"
      Exit Function
    End If
  
    For i = 1 To EntriesRead
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      '
      Result = PtrToInt(TempStr.XLo, BufPtr + (i - 1) * 4, 2)
      Result = PtrToInt(TempStr.XHi, BufPtr + (i - 1) * 4 + 2, 2)
      LSet TempPtr = TempStr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(UNArray(0), TempPtr.X)
      UName = Left(UNArray, StrLen(TempPtr.X))
      Users = Users & UName & ","
      Debug.Print "User: " & UName
    Next i
  Loop Until EntriesRead = TotalEntries

  If Right(Users, 1) = "," Then
      Users = Mid(Users, 1, Len(Users) - 1)
  End If
  ' The above condition is only valid for reading accounts on Windows NT,
  ' but is not OK for OS/2 or LanMan
  Result = NetApiBufferFree(BufPtr) ' Don't leak memory
End Function


Private Function GetPrimaryDCName(ByVal MName As String, ByVal DName As String) As String
  Dim Result As Long, DCName As String, DCNPtr As Long
  Dim DNArray() As Byte, MNArray() As Byte, DCNArray(100) As Byte
  MNArray = MName & vbNullChar
  DNArray = DName & vbNullChar
  Result = NetGetDCName(MNArray(0), DNArray(0), DCNPtr)

  If Result <> 0 Then
    Debug.Print "Error: " & Result
    Exit Function
  End If
  Result = PtrToStr(DCNArray(0), DCNPtr)
  Result = NetApiBufferFree(DCNPtr)
  DCName = DCNArray()

  If InStr(DCName, Chr(0)) > 0 Then
    GetPrimaryDCName = Mid(DCName, 1, InStr(DCName, Chr(0)) - 1)
  Else
    GetPrimaryDCName = DCName
  End If
End Function


Function AddUser(ByVal SName As String, ByVal UName As String, _
  ByVal PWD As String) As Long
  Dim Result As Long, UNPtr As Long, PWDPtr As Long, ParmError As Long
  Dim SNArray() As Byte, UNArray() As Byte, PWDArray() As Byte
  Dim UserStruct As TUser1
  '
  ' Move to byte arrays
  '
  SNArray = SName & vbNullChar
  UNArray = UName & vbNullChar
  PWDArray = PWD & vbNullChar
  '
  ' Allocate buffer space
  '
  Result = NetAPIBufferAllocate(UBound(UNArray) + 1, UNPtr)
  Result = NetAPIBufferAllocate(UBound(PWDArray) + 1, PWDPtr)
  '
  ' Copy arrays to the buffer
  '
  Result = StrToPtr(UNPtr, UNArray(0))
  Result = StrToPtr(PWDPtr, PWDArray(0))
  '
  ' Fill the structure
  '
  With UserStruct
    .ptrName = UNPtr
    .ptrPassword = PWDPtr
    .dwPasswordAge = 3
    .dwPriv = USER_PRIV_USER
    .ptrHomeDir = 0
    .ptrComment = 0
    .dwFlags = UF_NORMAL_ACCOUNT Or UF_SCRIPT
    .ptrScriptPath = 0
  End With
  '
  ' Add the user
  '
  Result = NetUserAdd1(SNArray(0), 1, UserStruct, ParmError)
  AddUser = Result


  If Result <> 0 Then
    Debug.Print "Error " & Result & " in parameter " & ParmError _
    & " when adding user " & UName
  End If
  '
  ' Release buffers from memory
  '
  Result = NetApiBufferFree(UNPtr)
  Result = NetApiBufferFree(PWDPtr)
End Function

Public Function GetPDC(ByVal strDomain As String) As String
  'DomainName = strDomain
  PrimaryDomainController = GetPrimaryDCName("", strDomain)
  GetPDC = PrimaryDomainController
End Function

Public Function GetUserGroupList(ByVal strNTGroup As String) As String
  UserList = ""
  NTGroup = strNTGroup
  Call EnumerateUsers(PrimaryDomainController, NTGroup, UserList)
  GetUserGroupList = UserList
End Function

Public Function GetAllUsers() As String
  UserList = ""
  Call EnumerateUsers(PrimaryDomainController, "", UserList)
  GetAllUsers = UserList
End Function

Public Function UserInGroup(Optional ByVal strUserName As String) As String
  UserList = ""
  Call EnumerateUsers(PrimaryDomainController, NTGroup, UserList)

  If Right(UserList, 1) = "," Then
    UserList = Mid(UserList, 1, Len(UserList) - 1)
  End If
  UserInGroup = InStr(LCase(UserList), LCase(strUserName)) > 0
End Function

Public Function GetComputerName() As String
  Dim lReturn As Long, buffer(512) As Byte
  Dim i As Integer
  Dim wk101 As WKSTA_INFO_101
  Dim pwk101 As Long
  Dim sComputerName As String
    
  lReturn = NetWkstaGetInfo(ByVal 0&, 101, pwk101)
  RtlMoveMemory wk101, ByVal pwk101, Len(wk101)
  lstrcpyW buffer(0), wk101.wki101_computername
  i = 0
  Do While buffer(i) <> 0
    sComputerName = sComputerName & Chr(buffer(i))
    i = i + 2
  Loop
  lReturn = NetApiBufferFree(pwk101)
  
  GetComputerName = sComputerName
End Function

Public Function GetDomainName() As String
  Dim lReturn As Long, buffer(512) As Byte
  Dim i As Integer
  Dim wk1 As WKSTA_USER_INFO_1
  Dim pwk1 As Long
  Dim sLogonDomain As String
    
  lReturn = NetWkstaUserGetInfo(ByVal 0&, 1, pwk1)
  RtlMoveMemory wk1, ByVal pwk1, Len(wk1)
  lstrcpyW buffer(0), wk1.wkui1_logon_domain
  i = 0
  Do While buffer(i) <> 0
    sLogonDomain = sLogonDomain & Chr(buffer(i))
    i = i + 2
  Loop
  lReturn = NetApiBufferFree(pwk1)
  
  GetDomainName = sLogonDomain
  
End Function

