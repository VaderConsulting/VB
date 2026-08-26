Attribute VB_Name = "MOD_Netuser"
'The type declarations came from the Microsoft Knowledge Base
   Option Explicit
   Option Base 0     ' Important assumption for this code

   Type MungeLong
     X As Long
     Dummy As Integer
   End Type

   Type MungeInt
     XLo As Integer
     XHi As Integer
     Dummy As Integer
   End Type

   Type TUser0                    ' Level 0
     ptrName As Long
   End Type

   Type TUser1                    ' Level 1
     ptrName As Long
     ptrPassword As Long
     dwPasswordAge As Long
     dwPriv As Long
     ptrHomeDir As Long
     ptrComment As Long
     dwFlags As Long
     ptrScriptPath As Long
   End Type

'this is the only variable that I added
    Global GLBServer As String
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
   Const UF_NORMAL_ACCOUNT = &H200     ' Needs to be ORed with the
                                       ' other flags

   '
   ' for lFilter
   '
   Const FILTER_NORMAL_ACCOUNT = &H2

'All of the declares were copied from the microsoft knowledge base
   Declare Function NetGetDCName Lib "NETAPI32.DLL" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
   Declare Function NetUserDel Lib "NETAPI32.DLL" (ServerName As Byte, Username As Byte) As Long
   Declare Function NetGroupAddUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, Username As Byte) As Long
   Declare Function NetGroupDelUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, Username As Byte) As Long
   ' Add using Level 1 user structure
   Declare Function NetUserAdd1 Lib "NETAPI32.DLL" Alias "NetUserAdd" (ServerName As Byte, ByVal Level As Long, Buffer As TUser1, ParmError As Long) As Long
   ' Enumerate using Level 0 user structure
   Declare Function NetUserEnum0 Lib "NETAPI32.DLL" Alias "NetUserEnum" (ServerName As Byte, ByVal Level As Long, ByVal lFilter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
   Declare Function NetGroupEnumUsers0 Lib "NETAPI32.DLL" Alias "NetGroupGetUsers" (ServerName As Byte, GroupName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
   Declare Function NetGroupEnum0 Lib "NETAPI32.DLL" Alias "NetGroupEnum" (ServerName As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
   Declare Function NetUserGetGroups0 Lib "NETAPI32.DLL" Alias "NetUserGetGroups" (ServerName As Byte, Username As Byte, ByVal Level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long
   Declare Function NetAPIBufferFree Lib "NETAPI32.DLL" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
   Declare Function NetAPIBufferAllocate Lib "NETAPI32.DLL" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
   Declare Function PtrToStr Lib "Kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
   Declare Function StrToPtr Lib "Kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
   Declare Function PtrToInt Lib "Kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
   Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

   Function AddUserToGroup(ByVal SName As String, ByVal Gname As String, ByVal Uname As String) As Long
   '  this is microsofts
   ' This only adds users to global groups - not to local groups
   '
   Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, Result As Long
     SNArray = SName & vbNullChar
     GNArray = Gname & vbNullChar
     UNArray = Uname & vbNullChar
     Result = NetGroupAddUser(SNArray(0), GNArray(0), UNArray(0))
     If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & Gname & "'"
     AddUserToGroup = Result
   End Function

   Function DelUser(ByVal SName As String, ByVal Uname As String) As Long
   'this is microsofts
   Dim UNArray() As Byte, SNArray() As Byte
     UNArray = Uname & vbNullChar
     SNArray = SName & vbNullChar
     DelUser = NetUserDel(SNArray(0), UNArray(0))
   End Function

   Function DelUserFromGroup(ByVal SName As String, ByVal Gname As String, ByVal Uname As String) As Long
   ' again--microsofts
   ' This only deletes users from global groups - not local groups
   '
   Dim SNArray() As Byte, GNArray() As Byte, UNArray() As Byte, Result As Long
     SNArray = SName & vbNullChar
     GNArray = Gname & vbNullChar
     UNArray = Uname & vbNullChar
     Result = NetGroupDelUser(SNArray(0), GNArray(0), UNArray(0))
     If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & Gname & "'"
     DelUserFromGroup = Result
   End Function

   Function EnumerateGroups(ByVal SName As String, ByVal Uname As String, GroupArray() As String) As Long
   ' this is microsofts with the exception to pass the information obtained
   ' back in the form of an array rather than debug.print it
   ' Enumerates global groups only - not local groups
   '
   ' The buffer is filled from the left with pointers to user names that
   ' are filled from the right side. For example:
   '
   '     ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
   '     ^-------------- BufPtr buffer ----------------^
   '
   ' On NT, TotalEntries is the number of entries left to be read including
   ' the currently read entries.
   '
   ' On LanMan and OS/2, it is the total number of entries, period. Code
   ' would have to be changed to reflect this if the Domain controller
   ' wasn't an NT machine.
   '
   ' BufPtr gets the address of the buffer (or ptr1 - add 4 to BufPtr for
   ' each additional pointer)
   '
   Dim Result As Long, BufPtr As Long, EntriesRead As Long, _
   TotalEntries As Long, ResumeHandle As Long, BufLen As Long, _
   SNArray() As Byte, GNArray(99) As Byte, UNArray() As Byte, _
   Gname As String, i As Integer, UNPtr As Long, _
   TempPtr As MungeLong, TempStr As MungeInt
    Dim base As Long
     SNArray = SName & vbNullChar       ' Move to byte array
     UNArray = Uname & vbNullChar       ' Move to Byte array
     BufLen = 255                       ' Buffer size
     ResumeHandle = 0                   ' Start with the first entry
    base = 0
     Do
       If Uname = "" Then
         Result = NetGroupEnum0(SNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
       Else
         Result = NetUserGetGroups0(SNArray(0), UNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries)
       End If
       EnumerateGroups = Result
       If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                                ' required
         Debug.Print "Error " & Result & " enumerating group " & EntriesRead & " of " & TotalEntries
         EnumerateGroups = 123456
         Exit Function
       End If
       For i = 1 To EntriesRead
         ' Get pointer to string from beginning of buffer
         ' Copy 4 byte block of memory in 2 steps
         Result = PtrToInt(TempStr.XLo, BufPtr + (i - 1) * 4, 2)
         Result = PtrToInt(TempStr.XHi, BufPtr + (i - 1) * 4 + 2, 2)
         LSet TempPtr = TempStr ' munge 2 Integers to a Long
         ' Copy string to array and convert to a string
         Result = PtrToStr(GNArray(0), TempPtr.X)
         Gname = Left(GNArray, StrLen(TempPtr.X))
         base = base + 1
         ReDim Preserve GroupArray(base)
         GroupArray(base) = Gname
       Next i
     Loop Until EntriesRead = TotalEntries

   ' The above condition only valid for reading accounts on NT
   ' but not OK for OS/2 or LanMan

     Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory

   End Function

   Function EnumerateUsers(ByVal SName As String, ByVal Gname As String, UserArray() As String) As Long
   '  microsofts with the same note as above (about the array)
   ' If a group name is specified, it must be a global group
   ' and not a local group.
   '
   ' The buffer is filled from the left with pointers to user names that
   ' are filled from the right side. For example:
   '
   '     ptr1|ptr2|...|ptrn|<garbage>|strn|...|str2|str1
   '     ^-------------- BufPtr buffer ----------------^
   '
   ' On Windows NT, TotalEntries is the number of entries left to be read,
   ' including the currently read entries.
   ' On LanMan and OS/2, it is the total number of entries, period.  Code
   ' would have to be changed to reflect this if the Domain controller
   ' wasn't an NT machine.
   '
   ' BufPtr gets the address of the buffer (or ptr1 - add 4 to BufPtr for
   ' each additional pointer)
   '
   ' SName should be "\\servername"
   '
   Dim Result As Long
   Dim BufPtr As Long
   Dim EntriesRead As Long
   Dim TotalEntries As Long
   Dim ResumeHandle As Long
   Dim BufLen As Long
   Dim SNArray() As Byte
   Dim GNArray() As Byte
   Dim UNArray(99) As Byte
   Dim Uname As String
   Dim i As Integer
   Dim UNPtr As Long
   Dim TempPtr As MungeLong
   Dim TempStr As MungeInt
    Dim base As Long
     SNArray = SName & vbNullChar       ' Move to byte array
     GNArray = Gname & vbNullChar       ' Move to Byte array
     BufLen = 255                       ' Buffer size
     ResumeHandle = 0                   ' Start with the first entry
    base = 0
     Do
       If Gname = "" Then
         Result = NetUserEnum0(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
       Else
         Result = NetGroupEnumUsers0(SNArray(0), GNArray(0), 0, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
       End If
       EnumerateUsers = Result
       If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                                ' required
         Debug.Print "Error " & Result & " enumerating user " & EntriesRead & " of " & TotalEntries
         If Result = 2220 Then Debug.Print "There is no **GLOBAL** group '" & Gname & "'"
         EnumerateUsers = 123456
         Exit Function
       End If
       For i = 1 To EntriesRead
         ' Get pointer to string from beginning of buffer
         ' Copy 4-byte block of memory in 2 steps
         Result = PtrToInt(TempStr.XLo, BufPtr + (i - 1) * 4, 2)
         Result = PtrToInt(TempStr.XHi, BufPtr + (i - 1) * 4 + 2, 2)
         LSet TempPtr = TempStr ' munge 2 integers into a Long
         ' Copy string to array
         Result = PtrToStr(UNArray(0), TempPtr.X)
         Uname = Left(UNArray, StrLen(TempPtr.X))
         base = base + 1
         ReDim Preserve UserArray(base)
         UserArray(base) = Uname
       Next i
     Loop Until EntriesRead = TotalEntries
   ' The above condition is only valid for reading accounts on Windows NT,
   ' but is not OK for OS/2 or LanMan

     Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory

   End Function

   Function GetPrimaryDCName(ByVal MName As String, ByVal DName As String) As String
   ' microsofts--if your starting--DC stands for Domain Controler
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
     Result = NetAPIBufferFree(DCNPtr)
     DCName = DCNArray()
     GetPrimaryDCName = DCName
   End Function

   Function AddUser(ByVal SName As String, ByVal Uname As String, ByVal PWD As String) As Long
   Dim Result As Long, UNPtr As Long, PWDPtr As Long, ParmError As Long
   Dim SNArray() As Byte, UNArray() As Byte, PWDArray() As Byte
' microsoft's
   Dim UserStruct As TUser1
   '
   ' Move to byte arrays
   '
     SNArray = SName & vbNullChar
     UNArray = Uname & vbNullChar
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
       Debug.Print "Error " & Result & " in parameter " & ParmError & " when adding user " & Uname
     End If
   '
   ' Release buffers from memory
   '
     Result = NetAPIBufferFree(UNPtr)
     Result = NetAPIBufferFree(PWDPtr)

   End Function

Function IsMember(Uname As String, GNames() As String) As Boolean
    ' this is mine all mine!  (Boy, I sure contributed a lot to this project! :)
    ' this function will return a boolean value stating indicating whether a user
    ' is a member of a particular group or not.
    ' Robert May 1-30-98
    On Error Resume Next
    Dim res
    Dim found As Boolean
    Dim tempgroups() As String
    Dim i As Integer
    Dim a As Integer
    found = False
    If EnumerateGroups(GLBServer, Uname, tempgroups()) <> 123456 Then 'I set the return value to 123456 so I can tell if nothing came back
        For a = 0 To UBound(GNames)
            For i = 1 To UBound(tempgroups)
                If UCase$(tempgroups(i)) = UCase$(GNames(a)) Then
                    found = True
                End If
            Next i
        Next a
    End If
    IsMember = found
        
End Function
Sub POPLST(DeArray() As String, LSTBOX As Object, DGroups() As String)
    ' This procedure will populate a list box using the array's obtained from
    ' the enumerate functions
    ' Robert May 1-30-98
    On Error Resume Next
    Dim drop As Boolean
    Dim i As Integer
    Dim a As Integer
    For i = 1 To UBound(DeArray)
        drop = False
        For a = 0 To UBound(DGroups)
            If DGroups(a) = DeArray(i) Then drop = True
        Next a
        If drop <> True Then LSTBOX.AddItem DeArray(i)
    Next i
End Sub
Sub ShowUsers(Gname As String, CurUser As String, LSTBOX As Object)
    ' this prodcedure get's the users based on the group name (gname) passed
    ' to it and then populates the list box.  Also checks privledges to see
    ' just how much info is returned
    ' Robert May 1-30-98
    Dim drops() As String
    Dim tempusers() As String
    Dim i As Integer
    Dim res As Boolean
    Dim AdminGs() As String
    LSTBOX.Clear
    EnumerateUsers GLBServer, Gname, tempusers
    ReDim Preserve AdminGs(0)
    AdminGs(0) = "Domain Admins"
    res = IsMember(CurUser, AdminGs)  'if your a member of the admins group, you get to see all of the groups, not some.
    ReDim Preserve AdminGs(1)
    AdminGs(0) = "Domain Admins"   ' these are the groups we don't want to show.
    AdminGs(1) = "Management"
    If res = False Then     ' if users are Admins, and our user dosen't have admin privledges, we won't let him to anything to our accounts (we don't show up!)
        For i = 0 To UBound(tempusers)
            If IsMember(tempusers(i), AdminGs) Then
                ReDim Preserve drops(i)
                drops(i) = tempusers(i)
            End If
        Next i
    Else
        ReDim drops(0)
        drops(0) = "none"
    End If
    POPLST tempusers, LSTBOX, drops  ' populate list boxes

End Sub
Sub ShowGroups(Uname As String, CurUser As String, LSTBOX As Object)
    ' basically the same as above, only shows the groups instead of users
    ' Robert May 1-30-98
    Dim drops() As String
    Dim tempgroups() As String
    Dim i As Integer
    Dim res As Boolean
    Dim AdminGs() As String
    LSTBOX.Clear
    EnumerateGroups GLBServer, Uname, tempgroups
    ReDim Preserve AdminGs(0)
    AdminGs(0) = "Domain Admins"
    res = IsMember(CurUser, AdminGs)
    If res = False Then
        ReDim drops(4)
        drops(0) = "Domain Admins"
        drops(1) = "Management"
        drops(2) = "RTC"
        drops(3) = "Domain Guests"
        drops(4) = "Domain Users"
    Else
        ReDim drops(0)
        drops(0) = "none"
    End If
    POPLST tempgroups, LSTBOX, drops

End Sub
