Attribute VB_Name = "modNT"
'Name:          NetGroup.Bas
'Description:   Network Group Account Functions for Windows NT
'Dependencies:  NetDecs.Bas
'               NetErr.BAS
'Included
'Functions:     NetAddGroup
'               NetDelGroup
'               NetAddUserToGroup
'               NetDelUserFromGroup
'               NetListUsersPerGroupArray
'               NetListUsersPerGroupFile
'               NetListGroupsPerUserArray
'------------------------------------------------------------------------

    Public Type MungeLong
      X As Long
      Dummy As Integer
    End Type
    
    Public Type MungeInt
      XLo As Integer
      XHi As Integer
      Dummy As Integer
    End Type



Public Sub NetAddGroup(ByVal ServerName$, ByVal GroupName$, ByVal Comment$)
'THIS FUNCTION ADDS A GLOBAL GROUP TO THE NT DOMAIN
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayGroupName() As Byte, arrayComment() As Byte
Dim GroupStructLevel1 As Level_1_Group_Structure_Type

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(GroupName$, arrayGroupName)
    Call ConvStringToByteArray(Comment$, arrayComment)
    
    Call ConvArrayToPointer(arrayGroupName, ptrGroupName&)
    Call ConvArrayToPointer(arrayComment, ptrComment&)
    
    With GroupStructLevel1
        .ptrName = ptrGroupName&
        .ptrComment = ptrComment&
    End With 'GroupStructLevel1

    Const GROUP_LEVEL_1 = 1
    Reply& = NetGroupAdd1(arrayServerName(0), GROUP_LEVEL_1, GroupStructLevel1, ParmError&)
    Call NetReturnError(Reply&, GroupName$)
    
End Sub

Public Sub NetDelGroup(ByVal ServerName$, ByVal GroupName$)
'THIS FUNCTION DELETES A GLOBAL GROUP
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayGroupName() As Byte

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(GroupName$, arrayGroupName)
    
    Reply& = NetGroupDel(arrayServerName(0), arrayGroupName(0))
    Call NetReturnError(Reply&, GroupName$)

End Sub

Public Sub NetAddUserToGroup(ByVal ServerName$, ByVal GroupName$, ByVal UserName$)
'THIS FUNCTION ADDS A USER TO A GLOBAL GROUP
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayGroupName() As Byte, arrayUserName() As Byte

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(GroupName$, arrayGroupName)
    Call ConvStringToByteArray(UserName$, arrayUserName)
  
    Result& = NetGroupAddUser(arrayServerName(0), arrayGroupName(0), arrayUserName(0))
    Call NetReturnError(Reply&, UserName$)

End Sub



Public Sub NetDelUserFromGroup(ByVal ServerName$, ByVal GroupName$, ByVal UserName$)
'THIS FUNCTION DELETES A USER FROM A GROUP
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayGroupName() As Byte, arrayUserName() As Byte

    TripError = False
    
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(GroupName$, arrayGroupName)
    Call ConvStringToByteArray(UserName$, arrayUserName)
  
    Reply& = NetGroupDelUser(arrayServerName(0), arrayGroupName(0), arrayUserName(0))
    Call NetReturnError(Reply&, UserName$)

End Sub

'Public Function NetListUsersPerGroupArray(ByVal ServerName$, ByVal GroupName$) As stringarray
'THIS FUNCTION TAKES A SERVER AND A GROUP AND LIST IT'S USERS
'THIS FUNCTION RETURNS AN ARRAY OF STRINGS AS "NetListUsersPerGroupArray.List(x)"
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

'Dim arrayServerName() As Byte, arrayGroupName() As Byte, arrayUserName(CAP) As Byte
'Dim TempPtr As MungeLong, TempStr As MungeInt

'Const FILTER_NORMAL_ACCOUNT = &H2
'    TripError = False
'
'    Call ConvStringToByteArray(ServerName$, arrayServerName)
'    Call ConvStringToByteArray(GroupName$, arrayGroupName)
'
'    BufferLength& = 255                       ' Buffer size
'    ResumeHandle& = 0                   ' Start with the first entry
'
'    xIndex% = 1
'    Do
'      If GroupName$ = "" Then
'        Reply& = NetUserEnum0(arrayServerName(0), 0, FILTER_NORMAL_ACCOUNT, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&, ResumeHandle&)
'      Else
'        Reply& = NetGroupEnumUsers0(arrayServerName(0), arrayGroupName(0), 0, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&, ResumeHandle&)
'      End If
'      EnumerateUsers = Reply&
'      If Reply& <> 0 And Reply& <> 234 Then    ' 234 means multiple reads required
'        Call NetReturnError(Reply&, GroupName$)
'        Exit Function
'      End If
'      For counter% = 1 To EntriesRead&
'        ' Get pointer to string from beginning of buffer
'        Reply& = PtrToInt(TempStr.XLo, ptrBuffer& + (counter% - 1) * 4, 2)      ' Doing this to copy a 4-byte block of memory to a Long
'        Reply& = PtrToInt(TempStr.XHi, ptrBuffer& + (counter% - 1) * 4 + 2, 2)
'        LSet TempPtr = TempStr
'        ' Copy string to array
'        Reply& = PtrToStr(arrayUserName(0), TempPtr.X)
'        UserName$ = Left(arrayUserName, StrLen(TempPtr.X))
'        NetListUsersPerGroupArray.List(xIndex% - 1) = UserName$
'        xIndex% = xIndex% + 1
'      Next
'    Loop Until EntriesRead& = TotalEntries&     ' This condition only valid for reading accounts on NT - but not OK for OS/2 or LanMan
'
'    Call ConvPointerToNothing(ptrBuffer&)
'End Function

Public Sub NetListUsersPerGroupFile(ByVal ServerName$, ByVal GroupName$)
'THIS FUNCTION TAKES A SERVER AND A GROUP AND LIST IT'S USERS
'THIS FUNCTION WRITES TO A RANDOM ACCESS FILE (FIELD LEN = 45) C:\USERS.TXT
'YOU MUST USE:
' Open "C:\users.txt" for random access Read as #1 Len=45
'TO READ THE RECORDS.
'THE FIRST RECORD, BY THE WAY, IS THE NUMBER OF RECORDS TO READ.
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

Dim arrayServerName() As Byte, arrayGroupName() As Byte, arrayUserName(99) As Byte
Dim TempPtr As MungeLong, TempStr As MungeInt
Dim BufferLength&, ResumeHandle&, xIndex%, ptrBuffer&, EntriesRead&, TotalEntries&, Reply&, counter%, u$
Dim EnumerateUsers As Long
Dim UserName As String * 45

TripError = False

Const FILTER_NORMAL_ACCOUNT = &H2
   
    Call ConvStringToByteArray(ServerName$, arrayServerName)
    Call ConvStringToByteArray(GroupName$, arrayGroupName)
  
    BufferLength& = 255                       ' Buffer size
    ResumeHandle& = 0                   ' Start with the first entry

     On Error Resume Next
     Kill "C:\USERS.TXT"
     On Error GoTo 0
     Open "C:\USERS.TXT" For Random Access Write As #1 Len = 45
     xIndex% = 0
   Do
      If GroupName$ = "" Then
        Reply& = NetUserEnum0(arrayServerName(0), 0, FILTER_NORMAL_ACCOUNT, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&, ResumeHandle&)
      Else
        Reply& = NetGroupEnumUsers0(arrayServerName(0), arrayGroupName(0), 0, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&, ResumeHandle&)
      End If
      EnumerateUsers = Reply&
      If Reply& <> 0 And Reply& <> 234 Then    ' 234 means multiple reads required
        Call NetReturnError(Reply&, GroupName$)
        Exit Sub
      End If
      For counter% = 1 To EntriesRead&
        ' Get pointer to string from beginning of buffer
        Reply& = PtrToInt(TempStr.XLo, ptrBuffer& + (counter% - 1) * 4, 2)      ' Doing this to copy a 4-byte block of memory to a Long
        Reply& = PtrToInt(TempStr.XHi, ptrBuffer& + (counter% - 1) * 4 + 2, 2)
        LSet TempPtr = TempStr
        ' Copy string to array
        Reply& = PtrToArray(arrayUserName(0), TempPtr.X)
        u$ = Trim(Left(arrayUserName, StrLen(TempPtr.X)))
        UserName = Left$(u$, 45)
        If InStr(UserName, "$") Then  'This is a computer name or a process, not to be considered a user!
            'Do Nothing
        Else
            If Trim(UserName) <> "" Then
                xIndex% = xIndex% + 1
                Put #1, xIndex%, UserName
            End If
        End If
      Next
    Loop Until EntriesRead& = TotalEntries&     ' This condition only valid for reading accounts on NT - but not OK for OS/2 or LanMan
    Call ConvPointerToNothing(ptrBuffer&)
    Close #1
    
Dim b As String * 45
Dim a As String * 45
Open "C:\USERS.TXT" For Random As #1 Len = 45
Get #1, 1, a
Put #1, xIndex% + 1, a
b = Str(xIndex% + 1)
Put #1, 1, b
Close #1

End Sub

'Public Function NetListGroupsPerUserArray(ByVal ServerName$, ByVal UserName$) As stringarray
'THIS FUNCTION TAKES A SERVER AND A USER AND LIST IT'S GROUPS
'THIS FUNCTION RETURNS AN ARRAY OF STRINGS AS "NetListGroupsPerUserArray.List(x)"
'SERVERNAME MUST BE UNC FORMAT AS IN \\PDC AND MUST BE PDC'S COMPUTER NAME.

'Dim arrayServerName() As Byte, arrayGroupName(99) As Byte, arrayUserName() As Byte
'Dim TempPtr As MungeLong, TempStr As MungeInt

'    Call ConvStringToByteArray(ServerName$, arrayServerName)
'    Call ConvStringToByteArray(UserName$, arrayUserName)

'    TripError = False
'
'    BufferLength& = 255  'Buffer Size
'    ResumeHandle& = 0  ' Start with the first entry
'
'    xIndex% = 1
'    Do
'        If UserName$ = "" Then
'          Reply& = NetGroupEnum0(arrayServerName(0), 0, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&, ResumeHandle&)
'        Else
'          Reply& = NetUserGetGroups0(arrayServerName(0), arrayUserName(0), 0, ptrBuffer&, BufferLength&, EntriesRead&, TotalEntries&)
'        End If
'
'        If Reply& <> 0 And Reply& <> 234 Then    ' 234 means multiple reads required
'            Call NetReturnError(Reply&, UserName$)
'            Exit Function
'        End If
'
'        For counter% = 1 To EntriesRead&
'          ' Get pointer to string from beginning of buffer
'          Reply& = PtrToInt(TempStr.XLo, ptrBuffer& + (counter% - 1) * 4, 2)      ' Doing this to copy a 4-byte block of memory to a Long
'          Reply& = PtrToInt(TempStr.XHi, ptrBuffer& + (counter% - 1) * 4 + 2, 2)
'          LSet TempPtr = TempStr
'          ' Copy string to array
'
'          Reply& = PtrToStr(arrayGroupName(0), TempPtr.X)
'          GroupName$ = Left(arrayGroupName, StrLen(TempPtr.X))
'          If xIndex% > CAP Then
'            EntriesRead& = TotalEntries&
'            Exit For
'        End If
'          If xIndex% = 1 Then FirstName$ = GroupName$
'          If xIndex% <> 1 And GroupName$ = FirstName$ Then
'            EntriesRead& = TotalEntries&
'            Exit For
'          End If
'          NetListGroupsPerUserArray.List(xIndex% - 1) = GroupName$
'        xIndex% = xIndex% + 1
'        Next
'    Loop Until EntriesRead& = TotalEntries&
'
'    Call ConvPointerToNothing(ptrBuffer&)

'End Function








