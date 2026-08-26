Attribute VB_Name = "NTAPI"
'All variables must be declared
Option Explicit
Option Base 0     ' Important assumption for this code
                  ' Arrays starts at 0
 
'****************************************
' Private Type Declarations
'****************************************
Private Type MungeLong
  x As Long
  Dummy As Integer
End Type
     
Private Type MungeInt
  XLo As Integer
  XHi As Integer
  Dummy As Integer
End Type

Private Type TUser1006
  ptrHomeDir As Long
End Type

'****************************************
' Constant declarations
'****************************************

Public Const UF_SCRIPT = &H1
Public Const UF_ACCOUNTDISABLE = &H2
Public Const UF_HOMEDIR_REQUIRED = &H8
Public Const UF_LOCKOUT = &H10
Public Const UF_PASSWD_NOTREQD = &H20
Public Const UF_PASSWD_CANT_CHANGE = &H40
        
Public Const UF_TEMP_DUPLICATE_ACCOUNT = &H100
Public Const UF_NORMAL_ACCOUNT = &H200
Public Const UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
Public Const UF_WORKSTATION_TRUST_ACCOUNT = &H1000
Public Const UF_SERVER_TRUST_ACCOUNT = &H2000
        
Public Const UF_DONT_EXPIRE_PASSWD = &H10000
Public Const UF_MNS_LOGON_ACCOUNT = &H20000

Public Const AF_OP_PRINT = 1
Public Const AF_OP_COMM = 2
Public Const AF_OP_SERVER = 4
Public Const AF_OP_ACCOUNTS = 8

Public Const DateFormat As String = "dd/mm/yyyy hh:nn:ss"

'****************************************
' API function declarations
'****************************************
Private Declare Function NetShareDel Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByRef netname As Byte, _
   ByVal reserved As Long) As Long

Private Declare Function NetUserEnum Lib "netapi32.dll" _
 _
  (ByRef servername As Byte, _
   ByVal level As Long, _
   ByVal lFilter As Long, _
   ByRef buffer As Long, _
   ByVal prefmaxlen As Long, _
   ByRef entriesread As Long, _
   ByRef totalentries As Long, _
   ByRef ResumeHandle As Long) As Long

Private Declare Function NetGroupEnumUsers Lib "netapi32.dll" Alias _
     "NetGroupGetUsers" _
  (ByRef servername As Byte, _
   ByRef GroupName As Byte, _
   ByVal level As Long, _
   ByRef buffer As Long, _
   ByVal prefmaxlen As Long, _
   ByRef entriesread As Long, _
   ByRef totalentries As Long, _
   ByRef ResumeHandle As Long) As Long
   
Private Declare Function NetUserGetGroups Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByRef username As Byte, _
   ByVal level As Long, _
   ByRef buffer As Long, _
   ByVal prefmaxlen As Long, _
   ByRef entriesread As Long, _
   ByRef totalentries As Long) As Long
   
Private Declare Function NetQueryDisplayInformation Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByVal level As Long, _
   ByVal Index As Long, _
   ByVal EntriesRequested As Long, _
   ByVal PreferredMaximumLength As Long, _
   ByRef ReturnedEntryCount As Long, _
   ByRef SortedBuffer As Long) As Long
   
Private Declare Function NetUserGetInfo Lib "NETAPI32" _
  (ByRef servername As Byte, _
   ByRef username As Byte, _
   ByVal level As Long, _
   ByRef buffer As Long) As Long
   
Private Declare Function NetUserSetInfo Lib "NETAPI32" _
  (ByRef servername As Byte, _
   ByRef username As Byte, _
   ByVal level As Long, _
   ByRef buffer As TUser1006, _
   ByRef parm_err As Long) As Long
   
Private Declare Function NetShareGetInfo Lib "NETAPI32" _
  (ByRef servername As Byte, _
   ByRef netname As Byte, _
   ByVal level As Long, _
   ByRef buffer As Long) As Long
   
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias _
     "NetApiBufferFree" (ByVal Ptr As Long) As Long
      
Private Declare Function NetAPIBufferAllocate Lib "netapi32.dll" Alias _
     "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long
      
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" _
  (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
    
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" _
     (RetVal As Byte, ByVal Ptr As Long) As Long
    
Private Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" _
     (ByVal Ptr As Long, Source As Byte) As Long
    
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" _
     (ByVal Ptr As Long) As Long
  
Private Declare Function GetDiskFreeSpace Lib "kernel32" _
 Alias "GetDiskFreeSpaceA" _
 (ByVal lpRootPathName As String, _
 lpSectorsPerCluster As Long, _
 lpBytesPerSector As Long, _
 lpNumberOfFreeClusters As Long, _
 lpTotalNumberOfClusters As Long) As Long
  
Private Declare Function NetGetDCName Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByRef DomainName As Byte, _
   ByRef buffer As Long) As Long
  
Private Declare Function NetGroupAddUser Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByRef GroupName As Byte, _
   ByRef username As Byte) As Long
   
Private Declare Function NetGroupDelUser Lib "netapi32.dll" _
  (ByRef servername As Byte, _
   ByRef GroupName As Byte, _
   ByRef username As Byte) As Long

'****************************************
' Public Methods
'****************************************

Public Function GetPDC(Server As String, domain As String, PDC As String) As Long
  Dim Result As Long
  Dim SNArray() As Byte
  Dim DArray() As Byte
  Dim DCNPtr As Long
  Dim STRArray(100) As Byte
    
  SNArray = Server & vbNullChar      ' Move to byte array
  DArray = domain & vbNullChar       ' Move to byte array
           
  Result = NetGetDCName(SNArray(0), _
     DArray(0), DCNPtr)
  
  GetPDC = Result
    
  If Result = 0 Then
    Result = PtrToStr(STRArray(0), DCNPtr)
    PDC = Left(STRArray(), StrLen(DCNPtr))
  Else
    PDC = ""
  End If
  NetAPIBufferFree (DCNPtr)
End Function

Public Function AddUser(domain As String, Group As String, User As String) As Boolean
  Dim Result As Long
  Dim PDC As String
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim UNArray() As Byte
  Dim Ok As Boolean
  
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  SNArray = PDC & vbNullChar         ' Move to byte array
  GNArray = Group & vbNullChar       ' Move to byte array
  UNArray = User & vbNullChar        ' Move to byte array
    
  Result = NetGroupAddUser(SNArray(0), GNArray(0), UNArray(0))
  If Result <> 0 Then GoTo HandleError
    
  AddUser = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  AddUser = False
  GoTo ExitHere

End Function

Public Function DelUser(domain As String, Group As String, User As String) As Boolean
  Dim Result As Long
  Dim PDC As String
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim UNArray() As Byte
  Dim Ok As Boolean
  Dim Stamp As String
  
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  SNArray = PDC & vbNullChar         ' Move to byte array
  GNArray = Group & vbNullChar         ' Move to byte array
  UNArray = User & vbNullChar         ' Move to byte array
  
  Result = NetGroupDelUser(SNArray(0), GNArray(0), UNArray(0))
  If Result <> 0 Then GoTo HandleError
  
  DelUser = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  DelUser = False
  GoTo ExitHere
RuntimeError:
  Resume HandleError
End Function


Function QuickEnumerate(domain As String, level As Long, _
      Size As Long, Data() As String) As Boolean
' To enumerate Users specify Level = 1
' To enumerate Machines specify Level = 2 (not implemented)
' To enumerate Groups specify Level = 3

' Returns an array holding strings og usernames ogr groupnames
' Groupnames are returned as a string where the first 20 letters are groupname,
' and the last letters is a groupdescription

  On Error GoTo RuntimeError
  Dim APIResult As Long
  Dim Result As Long
  Dim PDC As String
  Dim SNArray() As Byte
  Dim EntriesRequested As Long
  Dim PreferredMaximumLength As Long
  Dim ReturnedEntryCount As Long
  Dim SortedBuffer As Long
  Dim TempPtr As MungeLong
  Dim tempstr As MungeInt
  Dim STRArray(500) As Byte
  Dim i As Integer
  Dim Index As Long
  Dim NextIndex As Long
  Dim MoreData As Boolean
  'Dim InternalData() As String
  Dim ArrayRoom As String
  Dim name As String
  Dim comment As String
       
  ReDim Data(1 To 8000)
  ArrayRoom = 8000 'If arraydefinition changes size, also change this number
  If Not (level = 1 Or level = 3) Then GoTo HandleError
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  Size = 0
  SNArray = PDC & vbNullChar         ' Move to byte array
  Index = 0                          ' Start with the first entry
  EntriesRequested = 500
  PreferredMaximumLength = 6000      ' Buffer size:
                                     ' Specifies the size of the chunks
                                     ' used to transfer user data from the
                                     ' NT directory database
                                     ' Try experimenting
  Do
    APIResult = NetQueryDisplayInformation(SNArray(0), level, Index, _
         EntriesRequested, PreferredMaximumLength, ReturnedEntryCount, _
         SortedBuffer)
    
    If APIResult <> 0 And APIResult <> 234 Then    ' 234 means multiple reads
                                                   ' required
      GoTo HandleError
    End If
         
    For i = 1 To ReturnedEntryCount
      Size = Size + 1
      If Size > ArrayRoom Then
        ArrayRoom = ArrayRoom + 2000
        ReDim Preserve Data(1 To ArrayRoom)
      End If
      Select Case level
        Case Is = 1
         
          ' Get pointer to string from beginning of buffer
          ' Copy 4-byte block of memory in 2 steps
          Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 24, 2)
          Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 24 + 2, 2)
          LSet TempPtr = tempstr ' munge 2 integers into a Long
          ' Copy string to array
          Result = PtrToStr(STRArray(0), TempPtr.x)
          Data(Size) = Left(STRArray, StrLen(TempPtr.x))
          
          ' Get pointer to string from beginning of buffer
          ' Copy 4-byte block of memory in 2 steps
          Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 24 + 20, 2)
          Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 24 + 22, 2)
          LSet TempPtr = tempstr ' munge 2 integers into a Long
          NextIndex = TempPtr.x
                    
        Case Is = 3
          ' Get pointer to string from beginning of buffer
          ' Copy 4-byte block of memory in 2 steps
          Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20, 2)
          Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 2, 2)
          LSet TempPtr = tempstr ' munge 2 integers into a Long
          ' Copy string to array
          Result = PtrToStr(STRArray(0), TempPtr.x)
          name = Left(STRArray, StrLen(TempPtr.x))
          
          ' Get pointer to string from beginning of buffer
          ' Copy 4-byte block of memory in 2 steps
          Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 4, 2)
          Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 6, 2)
          LSet TempPtr = tempstr ' munge 2 integers into a Long
          ' Copy string to array
          Result = PtrToStr(STRArray(0), TempPtr.x)
          comment = Left(STRArray, StrLen(TempPtr.x))
                    
          Data(Size) = "1234567890123456789012"
          LSet Data(Size) = name
          Data(Size) = Data(Size) & comment
                    
          ' Get pointer to string from beginning of buffer
          ' Copy 4-byte block of memory in 2 steps
          Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 16, 2)
          Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 18, 2)
          LSet TempPtr = tempstr ' munge 2 integers into a Long
          NextIndex = TempPtr.x
                    
      End Select
    Next i
    Result = NetAPIBufferFree(SortedBuffer)         ' Don't leak memory
    'If ReturnedEntryCount = 0 Then GoTo Handleerror ' In the odd event that there are no entries
    Index = NextIndex
  Loop Until APIResult = 0          'No more data
  
 ' ReDim Data(1 To Size)
 ' For i = 1 To Size           ' I am not sure whether this is a smart move, but it has
                              ' been made to insure that only real date is passed
                              ' accros the DCOM connection, and not empty array fields
  '  Data(i) = InternalData(i)
  'Next i
  
  If Size > 0 Then
    ReDim Preserve Data(1 To Size)  ' I am not sure whether this is a smart move, but it has
  Else                              ' been made to insure that only real date is passed
    ReDim Preserve Data(1 To 1)     ' accros the DCOM connection, and not empty array fields
  End If
  
  QuickEnumerate = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  QuickEnumerate = False
  GoTo ExitHere
RuntimeError:
  Resume HandleError
End Function

Function EnumerateGroupMembers(domain As String, Group As String, _
     Size As Long, Data() As String) As Boolean
  
  On Error GoTo RuntimeError

  Dim Ok As Boolean
  Dim Result As Long
  Dim bufptr As Long
  Dim entriesread As Long
  Dim totalentries As Long
  Dim ResumeHandle As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim GNArray() As Byte
  Dim STRArray(99) As Byte
  Dim PDC As String
  Dim i As Integer
  Dim TempPtr As MungeLong
  Dim tempstr As MungeInt
  'Dim InternalData() As String
  Dim ArrayRoom As Long
    
  ReDim Data(1 To 5000)
  ArrayRoom = 5000 'If arraydefinition changes size, also change this number
    
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  
  SNArray = PDC & vbNullChar         ' Move to byte array
  GNArray = Group & vbNullChar       ' Move to Byte array
  
  BufLen = 1000                      ' Buffer size:
                                     ' Specifies the size of the chunks
                                     ' used to transfer user data from the
                                     ' NT directory database
                                     ' Try experimenting
  ResumeHandle = 0                   ' Start with the first entry
  Size = 0
  Do
    Result = NetGroupEnumUsers(SNArray(0), GNArray(0), 0, bufptr, _
      BufLen, entriesread, totalentries, ResumeHandle)
    
    
    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads
                                             ' required
      GoTo HandleError
    End If
      
    For i = 1 To entriesread
      Size = Size + 1
      If Size > ArrayRoom Then
        ArrayRoom = ArrayRoom + 1000
        ReDim Preserve Data(1 To ArrayRoom)
      End If
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + (i - 1) * 4, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + (i - 1) * 4 + 2, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      Data(Size) = Left(STRArray, StrLen(TempPtr.x))
    Next i
        
    Result = NetAPIBufferFree(bufptr)         ' Don't leak memory
           
  Loop Until entriesread = totalentries
  
  'ReDim Data(1 To Size)
  'For i = 1 To Size          ' I am not sure whether this is a smart move, but it has
                              ' been made to insure that only real date is passed
                              ' accros the DCOM connection
  '  Data(i) = InternalData(i)
  'Next i
  If Size > 0 Then
    ReDim Preserve Data(1 To Size)  ' I am not sure whether this is a smart move, but it has
  Else                              ' been made to insure that only real date is passed
    ReDim Preserve Data(1 To 1)     ' accros the DCOM connection
  End If
  
  EnumerateGroupMembers = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  EnumerateGroupMembers = False
  GoTo ExitHere
RuntimeError:
  Resume HandleError
End Function

Function EnumerateUsersGroupMembships(domain As String, User As String, _
     Size As Long, Data() As String) As Boolean
  
  On Error GoTo RuntimeError

  Dim Ok As Boolean
  Dim Result As Long
  Dim bufptr As Long
  Dim entriesread As Long
  Dim totalentries As Long
  Dim BufLen As Long
  Dim SNArray() As Byte
  Dim UNArray() As Byte
  Dim STRArray(99) As Byte
  Dim PDC As String
  Dim i As Integer
  Dim TempPtr As MungeLong
  Dim tempstr As MungeInt
  'Dim InternalData() As String
  Dim ArrayRoom As Long
    
  ReDim Data(1 To 500)
  ArrayRoom = 500 'If arraydefinition changes size, also change this number
    
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  
  SNArray = PDC & vbNullChar         ' Move to byte array
  UNArray = User & vbNullChar       ' Move to Byte array
  
  BufLen = 1000                      ' Buffer size:
                                     ' Specifies the size of the chunks
                                     ' used to transfer user data from the
                                     ' NT directory database
                                     ' Try experimenting
  Size = 0
  'do
    Result = NetUserGetGroups(SNArray(0), UNArray(0), 0, bufptr, _
      BufLen, entriesread, totalentries)
    
    
    If Result <> 0 Then
      GoTo HandleError
    End If
      
    For i = 1 To entriesread
      Size = Size + 1
      If Size > ArrayRoom Then
        ArrayRoom = ArrayRoom + 200
        ReDim Preserve Data(1 To ArrayRoom)
      End If
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + (i - 1) * 4, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + (i - 1) * 4 + 2, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      Data(Size) = Left(STRArray, StrLen(TempPtr.x))
    Next i
        
    Result = NetAPIBufferFree(bufptr)         ' Don't leak memory
           
  'Loop Until entriesread = totalentries
  
  If Size > 0 Then
    ReDim Preserve Data(1 To Size)    ' I am not sure whether this is a smart move, but it has
  Else                                ' been made to insure that only real date is passed
    ReDim Preserve Data(1 To 1)    ' accros the DCOM connection
  End If
  
  If entriesread <> totalentries Then GoTo HandleError
  EnumerateUsersGroupMembships = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  EnumerateUsersGroupMembships = False
  GoTo ExitHere
RuntimeError:
  Resume HandleError
End Function

Public Function GetUserInfo(domain As String, User As String, Info As User_Info_3) As Boolean

  'Not all fields in res are set by this function yet

  Dim Result As Long
  Dim bufptr As Long
  Dim str As String
  Dim SNArray() As Byte
  Dim UNArray() As Byte
  Dim PDC As String
  Dim STRArray(500) As Byte
  
  Dim TempPtr As MungeLong
  Dim tempstr As MungeInt
  
  Dim DateVar As Date
  Dim TimeStart As Date
  
  Set Info = New User_Info_3
        
  TimeStart = DateSerial(1970, 1, 1) + TimeSerial(1, 0, 0)
  
  Result = GetPDC("", domain, PDC)
  If Result <> 0 Then GoTo HandleError
  
  SNArray = PDC & vbNullChar       ' Move to byte array
  UNArray = User & vbNullChar       ' Move to byte array
           
  Result = NetUserGetInfo(SNArray(0), _
     UNArray(0), 3, bufptr)
  
  If Result = 0 Then
    With Info
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 0, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 2, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .name = Left(STRArray, StrLen(TempPtr.x))
      
      .password = "****************************"
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 8, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 10, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      
      DateVar = DateAdd("s", -TempPtr.x, Now)
      .password_last_set = Format(DateVar, DateFormat)
      
        
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 12, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 14, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      Select Case TempPtr.x
        Case 0
          .priv = "Guest"
        Case 1
          .priv = "User"
        Case 2
          .priv = "Administrator"
        Case 3
          .priv = "Mask"
        Case Else
          .priv = "Unknown"
      End Select
      
          
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 16, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 18, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .home_dir = Left(STRArray, StrLen(TempPtr.x))
  
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 20, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 22, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .comment = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 24, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 26, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .flags = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 28, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 30, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .script_path = Left(STRArray, StrLen(TempPtr.x))
  
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 32, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 34, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .auth_flags = TempPtr.x
                  
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 36, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 38, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .full_name = Left(STRArray, StrLen(TempPtr.x))
  
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 40, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 42, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .usr_comment = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 44, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 46, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .parms = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 48, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 50, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .workstations_allowed = Left(STRArray, StrLen(TempPtr.x))
      If .workstations_allowed = "" Then .workstations_allowed = "All"
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 52, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 54, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      If TempPtr.x = 0 Then
        .last_logon = "Never"
      Else
        DateVar = DateAdd("s", TempPtr.x, TimeStart)
        .last_logon = Format(DateVar, DateFormat)
      End If
            
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 56, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 58, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      If TempPtr.x = 0 Then
        .last_logoff = "Unknown"
      Else
        DateVar = DateAdd("s", TempPtr.x, TimeStart)
        .last_logoff = Format(DateVar, DateFormat)
      End If
            
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 60, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 62, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      If TempPtr.x = -1 Then
        .acct_expires = "Never"
      Else
        DateVar = DateAdd("s", TempPtr.x, TimeStart)
        .acct_expires = Format(DateVar, DateFormat)
      End If
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 64, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 66, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      If TempPtr.x = -1 Then
        .max_storage = "Unlimited"
      Else
        .max_storage = TempPtr.x
      End If
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 68, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 70, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .units_per_week = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 72, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 74, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .logon_hours = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 76, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 78, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .bad_pw_count = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 80, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 82, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .num_logons = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 84, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 86, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .logon_server = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 88, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 90, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .country_code = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 92, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 94, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .code_page = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 96, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 98, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .user_id = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 100, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 102, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      .primary_group_id = TempPtr.x
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 104, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 106, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .profile = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 108, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 110, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      ' Copy string to array
      Result = PtrToStr(STRArray(0), TempPtr.x)
      .home_dir_drive = Left(STRArray, StrLen(TempPtr.x))
      
      ' Get pointer to string from beginning of buffer
      ' Copy 4-byte block of memory in 2 steps
      Result = PtrToInt(tempstr.XLo, bufptr + 112, 2)
      Result = PtrToInt(tempstr.XHi, bufptr + 114, 2)
      LSet TempPtr = tempstr ' munge 2 integers into a Long
      If TempPtr.x = 0 Then
        .password_expired = "No"
      Else
        .password_expired = "Yes"
      End If
      
    End With
  Else
    GoTo HandleError
  End If
  Result = NetAPIBufferFree(bufptr)         ' Don't leak memory
  
  GetUserInfo = True
ExitHere:
  Exit Function
HandleError:
  On Error Resume Next
  GetUserInfo = False
  GoTo ExitHere
RuntimeError:
  Resume HandleError
  
End Function

