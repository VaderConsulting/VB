'Name:		NetDecs.Bas
'Description:	Network API Declarations & Functions for Windows NT
'Dependencies:	none
'Included
'Functions:	ConvPointerToArray
'		ConvStringToByteArray
'		ConvArrayToPointer
'		ConvPointerToNothing
'		ExtractFieldFromStruct
'------------------------------------------------------------------------
Attribute VB_Name = "NETAPI"
Option Base 0                  ' Important assumption for this code, MEANS ARRAYS START AT 0

Public Const CAP = 400

Global TripError As Boolean 'YOU CAN WATCH THIS VALUE TO SEE IF ERROR HAS OCCURRED IN ANY OF THE FUNCTIONS

Type UserRecord
    Name As String
    FullName As String
    Description As String
End Type 'UserRecord

Type StringArray
    List(CAP) As String
End Type

Type MungeLong
  X As Long
  Dummy As Integer
End Type

Type MungeInt
  XLo As Integer
  XHi As Integer
  Dummy As Integer
End Type

Type Level_1_Group_Structure_Type
    ptrName As Long
    ptrComment As Long
End Type 'Level_1_Group_Structure_Type

Type Level_2_User_Structure_Type
    ptrName As Long
    ptrPassword As Long
    dwPasswordAge As Long
    dwPriv As Long
    ptrHomeDir As Long
    ptrComment As Long
    dwFlags As Long
    ptrScriptPath As Long
    dwAuthFlags As Long
    ptrFullName As Long
    ptrComment2 As Long
    ptrParms As Long
    ptrWorkstations As Long
    dwLastLogon As Long
    dwLastLogoff As Long
    dwAccountExpires As Long
    dwMaxStorage As Long
    dwUnitsPerWeek As Long
    pbLogonHours As Long
    dwBadPasswordCount As Long
    dwNumLogons As Long
    ptrLogonServer As Long
    dwCountryCode As Long
    dwCodePage As Long
End Type 'Level_2_User_Structure_Type

Type HomeDriveType
    ptrHomeDrive As Long
End Type 'HomeDrive Type

Type LoginScriptType
    ptrScriptPath As Long
End Type 'LoginScriptType

Type PlayfulStructType
    ptrVar As Long
End Type 'PlayfulStructType

Declare Function NetGetDCName Lib "NETAPI32.DLL" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Declare Function NetUserChangePassword Lib "NETAPI32.DLL" (ByVal DomainName As String, ByVal UserName As String, ByVal OldPassword As String, ByVal NewPassword As String) As Long
Declare Function NetUserResetPassword Lib "NETAPI32.DLL" Alias "NetUserSetInfo" (ByVal ServerName As String, ByVal UserName As String, ByVal level As Long, Buffer As PlayfulStructType, ParmError As Long) As Long
Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" (lpName As Any, ByVal lpUserName$, lpnLength&)

Declare Function NetUserAdd2 Lib "NETAPI32.DLL" Alias "NetUserAdd" (ServerName As Byte, ByVal level As Long, Buffer As Level_2_User_Structure_Type, ParmError As Long) As Long
Declare Function NetUserDel Lib "NETAPI32.DLL" (ServerName As Byte, UserName As Byte) As Long
Declare Function NetUserSetInfoHomeDrive Lib "NETAPI32.DLL" Alias "NetUserSetInfo" (ServerName As Byte, UserName As Byte, ByVal level As Long, Buffer As HomeDriveType, ParmError As Long) As Long
Declare Function NetUserSetInfoLogonScript Lib "NETAPI32.DLL" Alias "NetUserSetInfo" (ServerName As Byte, UserName As Byte, ByVal level As Long, Buffer As LoginScriptType, ParmError As Long) As Long
Declare Function NetUserSetInfo Lib "NETAPI32.DLL" (ServerName As Byte, UserName As Byte, ByVal level As Long, Buffer As PlayfulStructType, ParmError As Long) As Long
Declare Function NetUserGetInfo Lib "NETAPI32.DLL" (ServerName As Byte, UserName As Byte, ByVal level As Long, Buffer As Long) As Long
Declare Function NetUserGetGroups0 Lib "NETAPI32.DLL" Alias "NetUserGetGroups" (ServerName As Byte, UserName As Byte, ByVal level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long) As Long
Declare Function NetUserEnum0 Lib "NETAPI32.DLL" Alias "NetUserEnum" (ServerName As Byte, ByVal level As Long, ByVal lFilter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long

Declare Function NetGroupEnumUsers0 Lib "NETAPI32.DLL" Alias "NetGroupGetUsers" (ServerName As Byte, GroupName As Byte, ByVal level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Declare Function NetGroupEnum0 Lib "NETAPI32.DLL" Alias "NetGroupEnum" (ServerName As Byte, ByVal level As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHandle As Long) As Long
Declare Function NetGroupAdd1 Lib "NETAPI32.DLL" Alias "NetGroupAdd" (ServerName As Byte, ByVal level As Long, Buffer As Level_1_Group_Structure_Type, ParmError As Long) As Long
Declare Function NetGroupDel Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte) As Long
Declare Function NetGroupAddUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long
Declare Function NetGroupDelUser Lib "NETAPI32.DLL" (ServerName As Byte, GroupName As Byte, UserName As Byte) As Long

Declare Function NetAPIBufferFree Lib "NETAPI32.DLL" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Declare Function NetAPIBufferAllocate Lib "NETAPI32.DLL" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long

Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long
Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Declare Function PtrToArray Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Declare Function GetStrPtr Lib "kernel32" Alias "lstrcpyW" (Dest As Byte, Source As Byte) As Long

Public Sub ConvPointerToArray(p&, a() As Byte)
'THIS FUNCTION CONVERTS A POINTER TO AN ARRAY
'a=byte array
'p=pointer
	Dim Reply&
	Reply& = PtrToArray(a(0), p&)
End Sub


Public Sub ConvStringToByteArray(s$, a() As Byte)
'THIS FUNCTION CONVERTS A STRING TO A BYTE ARRAY
's=string
'a=byte array
    a = s$ & vbNullChar
End Sub


Public Sub ConvArrayToPointer(a() As Byte, p&)
'THIS FUNCTION CONVERTS A BYTE ARRAY TO A POINTER
'a=byte array
'p=pointer
    Reply& = NetAPIBufferAllocate(UBound(a) + 1, p&)
    p& = StrToPtr(p&, a(0))
End Sub


Public Sub ConvPointerToNothing(p&)
'THIS FUNCTION FREES UP A POINTER
'THIS FUNCTION HAS A FUNNY NAME SO IT SORTS TO THE TOP WITH
'THE OTHER POINTER AND STRING FUNCTIONS

    Reply& = NetAPIBufferFree(p&)
End Sub

Public Function ExtractFieldFromStruct(struct&, field_num%) As String
'THIS FUNCTION EXTRACTS A FIELD OUT OF A STRUCTURE THAT IS ACCESSED BY ITS POINTER
'NO VB EQUIVALENT IS POSSIBLE, SO IT MUST BE "MUNGED" (PUT A SQUARE PEG IN ROUND HOLE)
'
'DOWNSIDE: YOU MUST RETRIEVE struct& EACH TYPE YOU WANT ANOTHER FIELD
'I REPEAT--THIS FUNCTION WILL LOCK YOUR SYSTEM UP IF YOU USE IT TO RETRIEVE MORE THAN ONE FIELD
'WITHOUT REALLOCATING struct&

Dim arrayField(255) As Byte, Reply&
Dim TempPtr As MungeLong, TempStr As MungeInt
    
    ' Get pointer to string from beginning of buffer
    Reply& = PtrToInt(TempStr.XLo, struct& + (field_num% - 1) * 4, 2)      ' Doing this to copy a 4-byte block of memory to a Long
    Reply& = PtrToInt(TempStr.XHi, struct& + (field_num% - 1) * 4 + 2, 2)
    LSet TempPtr = TempStr
    ' Copy string to array
    Reply& = PtrToArray(arrayField(0), TempPtr.X)
    ExtractFieldFromStruct = Trim(Left(arrayField, StrLen(TempPtr.X)))
End Function
