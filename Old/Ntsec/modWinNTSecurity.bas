Attribute VB_Name = "modWinNTSecurity"
Option Explicit

Public LastError As Long
Public LastErrorMsg As String

Enum ACCESS_MODE
    NOT_USED_ACCESS = 0
    GRANT_ACCESS
    SET_ACCESS
    DENY_ACCESS
    REVOKE_ACCESS
    SET_AUDIT_SUCCESS
    SET_AUDIT_FAILURE
End Enum

Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0&
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    'SE_DS_OBJECT
    'SE_DS_OBJECT_ALL
    'SE_PROVIDER_DEFINED_OBJECT
End Enum

Enum MULTIPLE_TRUSTEE_OPERATION
    NO_MULTIPLE_TRUSTEE
    TRUSTEE_IS_IMPERSONATE
End Enum

Enum TRUSTEE_FORM
    TRUSTEE_IS_SID
    TRUSTEE_IS_NAME
End Enum

Enum TRUSTEE_TYPE
    TRUSTEE_IS_UNKNOWN
    TRUSTEE_IS_USER
    TRUSTEE_IS_GROUP
End Enum

Type TRUSTEE
    pMultipleTrustee            As Long
    MultipleTrusteeOperation    As MULTIPLE_TRUSTEE_OPERATION
    TrusteeForm                 As TRUSTEE_FORM
    TrusteeType                 As TRUSTEE_TYPE
    ptstrName                   As String
End Type

Type EXPLICIT_ACCESS
    grfAccessPermissions        As Long
    grfAccessMode               As ACCESS_MODE
    grfInheritance              As Long
    TRUSTEE                     As TRUSTEE
End Type

Type AceArray
    List() As EXPLICIT_ACCESS
End Type

Public Const OWNER_SECURITY_INFORMATION = &H1
Public Const GROUP_SECURITY_INFORMATION = &H2
Public Const DACL_SECURITY_INFORMATION = &H4
Public Const SACL_SECURITY_INFORMATION = &H8

'Generic AccessRights
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000

'Common AccessRights combinations
Public Const COMMON_ADD = &H1201B6
Public Const COMMON_ADD_READ = &H1201BF
Public Const COMMON_READ = &H1200A9
Public Const COMMON_CHANGE = &H1301BF
Public Const COMMON_FULL_CONTROL = GENERIC_ALL
Public Const COMMON_ALL = &H1F01FF
'Inheritance flags
Public Const NO_INHERITANCE = &H0
Public Const OBJECT_INHERIT_ACE = &H1
Public Const CONTAINER_INHERIT_ACE = &H2
Public Const NO_PROPAGATE_INHERIT_ACE = &H4
Public Const INHERIT_ONLY_ACE = &H8
Public Const INHERITED_ACE = &H10
Public Const VALID_INHERIT_FLAGS = &H1F



Declare Function GetNamedSecurityInfo Lib "advapi32.dll" Alias "GetNamedSecurityInfoA" ( _
    ByVal pObjectName As String, _
    ByVal ObjectType As SE_OBJECT_TYPE, _
    ByVal SecurityInfo As Long, _
    ppsidOwner As Long, _
    ppsidGroup As Long, _
    ppDacl As Long, _
    ppSacl As Long, _
    ppSecurityDescriptor As Long) As Long

Declare Sub BuildExplicitAccessWithName Lib "advapi32.dll" Alias "BuildExplicitAccessWithNameA" ( _
    pExplicitAccess As EXPLICIT_ACCESS, _
    ByVal pTrusteeName As String, _
    ByVal AccessPermissions As Long, _
    ByVal AccessMode As ACCESS_MODE, _
    ByVal Inheritance As Long)

Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" ( _
    ByVal cCountOfExplicitEntries As Long, _
    pListOfExplicitEntries As EXPLICIT_ACCESS, _
    ByVal OldAcl As Long, _
    NewAcl As Long) As Long
 
Declare Function SetNamedSecurityInfo Lib "advapi32.dll" Alias "SetNamedSecurityInfoA" ( _
    ByVal pObjectName As String, _
    ByVal ObjectType As SE_OBJECT_TYPE, _
    ByVal SecurityInfo As Long, _
    psidOwner As Long, _
    psidGroup As Long, _
    ByVal pDACL As Long, _
    pSacl As Long) As Long

Declare Function GetExplicitEntriesFromAcl Lib "advapi32.dll" Alias "GetExplicitEntriesFromAclA" ( _
    ByVal pacl As Long, _
    pcCountOfExplicitEntries As Long, _
    pListOfExplicitEntries As Long) As Long

Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Declare Function GetLastError Lib "kernel32.dll" () As Long
Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (RetVal As Long, ByVal Ptr As Long, ByVal nCharCount As Long)
Declare Function CopyMemory Lib "kernel32" Alias "lstrcpynW" (RetVal As Long, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStrA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
  
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const LANG_USER_DEFAULT = &H400&

Function AccessRights(ByVal grfAccessMode As Long, ByVal grfAccessPermissions As Long)
    If grfAccessPermissions = COMMON_FULL_CONTROL Then
        If grfAccessMode <> DENY_ACCESS Then
            AccessRights = "Full Control"
        Else
            AccessRights = "None"
        End If
        Exit Function
    End If
    
    If grfAccessPermissions = COMMON_ALL Then
        If grfAccessMode <> DENY_ACCESS Then
            AccessRights = "All"
        Else
            AccessRights = "None"
        End If
        Exit Function
    End If
    
    If grfAccessPermissions = COMMON_ADD Then
        AccessRights = "Add "
        Exit Function
    End If
    
    If grfAccessPermissions = COMMON_ADD_READ Then
        AccessRights = "Add & Read"
        Exit Function
    End If
    
    If grfAccessPermissions = COMMON_CHANGE Then
        AccessRights = "Change"
        Exit Function
    End If
    
    If grfAccessPermissions = COMMON_READ Then
        AccessRights = "Read"
        Exit Function
    End If
    
End Function

Function AddAccessControlElement(ByVal lpObjectName, ByVal ObjectType As SE_OBJECT_TYPE, ByVal TrusteeName As String, ByVal AccessPermissions As Long, ByVal AccessMode As ACCESS_MODE) As Boolean
    Dim dwRes As Long
    Dim pOldDACL As Long
    Dim pNewDACL As Long
    Dim pSD As Long
    Dim ExplicitAccess As EXPLICIT_ACCESS
    Dim ea As EXPLICIT_ACCESS
    Dim I As Integer
    
    If lpObjectName = Empty Then
        AddAccessControlElement = False
        Exit Function
    End If
        
    'get a pointer to the existing DACL
    dwRes = GetNamedSecurityInfo(lpObjectName, ObjectType, _
            DACL_SECURITY_INFORMATION, _
            0&, 0&, pOldDACL, 0&, pSD)
            
    If dwRes <> 0 Then
        Call SetError(dwRes)
        AddAccessControlElement = False
        Exit Function
    End If
    
    'initialize an EXPLICIT_ACCESS structure to allow access
    Call ZeroMemory(ea, Len(ea))
    Call BuildExplicitAccessWithName(ea, TrusteeName & vbNullChar, AccessPermissions, _
            AccessMode, NO_INHERITANCE)
            
    'Create a new ACL by merging the EXPLICIT_ACCESS structure
    'with the existing DACL
    dwRes = SetEntriesInAcl(1, ea, pOldDACL, pNewDACL)
    If dwRes <> 0 Then
        Call SetError(dwRes)
        AddAccessControlElement = False
        Exit Function
    End If
    
    'Attach the new ACL as the object's DACL
    dwRes = SetNamedSecurityInfo(lpObjectName, ObjectType, _
            DACL_SECURITY_INFORMATION, _
            0&, 0&, pNewDACL, 0&)
    
    If dwRes <> 0 Then
        Call SetError(dwRes)
        AddAccessControlElement = False
        Exit Function
    End If
    
    AddAccessControlElement = True
    
    If pSD <> 0 Then dwRes = LocalFree(pSD)
    If pNewDACL <> 0 Then dwRes = LocalFree(pNewDACL)
End Function
Function GetAccessControlElements(ByVal lpObjectName, ByVal ObjectType As SE_OBJECT_TYPE) As AceArray
    Dim dwRes As Long
    Dim pDACL As Long
    Dim pSD As Long
    Dim ExplicitAccess As EXPLICIT_ACCESS
    Dim pcCountOfExplicitEntries As Long
    Dim pListOfExplicitEntries As Long
    Dim ACE() As Long
    Dim I As Integer
    
    If lpObjectName = Empty Then Exit Function
    
    'get a pointer to the existing DACL
    dwRes = GetNamedSecurityInfo(lpObjectName, ObjectType, _
            DACL_SECURITY_INFORMATION, _
            0&, 0&, pDACL, 0&, pSD)
            
    If dwRes <> 0 Then Call SetError(dwRes): Exit Function
    
    'Retrieve an array of EXPLICIT_ACCESS structures from the ACL
    dwRes = GetExplicitEntriesFromAcl(pDACL, pcCountOfExplicitEntries, pListOfExplicitEntries)
    dwRes = Val("&H" & Right$(Hex$(dwRes), 4))
    If dwRes <> 0 Then Call SetError(dwRes): Exit Function
    
    ReDim ACE((Len(ExplicitAccess) / 4) * pcCountOfExplicitEntries)
    Call MoveMemory(ACE(0), pListOfExplicitEntries, Len(ExplicitAccess) * pcCountOfExplicitEntries)
    
    ReDim GetAccessControlElements.List(pcCountOfExplicitEntries - 1)
    
    For I = 0 To pcCountOfExplicitEntries - 1
        With GetAccessControlElements.List(I)
            .grfAccessPermissions = ACE(I * 8 + 0)
            .grfAccessMode = ACE(I * 8 + 1)
            .grfInheritance = ACE(I * 8 + 2)
            .TRUSTEE.pMultipleTrustee = ACE(I * 8 + 3)
            .TRUSTEE.MultipleTrusteeOperation = ACE(I * 8 + 4)
            .TRUSTEE.TrusteeForm = ACE(I * 8 + 5)
            .TRUSTEE.TrusteeType = ACE(I * 8 + 6)
            .TRUSTEE.ptstrName = PointerToString(ACE(I * 8 + 7))
        End With
    Next I
    
    Erase ACE
    If pSD <> 0 Then dwRes = LocalFree(pSD)
    If pDACL <> 0 Then dwRes = LocalFree(pDACL)
    If pListOfExplicitEntries <> 0 Then dwRes = LocalFree(pListOfExplicitEntries)
End Function


Sub Main()
    Dim ACEs As AceArray
    Dim I As Integer
    
    
    'If AddAccessControlElement("d:\SecTest", SE_FILE_OBJECT, "heelal\acn3520", GENERIC_ALL) Then MsgBox "Ok."
    
    ACEs = GetAccessControlElements("d:\SecTest", SE_FILE_OBJECT)
    With ACEs
        For I = 0 To UBound(.List)
            With .List(I)
                Debug.Print .TRUSTEE.TrusteeForm & " - " & .TRUSTEE.TrusteeType & " - " & .TRUSTEE.ptstrName & " - " & AccessRights(.grfAccessMode, .grfAccessPermissions)
            End With
        Next I
    End With
        

End Sub

Function PointerToString(ByVal Pointer) As String
    Dim StringValue As String
    Dim NullPos As Long
    Dim Temp As Long
    
    ' Copy string to array and convert to a string
    If Pointer > 0 And StrLen(Pointer) > 0 Then
        StringValue = Space$(StrLen(Pointer) + 50)
        Temp = PtrToStrA(StringValue, Pointer)
        NullPos = InStr(StringValue, Chr$(0))
        If NullPos > 0 Then
            PointerToString = Left$(StringValue, NullPos - 1) 'Lose the null terminator...
        Else
            PointerToString = StringValue 'Just pass the string...
        End If
    Else
        PointerToString = ""
    End If
End Function


Sub SetError(ByVal dwErrCode As Long)
    Static sMsgBuf As String * 257, dwLen As Long
    
    dwLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM _
                Or FORMAT_MESSAGE_IGNORE_INSERTS _
                Or FORMAT_MESSAGE_MAX_WIDTH_MASK, ByVal 0&, _
                dwErrCode, LANG_USER_DEFAULT, _
                ByVal sMsgBuf, 256&, 0&)
    
    LastError = dwErrCode
    If dwLen Then LastErrorMsg = Left$(sMsgBuf, dwLen)
End Sub


