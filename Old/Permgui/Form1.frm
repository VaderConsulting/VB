VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify access..."
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optType 
      Caption         =   "Custom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton optType 
      Caption         =   "Standard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fmeCustom 
      Caption         =   "Custom"
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3015
      Begin VB.OptionButton optSecurity 
         Caption         =   "Other"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optSecurity 
         Caption         =   "Access not specified"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optSecurity 
         Caption         =   "Full Control"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame fmeAccess 
         Caption         =   "Other"
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2175
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Read"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Write"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Execute"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Delete"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Change Permission"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Take Ownership"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   1815
         End
      End
   End
   Begin VB.Frame fmeStandard 
      Caption         =   "Standard"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3615
      Begin VB.OptionButton optStandard 
         Caption         =   "Full Control"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optStandard 
         Caption         =   "Change"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optStandard 
         Caption         =   "Read Only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fmeScope 
      Caption         =   "Scope"
      Height          =   855
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   1335
      Begin VB.CheckBox chkScope 
         Caption         =   "Folders"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkScope 
         Caption         =   "Files"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\temp\security"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      Caption         =   "Enter Path to modify security.."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CODE Stolen from MSDN  and modifed to take Comandline
'makes folder or file readonly for everyone without talking to Domain controller
'usage Readonly /f:File/FolderName
'compile as PCODE

Option Explicit

Private Type SID_IDENTIFIER_AUTHORITY
    Value(5) As Byte '6 bytes
End Type

Private Type ACE_HEADER
    AceType As Byte
    AceFlags As Byte
    AceSize As Integer
End Type

Private Type ACCESS_ALLOWED_ACE
    Header As ACE_HEADER
    Mask As Long
    SidStart As Long
End Type

Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    ACLSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Integer
    Owner As Long
    Group As Long
    sACL As ACL
    dacl As ACL
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type ACL_SIZE_INFORMATION
    AceCount As Long
    AclBytesInUse As Long
    AclBytesFree As Long
End Type


Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, ByVal dwRevision As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" ( _
    pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, _
    ByVal nSubAuthorityCount As Byte, _
    ByVal nSubAuthority0 As Long, _
    ByVal nSubAuthority1 As Long, _
    ByVal nSubAuthority2 As Long, _
    ByVal nSubAuthority3 As Long, _
    ByVal nSubAuthority4 As Long, _
    ByVal nSubAuthority5 As Long, _
    ByVal nSubAuthority6 As Long, _
    ByVal nSubAuthority7 As Long, _
    pSID As Long) _
As Long  ' pSid above in AllocateAndInitializeSid: pass pointer byref
' pSid in GetLengthSid below is dereferenced pass byval
Private Declare Function GetLengthSid Lib "advapi32.dll" (ByVal pSID As Long) As Long
' pSid is dereferenced in FreeSid pass byval
Private Declare Sub FreeSid Lib "advapi32.dll" (ByVal pSID As Long)
' pSid  is dereferenced CopySid pass byval
Private Declare Function CopySid Lib "advapi32.dll" (ByVal nDestinationSidLength As Long, pDestinationSid As Byte, ByVal pSID As Long) As Long
Private Declare Function InitializeAcl Lib "advapi32.dll" (pAcl As Byte, ByVal nAclLength As Long, ByVal dwAclRevision As Long) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As Long, ByVal bDaclPresent As Long, pDacl As Byte, ByVal bDaclDefaulted As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AddAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, pAceList As Byte, ByVal nAceListLength As Long) As Long
Private Declare Function AddAce2 Lib "advapi32.dll" Alias "AddAce" (ByVal pAcl As Long, ByVal dwAceRevision As Long, ByVal dwStartingAceIndex As Long, ByVal pAceList As Long, ByVal nAceListLength As Long) As Long
Private Declare Function SetFileSecurity Lib "advapi32.dll" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, ByVal pSecurityDescriptor As Long) As Long

Private Const ACCESS_ALLOWED_ACE_TYPE = &H0
Private Const ACCESS_DENIED_ACE_TYPE = &H1
Private Const ACCESS_SYSTEM_SECURITY = &H1000000
Private Const ACL_REVISION = (2)
Private Const ACL_REVISION1 = (1)
Private Const ACL_REVISION2 = (2)
Private Const AclRevisionInformation = 1
Private Const AclSizeInformation = 2
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20
Private Const SECURITY_CONTEXT_TRACKING = &H40000
Private Const SECURITY_CREATOR_GROUP_RID = &H1
Private Const SECURITY_CREATOR_OWNER_RID = &H0
Private Const SECURITY_DESCRIPTOR_MIN_LENGTH = (20)
Private Const SECURITY_DESCRIPTOR_REVISION = (1)
Private Const SECURITY_DESCRIPTOR_REVISION1 = (1)
Private Const SECURITY_EFFECTIVE_ONLY = &H80000
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_LOCAL_RID = &H0
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_NULL_RID = &H0
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_SQOS_PRESENT = &H100000
Private Const SECURITY_VALID_SQOS_FLAGS = &H1F0000
Private Const SECURITY_WORLD_RID = &H0
Private Const SECURITY_ANONYMOUS = 1
Private Const SECURITY_IDENTIFICATION = 2
Private Const SECURITY_NULL_SID_AUTHORITY = 0
Private Const SECURITY_WORLD_SID_AUTHORITY = 1
Private Const SECURITY_LOCAL_SID_AUTHORITY = 2
Private Const SECURITY_CREATOR_SID_AUTHORITY = 3
Private Const SECURITY_NT_AUTHORITY As Long = 5&
Private Const DACL_SECURITY_INFORMATION As Long = &H4
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_GROUP_RID_ADMINS = &H200
Private Const DOMAIN_GROUP_RID_GUESTS = &H202
Private Const DOMAIN_GROUP_RID_USERS = &H201
Private Const DOMAIN_USER_RID_ADMIN = &H1F4
Private Const DOMAIN_USER_RID_GUEST = &H1F5
Private Const SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege" & vbNullChar
 
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const ERROR_SUCCESS = 0&
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const OWNER_SECURITY_INFORMATION As Long = &H1&
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const LANG_USER_DEFAULT = &H400&

Private Const GENERIC_READ = &H80000000            ' Read only
Private Const GENERIC_ALL = &H10000000             ' Full Control
Private Const INHERIT_ONLY_ACE = &H8
Private Const OBJECT_INHERIT_ACE = &H1
Private Const GENERIC_EXECUTE = &H20000000
Private Const MAXDWORD = &HFFFFFFFF
Private Const CONTAINER_INHERIT_ACE = &H2
Private Const FILE_READ_ATTRIBUTES = &H80
Private Const FILE_READ_EA = &H8
Private Const FILE_READ_DATA = &H1
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const DELETE = &H10000
Private Const GENERIC_WRITE = &H40000000
Private Const READONLY = &HA0000000

' ***************** Constants for AddAccessAllowed Ace code ***************

Private Const GMEM_MOVEABLE = &H2
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)
Private Const SD_SIZE = (65536 + SECURITY_DESCRIPTOR_MIN_LENGTH)
Private Const SidTypeUser = 1

' New API Declarations for AddAccessAllowed Ace code

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, Sid As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function GetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As Byte, lpbDaclPresent As Long, pDacl As Long, lpbDaclDefaulted As Long) As Long
Private Declare Function GetFileSecurityN Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, ByVal pSecurityDescriptor As Long, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function GetFileSecurity Lib "advapi32.dll" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As Byte, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function GetAclInformation Lib "advapi32.dll" (ByVal pAcl As Long, pAclInformation As Any, ByVal nAclInformationLength As Long, ByVal dwAclInformationClass As Long) As Long
Private Declare Function GetAce Lib "advapi32.dll" (ByVal pAcl As Long, ByVal dwAceIndex As Long, pace As Any) As Long
Private Declare Function AddAccessAllowedAce Lib "advapi32.dll" (pAcl As Byte, ByVal dwAceRevision As Long, ByVal AccessMask As Long, pSID As Byte) As Long

Sub Form_Load()
    Dim lpos As Integer, rpos As Integer, retval As Boolean, sFolder As String
    Dim UserProfilePath As String, CommandLine As String
    UserProfilePath = UCase(Environ$("UserProfile"))
    'Thanks Dave for this bit
    ' command line = "/f:%UserProfile%\Desktop" , where /f: indicates the filename to import
    CommandLine = UCase(Command)
    If CommandLine = "" Then
        'frmMain.Caption = "Enter folder - click btn to change permissions"
    Else
        Me.Hide
        If InStr(1, CommandLine, "/F:") > 0 Then
            lpos = InStr(1, CommandLine, "/f:") + 4
            If rpos = 0 Then rpos = Len(CommandLine)
            sFolder = Mid(CommandLine, lpos, rpos)
            sFolder = Replace(sFolder, "%USERPROFILE%", UserProfilePath)
            retval = AlterPermissions(sFolder)
            End
        End If
    End If
End Sub

Function AlterPermissions(sFolder As String) As Boolean
    Dim result1 As Boolean, result2 As Boolean, result3 As Boolean, result4 As Boolean
    
    ' FirstAddReadMask = Files                      SecondAddReadMask = Folders
    
    ' Everyone Read Only
    result1 = ReplacePermissions(sFolder, 1, SECURITY_WORLD_RID, 0, 0, 0, SECURITY_WORLD_SID_AUTHORITY, READONLY, READONLY)
    
    ' System Full Control
    'result2 = ReplacePermissions(sFolder, 1, SECURITY_LOCAL_SYSTEM_RID, 0, 0, 0, SECURITY_NT_AUTHORITY, GENERIC_ALL, GENERIC_ALL)
    
    ' Administrators Full Control
    'result3 = ReplacePermissions(sFolder, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, SECURITY_NT_AUTHORITY, GENERIC_ALL, GENERIC_ALL)
    
    ' Test Combined RID's  -  does not work
    'result4 = ReplacePermissions(sFolder, 3, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, SECURITY_LOCAL_SYSTEM_RID, 0, SECURITY_NT_AUTHORITY, GENERIC_ALL, GENERIC_ALL)
    
    'AddPermissions (sFolder) ' <------ Screws app!
    AlterPermissions = result1 Or result2 Or result3
End Function

Private Function ReplacePermissions(sFName As String, RIDCount As Byte, RID1 As Long, RID2 As Long, RID3 As Long, RID4 As Long, Scope As Long, FirstAceMask As Long, SecondAceMask As Long) As Boolean
   
   ' This function will set permissions for Folder specified in sFName
   ' FirstAceMask will set the mask for the files in the folder
   ' SecondAceMask will set the mask for the folder and subfolders.
   
   Dim udtSidIdentifierAuthority As SID_IDENTIFIER_AUTHORITY
   Dim udtAccessAllowedAce As ACCESS_ALLOWED_ACE
   Dim pSID As Long, ACLSize As Long
   Dim lAceSize As Long
   Dim x As Long
   Dim I As Integer
      
   ' To assign permissions you need to be a member of the
   ' Administrators group and the folder must be on an NTFS partition.
   For I = 0 To 4
       udtSidIdentifierAuthority.Value(I) = 0
   Next I
   udtSidIdentifierAuthority.Value(5) = Scope ' Appropriate ??????_NT_AUTHORITY
      
   Dim psdl As Long ' used to get security descriptor pointer
   
   ' Initialize Security Descriptor - the first parameter is address of the
   ' security descriptor.
   x = InitializeSecurityDescriptor(VarPtr(psdl), SECURITY_DESCRIPTOR_REVISION)
   If x = 0 Then
       Debug.Print "InitializeSecurityDescriptor failed"
       Exit Function
   End If
   
   ' Allocate and initialize a Sid
   ' The first parameter is the Sid Authority Identifier which identifies
   '    the top level authority.
   ' The second parameter indicates that there are 2 subauthorities to
   '    be placed in the SID.
   ' The RID is the Relative Identifier found in the SID that identifies
   '    the user or group.
   ' The last parameter, pSid is a pointer to a pointer to the sid.
   '    It is passed byref here and later dereferenced passing byval.
   '    It is passed byval in GetLengthSid, CopySid and FreeSid.
   x = AllocateAndInitializeSid(udtSidIdentifierAuthority, RIDCount, RID1, RID2, RID3, RID4, 0, 0, 0, 0, pSID)
   If x = 0 Then
       Debug.Print "AllocateAndInitializeSid failed"
       Exit Function
   End If
   
   ' Calculate length of ACL
   Dim AnAcl As ACL, AnAAA As ACCESS_ALLOWED_ACE
   
   ACLSize = Len(AnAcl) + Len(AnAAA) - Len(x) + GetLengthSid(pSID) + Len(AnAAA) - Len(x) + GetLengthSid(pSID)
    
   ' Allocate memory for ACL
   ReDim bufACL(ACLSize - 1) As Byte
   
   ' Init the ACL
   ' Creates a new ACL structure.  This is a variable length structure.
   ' A byte array is created to hold the contiguous bytes.  The size of
   ' this buffer was calculated above in ACLSize.  In this call the pointer
   ' to the first element of the array is passed.
   ' The third parameter must be ACL_REVISION.
   x = InitializeAcl(bufACL(0), ACLSize, ACL_REVISION)
   If x = 0 Then
       Debug.Print "InitializeAcl failure"
       Exit Function
   End If
     
   ' Set values in AllowedAccessAce structure:
   ' It is easier to assign values using UDT notation.
   ' There are 2 types of ACEs.  Access allowed and Access denied.
   ' This field indicates that we are using an Access allowed ACE.
   udtAccessAllowedAce.Header.AceType = ACCESS_ALLOWED_ACE_TYPE
   ' AceFlags field specifies control flags.
   ' INHERIT_ONLY_ACE does not apply to containers (folder) but to objects
   ' in the container (files).
   ' OBJECT_INHERIT_ACE indicates that the ACE is inherited by
   ' non container objects such as files within the container object
   ' to which the ACE is assigned.
   udtAccessAllowedAce.Header.AceFlags = INHERIT_ONLY_ACE Or OBJECT_INHERIT_ACE
   ' Size of the ACE
   lAceSize = Len(AnAAA) - Len(x) + GetLengthSid(pSID)
   udtAccessAllowedAce.Header.AceSize = lAceSize
   ' The Mask specifies access rights granted to the ACE.
   ' See SDK documentation under ACCESS_MASK for the break down by bits.
   udtAccessAllowedAce.Mask = FirstAceMask
   ReDim bufAce(lAceSize - 1) As Byte
   ' Copy the AccessAllowedAce structure(UDT) to buffer.
   CopyMemory bufAce(0), udtAccessAllowedAce, lAceSize
   ' Copy sid to buffer where bufAce is the destination buffer -
   ' pSid is ptr to source SID.
   ' y(8) corresponds to SidStart field in AccessAllowedAce struct.
   x = CopySid(GetLengthSid(pSID), bufAce(8), pSID)
    If x = 0 Then
       Debug.Print "CopySid fail " & Err.LastDllError
       Exit Function
    End If
       
    ' Add an ACE to ACL.  This is done twice.
    ' This first AddAce call applies to files in the folder.
    ' The first parameter is a pointer to the variable length ACL which is
    '     passed using a pointer to the first element in a byte array.
    ' The second parameter needs to be ACL_REVISION.
    ' The third paramter specifies the position of the ACE in the
    '    ACL which in this case is at the end.
    ' The fourth parameter is a pointer to one or more ACEs.
    '    These ACEs would be placed in contiguous memory and are
    '    placed in a byte array the size which is placed in the last parameter.
    x = AddAce(bufACL(0), ACL_REVISION, MAXDWORD, bufAce(0), udtAccessAllowedAce.Header.AceSize)
    If x = 0 Then
        Debug.Print "First AddAce failed with error " & Err.LastDllError
    End If
    
    CopyMemory udtAccessAllowedAce, bufAce(0), udtAccessAllowedAce.Header.AceSize
        
    udtAccessAllowedAce.Mask = SecondAceMask
    udtAccessAllowedAce.Header.AceFlags = CONTAINER_INHERIT_ACE

    CopyMemory bufAce(0), udtAccessAllowedAce, udtAccessAllowedAce.Header.AceSize
   
    
    ' bufACL(0) - ptr to first element in byte array that
    ' contains contents of the acl structure - ACE gets added to this ACL
    ' which contains ACEs stored contiguously.
    ' This ACE applies to directories and subdirectories.
        
    x = AddAce(bufACL(0), ACL_REVISION, MAXDWORD, bufAce(0), udtAccessAllowedAce.Header.AceSize)
    If x = 0 Then
        Debug.Print "Second AddAce failed with error " & Err.LastDllError
    End If
        
    ' Set the DACL in the security descriptor
    ' The first paramter is the pointer to the security descriptor.
    ' The second paramter is boolean indicating the presence of a DACL
    ' in the Security Descriptor.
    ' The third parameter is the address of the DACL which is variable
    ' length and passed in a byte array.
    ' The fourth paramter indicates that DACL is created by user.
    x = SetSecurityDescriptorDacl(VarPtr(psdl), 1, bufACL(0), 0)
    If x = 0 Then
       Debug.Print "SetSecurityDescriptorDacl failure " & Err.LastDllError
       Exit Function
    End If
    
    Dim si As Long
    si = DACL_SECURITY_INFORMATION
   
    ' Set security on Folder object
    x = SetFileSecurity(sFName, si, VarPtr(psdl))
    If x = 0 Then
        Debug.Print "SetFileSecurity failure " & Err.LastDllError
        Exit Function
    End If
    
    ' Free Sid
    FreeSid pSID
    
    ReplacePermissions = True
    
End Function

Public Sub AddPermissions(sFilename As String)
   Dim lResult As Long                  ' Result of various API calls.
   Dim I As Integer                     ' Used in looping.
   Dim bUserSid(255) As Byte            ' This will contain your SID.
   Dim sSystemName As String            ' Name of this computer system.
   Dim lSystemNameLength As Long        ' Length of string that contains the name of this system.
   Dim lLengthUserName As Long          ' Max length of user name.
   Dim sUserName As String * 255        ' String to hold the current user name.
   Dim lUserSID As Long                 ' Used to hold the SID of the current user.
   Dim lUserSIDSize As Long             ' Size of the SID.
   Dim sDomainName As String * 255      ' Domain the user belongs to.
   Dim lDomainNameLength As Long        ' Length of domain name needed.
   Dim lSIDType As Long                 ' The type of SID info we are getting back.
   Dim sFileSD As SECURITY_DESCRIPTOR   ' SD of the file we want.
   Dim bSDBuf() As Byte                 ' Buffer that holds the security descriptor for this file.
   Dim lFileSDSize As Long              ' Size of the File SD.
   Dim lSizeNeeded As Long              ' Size needed for SD for file.
   Dim sNewSD As SECURITY_DESCRIPTOR    ' New security descriptor.
   Dim sACL As ACL                      ' Used in grabbing the DACL from the File SD.
   Dim lDaclPresent As Long             ' Used in grabbing the DACL from the File SD.
   Dim lDaclDefaulted As Long           ' Used in grabbing the DACL from the File SD.
   Dim sACLInfo As ACL_SIZE_INFORMATION ' Used in grabbing the ACL from the File SD.
   Dim lACLSize As Long                 ' Size of the ACL structure used to get the ACL from the File SD.
   Dim pAcl As Long                     ' Current ACL for this file.
   Dim lNewACLSize As Long              ' Size of new ACL to create.
   Dim bNewACL() As Byte                ' Buffer to hold new ACL.

   Dim sCurrentACE As ACCESS_ALLOWED_ACE ' Current ACE.
   Dim pCurrentAce As Byte               ' Our current ACE. ***** CHANGED from Long to Byte

   ' The first action taken is acquiring the name of the user
   ' who is currently logged onto this system. Take the user's
   ' name and grab its companion SID for future use.
   ' Use the GetUserName API to find out who is currently logged onto
   ' this system. Preset the length of the string to hold the
   ' returned user name from the "GetUserName" API.
   lLengthUserName = 255
   sUserName = Space(lLengthUserName)

   ' Call GetUserName to find out who is logged onto this system.
   lResult = GetUserName(sUserName, lLengthUserName)

   ' Return value of zero means the call failed; test for this before
   ' continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Retrieve the Current User Name"
      Exit Sub
   End If

   ' You now have the user name of the person who is logged onto this
   ' system. Using that information, get the SID of the user.  Get the SID of this
   ' user by using the LookupAccountName API. In order to use the SID
   ' of the current user account, call the LookupAccountName API
   ' twice. The first time is to get the required sizes of the SID
   ' and the DomainName string. The second call is to actually get
   ' the desired information.

   lResult = LookupAccountName(vbNullString, sUserName, bUserSid(0), 255, sDomainName, lDomainNameLength, lSIDType)

   ' Now set the sDomainName string buffer to its proper size before
   ' calling the API again.
   sDomainName = Space(lDomainNameLength)

   ' Call the LookupAccountName again to get the actual SID for user.
   lResult = LookupAccountName(vbNullString, sUserName, bUserSid(0), 255, sDomainName, lDomainNameLength, lSIDType)

   ' Return value of zero means the call to LookupAccountName failed;
   ' test for this before you continue.
     If (lResult = 0) Then
        Debug.Print "Error: Unable to Lookup the Current User Account: " & sUserName
        Exit Sub
     End If

   ' You now have the SID for the user who is logged on.
   ' The SID is of interest since it will get the security descriptor
   ' for the file that the user is interested in.

   lFileSDSize = Len(sFileSD)

   ' The GetFileSecurity API will retrieve the Security Descriptor
   ' for the file. However, you must call this API twice: once to get
   ' the proper size for the Security Descriptor and once to get the
   ' actual Security Descriptor information.

   lResult = GetFileSecurityN(sFilename, DACL_SECURITY_INFORMATION, 0, 0, lSizeNeeded)

   ' Redimension the Security Descriptor buffer to the proper size.
   ReDim bSDBuf(lSizeNeeded)

   ' Now get the actual Security Descriptor for the file.
   lResult = GetFileSecurity(sFilename, DACL_SECURITY_INFORMATION, bSDBuf(0), lSizeNeeded, lSizeNeeded)

   ' A return code of zero means the call failed; test for this before continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Get the File Security Descriptor"
      Exit Sub
   End If

    Dim psdl As Long ' used to get security descriptor pointer
   ' Call InitializeSecurityDescriptor to build a new SD for the file.
   'lResult = InitializeSecurityDescriptor(sNewSD, SECURITY_DESCRIPTOR_REVISION)
   lResult = InitializeSecurityDescriptor(VarPtr(psdl), SECURITY_DESCRIPTOR_REVISION)

   ' A return code of zero means the call failed; test for this before continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Initialize New Security Descriptor"
      Exit Sub
   End If

   ' You now have the file's SD and a new Security Descriptor
   ' that will replace the current one. Next, pull the DACL from
   ' the SD. To do so, call the GetSecurityDescriptorDacl API
   ' function.

   lResult = GetSecurityDescriptorDacl(bSDBuf(0), lDaclPresent, pAcl, lDaclDefaulted)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Get DACL from File Security Descriptor"
      Exit Sub
   End If

   ' You have the file's SD, and want to now pull the ACL from the
   ' SD. To do so, call the GetACLInformation API function.
   ' See if ACL exists for this file before getting the ACL
   ' information.
   If (lDaclPresent = False) Then
      Debug.Print "Error: No ACL Information Available for this File"
      Exit Sub
   End If

   ' Attempt to get the ACL from the file's Security Descriptor.
   lResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2&)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Get ACL from File Security Descriptor"
      Exit Sub
   End If

   ' Now that you have the ACL information, compute the new ACL size
   ' requirements.
   Dim pSID As Long
   pSID = VarPtr(bUserSid(0))
   'lNewACLSize = sACLInfo.AclBytesInUse + Len(sCurrentACE) + GetLengthSid(bUserSid(0)) - 4
   lNewACLSize = sACLInfo.AclBytesInUse + Len(sCurrentACE) + GetLengthSid(pSID) - 4
   'lNewACLSize = sACLInfo.AclBytesInUse + Len(sCurrentACE) - 4

   ' Resize our new ACL buffer to its proper size.
   ReDim bNewACL(lNewACLSize)

   ' Use the InitializeAcl API function call to initialize the new
   ' ACL.
   lResult = InitializeAcl(bNewACL(0), lNewACLSize, ACL_REVISION2)

   ' A return code of zero means the call failed; test for this
   ' before continuing.
   If (lResult = 0) Then
      Debug.Print "Error: Unable to Initialize New ACL"
      Exit Sub
   End If

   ' If a DACL is present, copy it to a new DACL.
   If (lDaclPresent) Then

      ' Copy the ACEs from the file to the new ACL.
      If (sACLInfo.AceCount > 0) Then
         ' Grab each ACE and stuff them into the new ACL.
         For I = 0 To (sACLInfo.AceCount - 1)
            ' Attempt to grab the next ACE.
            lResult = GetAce(pAcl, I, pCurrentAce)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
               Debug.Print "Error: Unable to Obtain ACE (" & I & ")"
               Exit Sub
            End If

            ' You have a pointer to the ACE. Place it into a structure, so you can get at its size.
            CopyMemory sCurrentACE, pCurrentAce, LenB(sCurrentACE)

            ' Now that you have the ACE, add it to the new ACL.
            ' *************** KILLS APP!!!!!!
            lResult = AddAce2(VarPtr(bNewACL(0)), ACL_REVISION, MAXDWORD, VarPtr(pCurrentAce), VarPtr(sCurrentACE.Header.AceSize))
            'lResult = AddAce(bNewACL(0), ACL_REVISION, MAXDWORD, pCurrentAce, sCurrentACE.Header.AceSize)

            ' Make sure you have the current ACE under question.
            If (lResult = 0) Then
               Debug.Print "Error: Unable to Add ACE to New ACL"
               Exit Sub
            End If
         Next I

         ' You have now rebuilt a new ACL and want to add it to the newly created DACL.
         lResult = AddAccessAllowedAce(bNewACL(0), ACL_REVISION2, GENERIC_READ, bUserSid(0))

         ' Make sure added the ACL to the DACL.
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Add ACL to DACL"
            Exit Sub
         End If

         ' Set the file's Security Descriptor to the new DACL.
         lResult = SetSecurityDescriptorDacl(VarPtr(sNewSD), 1, bNewACL(0), 0)

         ' Make sure you set the SD to the new DACL.
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Set New DACL to Security Descriptor"
            Exit Sub
         End If

         ' The final step is to add the Security Descriptor back to the file!
         lResult = SetFileSecurity(sFilename, DACL_SECURITY_INFORMATION, VarPtr(sNewSD))

         ' Make sure you added the Security Descriptor to the file!
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Set New Security Descriptor to File : " & sFilename
         Else
            Debug.Print "Updated Security Descriptor on File: " & sFilename
         End If

         ' Finally, show the current ACE count for this file.
         lResult = GetFileSecurityN(sFilename, DACL_SECURITY_INFORMATION, 0, 0, lSizeNeeded)

         ' Redimension the Security Descriptor buffer variable to the proper size.
         ReDim bSDBuf(lSizeNeeded)

         ' Now get the actual Security Descriptor for the file.
         lResult = GetFileSecurity(sFilename, DACL_SECURITY_INFORMATION, bSDBuf(0), lSizeNeeded, lSizeNeeded)

         ' Make sure you've got the SD for this file!
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Get the File Security Descriptor"
            Exit Sub
         End If

         ' Now grab the DACL once more...
         lResult = GetSecurityDescriptorDacl(bSDBuf(0), lDaclPresent, pAcl, lDaclDefaulted)

         ' A return code of zero means the call failed; test for this before continuing.
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Get DACL from File Security Descriptor"
            Exit Sub
         End If

         ' Once again, grab the ACL for this file.
         lResult = GetAclInformation(pAcl, sACLInfo, Len(sACLInfo), 2&)

         ' A return code of zero means the call failed; test for this before continuing.
         If (lResult = 0) Then
            Debug.Print "Error: Unable to Get ACL from File Security Descriptor"
            Exit Sub
         End If

         ' Now show the new ACE count for this file SD!!!
         Debug.Print "ACE Count for this File Is: " & sACLInfo.AceCount
      End If
   End If
End Sub


Private Sub CmdOK_Click()
    Dim retval As Boolean
    retval = AlterPermissions(txtPath)
End Sub



