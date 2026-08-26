Attribute VB_Name = "Module1"
Option Explicit
'
' Define ADS_RIGHTS_ENUM constants:
'
Public Const ADS_RIGHT_DELETE = &H10000
Public Const ADS_RIGHT_READ_CONTROL = &H20000
Public Const ADS_RIGHT_WRITE_DAC = &H40000
Public Const ADS_RIGHT_WRITE_OWNER = &H80000
Public Const ADS_RIGHT_SYNCHRONIZE = &H100000
Public Const ADS_RIGHT_ACCESS_SYSTEM_SECURITY = &H1000000
Public Const ADS_RIGHT_GENERIC_READ = &H80000000
Public Const ADS_RIGHT_GENERIC_WRITE = &H40000000
Public Const ADS_RIGHT_GENERIC_EXECUTE = &H20000000
Public Const ADS_RIGHT_GENERIC_ALL = &H10000000
Public Const ADS_RIGHT_DS_CREATE_CHILD = &H1
Public Const ADS_RIGHT_DS_DELETE_CHILD = &H2
Public Const ADS_RIGHT_ACTRL_DS_LIST = &H4
Public Const ADS_RIGHT_DS_SELF = &H8
Public Const ADS_RIGHT_DS_READ_PROP = &H10
Public Const ADS_RIGHT_DS_WRITE_PROP = &H20
Public Const ADS_RIGHT_DS_DELETE_TREE = &H40
Public Const ADS_RIGHT_DS_LIST_OBJECT = &H80
Public Const ADS_RIGHT_DS_CONTROL_ACCESS = &H100
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ADS_PATHTYPE_ENUM constants
Public Const ADS_PATH_FILE = 1
Public Const ADS_PATH_FILESHARE = 2
Public Const ADS_PATH_REGISTRY = 3
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ADS_SD_FORMAT constants
Public Const ADS_SD_FORMAT_IID = 1
Public Const ADS_SD_FORMAT_RAW = 2
Public Const ADS_SD_FORMAT_HEXSTRING = 3
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' Ace Type definitions
'
Public Const ADS_ACETYPE_ACCESS_ALLOWED = 0
Public Const ADS_ACETYPE_ACCESS_DENIED = &H1
Public Const ADS_ACETYPE_SYSTEM_AUDIT = &H2
Public Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
Public Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Public Const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT = &H7
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' Ace Flag Public Constants
'
Public Const ADS_ACEFLAG_UNKNOWN = &H1
Public Const ADS_ACEFLAG_INHERIT_ACE = &H2                ' Permission applies to child objects also
Public Const ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE = &H4
Public Const ADS_ACEFLAG_INHERIT_ONLY_ACE = &H8
Public Const ADS_ACEFLAG_INHERITED_ACE = &H10             ' Permission is inherited from the parent
Public Const ADS_ACEFLAG_VALID_INHERIT_FLAGS = &H1F
Public Const ADS_ACEFLAG_SUCCESSFUL_ACCESS = &H40
Public Const ADS_ACEFLAG_FAILED_ACCESS = &H80
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Flags Public Constants for AD objects
'
Public Const ADS_FLAG_OBJECT_TYPE_PRESENT = &H1
Public Const ADS_FLAG_INHERITED_OBJECT_TYPE_PRESENT = &H2
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
' From WinNT.h
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' File Specific Access Rights
'
Public Const DELETE = &H10000
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL = &HFFFF

Public STANDARD_RIGHTS_READ As Long
Public STANDARD_RIGHTS_WRITE As Long
Public STANDARD_RIGHTS_EXECUTE As Long

'
' AccessSystemAcl access type
'

Public Const ACCESS_SYSTEM_SECURITY = &H1000000

'
' MaximumAllowed access type
'

Public Const MAXIMUM_ALLOWED = &H2000000

'
' These are the generic rights.
'

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000           ' FULL CONTROL

'
' AccessMask constants for FILE ACEs
'
Public Const FILE_READ_DATA = &H1               ' file & pipe
Public Const FILE_LIST_DIRECTORY = &H1          ' directory
Public Const FILE_WRITE_DATA = &H2              ' file & pipe
Public Const FILE_ADD_FILE = &H2                ' directory
Public Const FILE_APPEND_DATA = &H4             ' file
Public Const FILE_ADD_SUBDIRECTORY = &H4        ' directory
Public Const FILE_CREATE_PIPE_INSTANCE = &H4    ' named pipe
Public Const FILE_READ_EA = &H8                 ' file & directory
Public Const FILE_WRITE_EA = &H10               ' file & directory
Public Const FILE_EXECUTE = &H20                ' file
Public Const FILE_TRAVERSE = &H20               ' directory
Public Const FILE_DELETE_CHILD = &H40           ' directory
Public Const FILE_READ_ATTRIBUTES = &H80        ' all
Public Const FILE_WRITE_ATTRIBUTES = &H100      ' all

Public FILE_ALL_ACCESS As Long
Public FILE_GENERIC_READ As Long
Public FILE_GENERIC_WRITE As Long
Public FILE_GENERIC_EXECUTE As Long
Public FILE_GENERIC_MODIFY As Long
Public FILE_GENERIC_LIST As Long

Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_DELETE = &H4
'
' AceFlags values for files
'
Public Const OBJECT_INHERIT_ACE = &H1
Public Const CONTAINER_INHERIT_ACE = &H2
Public Const NO_PROPAGATE_INHERIT_ACE = &H4
Public Const INHERIT_ONLY_ACE = &H8
Public Const INHERITED_ACE = &H10
'
'
'
Public Const SE_DACL_PROTECTED = &H1000

Sub Main()
    STANDARD_RIGHTS_READ = READ_CONTROL
    STANDARD_RIGHTS_WRITE = READ_CONTROL
    STANDARD_RIGHTS_EXECUTE = READ_CONTROL
    
    FILE_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF

    FILE_GENERIC_READ = STANDARD_RIGHTS_READ Or _
                        FILE_READ_DATA Or _
                        FILE_READ_ATTRIBUTES Or _
                        FILE_READ_EA Or _
                        SYNCHRONIZE
    
    FILE_GENERIC_WRITE = STANDARD_RIGHTS_WRITE Or _
                         FILE_WRITE_DATA Or _
                         FILE_WRITE_ATTRIBUTES Or _
                         FILE_WRITE_EA Or _
                         FILE_APPEND_DATA Or _
                         SYNCHRONIZE
    
    FILE_GENERIC_EXECUTE = STANDARD_RIGHTS_EXECUTE Or _
                           FILE_READ_ATTRIBUTES Or _
                           FILE_EXECUTE Or _
                           SYNCHRONIZE
    
    FILE_GENERIC_MODIFY = STANDARD_RIGHTS_READ Or _
                          STANDARD_RIGHTS_WRITE Or _
                          FILE_TRAVERSE Or _
                          FILE_LIST_DIRECTORY Or _
                          FILE_READ_ATTRIBUTES Or _
                          FILE_READ_EA Or _
                          FILE_ADD_FILE Or _
                          FILE_ADD_SUBDIRECTORY Or _
                          FILE_WRITE_ATTRIBUTES Or _
                          FILE_WRITE_EA Or _
                          DELETE Or _
                          ADS_RIGHT_READ_CONTROL Or _
                          SYNCHRONIZE
    
    FILE_GENERIC_LIST = FILE_TRAVERSE Or _
                        FILE_LIST_DIRECTORY Or _
                        FILE_READ_ATTRIBUTES Or _
                        FILE_READ_EA Or _
                        ADS_RIGHT_READ_CONTROL

    Load frmMain
End Sub
