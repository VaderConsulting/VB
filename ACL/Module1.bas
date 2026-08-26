Attribute VB_Name = "modACL"
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
Public Const ADS_ACEFLAG_INHERIT_ACE = &H2
Public Const ADS_ACEFLAG_NO_PROPAGATE_INHERIT_ACE = &H4
Public Const ADS_ACEFLAG_INHERIT_ONLY_ACE = &H8
Public Const ADS_ACEFLAG_INHERITED_ACE = &H10
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

STANDARD_RIGHTS_READ = READ_CONTROL
STANDARD_RIGHTS_WRITE = READ_CONTROL
STANDARD_RIGHTS_EXECUTE = READ_CONTROL
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
Public Const GENERIC_ALL = &H10000000

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

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Sub to reorder an ACL
' Comments in the subroutine explain how the ACL should be ordered.
' The IADsAccessControlList::AddAce method makes not attempt to properly
' order the ACE being added.
'
Sub ReorderDacl(dacl) ':( Missing Scope
    Dim sdutil As New adssecurityutility
  '
  ' Initialize all of the new ACLs
  '
  ' VBS methods of creating the ACL bins
  '

    Set newdacl = CreateObject("AccessControlList")
    Set ImpDenyDacl = CreateObject("AccessControlList")
    Set InheritedDacl = CreateObject("AccessControlList")
    Set ImpAllowDacl = CreateObject("AccessControlList")
    Set InhAllowDacl = CreateObject("AccessControlList")
    Set ImpDenyObjectDacl = CreateObject("AccessControlList")
    Set impAllowObjectDacl = CreateObject("AccessControlList")
    '
    ' Sift the DACL into 5 bins:
    ' Inherited Aces
    ' Implicit Deny Aces
    ' Implicit Deny Object Aces
    ' Implicit Allow Aces
    ' Implicit Allow object aces
    '
    For Each ace In dacl
        '
        ' Sort the orignal ACEs into thier appropriate
        ' ACLs
        '
        If ((ace.AceFlags And ADS_ACEFLAG_INHERITED_ACE) = ADS_ACEFLAG_INHERITED_ACE) Then
            '
            ' Don't really care about the order of inherited aces. Since we are
            ' adding them to the top of a new list, when they are added back
            ' to the Dacl for the object, they will be in the same order as
            ' they were orginally. Just a positive side affect of adding items
            ' of a LIFO ( Last In First Out) type list.
            '
            InheritedDacl.AddAce ace
          Else 'NOT ((ACE.ACEFLAGS...
            '
            ' We have an Implicit ACE, lets put it the proper pool
            '
            Select Case ace.AceType
              Case ADS_ACETYPE_ACCESS_ALLOWED
                '
                ' We have an implicit allow ace
                '
                ImpAllowDacl.AddAce ace
              Case ADS_ACETYPE_ACCESS_DENIED
                '
                ' We have a implicit Deny ACE
                '
                ImpDenyDacl.AddAce ace
              Case ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
                '
                ' We have an object allowed ace
                ' Does it apply to a property? or an Object?
                '
                impAllowObjectDacl.AddAce ace
              Case ADS_ACETYPE_ACCESS_DENIED_OBJECT
                '
                ' We have a object Deny ace
                '
                ImpDenyObjectDacl.AddAce ace
              Case Else
                '
                ' Missed a bin?
                '
            End Select
        End If
    Next ':( Repeat For-Variable: ACE
    '
    ' Combine the ACEs in the proper order
    ' Implicit Deny
    ' Implicit Deny Object
    ' Implicit Allow
    ' Implicit Allow Object
    ' Inherited aces
    '
    ' Implicit Deny
    '
    For Each ace In ImpDenyDacl
        newdacl.AddAce ace
    Next ':( Repeat For-Variable: ACE
    '
    ' Implicit Deny Object
    '
    For Each ace In ImpDenyObjectDacl
        newdacl.AddAce ace
    Next ':( Repeat For-Variable: ACE
    '
    ' Implicit Allow
    '
    For Each ace In ImpAllowDacl
        newdacl.AddAce ace
    Next ':( Repeat For-Variable: ACE
    '
    ' Implicit Allow Object
    '
    For Each ace In impAllowObjectDacl
        newdacl.AddAce ace
    Next ':( Repeat For-Variable: ACE
    '
    ' Inherited Aces
    '
    For Each ace In InheritedDacl
        newdacl.AddAce ace
    Next ':( Repeat For-Variable: ACE
    '
    ' Clean up
    '
    Set InheritedDacl = Nothing
    Set ImpAllowDacl = Nothing
    Set ImpDenyObjectDacl = Nothing
    Set ImpDenyDacl = Nothing
    '
    ' Set the appropriate revision level
    ' for the DACL
    '
    newdacl.AclRevision = dacl.AclRevision
    '
    ' Replace the Security Discriptor
    '
    Set dacl = Nothing
    Set dacl = newdacl

End Sub

'==========================================================================
'
'<<<<<<<<<<<<<<<<<<<<<<<<< Begin IADsSecurity Constants >>>>>>>>>>>>
'
' Supported on Win2k with ADsSecurity.dll registered on the client.
' NOTE: ADsSecurity.dll does not work properly on XP, use ADsSecurityUTility object
' instead.
'
'
'++++++++++++++++++ MAIN SCRIPT ++++++++++++++++++++++++++++++++++
'
Dim oAce ' variable for the new ACE':( Missing Scope
Dim oDacl ' variable for the DACL of the object':( Missing Scope
Dim oSD ' variable for the Security Descriptor of the object':( Missing Scope
Dim oADsSecurityUtility ' variable for the ADsSecurity object':( Missing Scope
'
' Create an ACE object
' and an IADsSecurity object
'
Set oAce = CreateObject("AccessControlEntry")
Set oADsSecurityUtility = CreateObject("ADsSecurity")
'
' Retreive the Security Descriptor for the given NTFS File path
' Using ADsSecurity.dll
'
Set oSD = oADsSecurityUtility.GetSecurityDescriptor("FILE://c:\testdir\test.txt")
WScript.Echo Hex$(oSD.Control)
'
' Setting the SE_DACL_PROTECTED bit causes the DACL not to inherit
'
oSD.Control = oSD.Control Or SE_DACL_PROTECTED
WScript.Echo Hex$(oSD.Control)
'
' Retrieve the Discretionary ACL for the Key
'
Set oDacl = oSD.DiscretionaryACL
'
' I used NewDacl originall, caused a major problem with the reorder
' routine. Had to change the name, NewSnacl was unique....
'
Set NewSnacl = CreateObject("AccessControlList")
For Each oAce In oDacl
    '
    ' The INHERITED_ACE bit is the one that tells
    If (oAce.AceFlags And INHERITED_ACE) Then
        oAce.AceFlags = oAce.AceFlags Xor INHERITED_ACE
    End If
    NewSnacl.AddAce oAce
Next ':( Repeat For-Variable: OACE
NewSnacl.AclRevision = oDacl.AclRevision
ReorderDacl NewSnacl
'
' Place the DACL back onto the SD
'
oSD.DiscretionaryACL = NewSnacl
oADsSecurityUtility.SetSecurityDescriptor oSD
'
' Clean up
'
Set oAce = Nothing
Set oDacl = Nothing
Set oSD = Nothing
Set oADsSecurityUtility = Nothing
WScript.Echo "<<<<<<<<<<<<< Done >>>>>>>>>>>>>>>>"

':) Ulli's VB Code Formatter V2.16.6 (2003-Oct-07 08:33) 164 + 201 = 365 Lines
