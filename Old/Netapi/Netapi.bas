Attribute VB_Name = "Module1"
    Option Explicit
 
   Public Type NETRESOURCE
           dwScope As Long
           dwType As Long
           dwDisplayType As Long
           dwUsage As Long
           lpLocalName As Long
           lpRemoteName As Long
           lpComment As Long
           lpProvider As Long
   End Type
 
   Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
       "WNetOpenEnumA" ( _
       ByVal dwScope As Long, _
       ByVal dwType As Long, _
       ByVal dwUsage As Long, _
       lpNetResource As Any, _
       lphEnum As Long) As Long
 
   Public Declare Function WNetEnumResource Lib "mpr.dll" Alias _
       "WNetEnumResourceA" ( _
       ByVal hEnum As Long, _
       lpcCount As Long, _
       ByVal lpBuffer As Long, _
       lpBufferSize As Long) As Long
 
   Public Declare Function WNetCloseEnum Lib "mpr.dll" ( _
       ByVal hEnum As Long) As Long
 
   ' RESOURCE ENUMERATION
   Public Const RESOURCE_CONNECTED = &H1
   Public Const RESOURCE_GLOBALNET = &H2
   Public Const RESOURCE_REMEMBERED = &H3
 
   Public Const RESOURCETYPE_ANY = &H0
   Public Const RESOURCETYPE_DISK = &H1
   Public Const RESOURCETYPE_PRINT = &H2
   Public Const RESOURCETYPE_UNKNOWN = &HFFFF
 
   Public Const RESOURCEUSAGE_CONNECTABLE = &H1
   Public Const RESOURCEUSAGE_CONTAINER = &H2
   Public Const RESOURCEUSAGE_RESERVED = &H80000000

Sub main()
    frmNetAPI.Show
End Sub
