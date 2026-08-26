Attribute VB_Name = "modMain"
Option Explicit

' Constants
Public Const SV_TYPE_SQLSERVER = 4
Public Const NERR_Success = 0

' Type
Public Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_name As Long
End Type

' Declares for API & Kernel calls
Public Declare Function NetServerEnum Lib "Netapi32.dll" _
    (ByVal serverName As String, _
    ByVal level As Long, _
    ByRef lBufPtr As Long, _
    ByRef prefMaxLen As Long, _
    entriesRead As Long, _
    totalEntries As Long, _
    ByVal serverType As Long, _
    ByVal domain As String, _
    resume_handle As Long) As Long
    ' Important: you MUST pass prefMaxLen by REFERENCE
    
Public Declare Sub RtlMoveMemory _
    Lib "kernel32" ( _
    Dest As Any, _
    Vsrc As Any, _
    ByVal lSize&)

Public Declare Sub lstrcpyW Lib "kernel32" _
    (vDest As Any, ByVal sSrc As Any)

Public Declare Function NetApiBufferFree Lib "Netapi32.dll" _
    (ByVal lpBuffer As Long) As Long


Public Function EnumSQLServers() As Variant
'
' Author: Scott McNair, Raymond James Consulting
'   Note: I'm NOT a VB programmer, I'm a DBA, so...
'   All you real programmers have fun!
'
'   I added all the documentation so that others could get a better
'   understanding of how all this actually works.
'   If you delete all the comments, there is only about
'   35 lines of code, making this a very "compact" function.
' --------------------------------------------------------------------
' I owe a lot of thanks to:
'   Andy Doran: whose original example got me started in the right direction
'       http://dspace.dial.pipex.com/andy.doran/
'   Brian Hensel from Raymond James Consutling, who enhanced Andy's
'       example to make it easier/simpler.'
' --------------------------------------------------------------------
' This function uses the API call NetSeverEnum (Netapi32.dll)
' to retrieve a list of SQL Servers
' - It will only list those SQL Server using NAMED-PIPES, and,
' - It will ONLY work from an NT workstation/server.
' It will query the Current/Default Domain Controller to get the list
'
' The API function NetServerEnum returns a list in an array
' that is identified by a pointer (lBuf).
' Since VB doesn't inherantly use "pointers" as does C,
' you have to use the value of lBuf and pass it to other
' Kernel32 calls to manipulate the data/memory.
'
' Through my investigation/disection, I found that:
' The NetServerEnum API call actually returns TWO arrays:
' 1) lbuf = POINTER to a array of "byteWs" that holds a list consisting of
'  "Numread" occurances of "SrvList" that are a user defined type/struct of:
'           sv100_platform_id As Long
'           sv100_name As Long
' 2) The second array is a "btyeW" list of the server NAMES in UNICODE
'   format.
' You get to the second array by using the value stored in sv100_name.
' This is a POINTER (memory address) to the begining element of
' each of the Server names, followed by a "null" to signify the end of that
' Server name
' ----------------------------------------------------------------------
' Function return value:
'   We return a variant array so that the caller can use the
'   inherant VB functions to enum the list when it is returned.
' ----------------------------------------------------------------------
Dim serverName As String    ' Server to run the query on. If left as
                            ' a null, will run from the local and query
                            ' the current sDomain controller
Dim lBuf As Long            ' a number that is the acutal memory address
                            ' of the returned info
Dim prefMaxLen As Long      ' DON'T Set! A "0" means give me everything
                            ' on the first call
Dim NumRead As Long         ' Number of SQL server names RETURNED in lBuf
Dim NumSrvrs As Long        ' Number of SQL servers that EXIST
Dim sDomain As String       ' sDomain to check for servers
Dim vRsm As Variant         ' "Resume-Handle" [not used]
Dim rslt As Long            ' error #/result of the API call

Dim vSqlList() As Variant   ' variant to hold the list we create
Dim lBufTmp As Long         ' 2nd pointer to be used for manipulation
Dim bBuffer(99) As Byte     ' temporary array to hold the Sql Server name in
                            ' UNICODE format
Dim sTmp As String          ' Temporary string for char conversion from
                            ' a "byteW"(UNICODE) to a string/byte (ANSI)
Dim SrvList As SERVER_INFO_100  ' This is a usertype/struct that holds
                                ' the pointers to the specific info we want.
Dim x As Integer
Dim y As Integer

Const SV_TYPE_SQLSERVER = 4 ' Defined in the APIText Viewer Constants
Const NERR_Success = 0

' Call for the list
' NetServerEnum is an API call from Netapi32.dll
' Declare this API call in a module
rslt = NetServerEnum(vbNullString, 100, lBuf, prefMaxLen, _
    NumRead, NumSrvrs, SV_TYPE_SQLSERVER, vbNullString, vRsm)
    
' If it don't work, bail
If (rslt <> NERR_Success) Then
    GoTo Exit_Func
End If

' You CAN'T modify lbuf, it is appearantly locked by the OS,
' So, copy the contents of lBuf to lBufTmp
' lBuf holds a LONG that is an Memory ADDRESS to a byteW array
lBufTmp = lBuf

' Loop through the lBuf Array and get the Server Names
' from the second array
For x = 0 To NumRead - 1
    ' redim the return/result variant(array) to hold one more item
    ReDim Preserve vSqlList(x)
    
    ' truncate the temp string
    sTmp = vbNullString
    
    ' Fill the SrvList struct with data from the buffer returned from
    ' the API call
    ' Note the "byVal" usage.
    ' The Rtl API call expects you to pass a pointer/memory address,
    ' by specifying "byval", it "tricks" the API call into using
    ' the number stored in the lBufTmp var as a memory address,
    ' instead of using the address of lBufTmp.
    RtlMoveMemory SrvList, ByVal lBufTmp, Len(SrvList)
                
    ' SqlSrvr.Name is a *ptr to a byteW array
    ' copy the bytes to a temporary byteW array for translation.
    lstrcpyW bBuffer(0), SrvList.sv100_name
    
    ' Now convert the "unicode bytes" to a string to get the name.
    ' If Buffer(y) = 0, the "0" is a string terminator,
    ' meaning end of this Server Name
    y = 0
    Do While bBuffer(y) <> 0
       sTmp = sTmp & Chr$(bBuffer(y))
        y = y + 2
    Loop

    ' Add the string to the Variant(array)
    vSqlList(x) = sTmp
    
    ' Increment lbuftmp to point to the NEXT SrvList item
    ' in the list returned by the API call.
    lBufTmp = lBufTmp + Len(SrvList)

Next x

Exit_Func:
    ' Avoid Memory leaks!
    If lBuf Then NetApiBufferFree (lBuf)
    ' Return the list to the caller
    EnumSQLServers = vSqlList

End Function

Public Function EnumSQLShort() As Variant
'
' This is the same as EnumSQLServers but with no documentation.
Dim serverName As String
Dim lBuf As Long
Dim prefMaxLen As Long
Dim NumRead As Long
Dim NumSrvrs As Long
Dim sDomain As String
Dim vRsm As Variant
Dim rslt As Long
Dim vSqlList() As Variant
Dim lBufTmp As Long
Dim bBuffer(99) As Byte
Dim sTmp As String
Dim SrvList As SERVER_INFO_100
Dim x As Integer
Dim y As Integer

rslt = NetServerEnum(vbNullString, 100, lBuf, prefMaxLen, _
    NumRead, NumSrvrs, SV_TYPE_SQLSERVER, vbNullString, vRsm)
If (rslt <> NERR_Success) Then
    GoTo Exit_Func
End If

lBufTmp = lBuf
For x = 0 To NumRead - 1
    ReDim Preserve vSqlList(x)
    sTmp = vbNullString
    RtlMoveMemory SrvList, ByVal lBufTmp, Len(SrvList)
    lstrcpyW bBuffer(0), SrvList.sv100_name
    y = 0
    Do While bBuffer(y) <> 0
       sTmp = sTmp & Chr$(bBuffer(y))
        y = y + 2
    Loop
    vSqlList(x) = sTmp
    lBufTmp = lBufTmp + Len(SrvList)
Next x

Exit_Func:
    ' Avoid Memory leaks!
    If lBuf Then NetApiBufferFree (lBuf)
    EnumSQLShort = vSqlList
End Function
