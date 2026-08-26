Attribute VB_Name = "modMain"
Option Explicit

' Enumerate each persistent connection.
' For each connection, separate the server and share names
' Then open a config file, and cross-reference to see if the server is listed in the config file
' If found, replace the current (Old) Server with the new (Replacement) Server

' Initial Version D. Robinson 2003
'
' V   DATE          Author              History
' 0.1 11 Feb 2003   D. Robinson
' 0.2 12 Feb 2003   D. Robinson         Add Printer stuff as FUTURE Enhancements
' 0.3 12 Feb 2003   D. Robinson         Modify to remove WMI dependencies.  Add Win32 API code instead.  Remove V0.2 enhancements (WSH based)


Const HKEY_USERS = &H80000003

Dim strApp_path As String
Dim bDebug As Boolean

'---------------------------------------------------------------------------------------
' Procedure : Main
' DateTime  : 01-04-2003 11:13
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Sub Main()
    Dim wshNetwork As Object
    Dim oDrives As Object
    Dim strSharePath As String
    Dim strServer As String
    Dim strShare As String
    Dim strNewServer As String
    Dim datRunStart As Date
    Dim lngRunTime As Long
    Dim bResult As Boolean
    Dim intLoop As Integer
    Dim strDriveLetter As String
    Dim strUsername As String
    
    datRunStart = Now
    
    ' Define Application Path.  Don't use app.path as there is sometimes an added backslash.
    ' This ensures consistency
    If Right(App.Path, 1) = "\" Then
        strApp_path = App.Path
    Else
        strApp_path = App.Path & "\"
    End If
    
    strUsername = Environ$("USERNAME")
    
    If LCase(Command$) = "debug" Then
        bDebug = True
        If Dir(strApp_path & "debug.txt") <> "" Then
            Kill strApp_path & "debug.txt"
        End If
    End If
    
    Set wshNetwork = CreateObject("Wscript.Network")
    
    Set oDrives = wshNetwork.EnumNetworkDrives
    For intLoop = 0 To oDrives.Count - 1 Step 2
        strDriveLetter = oDrives.Item(intLoop)
        strSharePath = oDrives.Item(intLoop + 1)
        If strDriveLetter <> "" Then
            ExpandPath strSharePath, strServer, strShare
            If bDebug Then LogEvent "Drive: " & strDriveLetter & vbTab & "Server: " & strServer & vbTab & "Share: " & strShare
            bResult = GetReplacementServer(strServer, strNewServer)
            If bResult = True And strNewServer <> "" Then
                If bDebug Then LogEvent "Attempting disconnect"
                DisConnect2 strDriveLetter, True
                If bDebug Then LogEvent "Attempting connect"
                Connect2 strDriveLetter, "\\" & strNewServer & "\" & strShare, strUsername, ""
            End If
        End If
    Next
     
    Set wshNetwork = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ExpandPath
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Function ExpandPath(strFullPath, strServer, strShare)
    Dim i, j
    i = InStr(3, strFullPath, "\")
    strServer = Mid(strFullPath, 3, i - 3)
    strShare = Mid(strFullPath, i + 1, 256)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetReplacementServer
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Function GetReplacementServer(strOldServer As String, strReplacementServer As String) As Boolean
    Dim strServerA As String
    Dim strServerB As String
    Dim strFileLine As String
    Dim strServers() As String
    Dim intPos As Integer
    Dim intCount As Integer
    Dim strFilename As String
    
    strFilename = strApp_path & "SwapServer.ini"
    
    ' Bail out if the config file doesn't exist
    If Dir(strFilename) = "" Then
        GetReplacementServer = False
        If bDebug Then LogEvent "ERROR!  No Config file found (" & strFilename & ")"
        Exit Function
    End If
    
    ' Open config file
    Open strFilename For Input As #1
    
    ' Loop through the file
    Do Until EOF(1)
        ' Get strServerA and strServerB.  They will be a single variable
        Line Input #1, strFileLine
        
        ' Remove leading and trailing spaces
        strFileLine = Trim(strFileLine)
        
        ' Ensure this line is valid, ie not a comment.
        ' A comment can start with ' or // or #
        If Left(strFileLine, 1) <> "'" And Left(strFileLine, 2) <> "//" And Left(strFileLine, 1) <> "#" And strFileLine <> "" Then
            ' Remove double backslashes
            strFileLine = Replace(strFileLine, "\\", "")
            strFileLine = Replace(strFileLine, vbTab, "    ")
            
            ' Remove comments
            If InStr(1, strFileLine, "\\") > 0 Or InStr(1, strFileLine, "'") > 0 Or InStr(1, strFileLine, "#") > 0 Or InStr(1, strFileLine, ";") > 0 Then
                intPos = InStr(1, strFileLine, "\\")
                If intPos > 0 Then
                    strFileLine = Left(strFileLine, intPos - 1)
                End If
                
                intPos = InStr(1, strFileLine, "'")
                If intPos > 0 Then
                    strFileLine = Left(strFileLine, intPos - 1)
                End If
                
                intPos = InStr(1, strFileLine, "#")
                If intPos > 0 Then
                    strFileLine = Left(strFileLine, intPos - 1)
                End If
                
                intPos = InStr(1, strFileLine, ";")
                If intPos > 0 Then
                    strFileLine = Left(strFileLine, intPos - 1)
                End If
            End If
            ' Split the listed servers into an array
            strServers() = Split(strFileLine, ",")
            
            intCount = intCount + 1
            
            ' Extract each server, removing leading and trailing spaces
            strServerA = Trim(strServers(0))
            strServerB = Trim(strServers(1))
            
            If LCase(strServerA) = LCase(strOldServer) Then      ' Match found.
                If bDebug Then LogEvent "MATCH! Found '" & strServerA & "'.  Replacement = '" & strServerB & "'."
                strReplacementServer = strServerB
                GetReplacementServer = True
                Close 1
                Exit Function
            End If
        End If
        ' Continue through until no more lines
    Loop
    ' Done
    Close 1
    ' If we get this far, there was no match
    If bDebug Then LogEvent "No match found for '" & strOldServer & "'"
End Function

'---------------------------------------------------------------------------------------
' Procedure : LogEvent
' DateTime  : 01-04-2003 11:14
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub LogEvent(strMessage As String)
    Open strApp_path & "debug.txt" For Append As #2
        Print #2, strMessage
    Close 2
End Sub

