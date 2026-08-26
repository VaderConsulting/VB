Attribute VB_Name = "Module1"
Option Explicit

Public Type vbPrinter
    OldServer As String
    OldQueue As String
    NewServer As String
    NewQueue As String
    bRemove As Boolean
    bReplace As Boolean
End Type

Public Const intNoOfTasks = 6                               ' Total number of tasks

Public oEvents As New Collection
Public oDictionary As Scripting.Dictionary
Public strApp_Path As String
Public strAppStatus As String
Public PrelogonApplications As Collection
Public PostlogonApplications As Collection

Public bDebugMode As Boolean
Public strUsername As String
Public strDomainname As String
Public strComputername As String
Public strComputerModel As String
Public strComputerOS As String
Public bIsTerminalServer As Boolean
Public strCentral As String
Public strSitename As String
Public strUserPath As String
Public strCommsMediaPolicyTitle As String
Public strMediaPolicyURL As String

Public bReportErrors As Boolean
Public bReportWarnings As Boolean
Public bReportInformation As Boolean
Public bReportToEventLog As Boolean

Public bWriteDebugMessagesToEventLog As Boolean


Public Enum EventLogError
    Error = vbLogEventTypeError             ' 1
    Warning = vbLogEventTypeWarning         ' 2
    Information = vbLogEventTypeInformation ' 4
    DebugInfo = 8
End Enum

Public Enum ApplicationStatus
    statusUnknown = 0
    StatusError = 1
    statusWarning = 2
    statusOK = 4
End Enum

'---------------------------------------------------------------------------------------
' Procedure : LogEvent
' DateTime  : 14-03-2003 10:34
' Author    : Dave Robinson
' Purpose   : Keep track of application events
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub LogEvent(strMessage As String, Optional EventType As EventLogError = Information)
    Dim bDoReport As Boolean
    
    On Error GoTo LogEvent_Error

    If bReportErrors And EventType = Error Then
        bDoReport = True
    End If
    
    If bReportWarnings And EventType = Warning Then
        bDoReport = True
    End If
    
    If bReportInformation And EventType = Information Then
        bDoReport = True
    End If
    
    If bDebugMode And EventType = DebugInfo Then
        bDoReport = True
    End If
    
    Debug.Print strMessage
    
    ' Only add this event if we have determined it should be (according to the config options)
    If bDoReport Then
        oEvents.Add Format(Now, "DD/MM/YYYY|HH:NN:SS AM/PM") & "|" & EventType & "|" & strMessage
        If (EventType <> Information) And EventType <> DebugInfo Then
            'App.LogEvent strMessage, EventType
            'frmMain.sbrStatus.SimpleText = strMessage
            SetStatus EventType
        ElseIf EventType = DebugInfo Then
            If bWriteDebugMessagesToEventLog Then
                'App.LogEvent strMessage, vbLogEventTypeInformation
                'frmMain.sbrStatus.SimpleText = strMessage
            End If
            SetStatus statusUnknown
        End If
    End If
    
    frmMain.Refresh

    On Error GoTo 0
    Exit Sub

LogEvent_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure LogEvent of Module modGlobal", Error
End Sub












'---------------------------------------------------------------------------------------
' Procedure : OpenPrinterList
' DateTime  : 12/03/2003 22:49
' Author    : Dave Robinson
' Purpose   : Open the Substitutes.xml file, storing all info into the supplied Dictionary object
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub OpenPrinterList(PrinterList() As vbPrinter)
    Dim oXMLDocument As MSXML2.DOMDocument
    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
    Dim oXMLNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode1 As MSXML2.IXMLDOMNode
    Dim oXMLChildNode2 As MSXML2.IXMLDOMNode
    Dim strServerRole As String
    Dim lPrinterNo As Long
        
    On Local Error GoTo OpenConfigError
    
    LogEvent "Inside OpenPrinterList subroutine", DebugInfo
    ' Clear all items from the array before we start adding more (potentially duplicate) items to it.
    Erase PrinterList()
    
    ' Create an XML Document to load the XML file into
    Set oXMLDocument = New MSXML2.DOMDocument30
    
    ' Check to ensure the config.xml file exists
    If Dir(strApp_Path & "substitutes.xml") = "" Then
        LogEvent "Logon could not find the config file 'substitutes.xml' under the current path (" & strApp_Path & ").", Warning
    Else
        ' Load the XML file
        oXMLDocument.Load (strApp_Path & "substitutes.xml")
        ' Reference the nodes
        Set oXMLNodeList = oXMLDocument.getElementsByTagName("Printers")
        
        ' Retrieve the individual values in the configuration file
        'Set oXMLValuelist = oXMLNodeList.Item(0).childNodes
        
        ' #######################################################
        For Each oXMLNode In oXMLNodeList                             ' Printers
            For Each oXMLChildNode1 In oXMLNode.childNodes            ' Printer
                lPrinterNo = lPrinterNo + 1
                ReDim Preserve PrinterList(lPrinterNo)
                For Each oXMLChildNode2 In oXMLChildNode1.childNodes
                    If oXMLChildNode2.Text <> "" Then
                        Debug.Print "node: " & oXMLChildNode2.nodeName & " Text: " & oXMLChildNode2.Text
                        Select Case LCase(oXMLChildNode2.nodeName)
                            Case "oldserver"
                                PrinterList(lPrinterNo).OldServer = oXMLChildNode2.Text
                            Case "oldqueue"
                                PrinterList(lPrinterNo).OldQueue = oXMLChildNode2.Text
                            Case "newserver"
                                PrinterList(lPrinterNo).NewServer = oXMLChildNode2.Text
                            Case "newqueue"
                                PrinterList(lPrinterNo).NewQueue = oXMLChildNode2.Text
                        End Select
                    End If
                Next
            Next
        Next
        ' #######################################################
        ' ********************************************************
        ' Load each value and add it to a dictionary of values
        'For Each oXMLNode In oXMLValuelist
        '    lPrinterNo = lPrinterNo + 1
        '    ReDim Preserve PrinterList(lPrinterNo)
        '    'PrinterList(lPrinterNo).Queue = oXMLNode.baseName
        '    'PrinterList(lPrinterNo).Queue = oXMLNode.Text
        '    'Dic.Add LCase(oXMLNode.baseName), oXMLNode.Text
        '    Debug.Print LCase(oXMLNode.baseName), oXMLNode.Text
        'Next
        ' ********************************************************
    End If
    
    ' Clean up
    Set oXMLDocument = Nothing
    
    Exit Sub
    
OpenConfigError:
    LogEvent "**** Logon encountered internal error " & Err.Number & " (" & Err.Description & ") in OpenPrinterList.", Error
    
    Set oXMLDocument = Nothing
    Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : NTSetDefaultPrinter
' DateTime  : 31-03-2003 13:57
' Author    : Dave Robinson
' Purpose   : Sets the default printer under Windows NT/2000/XP
'
'  V    Date    Author      History
' 1.0   31-03-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------

Public Sub NTSetDefaultPrinter(strPrinter As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    On Error GoTo NTSetDefaultPrinter_Error

    'Get the printer information for the currently selected printer in the list. The information is taken from the WIN.INI file.
    Buffer = Space(1024)
    PrinterName = strPrinter
    r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))

    'Parse the driver name and port name out of the buffer
    GetDriverAndPort Buffer, DriverName, PrinterPort

    If DriverName <> "" And PrinterPort <> "" Then
        SetDefaultPrinter PrinterName, DriverName, PrinterPort
    End If

    On Error GoTo 0
    Exit Sub

NTSetDefaultPrinter_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure NTSetDefaultPrinter of Module modGlobal", Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetDefaultPrinter
' DateTime  : 01-04-2003 15:25
' Author    : Dave Robinson
' Purpose   : Sets the default WINDOWS Printer
'
'  V    Date    Author      History
' 1.0   01-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Private Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    On Error GoTo SetDefaultPrinter_Error

    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")

    On Error GoTo 0
    Exit Sub

SetDefaultPrinter_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure SetDefaultPrinter of Module modGlobal", Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetDriverAndPort
' DateTime  : 31-03-2003 13:56
' Author    : Dave Robinson
' Purpose   : Return the
'
'  V    Date    Author      History
' 1.0   31-03-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, PrinterPort As String)
    Dim iDriver As Integer
    Dim iPort As Integer
    On Error GoTo GetDriverAndPort_Error

    DriverName = ""
    PrinterPort = ""

    'The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

        'Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        'The port name is the second entry after the driver name
        'separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            'Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If

    On Error GoTo 0
    Exit Sub

GetDriverAndPort_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDriverAndPort of Module modGlobal", Error
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UpdateStatus
' DateTime  : 12/03/2003 22:45
' Author    : Dave Robinson
' Purpose   : Update Form with status
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub UpdateStatus(intStatus As Integer)
    On Error GoTo UpdateStatusError
        
    Select Case intStatus
        Case 1
            LogEvent "Prelogon", Information
        Case 2
            LogEvent "Media Access Policy", Information
        Case 3
            LogEvent "Checking password age", Information
        Case 4
            LogEvent "Writing sitename to registry", Information
        Case 5
            LogEvent "Printer redirection", Information
        Case 6
            LogEvent "Postlogon", Information
        Case 99
            LogEvent "Sending data to collection point", Information
            ' Set the status equal to the number of tasks
            intStatus = intNoOfTasks
    End Select
    
    'MakeTransparent Me.hWnd, 255 - ((intStatus / intNoOfTasks) * 255)
    
    'frmMain.imgImage.Picture = frmMain.imlImages.ListImages.Item(intStatus).Picture
    
    ' Update the Progressbar
    'frmMain.pbrProgress.Value = (intStatus / intNoOfTasks) * 100
    Exit Sub
UpdateStatusError:
    LogEvent "**** Logon encountered internal error " & Err.Number & " (" & Err.Description & ") in UpdateStatus.", Error
    Err.Clear
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetServerFromPath
' DateTime  : 14-03-2003 10:29
' Author    : Dave Robinson
' Purpose   : Retrieve the server portion of the given UNC
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function GetServerFromPath(strUNCPath As String) As String
    Dim intTemp1 As Integer
    Dim strPath As String
    On Error GoTo GetServerFromPath_Error

    If InStr(1, strUNCPath, "\\") = 0 Then Exit Function
    intTemp1 = InStr(3, Trim(strUNCPath), "\")
    strPath = Left(Trim(strUNCPath), intTemp1 - 1)
    GetServerFromPath = UCase(Right(strPath, Len(strPath) - 2))

    On Error GoTo 0
    Exit Function

GetServerFromPath_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure GetServerFromPath of Module modGlobal", Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPathfromUNC
' DateTime  : 14-03-2003 10:29
' Author    : Dave Robinson
' Purpose   : Retrieve the path or share or printer (etc) portion of the given UNC
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function GetPathfromUNC(strUNCPath As String) As String
    Dim intTemp1 As Integer
    Dim strPath As String
    On Error GoTo GetPathfromUNC_Error

    If InStr(1, strUNCPath, "\\") = 0 Then Exit Function
    intTemp1 = InStr(3, Trim(strUNCPath), "\")
    strPath = Right(Trim(strUNCPath), Len(Trim(strUNCPath)) - intTemp1)
    GetPathfromUNC = UCase(strPath)

    On Error GoTo 0
    Exit Function

GetPathfromUNC_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure GetPathfromUNC of Module modGlobal", Error
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetStatus
' DateTime  : 14-03-2003 10:38
' Author    : Dave Robinson
' Purpose   : Set the application status
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub SetStatus(Status As ApplicationStatus)
    On Error GoTo SetStatus_Error

    Select Case Status
        Case ApplicationStatus.statusUnknown
            strAppStatus = "Unknown or Debug Mode"
        Case ApplicationStatus.statusOK
            strAppStatus = "OK"
        Case ApplicationStatus.StatusError
            strAppStatus = "Error"
        Case ApplicationStatus.statusWarning
            strAppStatus = "Warning"
    End Select

    On Error GoTo 0
    Exit Sub

SetStatus_Error:

    LogEvent "**** Error " & Err.Number & " (" & Err.Description & ") in procedure SetStatus of Module modGlobal", Error
End Sub


