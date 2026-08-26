Attribute VB_Name = "modMain"
Option Explicit

Dim mstrUsername As String
Dim mstrComputername As String

Public strAppStatus As String
Public oEvents As New Collection
Public bReportErrors As Boolean
Public bReportWarnings As Boolean
Public bReportInformation As Boolean
Public bReportToEventLog As Boolean
Public bWriteDebugMessagesToEventLog As Boolean
Public bDebugMode As Boolean

Public strApp_Path As String
Public oDictionary As Scripting.Dictionary

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
    
    ' Only add this event if we have determined it should be (according to the config options)
    If bDoReport Then
        oEvents.Add Format(Now, "DD/MM/YYYY|HH:NN:SS AM/PM") & "|" & EventType & "|" & strMessage
        If (EventType <> Information) And EventType <> DebugInfo Then
            App.LogEvent strMessage, EventType
            SetStatus EventType
        ElseIf EventType = DebugInfo Then
            If bWriteDebugMessagesToEventLog Then
                App.LogEvent strMessage, vbLogEventTypeInformation
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
' Procedure : OpenConfig
' DateTime  : 02-04-2003 08:05
' Author    : Dave Robinson
' Purpose   : Open the configuration file and return all info into the supplied dictionary object
'
'  V    Date    Author      History
' 1.0   02-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Sub OpenConfig(Dic As Scripting.Dictionary)
    Dim oXMLDocument As MSXML2.DOMDocument40
    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
    Dim oXMLNode As MSXML2.IXMLDOMNode
    Dim oXMLValueList As MSXML2.IXMLDOMNodeList
    Dim intHint As Integer
    
    On Error GoTo OpenConfig_Error

    ' Clear all items from the dictionary object before we start adding more (potentially duplicate) items to it.
    Dic.removeAll
    
    ' Create an XML Document to load the XML file into
    Set oXMLDocument = New MSXML2.DOMDocument40
    
    ' Check to ensure the config.xml file exists
    If Dir(strApp_Path & "config.xml") = "" Then
       frmMain.SCM.LogEvent svcEventWarning, 12, frmMain.SCM.DisplayName & " could not find the config file 'config.xml' under the current path (" & strApp_Path & ")"
    Else
        ' Load the XML file
        oXMLDocument.Load (strApp_Path & "config.xml")

        ' Reference the nodes
        Set oXMLNodeList = oXMLDocument.getElementsByTagName("Config")
        
        ' Retrieve the individual values in the configuration file
        Set oXMLValueList = oXMLNodeList.Item(0).childNodes
        
        ' ********************************************************
        ' Load each value and add it to a dictionary of values
        For Each oXMLNode In oXMLValueList
            Dic.Add LCase(oXMLNode.baseName), oXMLNode.Text
        Next
        ' ********************************************************
    End If
    
    ' Clean up
    Set oXMLDocument = Nothing
    
'OpenConfigError:
'   frmMain.SCM.LogEvent svcEventError, 13,frmMain.SCM.DisplayName & " encountered an internal error (" & Err.Number & ")in OpenConfig.  See the next entry for detailed information"
'
'    ' The following is used for error diagnosis only.
'    Dim strItems As String
'    Dim Item As Variant
'    Dim i As Integer
'    Dim DicItems() As String
'    ReDim DicItems(Dic.Count)
'
'    ' Build a list of Dictionary items
'    For Each Item In Dic.Items
'        i = i + 1
'        DicItems(i) = Item
'    Next
'
'    ' Build a list of dictionary keys, and add the items as well
'    i = 0
'    For Each Item In Dic.Keys
'        i = i + 1
'        DicItems(i) = Item & "=" & DicItems(i)
'    Next
'
'    ' Build a string from the items and keys to add to the Event Log
'    i = 0
'    strItems = vbCrLf
'    For i = 1 To Dic.Count
'        strItems = strItems & DicItems(i) & vbCrLf
'    Next
'   frmMain.SCM.LogEvent svcEventError, 14,frmMain.SCM.DisplayName & strItems
'    Set oXMLDocument = Nothing
'    Err.Clear

    On Error GoTo 0
    Exit Sub

OpenConfig_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenConfig of Module modMain", Error
End Sub

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

'---------------------------------------------------------------------------------------
' Procedure : WriteEvents
' DateTime  : 14-03-2003 10:38
' Author    : Dave Robinson
' Purpose   : Write application events to the status log file
'             Taken from DRS Application
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Sub WriteEvents()
    Dim oWriter As New MXXMLWriter ' <-- Can't be more specific.  Causes problems
    Dim oContentHandler As IVBSAXContentHandler
    Dim oDTDHandler As IVBSAXDTDHandler
    Dim oLexicalHandler As IVBSAXLexicalHandler
    Dim oDeclarationHandler As IVBSAXDeclHandler
    Dim oErrorHandler As IVBSAXErrorHandler
    Dim oSAXAttributes As New MSXML2.SAXAttributes   ' <--- Do not change this declaration.  Doing so causes problems with W2K.
    Dim iEvent As Integer
    Dim iPosition1 As Integer
    Dim iPosition2 As Integer
    Dim iPosition3 As Integer
    Dim datNow As Date
    Dim strFilename As String
    Dim strDate As String
    Dim strTime As String
    Dim strType As String
    Dim strMessage As String
    Dim intHint As Integer
    
    On Error GoTo WriteError
    
    ' These objects are required to create the XML document.
    Set oContentHandler = oWriter
    Set oDTDHandler = oWriter
    Set oLexicalHandler = oWriter
    Set oDeclarationHandler = oWriter
    Set oErrorHandler = oWriter
    
    mstrUsername = Environ$("USERNAME")
    mstrComputername = Environ$("COMPUTERNAME")
    
    ' Prevent the XML Declaration from being added - this is because we can't work with the UTF-16 encoding that VB produces by default.
    oWriter.omitXMLDeclaration = True
    'Cause indenting of the XML Document
    oWriter.indent = True
    ' Allow the XML Document to be used without a DTD
    oWriter.standalone = True
    ' Start the XML Document
    oContentHandler.startDocument
    ' Remove any existing attributes
    oSAXAttributes.Clear
    ' Add some attributes
    datNow = Now
    oSAXAttributes.addAttribute "", "", "LastStatus", "", strAppStatus
    oSAXAttributes.addAttribute "", "", "Date", "", Format(datNow, "DD/MM/YYYY")
    oSAXAttributes.addAttribute "", "", "Time", "", Format(datNow, "HH:NN:SS AM/PM")
    oSAXAttributes.addAttribute "", "", "Username", "", mstrUsername
    oSAXAttributes.addAttribute "", "", "Computername", "", mstrComputername
    ' Start the XML Root Element (this adds the attributes)
    oContentHandler.startElement "", "", App.ProductName, oSAXAttributes
    ' Remove the sttributes
    oSAXAttributes.Clear
    
    ' Create each element
    For iEvent = 1 To oEvents.Count
        ' Strip the date from the status message
        iPosition1 = InStr(1, oEvents.Item(iEvent), "|")
        strDate = Trim(Left(oEvents.Item(iEvent), iPosition1 - 1))
        oSAXAttributes.addAttribute "", "", "Date", "", strDate
        ' Strip the time from the status message
        iPosition2 = InStr(iPosition1 + 1, oEvents.Item(iEvent), "|") + 1
        strTime = Trim(Mid(oEvents.Item(iEvent), iPosition1 + 1, iPosition2 - (iPosition1 + 2)))
        oSAXAttributes.addAttribute "", "", "Time", "", strTime
        ' Strip the event type from the status message
        iPosition3 = InStr(iPosition2 + 1, oEvents.Item(iEvent), "|") + 1
        strType = Trim(Mid(oEvents.Item(iEvent), iPosition2, iPosition3 - (iPosition2 + 1)))
        oSAXAttributes.addAttribute "", "", "Type", "", strType
        ' Add these attributes to the element
        oContentHandler.startElement "", "", "Event", oSAXAttributes
        ' Add the message
        strMessage = Mid(oEvents.Item(iEvent), iPosition3, 1024)
        oContentHandler.characters strMessage
        ' Remove the attributes
        oSAXAttributes.Clear
        oContentHandler.endElement "", "", "Event"
    Next iEvent
    
    'Complete the Root Element
    oContentHandler.endElement "", "", App.ProductName
    ' Finish the XML Document
    oContentHandler.endDocument
    
    strFilename = "C:\Temp\" & App.ProductName & "-" & mstrUsername & "-" & Format(datNow, "DDMMYYYY") & "-" & Format(datNow, "HHNNSS AM/PM") & "-" & CreateGUID & ".xml"
    ' Write the XML file out to the final location
    Open strFilename For Output As #1
        Print #1, oWriter.output
    Close 1
    
    ' Clean up
    Set oSAXAttributes = Nothing
    Set oWriter = Nothing
    Set oContentHandler = Nothing
    Set oDTDHandler = Nothing
    Set oLexicalHandler = Nothing
    Set oDeclarationHandler = Nothing
    Set oErrorHandler = Nothing
    Exit Sub
WriteError:
    LogEvent "**** " & App.ProductName & " encountered internal error " & Err.Number & " (" & Err.Description & ") in WriteEvents:" & intHint & ".", Error
    
    Set oSAXAttributes = Nothing
    Set oWriter = Nothing
    Set oContentHandler = Nothing
    Set oDTDHandler = Nothing
    Set oLexicalHandler = Nothing
    Set oDeclarationHandler = Nothing
    Set oErrorHandler = Nothing
    Err.Clear
End Sub
