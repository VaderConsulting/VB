Attribute VB_Name = "modMain"
Option Explicit

Public Enum DRIMS
    Disabled = 0
    CDC = 1
    OPD = 2
    Sunrise = 3
    VincentEnfield = 4
    Standard = 5
    DRIMSRM = 5
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public bReportErrors As Boolean
Public bReportWarnings As Boolean
Public bReportInformation As Boolean
Public bReportToEventLog As Boolean
Public oEvents As New Collection
Public bDebugMode As Boolean
Public strAppStatus As String
Public strUsername As String
Public strComputername As String

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

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : CreateGUID
' DateTime  : 12/03/2003 22:52
' Author    : Dave Robinson
' Purpose   : Used to create a GUID for filenames
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function CreateGUID() As String

  Dim uGUID As GUID
  Dim sGUID As String
  Dim bGUID() As Byte
  Dim lLen As Long
  Dim RetVal As Long

    On Error GoTo CreateGUIDError

    lLen = 40
    bGUID = String$(lLen, 0)

    CoCreateGuid uGUID

    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)

    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then
        RetVal = RetVal - 1
    End If

    sGUID = Replace(sGUID, "{", "")
    sGUID = Replace(sGUID, "}", "")

    CreateGUID = Left$(sGUID, RetVal)
    CreateGUID = Replace(CreateGUID, Chr$(0), "")

    Exit Function

CreateGUIDError:
    Err.Clear

End Function

'---------------------------------------------------------------------------------------
' Procedure : EndApp
' DateTime  : 04-04-2003 09:08
' Author    : Dave Robinson
' Purpose   : Initiate shutdown
'
'  V    Date        Author          History
' 1.0   04-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Sub EndApp()

    On Error GoTo EndApp_Error

    WriteEvents
    End

    On Error GoTo 0 ':( Dead Code

Exit Sub

EndApp_Error:

    LogEvent "Error " & Err.Number & " (" & Err.Description & ") in procedure EndApp of Module modMain", Error

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetIniValue
' DateTime  : 03-04-2003 08:21
' Author    : Dave Robinson
' Purpose   : Retrieves a value from an ini file corresponding to the section and key name passed.
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function GetIniValue(lpSectionName As String, lpKeyName As String, defaultValue As String, inifile As String) As String

  Dim success As Long
  Dim nSize As Long
  Dim Ret As String

    On Error GoTo GetIniValue_Error

    'call the API with the parameters passed.
    'The return value is the length of the string in ret, including the terminating null. If a default value was passed, and the section or
    'key name are not in the file, that value is returned. If no default value was passed (""), then success will = 0 if not found.

    'Pad a string large enough to hold the data.
    Ret = Space$(2048)
    nSize = Len(Ret)
    success = GetPrivateProfileString(lpSectionName, lpKeyName, defaultValue, Ret, nSize, inifile)

    If success Then
        GetIniValue = Left$(Ret, success)
    End If

    On Error GoTo 0

Exit Function

GetIniValue_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetIniValue of Module modMain", vbCritical

End Function

'---------------------------------------------------------------------------------------
' Procedure : InGroup
' DateTime  : 03-04-2003 07:56
' Author    : Dave Robinson
' Purpose   : Checks if the user is in the specified group.  Returns true if true
'
'  V    Date        Author          History
' 1.0   03-04-2003  Dave Robinson   Initial Version
' 1.1   03-04-2003  Dave Robinson   Added Domain independence
'---------------------------------------------------------------------------------------
Public Function InGroup(strUsername As String, strGroupname As String) As Boolean

  Dim oUser As IADsUser, oGroup As IADsGroup
  Dim strDomainName As String

    On Error GoTo InGroup_Error

    strDomainName = Environ$("USERDOMAIN")

    If bDebugMode Then MsgBox "My Domain name is " & strDomainName ':( Expand Structure

    Set oUser = GetObject("WinNT://" & strDomainName & "/" & strUsername & ",user")

    For Each oGroup In oUser.Groups
        If LCase$(strGroupname) = LCase$(oGroup.Name) Then
            InGroup = True
            Exit For '>---> Next
        End If
    Next ':( Repeat For-Variable: OGROUP

    Set oUser = Nothing
    Set oGroup = Nothing

    On Error GoTo 0

Exit Function

InGroup_Error:

    MsgBox "**** Error " & Err.Number & " (" & Err.Description & ") in procedure InGroup of Module modGlobal", vbCritical
    Set oUser = Nothing
    Set oGroup = Nothing
    Err.Clear

End Function

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
        oEvents.Add Format$(Now, "DD/MM/YYYY|HH:NN:SS AM/PM") & "|" & EventType & "|" & strMessage
        If (EventType <> Information) And EventType <> DebugInfo Then
            App.LogEvent strMessage, EventType
            'frmMain.sbrStatus.SimpleText = strMessage
            SetStatus EventType
          ElseIf EventType = DebugInfo Then 'NOT (EVENTTYPE...
            If bWriteDebugMessagesToEventLog Then
                App.LogEvent strMessage, vbLogEventTypeInformation
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
' Procedure : SetDRIMS
' DateTime  : 03-04-2003 08:52
' Author    : Dave Robinson
' Purpose   : Set appropriate DRIMS library according to group membership, and with [optional] integration option.
'
' Business Logic:
' Integration\Library    CDC    other    Disabled
'     On                  7       1,3      7
'     Off                 7       4,6      7
'   Not spec.           2,4,6    1,2,3   4,5,6
'
'
' These numbers refer to tasks that must be performed, as follows:
' 1:  Copy files
' 2:  Copy Ini file
' 3:  Set Registry Values
' 4:  Delete files
' 5:  Copy Not_User Ini file
' 6:  Set registry keys blank
' 7:  End without further action
'
'  V    Date    Author      History
' 1.0   03-04-2003  Dave Robinson   Initial Version
'---------------------------------------------------------------------------------------
Public Function SetDRIMS(Library As DRIMS) As Boolean

  Dim oRegistry As REGTool5.Registry
  Dim strIniFilename As String
  Dim bIntegrationOn As Boolean
  Dim bIntegrationOff As Boolean
  Dim bDoTask1 As Boolean
  Dim bDoTask2 As Boolean
  Dim bDoTask3 As Boolean
  Dim bDoTask4 As Boolean
  Dim bDoTask5 As Boolean
  Dim bDoTask6 As Boolean
  Dim bDoTask7 As Boolean
  Dim bResult As Boolean

    On Error GoTo SetDRIMS_Error

    Screen.MousePointer = vbHourglass

    Set oRegistry = CreateObject("REGTool5.Registry")

    ' Check commandline, and setup integration options.....
    If InStr(1, LCase$(Command$), "on") <> 0 Then
        bIntegrationOn = True
        LogEvent "IntegrationOn = " & bIntegrationOn
    End If
    If InStr(1, LCase$(Command$), "off") <> 0 Then
        bIntegrationOn = False
        bIntegrationOff = True
        LogEvent "IntegrationOn = " & bIntegrationOn
        LogEvent "IntegrationOff = " & bIntegrationOff
    End If

    LogEvent "SetDRIMS Library = " & Library, DebugInfo

    Select Case Library
      Case DRIMS.OPD
        LogEvent "Inside SetDRIMS OPD", DebugInfo
        strIniFilename = "pcdocsOPDMain.ini"
        If Not bIntegrationOff Then bDoTask1 = True ':( Expand Structure
        If (Not bIntegrationOn) And (Not bIntegrationOff) Then bDoTask2 = True ':( Expand Structure
        If bIntegrationOn Then bDoTask3 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask4 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask6 = True ':( Expand Structure
      Case DRIMS.Standard, DRIMS.DRIMSRM
        LogEvent "Inside SetDRIMS Standard/DRIMSRM", DebugInfo
        strIniFilename = "pcdocsDRIMS.ini"
        If Not bIntegrationOff Then bDoTask1 = True ':( Expand Structure
        If (Not bIntegrationOn) And (Not bIntegrationOff) Then bDoTask2 = True ':( Expand Structure
        If bIntegrationOn Then bDoTask3 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask4 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask6 = True ':( Expand Structure
      Case DRIMS.Sunrise
        LogEvent "Inside SetDRIMS Sunrise", DebugInfo
        strIniFilename = "pcdocsSunProd.ini"
        If Not bIntegrationOff Then bDoTask1 = True ':( Expand Structure
        If (Not bIntegrationOn) And (Not bIntegrationOff) Then bDoTask2 = True ':( Expand Structure
        If bIntegrationOn Then bDoTask3 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask4 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask6 = True ':( Expand Structure
      Case DRIMS.VincentEnfield
        LogEvent "Inside SetDRIMS Vincent Enfield", DebugInfo
        strIniFilename = "pcdocsVEProd.ini"
        If Not bIntegrationOff Then bDoTask1 = True ':( Expand Structure
        If (Not bIntegrationOn) And (Not bIntegrationOff) Then bDoTask2 = True ':( Expand Structure
        If bIntegrationOn Then bDoTask3 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask4 = True ':( Expand Structure
        If bIntegrationOff Then bDoTask6 = True ':( Expand Structure
      Case DRIMS.Disabled
        LogEvent "Inside SetDRIMS Disable DRIMS", DebugInfo
        If bIntegrationOn Or bIntegrationOff Then
            bDoTask7 = True
          Else 'NOT BINTEGRATIONON...
            bDoTask4 = True
            bDoTask5 = True
            bDoTask6 = True
        End If
      Case DRIMS.CDC
        LogEvent "Inside SetDRIMS CDC", DebugInfo
        If bIntegrationOn Or bIntegrationOff Then
            bDoTask7 = True
          Else 'NOT BINTEGRATIONON...
            strIniFilename = "pcdocscdc.ini"
            bDoTask2 = True
            bDoTask4 = True
            bDoTask6 = True
        End If
    End Select

    If bDoTask1 Then
        LogEvent "Inside SetDRIMS Task1", DebugInfo
        If Dir("C:\Program Files\Microsoft Office\Office10\Startup\WordinXP.dot") <> "" Then
            Kill "C:\Program Files\Microsoft Office\Office10\Startup\WordinXP.dot"
        End If
        If Dir("C:\Program Files\Microsoft Office\Office10\XLStart\docsxlXP.xla") <> "" Then
            Kill "C:\Program Files\Microsoft Office\Office10\XLStart\docsxlXP.xla"
        End If
        If Dir("C:\Program Files\Adobe\Acrobat\Acrobat\Plug_Ins\acrodx32.api") <> "" Then
            Kill "C:\Program Files\Adobe\Acrobat\Acrobat\Plug_Ins\acrodx32.api"
        End If

        FileCopy "C:\Program Files\Drims\progs\Integration\WordinXP.dot", "C:\Program Files\Microsoft Office\Office10\Startup\WordinXP.dot"
        FileCopy "C:\Program Files\Drims\progs\Integration\docsxlXP.xla", "C:\Program Files\Microsoft Office\Office10\XLStart\docsxlXP.xla"
        FileCopy "C:\Program Files\Drims\progs\Integration\acrodx32.api", "C:\Program Files\Adobe\Acrobat\Acrobat\Plug_Ins\acrodx32.api"
    End If

    If bDoTask2 Then
        LogEvent "Inside SetDRIMS Task2", DebugInfo
        If Dir("C:\Program Files\Drims\progs\pcdocs.ini") <> "" Then
            Kill "C:\Program Files\Drims\progs\pcdocs.ini"
        End If

        FileCopy "C:\Program Files\Drims\progs\pcdocsini\" & strIniFilename, "C:\Program Files\Drims\progs\pcdocs.ini"
    End If

    If bDoTask3 Then
        LogEvent "Inside SetDRIMS Task3", DebugInfo
        oRegistry.UpdateKey REGToolRootTypes.HKEY_CLASSES_ROOT, "MS WORD\ODMA32", "", "PCDOCS"
        oRegistry.UpdateKey REGToolRootTypes.HKEY_CLASSES_ROOT, "VISIO\ODMA32", "", "PCDOCS"
        oRegistry.UpdateKey REGToolRootTypes.HKEY_CLASSES_ROOT, "MS POWERPOINT\ODMA32", "", "PCDOCS"
        oRegistry.UpdateKey REGToolRootTypes.HKEY_CLASSES_ROOT, "MSPROJECT\ODMA32", "", "PCDOCS"
        oRegistry.UpdateKey REGToolRootTypes.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchEIM.Module", "4.0;ExchEIM.dll;1"
        oRegistry.UpdateKey REGToolRootTypes.HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchINS.Module", "4.0;exchINS.dll;1"
    End If

    If bDoTask4 Then
        LogEvent "Inside SetDRIMS Task4", DebugInfo
        If Dir("C:\Program Files\Microsoft Office\Office10\Startup\WordinXP.dot") <> "" Then
            Kill "C:\Program Files\Microsoft Office\Office10\Startup\WordinXP.dot"
        End If
        If Dir("C:\Program Files\Microsoft Office\Office10\XLStart\docsxlXP.xla") <> "" Then
            Kill "C:\Program Files\Microsoft Office\Office10\XLStart\docsxlXP.xla"
        End If
        If Dir("C:\Program Files\Adobe\Acrobat\Acrobat\Plug_Ins\acrodx32.api") <> "" Then
            Kill "C:\Program Files\Adobe\Acrobat\Acrobat\Plug_Ins\acrodx32.api"
        End If
    End If

    If bDoTask5 Then
        LogEvent "Inside SetDRIMS Task5", DebugInfo
        If Dir("C:\Program Files\Drims\progs\pcdocs.ini") <> "" Then
            Kill "C:\Program Files\Drims\progs\pcdocs.ini"
        End If

        FileCopy "C:\Program Files\Drims\progs\pcdocsini\pcdocs_NOT_USER.ini", "C:\Program Files\Drims\progs\pcdocs.ini"
    End If

    If bDoTask6 Then
        LogEvent "Inside SetDRIMS Task6", DebugInfo
        DelRegValue HKEY_CLASSES_ROOT, "MS WORD\ODMA32", ""
        DelRegValue HKEY_CLASSES_ROOT, "VISIO\ODMA32", ""
        DelRegValue HKEY_CLASSES_ROOT, "MS POWERPOINT\ODMA32", ""
        DelRegValue HKEY_CLASSES_ROOT, "MSPROJECT\ODMA32", ""
        DelRegValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchEIM.Module"
        DelRegValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Exchange\Client\Extensions", "ExchINS.Module"
    End If

    If bDoTask7 Then
        LogEvent "Inside SetDRIMS Task7", DebugInfo
        SetDRIMS = True
        Exit Function '>---> Bottom
    End If

    SetDRIMS = True

    Screen.MousePointer = vbDefault

    Set oRegistry = Nothing
    On Error GoTo 0

Exit Function

SetDRIMS_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetDRIMS of Module modMain", vbCritical
    Screen.MousePointer = vbDefault
    Set oRegistry = Nothing

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
    oSAXAttributes.addAttribute "", "", "Date", "", Format$(datNow, "DD/MM/YYYY")
    oSAXAttributes.addAttribute "", "", "Time", "", Format$(datNow, "HH:NN:SS AM/PM")
    oSAXAttributes.addAttribute "", "", "Username", "", strUsername
    oSAXAttributes.addAttribute "", "", "Computername", "", strComputername
    ' Start the XML Root Element (this adds the attributes)
    oContentHandler.startElement "", "", "SetLibrary", oSAXAttributes
    ' Remove the sttributes
    oSAXAttributes.Clear

    ' Create each element
    For iEvent = 1 To oEvents.Count
        ' Strip the date from the status message
        iPosition1 = InStr(1, oEvents.Item(iEvent), "|")
        strDate = Trim$(Left$(oEvents.Item(iEvent), iPosition1 - 1))
        oSAXAttributes.addAttribute "", "", "Date", "", strDate
        ' Strip the time from the status message
        iPosition2 = InStr(iPosition1 + 1, oEvents.Item(iEvent), "|") + 1
        strTime = Trim$(Mid$(oEvents.Item(iEvent), iPosition1 + 1, iPosition2 - (iPosition1 + 2)))
        oSAXAttributes.addAttribute "", "", "Time", "", strTime
        ' Strip the event type from the status message
        iPosition3 = InStr(iPosition2 + 1, oEvents.Item(iEvent), "|") + 1
        strType = Trim$(Mid$(oEvents.Item(iEvent), iPosition2, iPosition3 - (iPosition2 + 1)))
        oSAXAttributes.addAttribute "", "", "Type", "", strType
        ' Add these attributes to the element
        oContentHandler.startElement "", "", "Event", oSAXAttributes
        ' Add the message
        strMessage = Mid$(oEvents.Item(iEvent), iPosition3, 1024)
        oContentHandler.characters strMessage
        ' Remove the attributes
        oSAXAttributes.Clear
        oContentHandler.endElement "", "", "Event"
    Next iEvent

    'Complete the Root Element
    oContentHandler.endElement "", "", "SetLibrary"
    ' Finish the XML Document
    oContentHandler.endDocument

    strFilename = "C:\Temp\SetLibraryResults-" & strUsername & "-" & Format$(datNow, "DDMMYYYY") & "-" & Format$(datNow, "HHNNSS AM/PM") & "-" & CreateGUID & ".xml"
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

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:26) 49 + 582 = 631 Lines
