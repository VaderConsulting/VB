VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Discover 2"
   ClientHeight    =   390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Audit in progress..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oWriter As New MXXMLWriter30
Public oContentHandler As IVBSAXContentHandler
Public oDTDHandler As IVBSAXDTDHandler
Public oLexicalHandler As IVBSAXLexicalHandler
Public oDeclarationHandler As IVBSAXDeclHandler
Public oErrorHandler As IVBSAXErrorHandler
Public strApp_Path As String

Public strNA1Filename As String
Public strNA2Filename As String
Public strNA3Filename As String
Public strNA4Filename As String
Public strXMLFilename As String
Public strErrFilename As String
Public intTransparency As Integer

Public intClasses As Integer

Private Sub Form_Load()
    Dim mcolClasses As New Collection
    Dim oSAXAttributes As New SAXAttributes
    Dim datStart As Date
    Dim dblRandom As Double
    Dim Transparency As Integer
    Dim s As String
    Dim ParentDirectory As String
    Dim ComputerName As String
    Dim strComputer As String
    Dim ErrFilename As String
        
    ' Any errors - goto the VBError label
    On Error GoTo VBError
    
    ' These objects are required to create the XML documents.
    Set oContentHandler = oWriter
    Set oDTDHandler = oWriter
    Set oLexicalHandler = oWriter
    Set oDeclarationHandler = oWriter
    Set oErrorHandler = oWriter
    
    ' generate an App_Path variable that is the path to this application, similar to App.Path, though it will always end with a backslash.
    If Right(App.Path, 1) <> "\" Then
        strApp_Path = App.Path & "\"
    Else
        strApp_Path = App.Path
    End If
    
    ' Show and refresh the form
    Me.Show
    Me.Refresh
    
    ' Close this app if it hasn't been told what to do
    If Command$ = "" Then End
    
    ' Set the default transparency
    intTransparency = 255
    
    ' Set the parent (or root) directory
    ParentDirectory = strApp_Path
    
    ' If there is nothing to do, close down
    If Dir(ParentDirectory & "\" & Command$ & "\*.na2") = "" Then End
    
    ' This is a demo
    'If Date > CDate("31/12/2002") Then
    '    End
    'End If
    
    ' Get the next file to process
    s = Dir(ParentDirectory & "\" & Command$ & "\*.na2")
    ComputerName = Left(s, InStr(1, s, ".") - 1)
    
    strNA1Filename = ParentDirectory & "\" & Command$ & "\" & ComputerName & ".na1"
    strNA2Filename = ParentDirectory & "\" & Command$ & "\" & s
    strNA3Filename = ParentDirectory & "\" & Command$ & "\" & ComputerName & ".na3"
    strNA4Filename = ParentDirectory & "\" & Command$ & "\" & ComputerName & ".na4"
    strXMLFilename = ParentDirectory & "\" & Command$ & "\" & ComputerName & ".xml"
    ErrFilename = ParentDirectory & "\" & Command$ & "\" & ComputerName & ".err"
    
    ' Rename the working file - set it so the next instance of this app won't use the same file
    Name strNA2Filename As strNA3Filename
    
    strComputer = ComputerName
    
    Me.Caption = Me.Caption & " [" & ComputerName & "]"
    
    ' Get the list of Classes to query
    GetClasses mcolClasses
    
    ' Prevent the XML Declaration from being added - this is because we can't work with the UTF-16 encoding that
    '   VB produces by default.
    oWriter.omitXMLDeclaration = True
    'Cause indenting of the XML Document
    oWriter.indent = True
    ' Allow the XML Document to be used without a DTD
    oWriter.standalone = True
    ' Start the XML Document
    oContentHandler.startDocument
    ' Remove any existing attributes
    oSAXAttributes.Clear
    ' Start the XML Root Element
    oContentHandler.startElement "", "", ComputerName, oSAXAttributes
    
    ' Perform the audit, adding elements to the XML Document as required.
    DoAudit mcolClasses, ComputerName
    
    'Complete the Root Element
    oContentHandler.endElement "", "", ComputerName
    ' Finish the XML Document
    oContentHandler.endDocument
    
    ' Write the XML file out to the final location
    Name strNA3Filename As strXMLFilename
    Open strXMLFilename For Output As #1
        Print #1, oWriter.output
    Close 1
    
    ' Clean up
    If Dir(ErrFilename) <> "" Then
        ' Delete error File
        Kill ErrFilename
    End If
    
    Set oWriter = Nothing
    End
VBError:
    End
End Sub

Private Function GetClasses(colClasses As Collection)
    Dim oXMLDocument As msxml2.DOMDocument30
    Dim oXMLNodeList As msxml2.IXMLDOMNodeList
    Dim oXMLNode As msxml2.IXMLDOMNode
    
    ' Any errors - goto the VBError label
    On Error GoTo VBError
    
    ' Create an XML Document to load the XML file into
    Set oXMLDocument = New msxml2.DOMDocument30
    
    ' Load the XML file
    oXMLDocument.Load (strApp_Path & "discover.xml")
    ' Go through this file, and get each class to extract
    For Each oXMLNode In oXMLDocument.documentElement.childNodes
        colClasses.Add oXMLNode.Text
        DoEvents
    Next
    
    ' Get the number of classes to query
    intClasses = colClasses.Count
    
    ' Clean up
    Set oXMLDocument = Nothing
    Exit Function
VBError:
    End
End Function

Private Sub DoAudit(colClasses As Collection, strComputername As String)
    Dim objWMIService As SWbemServices
    Dim oProperty As Variant
    Dim colProperties As Object
    Dim strComputer As String
    Dim varInstance As Variant
    Dim strClassName As String
    Dim strPropertyName As String
    Dim oSAXAttributes As New SAXAttributes
    Dim i As Integer
    Dim strArray As String
    Dim vItem As Variant
    Dim ErrFilename As String
    Dim NoOfErrors As Integer
    Dim intThisClass As Integer
    
    ' Any errors - goto the VBError label
    On Error GoTo VBError
    
    DoEvents
    ' Get a handle to the WMI object
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputername & "\root\cimv2")
    
    For Each vItem In colClasses
        intThisClass = intThisClass + 1
        strClassName = vItem
        ' Ensure the user is kept up-to-date
        UpdateUser "Audit in progress... " & strClassName
        ' Get a handle to the WMI Class in question
        Set colProperties = objWMIService.ExecQuery("Select * from " & strClassName)
        ' Start the XML Element
        oContentHandler.startElement "", "", strClassName, oSAXAttributes
        'Loop through each instance of this class
        For Each varInstance In colProperties
            ' Start an element to define the instance
            oContentHandler.startElement "", "", "Instance", oSAXAttributes
            'Loop through each property
            For Each oProperty In varInstance.Properties_
                ' For each property, create an XML Element
                oContentHandler.startElement "", "", oProperty.Name, oSAXAttributes
                ' Some properties are array values.  For these, go through each value and extract it.
                If oProperty.IsArray Then
                    If Not IsNull(oProperty) Then
                        strArray = ""
                        For i = 0 To UBound(oProperty.Value)
                            ' Separate each array value with the pipe character |
                            strArray = strArray & CStr(oProperty.Value(i)) & "|"
                        Next
                        ' Remove ending pipe character
                        If Right(strArray, 1) = "|" Then strArray = Left(strArray, Len(strArray) - 1)
                        oContentHandler.characters strArray
                    Else
                        oContentHandler.characters ""
                    End If
                Else
                    oContentHandler.characters oProperty & ""
                End If
                ' Finish the Property element
                oContentHandler.endElement "", "", oProperty.Name
                ' Remove the attributes (if any)
                oSAXAttributes.Clear
                DoEvents
            Next
            ' Finish the Instance element
            oContentHandler.endElement "", "", "Instance"
            DoEvents
        Next
        ' Complete this element
        oContentHandler.endElement "", "", strClassName
        DoEvents
        
        ' Set the forms transparency
        MakeTransparent Me.hWnd, intTransparency
        intTransparency = 255 - (intThisClass / intClasses) * 100
        Me.Refresh
    Next
    
    ' Clean up
    Set colProperties = Nothing
    Set objWMIService = Nothing
    Exit Sub
VBError:
    UpdateUser "Error " & Err.Number & " (" & Err.Description & ")"
    UpdateUser "//Error: " & Err.Number & " (" & Err.Description & ")"
    Err.Clear
    
    NoOfErrors = NoOfErrors + 1
    
    On Error Resume Next
        
    ' Delete pre-existing Error file
    If Dir(ErrFilename) <> "" Then
        Kill ErrFilename
    End If
    
    If Dir(ErrFilename) <> "" Then
        ' Could not delete Error file, so rename to NA1 file
        Name strNA3Filename As strNA1Filename
    Else
        ' No Error file, so rename NA3 file to Error file
        Name strNA3Filename As strErrFilename
    End If
    
    If NoOfErrors >= 5 Then
        Exit Sub
    End If
    Resume Next
End Sub

Private Sub UpdateUser(strMessage As String)
    ' Update the user as to the progress
    frmMain.lblInfo = strMessage
    DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Dir(strNA3Filename) <> "" Then
        Name strNA3Filename As strNA2Filename
    End If
End Sub
