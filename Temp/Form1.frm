VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl script 
      Left            =   2520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ListBox lstPackages 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ScriptLanguage
    VBScript = 0
    JavaScript = 1
End Enum

Private Sub Form_Load()
    Dim oMyXMLDocument As MSXML2.DOMDocument40  ' Final XML
    Dim oMyFragment As IXMLDOMDocumentFragment  ' Portion
    Dim oXML As MSXML2.DOMDocument40            ' Original
    
    Dim oFSO As Scripting.FileSystemObject
    
    Dim ScriptLine As String
    Dim ScriptContents As String
    
    Dim colFilePaths As New Collection
    
    'Set oRequiredPackages = CreateObject("Scripting.Dictionary")
    'Set oParentPackages = CreateObject("Scripting.Dictionary")
    
    'Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Create XML Documents to load the XML into
    Set oMyXMLDocument = New MSXML2.DOMDocument40
    Set oXML = New MSXML2.DOMDocument40
    
    Set oMyFragment = oMyXMLDocument.createDocumentFragment
    
    'lstPackages.AddItem "GUID3|Name3"
    'lstPackages.AddItem "GUID4|Name4"
    'CreateXMLFile "c:\temp\test.xml"
    
    LoadDocument "c:\temp\packages.xml", oXML
    LoadFragment "c:\temp\test.xml", oMyFragment
    MergeXML oMyFragment, oXML, oMyXMLDocument
    'SaveDocument
    
    'GetFiles oFSO.GetFolder("C:\Temp"), colFilePaths
    
    'ReadScript "c:\temp\instruction.txt", script, True
    

    'Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    Set oMyXMLDocument = Nothing
    Set oMyFragment = Nothing
    Set oXML = Nothing
End Sub

Sub GetFiles(oFolder As Folder, colPaths As Collection, Optional Recursive As Boolean = False)
    Dim SubFolder As Folder
    Dim colFiles As Files
    Dim oFile As File
    
    Set colFiles = oFolder.Files
    For Each oFile In colFiles
        colPaths.Add oFolder.Path & "\" & oFile.Name
    Next
    
    If Recursive Then
        For Each SubFolder In oFolder.SubFolders
            colPaths.Add SubFolder.Path
            GetFiles SubFolder, colPaths, Recursive
        Next
    End If
End Sub

'Private Sub ShowFiles(oFolder As Folder)
'    Dim f As Variant
'
'    For Each f In oFolder
'        Debug.Print f.Name
'
'    Next
'End Sub

Sub LoadDocument(strFilename As String, XML As MSXML2.DOMDocument40)
    XML.Load strFilename
End Sub

Sub LoadFragment(strFilename As String, Fragment As IXMLDOMDocumentFragment)
    Dim oTempXML As MSXML2.DOMDocument40
    Set oTempXML = New MSXML2.DOMDocument40
    
    oTempXML.Load strFilename
    
    
    
    Set oTempXML = Nothing
End Sub

Sub MergeXML(Fragment As IXMLDOMDocumentFragment, XML As MSXML2.DOMDocument40, XMLDocument As MSXML2.DOMDocument40)
    
End Sub

' This sub will only create an XML file from the contents of a listbox
' -------------------------------------------------------------------------------------------------------
Private Sub CreateXMLFile(strFilename As String, Optional bAppend As Boolean = False)
    Dim oWriter As New MXXMLWriter30
    Dim oContentHandler As IVBSAXContentHandler
    Dim oDTDHandler As IVBSAXDTDHandler
    Dim oLexicalHandler As IVBSAXLexicalHandler
    Dim oDeclarationHandler As IVBSAXDeclHandler
    Dim oErrorHandler As IVBSAXErrorHandler
    Dim oSAXAttributes As New SAXAttributes
    Dim strPackageGUID As String
    Dim strPackageName As String
    Dim l1 As Long
    Dim l2 As Long
    Dim XMLFilename As String
    
    On Error GoTo CreateXMLError
    
    ' These objects are required to create the XML document.
    Set oContentHandler = oWriter
    Set oDTDHandler = oWriter
    Set oLexicalHandler = oWriter
    Set oDeclarationHandler = oWriter
    Set oErrorHandler = oWriter
    
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
    
    oContentHandler.startElement "", "", "Packages", oSAXAttributes
    
    ' Create each element
    For l1 = 1 To lstPackages.ListCount
        l2 = InStr(1, lstPackages.List(l1 - 1), "|")
        strPackageGUID = Left(lstPackages.List(l1 - 1), l2 - 1)
        strPackageName = Mid(lstPackages.List(l1 - 1), l2 + 1, 1024)
        oSAXAttributes.addAttribute "", "", "GUID", "", strPackageGUID
        ' Add this attribute to the element
        oContentHandler.startElement "", "", "Package", oSAXAttributes
        ' Remove any existing attributes
        oSAXAttributes.Clear
        ' Add the package
        oContentHandler.startElement "", "", "Name", oSAXAttributes
        oContentHandler.characters strPackageName
        ' Remove the attributes
        oSAXAttributes.Clear
        oContentHandler.endElement "", "", "Name"
        oContentHandler.endElement "", "", "Package"
    Next l1
    
    'Complete the Root Element
    oContentHandler.endElement "", "", "Packages"
    ' Finish the XML Document
    oContentHandler.endDocument
    
    ' Write the XML file out to the final location
    Open strFilename For Output As #1
        Print #1, oWriter.output
    Close 1
    
    ' Clean up
    Set oWriter = Nothing
    Set oContentHandler = Nothing
    Set oDTDHandler = Nothing
    Set oLexicalHandler = Nothing
    Set oDeclarationHandler = Nothing
    Set oErrorHandler = Nothing
    Exit Sub
CreateXMLError:
    'LogError "PRS encountered internal error " & Err.Number & " (" & Err.Description & ") in CreateXMLFile", Error
    'LogEvent "**** PRS encountered internal error " & Err.Number & " (" & Err.Description & ") in CreateXMLFile"
    'SetStatus StatusError
    
    Set oWriter = Nothing
    Set oContentHandler = Nothing
    Set oDTDHandler = Nothing
    Set oLexicalHandler = Nothing
    Set oDeclarationHandler = Nothing
    Set oErrorHandler = Nothing
End Sub

' -------------------------------------------------------------------------------------------------------
' Load the given Script and load into the given script control
' Usage:  ReadScript "c:\helloworld.vbs", oScript, VBScript, False
'
Public Sub ReadScript(strPath As String, ScriptObject As ScriptControl, Optional Language As ScriptLanguage = VBScript, Optional AllowUi As Boolean = False)
    Dim f As Integer
    Dim ScriptLine As String
    Dim ScriptContents As String
    
    'On Error GoTo ReadScriptError
    
    f = FreeFile
    script.Reset
    
    ' Read the given file
    Open strPath For Input As #f

        Do Until EOF(f)
            Line Input #f, ScriptLine
            ScriptContents = ScriptContents & vbCrLf & ScriptLine
        Loop
    Close f
    
    ' This script is in the given language, so set the script control to use this language
    Select Case Language
        Case 0
            ScriptObject.Language = "VBScript"
        Case 1
            ScriptObject.Language = "JavaScript"
    End Select
    
    ' Set the AllowUI property as appropriate
    ScriptObject.AllowUi = AllowUi
    
    ' Add the contents of the file to the script control
    ScriptObject.AddCode ScriptContents
    Exit Sub
ReadScriptError:
    'LogError "PRS encountered internal error " & Err.Number & " (" & Err.Description & ") in ReadScript", Error
    'LogEvent "**** PRS encountered internal error " & Err.Number & " (" & Err.Description & ") in ReadScript"
    'SetStatus StatusError
    Err.Clear
End Sub
