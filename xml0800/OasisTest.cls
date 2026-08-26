VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OasisTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Total, NegPass, Failed, OutputFailed, Skipped As Integer

Public Validating As Boolean
Private testresults As DOMDocument
Private table_valid As IXMLDOMElement
Private table_valid_output As IXMLDOMElement
Private table_invalid As IXMLDOMElement
Private table_not_wf As IXMLDOMElement
Private table_informative As IXMLDOMElement

Public Cancel As Boolean
Public Quit As Boolean
Dim reportfile As String
Dim tempFile As String

Private Sub Class_Initialize()
    Total = 0
    Validating = False
    reportfile = App.Path & "\msxml3-sax.html"
    tempFile = App.Path & "\_temp.out"
    ClearFields
End Sub

Private Sub ClearFields()
    Dim i As Integer
    For i = 0 To Form1.TextResult.count - 1
        Form1.TextResult(i).text = ""
    Next
    Form1.Percent.Caption = ""
    Form1.Status.Caption = ""
    Form1.Percent.BackColor = &H8000000F
    Form1.FullReport.Enabled = False
End Sub


Public Sub run(testuri As String)

    ClearFields
        
    ' Make sure MSXML3 Beta Release is installed.
    On Error GoTo NoMsxml3
    Dim doc As DOMDocument
    Set doc = New DOMDocument
    
    ' Make sure the VBSAXXMLReader is available.
    Dim reader As VBSAXXMLReader30
    Set reader = New VBSAXXMLReader30

    ' Create a DOM document for storing the detailed results
    Set testresults = New DOMDocument
    Set table_valid = testresults.createElement("valid")
    Set table_valid_output = testresults.createElement("valid-output")
    Set table_invalid = testresults.createElement("invalid")
    Set table_not_wf = testresults.createElement("not-wf")
    Set table_informative = testresults.createElement("informative")
    
    ' Only support HTTP location for test suite at this time.
    ' This is explained in the CompareResults method.
    If (Mid(testuri, 1, 7) <> "http://") Then
        MsgBox "This test harness can only run conformance tests via HTTP."
        GoTo BadUri
    End If
    
    On Error GoTo Error
    '-------------------------------------------------------------
    ' Load up the test suite index into a DOM document
    doc.async = False
    If (Not doc.Load(testuri)) Then
        MsgBox "Error: " & doc.parseError.reason
        GoTo BadUri
    End If
    
    Dim tests As IXMLDOMNodeList
    Set tests = doc.selectNodes("//TEST")
    Dim node As IXMLDOMElement
    Set node = tests.nextNode()
    
    While Not node Is Nothing And Not Cancel
        Dim url As String
        RunTest testuri, node
        
        Form1.TextResult(0).text = Total
        Dim passed As Integer
        passed = Total - Failed - OutputFailed
        Form1.TextResult(1).text = passed
        Form1.TextResult(2).text = Failed + OutputFailed
        Form1.TextResult(3).text = NegPass
        Form1.TextResult(4).text = OutputFailed
        
        DoEvents
        
        Set node = tests.nextNode()
    Wend
    If Quit Then End
                
    If (Cancel) Then
        Form1.Status.Caption = "[Run Cancelled]"
    ElseIf (Total > 0) Then
        Form1.Status.Caption = ""
        Form1.Percent.Caption = FormatPercent(passed / Total, 0)
        Form1.Percent.BackColor = &HFFFFFF
        Form1.FullReport.Enabled = True
    End If
    Exit Sub
Error:
    MsgBox "Error: " & Err.Number & ": " & Err.Description
    Exit Sub
NoMsxml3:
    MsgBox "Error: This requires the MSXML3 Beta Release, but it does not seem to be installed.", vbOKOnly, "Error"
    Exit Sub
BadUri:
    Form1.Text1.SelStart = 0
    Form1.Text1.SelLength = Len(Form1.Text1.text)
    Form1.Text1.SetFocus
End Sub

Public Sub RunTest(docBase As String, element As IXMLDOMElement)

    Dim SaxOutput As TextStream
    
    ' Fetch basic test attributes from test element.
    Dim id, ttype, expectedOutput, uri As String
    id = element.getAttribute("ID")
    ttype = element.getAttribute("TYPE")
    expectedOutput = element.getAttribute("OUTPUT")
    uri = element.getAttribute("URI")

    ' NOTE: OASIS/NIST omitted the OUTPUT3 data cases, which should
    ' be used with validating parsers (which have special requirements
    ' for ignorable whitespace and unparsed entities -- untested here).
    ' ignoring ENTITIES for the moment (assuming they are always read)
    ' is this a positive (must-succeed) test?
    Dim positive As Boolean
    Dim table As IXMLDOMElement

    If (ttype = "not-wf") Then
        Set table = table_not_wf
        positive = False
    ElseIf (ttype = "error") Then
        Set table = table_informative
        positive = False
    ElseIf (ttype = "valid") Then
        Set table = table_valid
        positive = True
    ElseIf (ttype = "invalid") Then
        Set table = table_invalid
        If (Validating) Then
            positive = False
        Else
            positive = True
        End If
    End If

    ' update the relative URIs from the test database
    ' (DTD specifies they're relative to the document in
    ' which the URI is found)
    Dim base As String
    base = GetBaseURI(docBase, element)

    uri = base & "/" & uri
    Form1.Status.Caption = "Running '" & uri & "'" ' node.getAttribute("ID")
    
    If (expectedOutput = "" Or ttype <> "valid") Then
        expectedOutput = ""
    ElseIf (expectedOutput <> "") Then
        expectedOutput = base & "/" & expectedOutput
    End If

    ' report all results, but "error" cases don't affect conformance
    If (ttype <> "error") Then
        Total = Total + 1
        If (expectedOutput <> "") Then
            Total = Total + 1
        End If
    End If

    ' Setup the VBSAXXMLReader with all the right handlers and options
    Dim ContentHandler As ContentHandler
    Set ContentHandler = New ContentHandler
    Dim reader As VBSAXXMLReader30
    Set reader = New VBSAXXMLReader30
    Call reader.putFeature("http://xml.org/sax/features/external-parameter-entities", True)
    Call reader.putFeature("http://xml.org/sax/features/external-general-entities", True)
    Set reader.ContentHandler = ContentHandler
    Set reader.errorHandler = ContentHandler
    Set reader.dtdHandler = ContentHandler

    If (positive And expectedOutput <> "") Then
        ' Then we need to tell the ContentHandler to save the output.
        Dim fso As New FileSystemObject
        On Error Resume Next
        Kill tempFile
        On Error GoTo BadFile
        Set SaxOutput = fso.OpenTextFile(tempFile, ForWriting, True, TristateTrue)
        Set ContentHandler.SaxOutput = SaxOutput
    End If
    
    ' Now actually parse the test file.
    On Error GoTo Error
    reader.parseURL (uri)         ' Parse it
    
    If (positive And expectedOutput <> "") Then SaxOutput.Close

    ' Now let's check the results.
    If (Not positive) Then
        ' always report informative cases
        
        If (ttype = "error") Then
            Diagnose element, table, "INFORMATIVE", id, uri, ContentHandler.Error
        ' otherwise, succeeding is incorrect ...
        ElseIf (Not ContentHandler.FoundError) Then
            Failed = Failed + 1
            Diagnose element, table, "NEGFAIL", id, uri, ""
            
        ' and failing is correct, but the reason must be correct
        Else
            Diagnose element, table, "?PASS", id, uri, ContentHandler.Error
            NegPass = NegPass + 1
        End If
    ' positive tests that fail are easy too
    ElseIf (ContentHandler.FoundError) Then
        Failed = Failed + 1
        Diagnose element, table, "FAIL", id, uri, ContentHandler.Error
        If (expectedOutput <> "") Then
            Failed = Failed + 1 ' we also failed an output test then too !!
            Diagnose element, table_valid_output, "OUTPUT", id, uri, "[ input failed, no output to test ]"
        End If
    ElseIf (expectedOutput <> "") Then
        ' Now compare the output file to our in-memory
        ' version in contentHandler.SaxOutput
        Dim diff As String
        diff = CompareResults(tempFile, ContentHandler.DocType.XmlString, expectedOutput)
        If (diff <> "") Then
            OutputFailed = OutputFailed + 1
            Diagnose element, table_valid_output, "OUTPUT", id, uri, diff
        End If
        Kill tempFile
    End If
    Exit Sub
BadFile:
    Diagnose element, table_valid_output, "OUTPUT", id, uri, "Error creating output file: " & tempFile
    Exit Sub
Error:
    If (Not ContentHandler.FoundError) Then
        ContentHandler.FoundError = True
        ContentHandler.Error = "[ " & Err.Number & ": " & Err.Description & "]"
    End If
    Resume Next
End Sub

Private Function CompareResults(OutFile As String, DocType As String, ByVal url As String) As String
    
    ' Load the expected output (which is in utf-8)
    ' into a safearray
    
    Dim http As New XMLHTTP
    Call http.Open("GET", url, False)
    http.send
    If (http.statusText <> "OK") Then
        CompareResults = "[Error loading output file: " + url + "]"
        GoTo Error
    End If
    ' the nice thing is that the XMLHTTP object
    ' decodes this as utf-8 by default :-)
    Dim expected As String
    expected = http.responseText
       
    ' and load the sax output back into memory.
    ' (this is actually faster than building an in-memory
    ' VB string !!
    'On Error Resume Next
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(OutFile, ForReading, True, TristateTrue)
    Dim text As String
    text = DocType & ts.ReadAll  ' combine doctype and body for complete result.
    ts.Close
            
    ' compare with in memory string
    If (text <> expected) Then
        ' find where the difference is and report characters in the local area.
        Dim i, j, k, x, y
        Dim ch As String
        i = 1
        j = 1
        k = 1
        x = Len(text)
        y = Len(expected)
        While (i <= x And i <= y And Mid(text, i, 1) = Mid(expected, i, 1))
            i = i + 1
        Wend
        j = i
        ' scan ahead to next matching char
        While (j <= x And i < y And Mid(text, j, 1) <> Mid(expected, i + 1, 1))
            j = j + 1
        Wend
        k = i
        ' scan ahead to the other way
        While (k <= y And i < x And Mid(text, i + 1, 1) <> Mid(expected, k, 1))
            k = k + 1
        Wend
        If (k < j) Then j = k
        
        ' now we have the change pretty much contained with i and j
        ' now we need to add reasonable context, so we scan backwards for
        ' the previous space or syntax character and same forwards.
        Const syntax = "<>&=""'"
        While (i > 1)
            ch = Mid(text, i, 1)
            If (InStr(syntax, ch) > 0) Then
                GoTo 12
            End If
            i = i - 1
        Wend
12:
        ' now do the same thing scanning forwards, unlike scanning backwards we
        ' have to scan both buffers forwards because they are different.
        k = j
        While (j < x)
            ch = Mid(text, j, 1)
            If (InStr(syntax, ch) > 0) Then
                j = j - 1
                GoTo 13
            End If
            j = j + 1
        Wend
13:
        While (k < y)
            ch = Mid(expected, k, 1)
            If (InStr(syntax, ch) > 0) Then
                k = k - 1
                GoTo 14
            End If
            k = k + 1
        Wend
14:
        x = j - i + 1
        y = k - i + 1
        CompareResults = """" & Mid(text, i, x) & """ != """ & Mid(expected, i, y) & """"
        GoTo Error
    End If
    
    CompareResults = ""
    Exit Function
Error:
End Function

Private Function PruneURI(uri) As String
    Dim temp As Integer
    temp = InStrRev(uri, "/")
    If (temp = 0) Then temp = InStrRev(uri, "\")
    If (temp = 0) Then
        PruneURI = ""
    Else
        PruneURI = Mid$(uri, 1, temp - 1)
    End If
End Function


' This function tries to determine if a URI is absolute
Function IsAbsoluteURI(uri) As Boolean
    '  search for the first forward or reverse slash
    Dim p, q As Integer
    p = InStr(uri, "/")
    If (p = 0) Then p = InStr(uri, "\")
    If (p = 0) Then
        IsAbsoluteURI = False
    ElseIf (p = 1) Then
        IsAbsoluteURI = True
    '  preceding colon, is absolute
    ElseIf (":" = Mid$(uri, p - 1, 1)) Then
        IsAbsoluteURI = True
    Else
        IsAbsoluteURI = False
    End If
End Function

' Return the node's base URI, suitable for appending "/foo/bar.xml"
' Constructs an "xml:base" URI as it's searching for the real root
' node of the this entity (or document).  If it's absolute, it'll
' be returned in certain cases.
Function GetBaseURI(ByVal docBase As String, ByVal node As IXMLDOMNode) As String

    Dim xmlBase As String
    Dim realBase As String

    docBase = Replace(docBase, "\", "/")
outer:
    While Not node Is Nothing
        If (node.nodeType = NODE_ENTITY_REFERENCE) Then
            Dim entity As IXMLDOMEntity
            Set entity = node.ownerDocument.DocType.entities.getNamedItem(node.nodeName)

            ' DOM internal error?
            If (entity Is Nothing) Then
                MsgBox "** Can't find entity!!"
            ElseIf (entity.systemId = "" Or entity.systemId = "") Then
                MsgBox "** Bogus system ID!!"
            ElseIf (realBase = "") Then
                realBase = entity.systemId
            Else
                ' prepend to existing relative URI
                Dim entityBase As String
                entityBase = entity.systemId
                entityBase = PruneURI(entityBase)
                entityBase = entityBase & "/"
                realBase = entityBase & realBase
            End If
            If (IsAbsoluteURI(realBase)) Then
                docBase = realBase
                realBase = ""
            End If
        End If
        Set node = node.parentNode
    Wend

    ' found entity ref sequence and an absolute URI at the top?
    docBase = PruneURI(docBase)
    If (realBase <> "" And IsAbsoluteURI(docBase)) Then
        GetBaseURI = PruneURI(docBase & "/" & realBase)
    ' how about absolute URIs in xml:base hint sequences?
    ElseIf (xmlBase <> "" And IsAbsoluteURI(xmlBase)) Then
        GetBaseURI = xmlBase
    ' entity ref sequence with relative URI at top?
    ' sometimes the out-of-band PWD-style info won't change.
    ElseIf (realBase <> "") Then
        GetBaseURI = PruneURI(docBase & "/" & realBase)
    Else
        ' already pruned, doing it again would lose a part of the path
        GetBaseURI = docBase
    End If
    
End Function

' Save the diagnostic into the specified (X)HTML table.
Private Sub Diagnose(element As IXMLDOMElement, table As IXMLDOMElement, ByVal label As String, ByVal id As String, ByVal uri As String, ByVal parseError As String)
    Dim tr, td As IXMLDOMElement
    Set tr = testresults.createElement("tr")
    Dim text As String

    ' if (instr(parseError, "
    
    ' first column is test section/chapter (if known)
    ' Output tests do not report the XML spec section.
    If (label <> "OUTPUT") Then
        text = element.getAttribute("SECTIONS")
        Set td = testresults.createElement("td")
        td.text = text
        Call tr.appendChild(td)
    End If
    
    ' second column is test ID
    Set td = testresults.createElement("td")
    td.text = id
    Call tr.appendChild(td)

    ' third column is test description -- we drop any formatting
    ' hack:  if no element, we skip this
    ' NOTE:  "element.text" is specific to MSXML
    Set td = testresults.createElement("td")
    td.text = element.text
    Call tr.appendChild(td)

    ' fourth column is the diagnostic reason
    ' Output tests do not report the XML spec section.
    If (label <> "OUTPUT") Then
        ' if label is "INFORMATIVE", conformance is not at issue
        Set td = testresults.createElement("td")
        If (label <> "INFORMATIVE") Then
            If (parseError <> "") Then
                text = "(fatal) " & parseError
            Else
                text = "[wrongly accepted]"
            End If
        Else
            If (parseError <> "") Then
                text = parseError
            Else
                text = "[accepted]"
            End If
        End If
    Else
        text = parseError
    End If
    
    ' XXX if text starts as null or empty, use &nbsp;
    ' or better yet its char ref equivalent ... empty
    ' html table cells often render oddly

    ' highight diagnostic as an error, unless it's a
    ' "?PASS" (negative test that failed -- maybe for the right reason)
    ' or is purely informative (conformance shouldn't be an issue)

    If (label <> "?PASS" And label <> "INFORMATIVE") Then
        td.text = text
        Call td.setAttribute("hilite", "true")
    Else
        td.text = text
    End If
    Call tr.appendChild(td)

    ' update that particular result table
    ' XXX if we know the section & rule, insert it directly into
    ' the appropriate part of that table, so it's always sorted

    table.appendChild (tr)

End Sub

Sub GenerateReport()
    
    Dim xml As String
    xml = xml & "<results>"
    
    ' info about the parser and what mode it was in and the
    ' test run date, etc.
    xml = xml & "<name>MSXML3 SAX from VB</name>"
    xml = xml & "<description>Microsoft MSXML3 Beta Release, SAX interface from Visual Basic 6.0</description>"
    xml = xml & "<class>MSXML2.VBSAXXMLReader</class>"
    xml = xml & "<general-entities>included</general-entities>"
    xml = xml & "<parameter-entities>included</parameter-entities>"
    xml = xml & "<type>Non Validating</type>"
    xml = xml & "<date>" & FormatDateTime(Date, vbLongDate) & " " & FormatDateTime(Time, vbLongTime) & "</date>"
    xml = xml & "<harness>VB SAX Test Harness 1.0</harness>"
    xml = xml & "<runtime>Microsoft VB 6.0</runtime>"
    xml = xml & "<os>" & Environ("OS") & "</os>"
    xml = xml & "<testsuite>OASIS draft v1.0; with 5 Feb 2000 updates by David Brownell and 15 Aug 2000 updated by Chris Lovett</testsuite>"
    
    ' ... test result info
    Dim passed As Integer
    passed = Total - Failed - OutputFailed
    
    xml = xml & "<percent>" & FormatPercent(passed / Total, 0) & "</percent>"
    xml = xml & "<passed>" & passed & "</passed>"
    xml = xml & "<negpass>" & NegPass & "</negpass>"
    xml = xml & "<failed>" & (Failed + OutputFailed) & "</failed>"
    xml = xml & "<skipped>" & Skipped & "</skipped>"
    
    xml = xml & "</results>"
    
    If (Not testresults.loadXML(xml)) Then
        MsgBox "woops!"
    End If
    
    Dim Root As IXMLDOMElement
    Set Root = testresults.documentElement
    
    Call Root.appendChild(table_valid.cloneNode(True))
    Call Root.appendChild(table_valid_output.cloneNode(True))
    Call Root.appendChild(table_invalid.cloneNode(True))
    Call Root.appendChild(table_not_wf.cloneNode(True))
    Call Root.appendChild(table_informative.cloneNode(True))
    
    Dim xsl As New DOMDocument
    Dim xslfile As String
    xslfile = App.Path & "\" & "template.xsl"
    Dim res As Variant
    If (Not xsl.Load(xslfile)) Then
        res = LoadResData(101, "XSL")
        xsl.preserveWhiteSpace = True
        If (Not xsl.Load(res)) Then
            MsgBox "Internal error loading XSL resource"
        End If
        xsl.Save (xslfile)
    End If
    
    Dim result As String
    result = testresults.transformNode(xsl)
    
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(reportfile, ForWriting, True, TristateTrue)
    ts.Write (result)
    ts.Close

    Call Shell("""c:\Program Files\Internet Explorer\IEXPLORE.EXE"" " & reportfile, vbNormalFocus)

End Sub



