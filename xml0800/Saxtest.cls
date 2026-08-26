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

Public Total As Integer
Public Validating As Boolean

Public Sub Run(testuri As String)
    Dim doc As New MSXML2.DOMDocument
    If (Not doc.Load(testuri)) Then
        MsgBox "Error: " & doc.parseError.reason
        GoTo Error
    End If
    
    Dim found As IXMLDOMNodeList
    Set found = doc.selectNodes("//TEST")
    Dim Total As Integer
    Total = found.length
    Set found = doc.selectNodes("//TEST[@OUTPUT]")
    Total = Total + found.length
    Set found = doc.selectNodes("//TEST[@TYPE='error']")
    Total = Total - found.length
           
    Set found = doc.selectNodes("//TEST")
    Dim node As IXMLDOMElement
    Set node = found.nextNode()
    While Not node Is Nothing
        Dim url As String
        RunTest testuri, node
        DoEvents
        'SAX.RunTest (node)
        Set node = found.nextNode()
    Wend
Error:
    Exit Sub
End Sub

Public Sub RunTest(docBase As String, element As IXMLDOMElement)
    ' where will output diagnostic go?
    'var table;

    ' fetch basic test attributes ...
    Dim id, ttype, output, uri As String
    id = element.getAttribute("ID")
    ttype = element.getAttribute("TYPE")
    output = element.getAttribute("OUTPUT")
    uri = element.getAttribute("URI")

    ' NOTE: OASIS/NIST omitted the OUTPUT3 data cases, which should
    ' be used with validating parsers (which have special requirements
    ' for ignorable whitespace and unparsed entities -- untested here).
    ' ignoring ENTITIES for the moment (assuming they are always read)
    ' is this a positive (must-succeed) test?
    Dim positive As Boolean

    If (ttype = "not-wf") Then
        'table = table_not_wf
        positive = False
    ElseIf (ttype = "error") Then
        'table = table_informative
        positive = False
    ElseIf (ttype = "valid") Then
        'table = table_valid
        positive = True
    ElseIf (ttype = "invalid") Then
        'table = table_invalid
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
    Form1.Prompt.Caption = "Running '" & uri & "'" ' node.getAttribute("ID")
    
    If (output = "" Or ttype <> "valid") Then
        output = Null
    ElseIf (output <> Null) Then
        output = base & "/" & output
    End If

    ' report all results, but "error" cases don't affect conformance
    If (ttype <> "error") Then
        Total = Total + 1
        If (output <> "") Then
            Total = Total + 1
        End If
    End If

    ' Parse the test case
    ' NOTE:  this relies on MSXML-specific
    ' methods to load XML text and detect errors
  
    Dim reader As MSXML2.VBSAXXMLReader
    Dim contentHandler As contentHandler
    Dim errorHandler As errorHandler
    
    Set reader = New MSXML2.VBSAXXMLReader

    ' These ones will receive parsing events and errors
    Set contentHandler = New contentHandler
    Set errorHandler = New errorHandler

    Set reader.contentHandler = contentHandler     ' And they should
    Set reader.errorHandler = errorHandler         ' work together

    On Error GoTo Error            ' Trap errors

    'reader.parseURL (uri)         ' Parse it
    
Error:

End Sub

Private Sub Class_Initialize()
    Total = 0
    Validating = False
End Sub

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
            Set entity = node.ownerDocument.doctype.entities.getNamedItem(node.nodeName)

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

