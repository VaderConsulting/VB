VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVBSAXContentHandler
Implements IVBSAXDTDHandler
Implements IVBSAXErrorHandler

Public SaxOutput As TextStream
Private Level As Integer
Public FoundError As Boolean
Public Error As String
Public DocType As DocType

Private Sub WriteEscaped(text As String)
    Dim i As Integer
    For i = 1 To Len(text)
        Dim ch As String
        ch = Mid(text, i, 1)
        Select Case ch
        Case "&"
            SaxOutput.Write ("&amp;")
        Case "<"
            SaxOutput.Write ("&lt;")
        Case ">"
            SaxOutput.Write ("&gt;")
        Case """"
            SaxOutput.Write ("&quot;")
        Case Chr(9)
            SaxOutput.Write ("&#9;")
        Case Chr(10)
            SaxOutput.Write ("&#10;")
        Case Chr(13)
            SaxOutput.Write ("&#13;")
        Case Else
            SaxOutput.Write (ch)
        End Select
    Next
End Sub

Private Sub Class_Initialize()
    Set SaxOutput = Nothing
    Level = 0
    Set DocType = New DocType
End Sub

Private Sub IVBSAXContentHandler_startElement(ByVal strNamespaceURI As String, ByVal strLocalName As String, ByVal strQName As String, ByVal Attributes As IVBSAXAttributes)
    Dim i As Integer
    Dim count As Integer
    If (Not SaxOutput Is Nothing) Then
        If (Level = 0) Then DocType.Name = strQName
        Level = Level + 1
        Dim starttag As String
        SaxOutput.Write ("<")
        SaxOutput.Write (strQName)
        count = Attributes.length
        If (count > 0) Then
            ' Now to canonicalize the output, we need to sort the attributes.
            If (Attributes.length > 1) Then
                Dim names() As String
                ReDim names(count)
                For i = 0 To count - 1
                    names(i) = Attributes.getQName(i)
                Next
                
                Dim s As Sorter
                Set s = New Sorter
                Call s.Quicksort(names, 0, count - 1)
                For i = 0 To count - 1
                    SaxOutput.Write (" ")
                    SaxOutput.Write (names(i))
                    SaxOutput.Write ("=""")
                    WriteEscaped (Attributes.getValueFromQName(names(i)))
                    SaxOutput.Write ("""")
                Next
            Else
                ' optimization for single attributes case.
                SaxOutput.Write (" ")
                SaxOutput.Write (Attributes.getQName(0))
                SaxOutput.Write ("=""")
                WriteEscaped (Attributes.getValue(0))
                SaxOutput.Write ("""")
            End If
        End If
        SaxOutput.Write (">")
    End If
End Sub

Private Sub IVBSAXContentHandler_endElement(ByVal strNamespaceURI As String, ByVal strLocalName As String, ByVal strQName As String)
    If (Not SaxOutput Is Nothing) Then
        Level = Level - 1
        SaxOutput.Write ("</")
        SaxOutput.Write (strQName)
        SaxOutput.Write (">")
    End If
End Sub

Private Sub IVBSAXContentHandler_characters(ByVal text As String)
    If (Not SaxOutput Is Nothing) Then
        WriteEscaped (text)
    End If
End Sub

Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As IVBSAXLocator)
End Property

Private Sub IVBSAXContentHandler_endDocument()
End Sub

Private Sub IVBSAXContentHandler_endPrefixMapping(ByVal strPrefix As String)
    
End Sub

Private Sub IVBSAXContentHandler_ignorableWhitespace(ByVal strChars As String)
    If (Not SaxOutput Is Nothing) Then SaxOutput.Write (strChars)
End Sub

Private Sub IVBSAXContentHandler_processingInstruction(ByVal target As String, ByVal data As String)
    If (Not SaxOutput Is Nothing) Then
        SaxOutput.Write ("<?")
        SaxOutput.Write (target)
        SaxOutput.Write (" ")
        SaxOutput.Write (data)
        SaxOutput.Write ("?>")
    End If
End Sub

Private Sub IVBSAXContentHandler_skippedEntity(ByVal strName As String)

End Sub

Private Sub IVBSAXContentHandler_startDocument()

End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping(ByVal strPrefix As String, ByVal strURI As String)
End Sub

Private Sub IVBSAXDTDHandler_notationDecl(ByVal strName As String, ByVal strPublicId As String, ByVal strSystemId As String)
    If (Not SaxOutput Is Nothing) Then
        Dim notation As String
        notation = "<!NOTATION "
        notation = notation & strName
        If (strPublicId <> "") Then
            notation = notation & " PUBLIC '" & strPublicId & "'"
        End If
        If (strPublicId = "" And strSystemId <> "") Then
            notation = notation & " SYSTEM"
        End If
        If (strSystemId <> "") Then
            notation = notation & " '" & strSystemId & "'"
        End If
        notation = notation & ">"
        Call DocType.AddNotation(strName, notation)
    End If
End Sub

Private Sub IVBSAXDTDHandler_unparsedEntityDecl(ByVal strName As String, ByVal strPublicId As String, ByVal strSystemId As String, ByVal strNotationName As String)
End Sub

Private Sub IVBSAXErrorHandler_fatalError(ByVal lctr As IVBSAXLocator, ByVal msg As String, ByVal errCode As Long)
    Dim errmsg As String
    errmsg = "Error"
    If (lctr.lineNumber <> 0) Then
        errmsg = errmsg & " on line " & lctr.lineNumber
    End If
    If (lctr.columnNumber <> 0) Then
        errmsg = errmsg & " position " & lctr.columnNumber
    End If
    Error = errmsg & ": " & msg
    FoundError = True
End Sub

Private Sub IVBSAXErrorHandler_error(ByVal lctr As IVBSAXLocator, ByVal msg As String, ByVal errCode As Long)
End Sub

Private Sub IVBSAXErrorHandler_warning(ByVal lctr As IVBSAXLocator, ByVal msg As String, ByVal errCode As Long)
End Sub

