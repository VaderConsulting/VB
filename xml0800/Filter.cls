VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The next line is important – it says what we implement
Implements IVBSAXContentHandler
Implements IVBSAXErrorHandler

' Statistics
Public Elements As Long
Public AttributeNodes As Long
Public TextNodes As Long
Public TextChars As Long
Public NameChars As Long

Public NamesToUpper As Boolean
Public NamesToLower As Boolean
Public NamesToProper As Boolean
Public AttrsToElements As Boolean
Public Format As Boolean
Public Compress As Boolean
Public Indentation As Integer
Public IndentChar As String

Private OutputStream As TextStream
Public wsstr As String
Private Level As Integer
Private content() As Integer
Private contentsize As Integer
Private First As Boolean

Const CONTENT_EMPTY = 0
Const CONTENT_MIXED = 1
Const CONTENT_ELEMENT = 2

Public Sub run(url As String, OutFile As String)
    If (url = OutFile) Then
        MsgBox ("Output file cannot be the same as the input file")
        GoTo Error
    End If
    Dim fso As New FileSystemObject
    On Error Resume Next
    fso.DeleteFile (OutFile)
    On Error GoTo BadFile
    If (OutFile <> "") Then
        Set OutputStream = fso.OpenTextFile(OutFile, ForWriting, True)
    End If
    
    Form1.Status.Caption = "Running..."
    Form1.Command1.Enabled = False
    DoEvents
    
    On Error GoTo ParseErr
    Dim reader As VBSAXXMLReader30
    Set reader = New VBSAXXMLReader30

    Set reader.ContentHandler = Me
    Set reader.errorHandler = Me
    
    reader.parseURL (url)         ' Parse it
ParseErr:
    Form1.Status.Caption = ""
    Form1.Command1.Enabled = True
    If (Not OutputStream Is Nothing) Then OutputStream.Close
    Exit Sub
BadFile:
    MsgBox "Error opening file: " & Err.Description
Error:
    Form1.OutFile.SelLength = Len(OutFile)
    Form1.OutFile.SetFocus
End Sub

Private Function EscContent(value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    EscContent = value
End Function

Private Sub Class_Initialize()
    wsstr = " " & Chr(9) & Chr(10) & Chr(13)
    ReDim content(10)
    contentsize = 10
    First = True
End Sub

Private Sub Push()
    If (Level = contentsize) Then
        contentsize = contentsize + 10
        ReDim content(contentsize)
    End If
    Level = Level + 1
    content(Level) = CONTENT_EMPTY
End Sub

Private Sub Pop()
    Level = Level - 1
End Sub

Private Sub IVBSAXContentHandler_characters(ByVal strChars As String)
    If (Not OutputStream Is Nothing) Then
        content(Level) = CONTENT_MIXED
        If (Compress) Then
            Dim i As Integer
            For i = 1 To Len(strChars)
                If (InStr(wsstr, Mid(strChars, i, 1)) = 0) Then
                    GoTo output
                End If
            Next
            Exit Sub ' then it is all whitespace.
        End If
output:
        OutputStream.Write (EscContent(strChars))
    Else
        TextNodes = TextNodes + 1
        TextChars = TextChars + Len(strChars)
    End If
End Sub

Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As IVBSAXLocator)

End Property

Private Sub IVBSAXContentHandler_endDocument()

End Sub

Private Sub IVBSAXContentHandler_endElement(ByVal strNamespaceURI As String, ByVal strLocalName As String, ByVal strQName As String)
    NameChars = NameChars + Len(strQName)
    If (Not OutputStream Is Nothing) Then
        Dim saved As Integer
        saved = content(Level)
        Level = Level - 1
        If (saved <> CONTENT_EMPTY) Then     ' empty tags do not need indenting
            Call WriteIndent(Level, saved)
        End If
        OutputStream.Write ("</" & FilterName(strQName) & ">")
    End If
End Sub

Private Sub IVBSAXContentHandler_endPrefixMapping(ByVal strPrefix As String)

End Sub

Private Sub IVBSAXContentHandler_ignorableWhitespace(ByVal strChars As String)

End Sub

Private Sub IVBSAXContentHandler_processingInstruction(ByVal strTarget As String, ByVal strData As String)
    If (Not OutputStream Is Nothing) Then
        Call WriteIndent(Level, content(Level))
        OutputStream.Write ("<?" & strTarget & " " & strData & "?>")
        First = False
    End If
End Sub

Private Sub IVBSAXContentHandler_skippedEntity(ByVal strName As String)

End Sub

Private Sub IVBSAXContentHandler_startDocument()

End Sub

Private Function FilterName(Name As String) As String
    If (NamesToUpper) Then
        FilterName = StrConv(Name, vbUpperCase)
    ElseIf (NamesToLower) Then
        FilterName = StrConv(Name, vbLowerCase)
    ElseIf (NamesToProper) Then
        FilterName = StrConv(Name, vbProperCase)
    Else
        FilterName = Name
    End If
End Function

Private Function EscAttr(value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, """", "&quot;")
    EscAttr = value
End Function

Private Sub WriteIndent(Level As Integer, content As Integer)
    If (Format And (Not First) And content <> CONTENT_MIXED) Then
        OutputStream.WriteLine
        OutputStream.Write (String(Level * Indentation, IndentChar))
    End If
End Sub

Private Sub IVBSAXContentHandler_startElement(ByVal strNamespaceURI As String, ByVal strLocalName As String, ByVal strQName As String, ByVal Attributes As IVBSAXAttributes)
    Dim i As Integer

    Elements = Elements + 1
    NameChars = NameChars + Len(strQName)
    AttributeNodes = AttributeNodes + Attributes.length
    For i = 0 To Attributes.length - 1
        NameChars = NameChars + Len(Attributes.getQName(i))
        TextChars = TextChars + Len(Attributes.getValue(i))
    Next
    
    If (Not OutputStream Is Nothing) Then
        If (content(Level) = CONTENT_EMPTY) Then
            content(Level) = CONTENT_ELEMENT ' true means tag contains something
        End If
        Call WriteIndent(Level, content(Level))
        Push
        OutputStream.Write ("<" & FilterName(strQName))
        Dim Name As String
        For i = 0 To Attributes.length - 1
            Name = Attributes.getQName(i)
            Dim ns As Boolean
            If (Name = "xmlns" Or Mid(Name, 1, 6) = "xmlns:") Then
                ns = True
            End If
            If (Not AttrsToElements Or ns) Then
                If (Not ns) Then Name = FilterName(Name)
                OutputStream.Write (" " & Name & "=""")
                OutputStream.Write (EscAttr(Attributes.getValue(i)))
                OutputStream.Write ("""")
            End If
        Next
        OutputStream.Write (">")
        First = False
        If (AttrsToElements) Then
            For i = 0 To Attributes.length - 1
                Name = Attributes.getQName(i)
                If (Name <> "xmlns" And Mid(Name, 1, 6) <> "xmlns:") Then
                    content(Level) = CONTENT_ELEMENT
                    Call WriteIndent(Level, content(Level))
                    OutputStream.Write ("<" & FilterName(Name) & ">")
                    OutputStream.Write (EscContent(Attributes.getValue(i)))
                    OutputStream.Write ("</" & FilterName(Name) & ">")
                End If
            Next
        End If
    End If
End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping(ByVal strPrefix As String, ByVal strURI As String)

End Sub

Private Sub IVBSAXErrorHandler_error(ByVal oLocator As IVBSAXLocator, ByVal strError As String, ByVal nErrorCode As Long)

End Sub

Private Sub IVBSAXErrorHandler_fatalError(ByVal Locator As IVBSAXLocator, ByVal strError As String, ByVal nErrorCode As Long)
    Dim errmsg As String
    errmsg = "Error"
    If (Locator.lineNumber <> 0) Then
        errmsg = errmsg & " on line " & Locator.lineNumber
    End If
    If (Locator.columnNumber <> 0) Then
        errmsg = errmsg & " position " & Locator.columnNumber
    End If
    MsgBox errmsg & ": " & strError
End Sub

Private Sub IVBSAXErrorHandler_warning(ByVal oLocator As IVBSAXLocator, ByVal strError As String, ByVal nErrorCode As Long)

End Sub
