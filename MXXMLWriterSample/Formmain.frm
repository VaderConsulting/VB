VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "SAX Writer Usage Example"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandTryFile 
      Caption         =   "Try File"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox TextFileName 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "test.xml"
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton CommandTryDemo 
      Caption         =   "Try Demo"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox TextResult 
      Height          =   4695
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox TextSource 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "File name (for Try File only):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXMLReader As New SAXXMLReader
Dim oXMLWriter As New MXXMLWriter


Private Sub CommandExit_Click()
    End
End Sub

Private Sub CommandTryDemo_Click()
    ' We need these variables for typecasting the writer
    Dim oContentHandler As IVBSAXContentHandler
    Dim oDTDHandler As IVBSAXDTDHandler
    Dim oLexicalHandler As IVBSAXLexicalHandler
    Dim oDeclHandler As IVBSAXDeclHandler
    Dim oErrorHandler As IVBSAXErrorHandler
    
    ' That's just a helper
    Dim atrs As New SAXAttributes
    
    ' Set them all to writer, writer implements all these interfaces
    Set oContentHandler = oXMLWriter
    Set oDTDHandler = oXMLWriter
    Set oLexicalHandler = oXMLWriter
    Set oDeclHandler = oXMLWriter
    Set oErrorHandler = oXMLWriter
    
    ' Set parameters, clean the scene
    TextSource.Text = ""
    oXMLWriter.output = ""
       
    ' And manually call necessary events to generate XML file
    log "Content->startDocument"
    oContentHandler.startDocument
    log "Lexical->startDTD"
    oLexicalHandler.startDTD "MyDTD", "", "http://eureka.sample/mydtd.dtd"
        log "Decl->elementDecl"
        oDeclHandler.elementDecl "book", "title | descr"
        log "Decl->attributeDecl"
        oDeclHandler.attributeDecl "book", "author", "CDATA", "#IMPLIED", ""
        log "Decl->attributeDecl"
        oDeclHandler.attributeDecl "book", "ISBN", "CDATA", "#REQUIRED", "000000000"
        log "Decl->attributeDecl"
        oDeclHandler.attributeDecl "book", "cover", "(hard|soft)", "", "soft"
        log "Decl-elementDecl"
        oDeclHandler.elementDecl "title", "(#PCDATA)"
        log "Decl-elementDecl"
        oDeclHandler.elementDecl "descr", "(#PCDATA)"
    log "Lexical->endDTD"
    oLexicalHandler.endDTD
    log "Content->startElement"
    atrs.addAttribute "", "", "cover", "", "hard"
    oContentHandler.startElement "", "", "book", atrs
        log "Content->startElement"
        atrs.Clear
        oContentHandler.startElement "", "", "title", atrs
        log "Content->characters"
        oContentHandler.characters "On the Circular Problem of Quadratic Equations"
        log "Content->endElement"
        oContentHandler.endElement "", "", "title"
    log "Content->endElement"
    oContentHandler.endElement "", "", "book"
        
    TextResult.Text = oXMLWriter.output
End Sub

Private Sub CommandTryFile_Click()
    ' Set parameters, clean the scene
    TextSource.Text = ""
    oXMLWriter.output = ""
    oXMLWriter.omitXMLDeclaration = True
    On Error GoTo VBError
    oXMLReader.parseURL TextFileName.Text
    
    TextResult.Text = oXMLWriter.output
    Exit Sub
VBError:
    TextResult.Text = "Error: " & Err.LastDllError & " -- " & Err.Description & vbCrLf & "Check if file exists in the correct place"
End Sub

Private Sub Form_Load()
    Set oXMLReader.contentHandler = oXMLWriter
    Set oXMLReader.dtdHandler = oXMLWriter
    Set oXMLReader.errorHandler = oXMLWriter
    oXMLReader.putProperty "http://xml.org/sax/properties/declaration-handler", oXMLWriter
    oXMLReader.putProperty "http://xml.org/sax/properties/lexical-handler", oXMLWriter
End Sub

Private Sub log(msg As String)
    TextSource.Text = TextSource.Text & vbCrLf & msg
End Sub

