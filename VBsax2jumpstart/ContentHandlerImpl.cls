VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentHandlerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements IVBSAXContentHandler

Private Sub IVBSAXContentHandler_startElement(strNamespaceURI As String, strLocalName As String, strQName As String, ByVal attributes As MSXML2.IVBSAXAttributes)
    
    Dim i As Integer
    
    Form1.Text2.text = Form1.Text2.text & "<" & strLocalName
   
    For i = 0 To (attributes.length - 1)
        Form1.Text2.text = Form1.Text2.text & " " & attributes.getLocalName(i) & "=""" & attributes.getValue(i) & """"
    Next

    Form1.Text2.text = Form1.Text2.text & ">"

    If strLocalName = "qu" Then
        Err.Raise vbObjectError + 1, "ContentHandler.startElement", "Found element <qu>"
    End If

End Sub

Private Sub IVBSAXContentHandler_endElement(strNamespaceURI As String, strLocalName As String, strQName As String)
    
    Form1.Text2.text = Form1.Text2.text & "</" & strLocalName & ">"

End Sub

Private Sub IVBSAXContentHandler_characters(text As String)
    text = Replace(text, vbLf, vbCrLf)
    Form1.Text2.text = Form1.Text2.text & text
    
End Sub

Private Property Set IVBSAXContentHandler_documentLocator(ByVal RHS As MSXML2.IVBSAXLocator)

End Property

Private Sub IVBSAXContentHandler_endDocument()

End Sub

Private Sub IVBSAXContentHandler_endPrefixMapping(strPrefix As String)

End Sub

Private Sub IVBSAXContentHandler_ignorableWhitespace(strChars As String)

End Sub

Private Sub IVBSAXContentHandler_processingInstruction(target As String, data As String)
    
    Form1.Text2.text = Form1.Text2.text & "<?" & target & " " & data & ">"

End Sub

Private Sub IVBSAXContentHandler_skippedEntity(strName As String)

End Sub

Private Sub IVBSAXContentHandler_startDocument()

End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping(strPrefix As String, strURI As String)

End Sub
