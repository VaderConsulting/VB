VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public Sub run(url As String)

    ' Make sure the VBSAXXMLReader is available.
    Dim reader As VBSAXXMLReader30
    Set reader = New VBSAXXMLReader30
    
    ClearFields
    
    Dim handler As Filter
    Set handler = New Filter
    Set reader.ContentHandler = handler
    Set reader.errorHandler = handler
        
    Form1.Status.Caption = "Running..."
    Form1.Command1.Enabled = False
    DoEvents
    
    On Error GoTo Error
    reader.parseURL (url)         ' Parse it
    
    Form1.Text2(0).text = handler.Elements
    Form1.Text2(1).text = handler.AttributeNodes
    Form1.Text2(2).text = handler.TextNodes
    Form1.Text2(3).text = handler.TextChars
    Form1.Text2(4).text = handler.NameChars
    
    ' Now calculate a conservative estimate of the number of
    ' characters involved in markup, like element and attribute
    ' names and the special characters surrounding them (like <>=")
    Dim syntaxChars As Long
    syntaxChars = handler.NameChars + (handler.Elements * 2) + (handler.AttributeNodes * 3)
    Form1.Percent.Caption = FormatPercent(syntaxChars / (syntaxChars + handler.TextChars), 0)
    Form1.Percent.BackColor = &HFFFFFF
    Form1.Status.Caption = "Done"
Cleanup:
    Form1.Command1.Enabled = True
    Exit Sub
Error:
    Form1.Status.Caption = "Error: " & Err.Number & ": " & Err.Description
    GoTo Cleanup
End Sub

Private Sub ClearFields()
    Dim i As Integer
    For i = 0 To Form1.Text2.count - 1
        Form1.Text2(i).text = ""
    Next
    Form1.Percent.BackColor = &H8000000F
End Sub

