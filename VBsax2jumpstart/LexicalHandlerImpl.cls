VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LexicalHandlerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements IVBSAXLexicalHandler


Private Sub IVBSAXLexicalHandler_comment(strChars As String)

End Sub

Private Sub IVBSAXLexicalHandler_startDTD(strName As String, strPublicId As String, strSystemId As String)

End Sub

Private Sub IVBSAXLexicalHandler_endDTD()

End Sub

Private Sub IVBSAXLexicalHandler_startEntity(strName As String)

End Sub

Private Sub IVBSAXLexicalHandler_endEntity(strName As String)

End Sub

Private Sub IVBSAXLexicalHandler_startCDATA()

End Sub

Private Sub IVBSAXLexicalHandler_endCDATA()

End Sub

