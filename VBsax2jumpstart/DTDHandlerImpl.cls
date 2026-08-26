VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DTDHandlerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements IVBSAXDTDHandler

Private Sub IVBSAXDTDHandler_notationDecl(strName As String, strPublicId As String, strSystemId As String)

End Sub

Private Sub IVBSAXDTDHandler_unparsedEntityDecl(strName As String, strPublicId As String, strSystemId As String, strNotationName As String)

End Sub
