VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeclHandlerImpl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements IVBSAXDeclHandler


Private Sub IVBSAXDeclHandler_attributeDecl(strElementName As String, strAttributeName As String, strType As String, strValueDefault As String, strValue As String)

End Sub

Private Sub IVBSAXDeclHandler_elementDecl(strName As String, strModel As String)

End Sub

Private Sub IVBSAXDeclHandler_externalEntityDecl(strName As String, strPublicId As String, strSystemId As String)

End Sub

Private Sub IVBSAXDeclHandler_internalEntityDecl(strName As String, strValue As String)

End Sub
