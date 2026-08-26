VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************************************************
' Function: Create
' Usage:    Creates an XML Doc
' Author:   D Robinson
' Date:     23 October 2002
' Project:  NAT
' Client:   Internal
' Version:  1.0
'****************************************************************************
Public Function Create(bResult As Boolean) As DOMDocument
    ' Create objects
    Dim xDoc As DOMDocument30

    On Error GoTo Err
    
    bResult = False
    Set xDoc = CreateObject("Msxml2.DOMDocument")     ' Create new XML object
    ' ***********************************************************************************
    bResult = True
    
    Set Create = xDoc
    
    Set xDoc = Nothing
    Exit Function
Err:
    Set xDoc = Nothing
    Err.Clear
End Function

'****************************************************************************
' Function: Append
' Usage:    Appends to an XML Doc
' Author:   D Robinson
' Date:     23 October 2002
' Project:  NAT
' Client:   Internal
' Version:  1.0
'****************************************************************************
Public Function Append(strIn As String, strElement As String, XMLDoc As DOMDocument) As Boolean
    ' Create objects
    Dim xRootElement As IXMLDOMElement
    
    On Error GoTo Err
    
    Append = False
    
    Set xRootElement = XMLDoc.createElement(strElement)      ' Create Root element under XML Doc
    xRootElement.Text = strIn
    ' ***********************************************************************************
    XMLDoc.appendChild xRootElement                     ' Save to Document (in memory)
    Append = True
    
    Set xRootElement = Nothing
    Exit Function
Err:
    Set xRootElement = Nothing
    Err.Clear
End Function

'****************************************************************************
' Function: Read
' Usage:    Parses an XML Doc, returning a Dictionary object with the name/value pairs
' Author:   D Robinson
' Date:     23 October 2002
' Project:  NAT
' Client:   Internal
' Version:  1.0
'****************************************************************************
Public Function Read(XMLDoc As DOMDocument, dicOut As Dictionary) As Boolean
    ' Create objects
    Dim bResult As Boolean, n As Variant
    
    On Error GoTo Err
    
    Read = False
    XMLDoc.async = False                                        ' Wait for complete XML document to be loaded
    For Each n In XMLDoc.childNodes
        dicOut.Add n.nodeName, n.Text
    Next
    ' ***********************************************************************************
    Read = True
    Exit Function
Err:
    Err.Clear
End Function

'****************************************************************************
' Function: Load
' Usage:    Loads an XML Doc
' Author:   D Robinson
' Date:     23 October 2002
' Project:  NAT
' Client:   Internal
' Version:  1.0
'****************************************************************************
Public Function Load(bResult As Boolean, strFilename As String) As DOMDocument
    ' Create objects
    Dim xDoc As MSXML2.DOMDocument

    On Error GoTo Err
    
    bResult = False
    Set xDoc = CreateObject("Msxml2.DOMDocument")     ' Create new XML object
    xDoc.async = False
    bResult = xDoc.Load(LCase(strFilename))
    If bResult Then
        Set Load = xDoc
    End If
    ' ***********************************************************************************
    
    Set xDoc = Nothing
    Exit Function
Err:
    Set xDoc = Nothing
    Err.Clear
End Function

'****************************************************************************
' Function: Save
' Usage:    Saves an XML Doc
' Author:   D Robinson
' Date:     23 October 2002
' Project:  NAT
' Client:   Internal
' Version:  1.0
'****************************************************************************
Public Function Save(strFilename As String, XMLDoc As DOMDocument) As Boolean
    On Error GoTo Err
    
    XMLDoc.async = False
    XMLDoc.Save (LCase(strFilename))
    
    Save = True
    Exit Function
Err:
    Err.Clear
End Function

