VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reader"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim oXML As NAT.XML, xDoc As MSXML2.DOMDocument
    Dim XML As String, bResult As Boolean, n As Variant
    Dim myDictionary As New Scripting.Dictionary
    Dim i As Integer, Nodes As Variant, Keys As Variant
    
    myDictionary.CompareMode = TextCompare
    
    Set oXML = CreateObject("NAT.XML")
    Set xDoc = oXML.Load(bResult, "c:\XMLTest.XML")
    bResult = oXML.Append("myValue", "myNode", xDoc)
    oXML.save "c:\xmltest.xml", xDoc

    bResult = oXML.Read(xDoc, myDictionary)
    Nodes = myDictionary.Items
    Keys = myDictionary.Keys
    For i = 0 To myDictionary.Count - 1
        Debug.Print myDictionary.Keys(i), Nodes(i)
    Next
End Sub
