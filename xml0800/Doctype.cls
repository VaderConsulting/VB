VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class stores DTD information so that it can generate a string
' representation for comparison against expected output in the OASIS
' test suite.

' private store of notations so we can sort them when we're done.
Private Notations() As String
Private NotationDecls As Dictionary
Private NotationCount As Integer
Private NotationSize As Integer
Private NewLine As String
Public Name As String

Private Sub Class_Initialize()
    Set NotationDecls = New Dictionary
    NewLine = Chr(13) & Chr(10)
End Sub

Public Sub AddNotation(Name As String, blob As String)
    If (NotationCount = NotationSize) Then
        ReDim Notations(NotationCount + 10)
        NotationSize = NotationSize + 10
    End If
    Notations(NotationCount) = Name
    NotationCount = NotationCount + 1
    Call NotationDecls.Add(Name, blob)
End Sub

Public Property Get XmlString() As String
    Dim result As String
    If (NotationCount > 0) Then
        result = "<!DOCTYPE " & Name & " [" & NewLine
        Dim i As Integer
        Dim s As Sorter
        Set s = New Sorter
        Call s.Quicksort(Notations, 0, NotationCount - 1)
        For i = 0 To NotationCount - 1
            result = result & NotationDecls.Item(Notations(i)) & NewLine
        Next
        result = result & "]>" & NewLine
    End If
    XmlString = result
End Property

