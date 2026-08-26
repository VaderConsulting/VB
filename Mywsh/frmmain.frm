VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const GLOBALWSH = "global.wsh"

Private Sub Form_Load()
    Dim strFile As String       ' the script to execute
    
    ' no script specified
    strFile = LCase(Command$)
    If strFile = "" Then End
    
    ' Remove quotes (if any)
    Dim temp As String
    temp = Mid(strFile, 2, Len(strFile) - 2)
    strFile = temp

    ' VBS or JS?
    If Right(strFile, 4) = ".js" Then sc.Language = "JScript"
    
    ' Read GLOBAL.WSH if any
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(App.Path & "\" & GLOBALWSH)
    temp = f.ReadAll()
    f.Close
    Set fso = Nothing
    Set f = Nothing
    
        ' Split by VbCrLf (and skip item 0)
    Dim aObjects() As String
    Dim i As Integer
    aObjects = Split(temp, vbCrLf)
    For i = 1 To UBound(aObjects)
        Dim obj As Object
        ' Add objects to the script namespace
        Dim aElems() As String
        aElems = Split(aObjects(i), "=")
        Set obj = CreateObject(aElems(1))
        sc.AddObject aElems(0), obj
    Next
    
    ' Execute the script file
    Dim strScript As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(strFile)
    strScript = f.ReadAll()
    f.Close
    
    ' Execute
    sc.ExecuteStatement strScript
    
    ' Close
    Set f = Nothing
    Set fso = Nothing
    End
End Sub


