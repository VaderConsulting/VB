VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMain 
   Caption         =   "Scripting Control Demo"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtArguments 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   2100
      Width           =   2715
   End
   Begin VB.TextBox txtOutput 
      Height          =   1635
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   4575
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   2340
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdRunScript 
      Caption         =   "Run Script"
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   1740
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "Argument List:"
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   1740
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Output:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2340
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Demo of the Microsoft Scripting Control.
'Copyright 2002 CNET Networks. It may be freely
'distributed so long as this copyright notice remains.
'
'Editorial Disclaimer: Although all reasonable efforts have been
'made to ensure that this code is safe and correct, CNET Networks
'and Builder.com cannot be held responsible for any damages incurred
'by your use of this code. In other words, run at your own risk.

Option Explicit

Private Sub cmdRunScript_Click()
    Dim f As Long               'file pointer
    Dim strCode As String       'actual script code
    Dim strTemp As String       'temp variable
    Dim strFileName As String   'selected file name
    Dim strProcName As String   'procedure name of "main" sub in strFileName
    
    If List1.ListIndex <> -1 Then
        'This program assumes there will be a sub in the vbs
        'file with the same name as the file, this
        'doesn't necessarily have to be so. But it made
        'things easier for me in writing this demo
        strFileName = List1.List(List1.ListIndex)
        strProcName = Left(strFileName, InStr(strFileName, ".vbs") - 1)
        
        'Load the code from the selected script file.
        'Could actually be done in Form_Load for all
        'script files located.
        f = FreeFile()
        Open App.Path & "/" & strFileName For Input As f
        Do Until EOF(f)
            Line Input #f, strTemp
            strCode = strCode & vbCrLf & strTemp
        Loop
        Close f
        
        'Add all procedures found in the script file to the
        'script control's environment.
        ScriptControl1.AddCode (strCode)
        
        'Verify that the correct number of arguments are supplied
        'Multiple arguments should be comma-delimited
        On Error Resume Next
        If Len(Me.txtArguments & "") > 0 Then
            Err.Clear
            'try running the script with the arguments as entered.
            'if an incorrect number of arguments are supplied, the
            'control will raise an error 450 in our code, contrary
            'to documentation.
            ScriptControl1.Run strProcName, Me.txtArguments.Text
            If Err.Number = 450 Then
                MsgBox "The script you selected does not accept any arguments."
                Err.Clear
            End If
            
        Else
            'try running the script with no arguments.
            Err.Clear
            ScriptControl1.Run strProcName
            If Err.Number = 450 Then
                'Here's how to find out how many arguments a sub requires
                MsgBox "The script you selected requires " & Me.ScriptControl1.Procedures.Item(strProcName).NumArgs & " arguments."
                Err.Clear
            End If
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    Dim strDir As String                'Directory contents
    Dim oCustomers As New colCustomers  'play collection
    Dim oScriptOutput As New Output     'Output object for script
    
    'Get a list of all vbs script files in the app directory
    strDir = Dir(App.Path & "\*.vbs")
    Do Until strDir = ""
        Me.List1.AddItem strDir
        strDir = Dir()
    Loop
    
    'set up our scripts' play environment.
    Set oScriptOutput.OutputBox = Me.txtOutput
    'LoadData accepts a DSN name, a user id and a password
    'oCustomers.LoadData "L2KSDF00356SQL", "sa", ""
    'Make these objects available to scripts run with the
    'script control.
    'ScriptControl1.AddObject "Customers", oCustomers, True
    'ScriptControl1.AddObject "Output", oScriptOutput, True
    
End Sub

Private Sub ScriptControl1_Error()
    'Not really necessary unless ScriptControl1.AllowUI is set to false
    MsgBox "Script Error: " & Me.ScriptControl1.Error.Number & " - " & Me.ScriptControl1.Error.Description
End Sub

