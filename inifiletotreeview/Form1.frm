VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Ini file in Treeview"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Use the right mouse button to add or delete a section/key or to change the key value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu MnuMouse 
      Caption         =   "RightMouseButton"
      Visible         =   0   'False
      Begin VB.Menu MnuAddKey 
         Caption         =   "Add Key"
      End
      Begin VB.Menu MnuModifyKey 
         Caption         =   "Modify Key"
      End
      Begin VB.Menu MnuDeleteKey 
         Caption         =   "Delete Key"
      End
      Begin VB.Menu MnuAddSection 
         Caption         =   "Add Section"
      End
      Begin VB.Menu MnuDeleteSection 
         Caption         =   "Delete Section"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ObjectType
    otNone = 0
    otRoot = 1
    otSection = 2
    otKey = 3
End Enum

Private SourceNode As Node
Private SourceType As ObjectType

Dim strIniFile As String

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As Any, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long

Private Sub CmdAdd_Click()

    CmdAdd.Enabled = False
    CmdDelete.Enabled = False

End Sub

Private Sub Form_Load()
    
Dim strWindowsDir As String
Dim bSuccess As Boolean
    
    'strWindowsDir = GetWindowsDir()
    'strIniFile = strWindowsDir & "\Test.ini"
    strIniFile = App.Path & "\test.ini"
    If FileExists(strIniFile) Then
        LoadIniFileInTreeview (strIniFile)
    Else
    
        MsgBox "File 'Test.ini' is missing in application folder.", vbOKOnly, "Error ..."
        Unload Me
        Set Form1 = Nothing
    
    End If

End Sub

Private Sub LoadIniFileInTreeview(Filename As String)
    
Dim arrSection As Variant, x As Integer
Dim arrKeyName As Variant, y As Integer
Dim SectionName As String
Dim Keyname As String
Dim KeyValue As String
Dim nodx As Node

On Error Resume Next
    'Set root node:
    Set nodx = TreeView1.Nodes.Add(, , "root", "Sections")
    'Get the ini file section names
    arrSection = GetSectionNames(Filename)
        
    For x = 0 To UBound(arrSection)
        SectionName = arrSection(x)
        'store each section name under the root node :
        Set nodx = TreeView1.Nodes.Add("root", tvwChild, SectionName, SectionName)
        'Get the key names of each section and store the
        'arrKeyName array :
        arrKeyName = GetKeyNames(SectionName, Filename)
        For y = 0 To UBound(arrKeyName)
            Keyname = arrKeyName(y)
            'Get the key value of each key name :
            KeyValue = GetIniValue(SectionName, Keyname, "", Filename)
            If IsNumeric(Keyname) Then Keyname = CStr(Keyname)
            'store each keyname and key value under the
            'section name node.
            'Set key node to something like
            'section name<delimiter>key name=key value' :
            Set nodx = TreeView1.Nodes.Add(SectionName, tvwChild, _
                Keyname, Keyname & "=" & KeyValue)
        Next y
    Next x
    
End Sub

Public Function GetSectionNames(Filename As String) As Variant
        
Dim x As Integer
Dim FileNumber As Integer
Dim SectionName() As String
Dim Data As String
    'Open the ini file :
    x = 0
    FileNumber = FreeFile
    Open Filename For Input As #FileNumber
    'and read each line :
    Do While Not EOF(FileNumber)
        Line Input #1, Data
        'Store the line in the SectionNames array
        'only if it begins and ends with brackets ([]):
        If Left(Data, 1) = "[" And Right(Data, 1) = "]" Then
            ReDim Preserve SectionName(x)
            'Remove the brackets :
            SectionName(x) = Mid(Data, 2, Len(Data) - 2)
            x = x + 1
        End If
    Loop
    'Close the ini file :
    Close #FileNumber
    'Return the SectionName array :
    GetSectionNames = SectionName()

End Function

Public Function GetKeyNames(section As String, Filename As String) As Variant

Dim success As Long
Dim x As Integer
Dim nSize As Long
Dim lpKeyName As String
Dim ret As String
Dim Keyname() As String

    ret = Space$(2048)
    nSize = Len(ret)
    success = GetPrivateProfileString(section, vbNullString, "", ret, nSize, Filename)
    
    If success Then
    
        ret = Left$(ret, success)
        x = 0
        Do Until ret = ""
            lpKeyName = StripNulls(ret)
            ReDim Preserve Keyname(x)
            Keyname(x) = lpKeyName
            x = x + 1
        Loop
    
    End If
    
    'Return the KeyName array :
    GetKeyNames = Keyname()

End Function

Private Sub MnuAddKey_Click()

Dim answ As String
Dim Keyname() As String

    answ = InputBox("Key ?", "Add Key")
    If InStr(1, answ, "=") > 0 Then
        Keyname() = Split(answ, "=")
        SaveIniValue TreeView1.SelectedItem, Keyname(0), Keyname(1), strIniFile
        TreeView1.Nodes.Clear
        LoadIniFileInTreeview (strIniFile)
    Else
        MsgBox "Invalid Key", vbExclamation, "Add Key"
    End If

End Sub

Private Sub MnuAddSection_Click()
    
    Dim answ As String
    answ = InputBox("Section Name ?", "Add Section")
    If Len(Trim(answ)) > 0 Then
        SaveIniValue answ, "", "", strIniFile
        TreeView1.Nodes.Clear
        LoadIniFileInTreeview (strIniFile)
    End If
    
End Sub

Private Sub MnuDeleteKey_Click()

Dim Keyname() As String
    If MsgBox("Confirm to delete key '" & TreeView1.SelectedItem.Key & _
        "'", vbYesNo, "Delete...") = vbYes Then
        Keyname() = Split(TreeView1.SelectedItem, "=")
        DeleteIniValue TreeView1.SelectedItem.Parent, Keyname(0), strIniFile
        TreeView1.Nodes.Clear
        LoadIniFileInTreeview (strIniFile)
    Else
        Exit Sub
    End If

End Sub

Private Sub MnuDeleteSection_Click()
    
    If MsgBox("Confirm to delete section '" & TreeView1.SelectedItem.Key & _
        "'", vbYesNo, "Delete...") = vbYes Then
        DeleteIniSection TreeView1.SelectedItem, strIniFile
        TreeView1.Nodes.Clear
        LoadIniFileInTreeview (strIniFile)
    End If

End Sub

Private Sub MnuModifyKey_Click()

    TreeView1.StartLabelEdit

End Sub

Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
    
Dim Keyname As String
Dim KeyValue As String

    If InStr(1, NewString, "=") > 0 Then
        NewString = Trim(NewString)
        Keyname = Trim(Mid(NewString, 1, InStr(1, NewString, "=") - 1))
        KeyValue = Trim(Mid(NewString, InStr(1, NewString, "=") + 1))
        NewString = Keyname & "=" & KeyValue
        SaveIniValue TreeView1.SelectedItem.Parent, Keyname, KeyValue, strIniFile
    Else
        MsgBox "Invalid Key", vbExclamation, "Modify Key"
        NewString = TreeView1.SelectedItem.Text
    End If

End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
Dim nde As Node

    Set SourceNode = TreeView1.HitTest(x, y)
    'For Each nde In TreeView1.Nodes
    '    nde.Bold = False
    'Next
    SourceType = NodeType(SourceNode)
    If SourceType <> otNone Then
        Set TreeView1.SelectedItem = SourceNode
    Else 'if not clicked on a node
        Set SourceNode = TreeView1.SelectedItem
        SourceType = NodeType(SourceNode)
    End If
    'SourceNode.Bold = True
    If Button = vbRightButton Then ShowPopup

End Sub

Private Function NodeType(test_node As Node) As ObjectType
' ***********************************************
' Return the node's object type.
' ***********************************************
    If test_node Is Nothing Then
        NodeType = otNone
    Else
        If test_node.Key = "root" Then
            NodeType = otRoot
        ElseIf test_node.Parent.Key = "root" Then
            NodeType = otSection
        Else
            NodeType = otKey
        End If
    End If

End Function

' Display the appropriate popup menu items.
Private Sub ShowPopup()
    ' See what kind of node this is and
    ' hide the invalid popup menu items.
    Select Case SourceType
        Case otNone
            ' Allow add section.
            MnuDeleteKey.Enabled = False
            MnuAddSection.Enabled = True
            MnuAddKey.Enabled = False
            MnuModifyKey.Enabled = False
            MnuDeleteSection.Enabled = False
        Case otRoot
            ' Allow add section.
            MnuDeleteKey.Enabled = False
            MnuAddSection.Enabled = True
            MnuAddKey.Enabled = False
            MnuModifyKey.Enabled = False
            MnuDeleteSection.Enabled = False
        Case otSection
            ' Allow delete section & add key.
            MnuDeleteKey.Enabled = False
            MnuAddSection.Enabled = False
            MnuAddKey.Enabled = True
            MnuModifyKey.Enabled = False
            MnuDeleteSection.Enabled = True
        Case otKey
            ' Allow modify & delete key.
            MnuDeleteKey.Enabled = True
            MnuAddSection.Enabled = False
            MnuAddKey.Enabled = False
            MnuModifyKey.Enabled = True
            MnuDeleteSection.Enabled = False
    End Select
    ' Display the popup menu.
    PopupMenu MnuMouse

End Sub

