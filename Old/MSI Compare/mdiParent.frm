VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiParent 
   BackColor       =   &H8000000C&
   Caption         =   "MSI Explorer"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9675
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
         Begin VB.Menu mnuDatabase 
            Caption         =   "Database"
         End
         Begin VB.Menu mnuTransform 
            Caption         =   "Transform"
         End
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "mdiParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
' Process the Transform Command
Private Sub cmdTransform_Click(Index As Integer)
' #############################################################################
    ' Use MSI path for dialog if no previous MST path
    If MSTDirName(Index) = "" Then MSTDirName(Index) = MSIDirName(Index)
    
    With cdlTransform
        .Filename = ""
        .Filter = "Transform (*.mst)|*.mst"
        .InitDir = MSTDirName(Index)
        .CancelError = True
        On Error Resume Next
            Do Until .Filename <> "" Or Err <> 0
                .ShowOpen
            Loop
            If Err <> 0 Then Exit Sub
        On Error GoTo 0
        
        MSTDirName(Index) = GetFilePath(cdlTransform.Filename)
        
        MSTName(Index) = cdlTransform.Filename
        If MSTName(Index) <> "" Then
            lblFile(Index).Caption = MSIName(Index) & " (" & MSTName(Index) & ")"
            ApplyTransformtoIndex (Index)
        End If
    End With
End Sub

' #############################################################################
' Apply the selected Transform to the Database
Sub ApplyTransformtoIndex(Index As Integer)
' #############################################################################
    Dim Position As Integer
    Dim DatabaseName As String
    Dim TempMSIName As String
    Dim ReturnCode As Integer
    Dim TempDir As String
    
    TempDir = "c:\temp\msi_compare_temp\"
    
    On Error Resume Next
        MkDir "c:\temp"
        MkDir TempDir
    On Error GoTo 0
    
    ' For testing ONLY
    If NoOfTransforms(Index) > 0 Then
        Dim DummyFlag As Boolean
        DummyFlag = True
    End If
    
    ' Determine path to MSI
    Position = InStrRev(MSIName(Index), "\")
    DatabaseName = Mid(MSIName(Index), Position + 1, Len(MSIName(Index)) - Position)
    
    ' Add a number to the filename reflecting the version of the MSI we are using
    Position = InStrRev(DatabaseName, ".")
    If NoOfTransforms(Index) > 0 Then
        DatabaseName = Left(DatabaseName, Position - 2) & CStr(NoOfTransforms(Index)) & ".msi"
    Else
        DatabaseName = Left(DatabaseName, Position - 1) & CStr(NoOfTransforms(Index)) & ".msi"
    End If
    
    TempMSIName = TempDir & DatabaseName
    
    If Dir(TempMSIName) <> "" Then
        Kill TempMSIName
    End If
    
    ' Copy the existing MSI to a copy in the same dir
    Set MSI(Index) = Nothing
    FileCopy MSIName(Index), TempMSIName
    
    MSIName(Index) = TempMSIName
    
    Set MSI(Index) = OpenMSI(MSIName(Index), msiOpenDatabaseModeTransact, ReturnCode)
    
    If ReturnCode <> 0 Then
        LogError "ApplyTransformtoIndex", ReturnCode
    Else
    
    End If
    
    On Error Resume Next
        MSI(Index).ApplyTransform MSTName(Index), msiTransformErrorNone + msiTransformErrorAddExistingRow + msiTransformErrorDeleteNonExistingRow + msiTransformErrorUpdateNonExistingRow
        
        If Err <> 0 Then
            MsgBox "Error applying Transform - " & vbCrLf & vbCrLf & _
                   "Error code " & Err.Number & " (" & Err.Description & ")" & vbCrLf & vbCrLf & _
                   "The original MSI will now be loaded", vbCritical + vbOKOnly, "Error"
            
            ' Roll back !
            MSIDirName(Index) = GetFilePath(OriginalMSIName(Index))
            lblFile(Index).Caption = OriginalMSIName(Index)
            Set MSI(Index) = OpenMSI(OriginalMSIName(Index), msiOpenDatabaseModeTransact, ReturnCode)
            MSIName(Index) = OriginalMSIName(Index)
            PrepareTreeview (Index)
            Exit Sub
        End If
        MSI(Index).Commit
        
    On Error GoTo 0
    
    NoOfTransforms(Index) = NoOfTransforms(Index) + 1
    
    lblFile(Index).Caption = MSIName(Index) & " (" & MSTName(Index) & ")"
    PrepareTreeview (Index)
End Sub

' #############################################################################
' Determine the path for the given Component
Function GetPathFromComponent(Database As WindowsInstaller.Database, strComponent As String) As String
' #############################################################################
    Dim DatabaseView As View
    Dim DataRecord As Record
    Dim SQLString As String
    Dim KeyPath As String
    Dim ParentDir As String
    Dim FullPath As String
    Dim ComponentName As String
    Dim DirectoryName As String

    ' Get path to Component
    SQLString = "SELECT `Directory_` FROM `Component` WHERE `Component` = '" & strComponent & "'"
    DirectoryName = ReturnSingleValueFromDatabase(Database, SQLString)
    
    ' get Path to Directory
    ' This is a recursive retrieve that builds the path until no parent directory exists
    FullPath = DirectoryName
    Do
        ParentDir = GetParentDir(Database, DirectoryName)
        FullPath = ParentDir & "\" & FullPath
        DirectoryName = ParentDir
    Loop Until ParentDir = ""
    
    ' Remove the leading backslash
    If Left(FullPath, 1) = "\" Then
        FullPath = Right(FullPath, Len(FullPath) - 1)
    End If
    
    Set DatabaseView = Nothing
    GetPathFromComponent = FullPath
End Function

' #############################################################################
' Given a SQL string, returns a single value from the database
Function ReturnSingleValueFromDatabase(Database As WindowsInstaller.Database, SQLString) As String
' #############################################################################
    Dim DatabaseView As View
    Dim DataRecord As Record
    
    Set DatabaseView = Database.OpenView(SQLString)
    DatabaseView.Execute
    
    Set DataRecord = DatabaseView.Fetch
    If DataRecord Is Nothing Then Exit Function
    ReturnSingleValueFromDatabase = DataRecord.StringData(DataRecord.FieldCount)
    
    DatabaseView.Close
End Function

' #############################################################################
' Determine the parent Directory for a given Directory
Function GetParentDir(Database As WindowsInstaller.Database, strDirectory As String) As String
' #############################################################################
    Dim DatabaseView As View
    Dim DataRecord As Record
    Dim SQLString As String
    
    SQLString = "SELECT `Directory_Parent` FROM `Directory` WHERE `Directory` = '" & strDirectory & "'"
    
    ' Get Table view
    Set DatabaseView = Database.OpenView(SQLString)
    DatabaseView.Execute
    
    Set DataRecord = DatabaseView.Fetch
    If DataRecord Is Nothing Then Exit Function
    GetParentDir = DataRecord.StringData(DataRecord.FieldCount)
  
    DatabaseView.Close
    
    Set DatabaseView = Nothing
End Function

' #############################################################################
' Replace one string with another
Function ReplaceStrings(strOriginal As String) As String
' #############################################################################
    Dim i As Integer
    Dim NewString As String
    Dim Strings() As String
    Dim LeftString As String
    Dim RightString As String
    
    Do
        NewString = ReplacementString(i)
        Strings() = Split(NewString, "|")
        If NewString = "" Then Exit Do
        LeftString = Strings(0)
        RightString = Strings(1)
        strOriginal = Replace(strOriginal, LeftString, RightString)
        i = i + 1
    Loop
    ReplaceStrings = strOriginal
End Function

' #############################################################################
' LogError - output's the error
Sub LogError(Location As String, Errorcode As Integer)
' #############################################################################
    Debug.Print "Error " & Errorcode & " (" & Error(Errorcode) & ") in " & Location
End Sub

' #############################################################################
' Process the Open MSI File Menu command
Private Sub mnuDatabase_Click()
' #############################################################################
    Dim Index As Integer
    Dim MSIName As String
    
    Index = 0
    MSIName = DisplayOpenDialog()
    If MSIName <> "" Then
        Dim frmDatabase As New frmChild
        frmDatabase.MSIName = MSIName
        'MSIDirName = GetFilePath(MSIName)
        frmDatabase.Show
        frmDatabase.Refresh
    'End If
    
    'DisplayOpenDialog = cdlFiles.Filename
    'If cdlFiles.Filename <> "" Then
        'lblFile(ComponentIndex) = MSIName(ComponentIndex)
        'OriginalMSIName(ComponentIndex) = MSIName(ComponentIndex)
        'Set MSI(ComponentIndex) = OpenMSI(MSIName(ComponentIndex), msiOpenDatabaseModeReadOnly, ReturnCode)
        
        'If ReturnCode <> 0 Then
        '    LogError "DisplayOpenDialog", ReturnCode
        'Else
        '
        'End If
        
        'PrepareTreeview (frmDatabase)
    End If
    'If LCase(Right(MSIName(ComponentIndex), 3)) = "msi" Then
    '    cmdTransform(ComponentIndex).Enabled = True
    'Else
    '    cmdTransform(ComponentIndex).Enabled = False
    'End If
End Sub


' #############################################################################
' Display the Open dialog for the MSI
Function DisplayOpenDialog() As String
' #############################################################################
    Dim ReturnCode As Integer
    
    'If MSIDirName(ComponentIndex) = "" Then MSIDirName(ComponentIndex) = App.Path
    With cdlFiles
        .Filename = ""
        .Filter = "Database (*.msi)|*.msi|Merge Module (*.msm)|*.msm|All Files (*.*)|*.*"
        .InitDir = App.Path
        .CancelError = True
        On Error Resume Next
        Do Until .Filename <> "" Or Err <> 0
            .ShowOpen
        Loop
        
        If Err <> 0 Then Exit Function
        On Error GoTo 0
    End With
    DisplayOpenDialog = cdlFiles.Filename
End Function

' #############################################################################
' Prepare the treeview for adding nodes
Sub PrepareTreeview(frm As Object)
' #############################################################################
    Dim n As Node
    
    'frm.Show
    'frm.Refresh
    
    ' Setup Root keys
    frm.tvwFile.Nodes.Clear
    Set n = frm.tvwFile.Nodes.Add(, , LCase("Files"), "Files", "Folder")
    Set n = frm.tvwFile.Nodes.Add(, , LCase("Registry"), "Registry", "reg")
    Set n = frm.tvwFile.Nodes.Add(, , LCase("Components"), "Components", "wsi")
    Set n = frm.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKCR"), "HKCR", "Folder")
    Set n = frm.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKCU"), "HKCU", "Folder")
    Set n = frm.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKLM"), "HKLM", "Folder")
    
    ' Load Treeview with info from various tables
    'Load "File", Index
    'Load "Registry", Index
    'Load "Component", Index
    
    'lblStatus(Index).Caption = "Ready"
End Sub

' #############################################################################
' Load data from the specified table
Sub Load(strTable As String, Index As Integer)
' #############################################################################
    Dim SQLString As String
    Dim ErrorString As String
    
    lblStatus(Index).Caption = strTable & "..."
    
    Me.Refresh
    
    SQLString = "SELECT * FROM " & strTable & " ORDER BY Sequence"
    ErrorString = PopulateTreeview(tvwFile(Index), MSI(Index), SQLString)
End Sub

' #############################################################################
' PopulateTreeview - Retrieves MSI data and populates the treeview
Function PopulateTreeview(Tree As TreeView, Database As WindowsInstaller.Database, SQLString As String) As String
' #############################################################################
    Dim DatabaseView As View
    Dim DataRecord As Record
    Dim DataColumn
    Dim NoOfColumns As Integer
    Dim Column As Long
    Dim MSIData As String
    Dim ColumnNames As New Collection, strColumnNames As String
    Dim SQLString2 As String
    Dim TableName As String
    Dim ColumnName
    Dim i As Integer, n As Node
    Dim Data(255) As String
    Dim RegRoot As String, RegKey As String, RegValue As String, RegData As String, RegComponent As String
    Dim FilePath As String
    Dim File    As String, Component As String, Filename   As String, Filesize As String
    Dim Version As String, Language  As String, Attributes As String, Sequence As Long
    Dim myPath As String
    Dim myIcon As Long
    Dim hInstance As Long
    Dim TempPath As String
    
    ' Create path for temporary files
    TempPath = "c:\temp"
    On Error Resume Next
        MkDir TempPath
    On Error GoTo 0
    
    ' Get Column names
    TableName = GetTableName(SQLString)
    GetColumnNames Database, TableName, ColumnNames
    
    ' Get Table view
    On Error Resume Next
        
        Set DatabaseView = Database.OpenView(SQLString)
        
        If Err = -2147467259 Then
            PopulateTreeview = "Error in SQL string"
            Exit Function
        End If
        
    On Error GoTo 0
    DatabaseView.Execute
    
    For Each ColumnName In ColumnNames
        strColumnNames = strColumnNames & ColumnName & ","
    Next
    strColumnNames = Left(strColumnNames, Len(strColumnNames) - 1)
    
    ' Get Table data
    i = 0
    Do
        i = i + 1
        Set DataRecord = DatabaseView.Fetch
        If DataRecord Is Nothing Then
            Exit Do
        End If
        NoOfColumns = DataRecord.FieldCount
        For Column = 1 To NoOfColumns
            DataColumn = DataRecord.StringData(Column)
            Data(Column) = DataColumn
        Next
        
        If TableName = "Registry" Then
            ' Use Names of Registry Root's, not int values eg 'HKLM' instead of '2'
            RegRoot = HK(CInt(Data(2)))
            RegKey = Data(3)
            BuildNodesfromKey RegRoot & "\" & RegKey, Tree
            RegValue = Data(4)
            RegData = Data(5)
            RegComponent = Data(6)

            On Error Resume Next
                If RegValue <> "" Then
                    Set n = Tree.Nodes.Add(LCase(RegRoot & "\" & RegKey), tvwChild, LCase(RegRoot & "\" & RegKey & "\" & RegValue), RegValue, "Reg_sz")
                    Set n = Tree.Nodes.Add(LCase(RegRoot & "\" & RegKey & "\" & RegValue), tvwChild, LCase(RegRoot & "\" & RegKey & "\" & RegValue & "\" & RegData), RegData, "Reg_sz")
                End If
            On Error GoTo 0
            
        ElseIf TableName = "File" Then
            File = Data(1)
            Component = Data(2)
            Filename = Data(3)
            Filesize = Data(4)
            Version = Data(5)
            Language = Data(6)
            Attributes = Data(7)
            Sequence = Data(8)
            
            FilePath = "Files\" & GetPathFromComponent(MSI(0), Component)
            
            FilePath = ReplaceStrings(FilePath)
            
            BuildNodesfromPath FilePath, Tree
            
            ' ******************************************************
            ' Extract the icon for each file
            If CLng(Filesize) < LargeFileSize Then
                ' 1.  Extract the file to the TempPath so we can extract the icon
                Database.Export "File", TempPath, File
                ' 2.  Get the icon associated to this file
                myIcon = 0
                myPath = TempPath + "\" + File
                myIcon = ExtractAssociatedIconA(hInstance, myPath, 0)
                ' 3.  Delete the file we have extracted
                Kill myPath
                ' 4.  Remove the previous icon on the picture control
                Picture1.Cls
                ' 5.  Draw this icon on the invisible picture control
                DrawIcon Picture1.hdc, 0, 0, myIcon
                ' 6.  Destroy the icon we just extracted (to keep required resources low)
                DestroyIcon myIcon
            Else
                Picture1.Picture = imlLarge.ListImages.Item("Other").Picture
            End If
            On Error Resume Next
                ' 7.  Add the exported icon (now on the picture control) to the image list control
                imlMSI.ListImages.Add imlMSI.ListImages.Count + 1, File, Picture1.Image
            On Error GoTo 0
            ' ******************************************************
            Set n = Tree.Nodes.Add(LCase(FilePath), tvwChild, LCase(FilePath & "\" & File), File, File)
            
            File = ""
            Component = ""
            Filename = ""
            Filesize = ""
            Version = ""
            Language = ""
            Attributes = ""
            Sequence = 0
        ElseIf TableName = "Component" Then
            Debug.Print "Component!"
        End If
        Set n = Nothing
    Loop
  
    DatabaseView.Close
    
    Set DatabaseView = Nothing
    PopulateTreeview = "OK"
End Function

' #############################################################################
' Determines the table name from the SQL String
Function GetTableName(SQLString) As String
' #############################################################################
    Dim P As Integer, Q As Integer
    If InStr(1, UCase(SQLString), "FROM", vbTextCompare) = 0 Then Exit Function
    P = InStr(1, UCase(SQLString), "FROM", vbTextCompare) + 5
    Q = InStr(P, UCase(SQLString), " ", vbTextCompare)
    If Q = 0 Then Q = Len(SQLString) + 1
    GetTableName = Mid(SQLString, P, Q - P)
End Function

' #############################################################################
' GetColumnNames - returns a collection of column names for the given table
Function GetColumnNames(Database As WindowsInstaller.Database, TableName As String, ColumnNames As Collection)
' #############################################################################
    Dim DatabaseView As View
    Dim DataRecord As Record
    Dim ColumnData As String
    Dim SQLString As String
    
    SQLString = "SELECT `Table`, `Number`, `Name` FROM `_Columns` WHERE `Table` = '" & TableName & "' ORDER BY `Number`"
    
    ' Get Table view
    Set DatabaseView = Database.OpenView(SQLString)
    DatabaseView.Execute
    
    ' Get Table data
    Do
      Set DataRecord = DatabaseView.Fetch
      If DataRecord Is Nothing Then Exit Do
      ColumnData = DataRecord.StringData(DataRecord.FieldCount)
      ColumnNames.Add ColumnData, ColumnData
    Loop
  
    DatabaseView.Close
    
    Set DatabaseView = Nothing
End Function

' #############################################################################
' Creates each node in the tree in a hierarchial manner (registry)
Function BuildNodesfromKey(RegistryKey As String, Tree As TreeView)
' #############################################################################
    Dim Keys() As String, Key
    Dim n As Node, Parent As String
    Dim strKey As String
    Dim NumOfKeys As Integer, i As Integer
    
    Parent = Left(RegistryKey, InStr(1, RegistryKey, "\") - 1)
    'On Error Resume Next
    Keys() = Split(RegistryKey, "\")
    NumOfKeys = UBound(Keys())
    For Each Key In Keys()
        i = i + 1
        'If i > 1 Then
            'On Error Resume Next
            If i < NumOfKeys + 1 Then
                Set n = Tree.Nodes.Add(LCase(Parent), tvwChild, LCase(Parent & "\" & Key), Key, "Folder")
            Else
                Set n = Tree.Nodes.Add(LCase(Parent), tvwChild, LCase(Parent & "\" & Key), Key, "reg_sz")
            End If
            'On Error GoTo 0
            Parent = Parent & "\" & Key
        'End If
    Next
End Function

' #############################################################################
' Creates each node in the tree in a hierarchial manner (registry)
Function BuildNodesfromPath(FullPath As String, Tree As TreeView)
' #############################################################################
    Dim Paths() As String, Path
    Dim n As Node, Parent As String
    Dim strPath As String
    Dim NumOfPaths As Integer, i As Integer
    
    If FullPath = "" Then Exit Function
    
    Parent = Left(FullPath, InStr(1, FullPath, "\") - 1)
    On Error Resume Next
    Paths() = Split(FullPath, "\")
    NumOfPaths = UBound(Paths())
    For Each Path In Paths()
        i = i + 1
        If i > 1 Then
            On Error Resume Next
                Set n = Tree.Nodes.Add(LCase(Parent), tvwChild, LCase(Parent & "\" & Path), Path, "Folder")
            On Error GoTo 0
            Parent = Parent & "\" & Path
        End If
    Next
End Function

' #############################################################################
' Process the Application exit Menu Command
Private Sub mnuExit_Click()
' #############################################################################
    ExitApp
End Sub

' #############################################################################
' Process the Application close Command
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' #############################################################################
    Cancel = True
    ExitApp
End Sub

' #############################################################################
' Application ends here
Sub ExitApp()
' #############################################################################
    
    
    End
End Sub

