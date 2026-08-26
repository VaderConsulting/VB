VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChild 
   Caption         =   "Database"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   6765
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   2160
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ImageList imlLarge 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":030A
            Key             =   "Other"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMSI 
      Left            =   720
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":2ABC
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":9DC6
            Key             =   "reg"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":A218
            Key             =   "mst"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":A66A
            Key             =   "wsi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":A984
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":ADD6
            Key             =   "fnt"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B228
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B67A
            Key             =   "reg_bin"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B7D4
            Key             =   "reg_sz"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChild.frx":B92E
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4590
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwFile 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8070
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlMSI"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MSI                 As WindowsInstaller.Database
Public MSIName          As String
Dim OriginalMSIName     As String
Dim MSTName             As String
Dim MSIDirName          As String
Dim MSTDirName          As String
Dim NoOfTransforms      As Integer
Dim DatabaseNo          As Integer

' #############################################################################
' Start here
Private Sub Form_Load()
' #############################################################################
    Dim ReturnCode As Integer
    OriginalMSIName = MSIName
    MSIDirName = GetFilePath(MSIName)
    Set MSI = OpenMSI(MSIName, msiOpenDatabaseModeReadOnly, ReturnCode)
    PrepareTreeview
    NoOfTransforms = 0
End Sub

' #############################################################################
' Processed every time the form is activated
Private Sub Form_Activate()
' #############################################################################
    DoResize
        
    If LCase(Right(MSIName, 3)) = "msi" Then
        mdiParent.mnuTransform.Enabled = True
        mdiParent.Caption = "Database: (" & MSIName & ")"
        'Me.Caption = "Database: (" & MSIName & ")"
    Else
        mdiParent.mnuTransform.Enabled = False
        mdiParent.Caption = "Database: (" & MSIName & " & " & MSTName & ")"
        'Me.Caption = "Database: (" & MSIName & " & " & MSTName & ")"
    End If
End Sub

' #############################################################################
' Prepare the treeview for adding nodes
Sub PrepareTreeview()
' #############################################################################
    Dim n As Node
    
    Me.Show
    Me.Refresh
    
    ' Setup Root keys
    Me.tvwFile.Nodes.Clear
    Set n = Me.tvwFile.Nodes.Add(, , LCase("Files"), "Files", "Folder")
    Set n = Me.tvwFile.Nodes.Add(, , LCase("Registry"), "Registry", "reg")
    Set n = Me.tvwFile.Nodes.Add(, , LCase("Components"), "Components", "wsi")
    Set n = Me.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKCR"), "HKCR", "Folder")
    Set n = Me.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKCU"), "HKCU", "Folder")
    Set n = Me.tvwFile.Nodes.Add(LCase("Registry"), tvwChild, LCase("HKLM"), "HKLM", "Folder")
    
    ' Load Treeview with info from various tables
    Load "File" ', Index
    Load "Registry" ', Index
    Load "Component" ', Index
    
    Me.sbrStatus.SimpleText = "Ready"
End Sub

' #############################################################################
' Load data from the specified table
Sub Load(strTable As String)
' #############################################################################
    Dim SQLString As String
    Dim ErrorString As String
    
    Me.sbrStatus.SimpleText = strTable & "..."
    
    Me.Refresh
    
    SQLString = "SELECT * FROM " & strTable & " ORDER BY Sequence"
    ErrorString = PopulateTreeview(MSI, SQLString)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' #############################################################################
' PopulateTreeview - Retrieves MSI data and populates the treeview
Function PopulateTreeview(Database As WindowsInstaller.Database, SQLString As String) As String
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
            BuildNodesfromKey RegRoot & "\" & RegKey, Me.tvwFile
            RegValue = Data(4)
            RegData = Data(5)
            RegComponent = Data(6)

            On Error Resume Next
                If RegValue <> "" Then
                    Set n = Me.tvwFile.Nodes.Add(LCase(RegRoot & "\" & RegKey), tvwChild, LCase(RegRoot & "\" & RegKey & "\" & RegValue), RegValue, "Reg_sz")
                    Set n = Me.tvwFile.Nodes.Add(LCase(RegRoot & "\" & RegKey & "\" & RegValue), tvwChild, LCase(RegRoot & "\" & RegKey & "\" & RegValue & "\" & RegData), RegData, "Reg_sz")
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
            
            FilePath = "Files\" & GetPathFromComponent(MSI, Component)
            
            FilePath = ReplaceStrings(FilePath)
            
            BuildNodesfromPath FilePath, Me.tvwFile
            
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
            Set n = Me.tvwFile.Nodes.Add(LCase(FilePath), tvwChild, LCase(FilePath & "\" & File), File, File)
            
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
            sbrStatus.SimpleText = "Loading " & Key
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
                sbrStatus.SimpleText = "Loading " & Path
            On Error GoTo 0
            Parent = Parent & "\" & Path
        End If
    Next
End Function

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set MSI = Nothing
End Sub

' #############################################################################
' Resize the treeview according to the form resize event
' #############################################################################
Sub DoResize()
    On Error Resume Next
    tvwFile.Width = Me.Width - 80
    tvwFile.Height = Me.Height - (sbrStatus.Height + 230)
    Exit Sub
End Sub

' #############################################################################
' Resize the treeview according to the form resize event
' #############################################################################
Private Sub Form_Resize()
    DoResize
End Sub
