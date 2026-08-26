Attribute VB_Name = "Module1"
Option Explicit

Public Const MaxNoOfOpenDatabases             As Integer = 2
Public Const LargeFileSize                    As Long = 1048576 ' bytes

'Public MSI(MaxNoOfOpenDatabases)              As WindowsInstaller.Database
'Public MSIName(MaxNoOfOpenDatabases)          As String
'Public OriginalMSIName(MaxNoOfOpenDatabases)  As String
'Public MSTName(MaxNoOfOpenDatabases)          As String
'Public MSIDirName(MaxNoOfOpenDatabases)       As String
'Public MSTDirName(MaxNoOfOpenDatabases)       As String
'Public NoOfTransforms(MaxNoOfOpenDatabases)   As Integer

Public ReplacementString(255)                 As String

Public HK(2)                                  As String

Public Const HKCR                             As Integer = 0
Public Const HKCU                             As Integer = 1
Public Const HKLM                             As Integer = 2

Public NoOfOpenDatabases                      As Integer

Public Declare Function ExtractAssociatedIconA Lib "shell32.dll" ( _
                                                           ByVal hInst As Long, _
                                                           ByVal lpIconPath As String, _
                                                           lpiIcon As Long) _
                                                           As Long
                                                           
Public Declare Function DrawIcon Lib "user32" ( _
                                         ByVal hdc As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal hIcon As Long) _
                                         As Long
                                         
Public Declare Function DestroyIcon Lib "user32" ( _
                                            ByVal hIcon As Long) _
                                            As Long
                                            
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" ( _
                                                                      ByVal hInst As Long, _
                                                                      ByVal lpszExeFileName As String, _
                                                                      ByVal nIconIndex As Long) _
                                                                      As Long

' #############################################################################
' Extract the path for the given full filename
Public Function GetFilePath(strFilename As String) As String
' #############################################################################
    Dim Position As Integer
    Position = InStrRev(strFilename, "\")
    GetFilePath = Left(strFilename, Position - 1)
End Function

' #############################################################################
' OpenMSI - Opens the given MSI filename, returning a Database object
Public Function OpenMSI(MSIPath As String, Mode As Integer, ReturnCode As Integer) As WindowsInstaller.Database
' #############################################################################
    Dim oInstaller As WindowsInstaller.Installer
    Set oInstaller = CreateObject("WindowsInstaller.Installer")
    On Error Resume Next
        Set OpenMSI = oInstaller.OpenDatabase(MSIPath, Mode)
        ReturnCode = Err.Number
    On Error GoTo 0
    Set oInstaller = Nothing
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


