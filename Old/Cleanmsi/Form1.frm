VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtMSI 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\Documents\Downloads\ACDSee.msi"
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public MSI         As WindowsInstaller.Database

Private Sub cmdGo_Click()
    Dim MSIName As String
    Dim SQLString As String
    Dim DatabaseView As View
    Dim DataRecord As Record
    Dim DataColumn As String, Data(20) As String
    Dim ColumnName, TableName As String
    Dim ColumnNames As New Collection, strColumnNames As String
    Dim i As Integer, NoOfColumns As Integer, Column As Integer

    MSIName = txtMSI.Text

    If Trim(MSIName) = "" Then Exit Sub

    Set MSI = OpenMSI(MSIName, msiOpenDatabaseModeReadOnly)
    
    'SQLString = "SELECT * FROM Registry ORDER BY Registry"
    SQLString = ""
    'SQLString = SQLString & "SELECT `Registry`.`Registry`, `Key`, `Name`, `Value` FROM `Registry` INNER JOIN `Registry` ALIAS `Reg2` "
    'SQLString = SQLString & "ON "
    'SQLString = SQLString & "`Registry`.`Key` = `Reg2`.`Key` " ' AND "
    'SQLString = SQLString & "`Registry`.`Name` = `Reg2`.`Name` AND "
    'SQLString = SQLString & "`Registry`.`Value` = `Reg2`.`Value` "
    'SQLString = SQLString & "GROUP BY `Registry`.`Registry`, `Registry`.`Key`, `Registry`.`Name`, `Registry`.`Value` "
    'SQLString = SQLString & "HAVING Count(`Registry`.`Key`) > 1 "
    'SQLString = SQLString & "ORDER BY `Registry`.`Registry`, `Registry`.`Key`, `Registry`.`Name`, `Registry`.`Value`"
    SQLString = "SELECT `Registry` FROM Registry WHERE `Registry` <> 'primarykey%' "
                 
    ' Get Column names
    TableName = GetTableName(SQLString)
    GetColumnNames MSI, TableName, ColumnNames

    ' Get Table view
    'On Error Resume Next

        Set DatabaseView = MSI.OpenView(SQLString)

    'On Error GoTo 0
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
            Debug.Print Data(Column)
        Next

    Loop

    DatabaseView.Close

    Set DatabaseView = Nothing
End Sub

' #############################################################################
' OpenMSI - Opens the given MSI filename, returning a Database object
Function OpenMSI(MSIPath As String, Mode As Integer) As WindowsInstaller.Database
' #############################################################################
    Dim oInstaller As WindowsInstaller.Installer
    Set oInstaller = CreateObject("WindowsInstaller.Installer")
    On Error Resume Next
        Set OpenMSI = oInstaller.OpenDatabase(MSIPath, Mode)
    On Error GoTo 0
    Set oInstaller = Nothing
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set MSI = Nothing
End Sub

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
    GetTableName = Replace(GetTableName, "`", "")
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
