VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NAT Reader Process"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   360
      Top             =   960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NAT\HOME.LOCAL\NAT.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\NAT\HOME.LOCAL\NAT.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim DomainName As String
    Dim ParentDirectory As String
    Dim SourceDirectory As String
    Dim s As String
    Dim ComputerName As String
    Dim DataType As String
    Dim DataEntry As String
    Dim Data As String
    Dim TableName As String
    Dim FieldName As String
    Dim DSN As String
    Dim adoConnection As ADODB.Connection
    Dim adoRS As ADODB.Recordset
    Dim SQL As String
    Dim HostID As Long
    
    Set adoConnection = CreateObject("ADODB.Connection")
    Set adoRS = CreateObject("ADODB.Recordset")
    
    If Command$ = "" Or App.PrevInstance Then
        GoTo ExitApp
    End If
    
    ParentDirectory = "C:\NAT"
    
    If Dir(ParentDirectory & "\NAT Template.mdb") = "" Then End
    
    DomainName = Command$
    SourceDirectory = ParentDirectory & "\" & DomainName
    
    DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SourceDirectory & "\NAT.mdb;Persist Security Info=False"
    
    If Dir(SourceDirectory & "\NAT.mdb") = "" Then
        FileCopy ParentDirectory & "\NAT Template.mdb", SourceDirectory & "\NAT.mdb"
    End If
    
    If Dir(SourceDirectory & "\*.na4") = "" Then
        GoTo ExitApp
    End If
    
    adoConnection.Open DSN

    s = Dir(SourceDirectory & "\*.na4")
    
    Do Until s = ""
        ComputerName = Left(s, InStr(1, s, ".") - 1)
        
        SQL = "DELETE FROM tblHosts WHERE Name = '" & ComputerName & "'"
        adoConnection.Execute SQL
    
        SQL = "INSERT INTO tblHosts (Name) VALUES ('" & ComputerName & "')"
        adoConnection.Execute SQL
        SQL = "SELECT ID from tblHosts Where Name = '" & ComputerName & "'"
        adoRS.Open SQL, adoConnection
        Do Until adoRS.EOF
            HostID = adoRS("ID")
            adoRS.MoveNext
        Loop
        adoRS.Close
        
        Open SourceDirectory & "\" & s For Input As #1
            Do Until EOF(1)
                Line Input #1, Data
                SQL = Replace(Data, "HOSTID", HostID)
                If Left(Data, 2) <> "//" Then
                    adoConnection.Execute SQL
                End If
            Loop
        Close 1
        Name SourceDirectory & "\" & s As SourceDirectory & "\" & ComputerName & ".nat"
        s = Dir
    Loop
    
    adoConnection.Close
    
ExitApp:
    
    Set adoRS = Nothing
    Set adoConnection = Nothing
    End
End Sub
