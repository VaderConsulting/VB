VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NAT Collection Process"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim DomainLDAPPath As String
    Dim DomainName As String
    Dim App_Path As String
    Dim objConnection As ADODB.Connection
    Dim objCommand As ADODB.Command
    Dim objRecordset As ADODB.Recordset
    Dim Computername As String, ADPath As String
    Dim ParentDirectory As String
    
    Me.Show
    Me.Refresh
    lblInfo.Caption = "Starting up..."
    Me.Refresh
    
    ParentDirectory = "C:\NAT\"
    
    Const ADS_SCOPE_SUBTREE = 2
    On Error Resume Next
    
    ' Get Domain Name (Environment variable method)
    DomainName = UCase(Command$)
    DomainLDAPPath = "LDAP://DC=" & Replace(DomainName, ".", ",DC=")
    
    ' Ensure the App_Path variable ends with a '\'
    If Right(App.Path, 1) <> "\" Then
        App_Path = App.Path & "\"
    Else
        App_Path = App.Path
    End If
    
    ' If required, create parent directory
    If Dir(ParentDirectory, vbDirectory) = "" Then
        MkDir ParentDirectory
    End If
    
    ' if required, create directory to hold machine information (per domain)
    If Dir(ParentDirectory & DomainName, vbDirectory) = "" Then
        MkDir ParentDirectory & DomainName
    End If
    
    ' Get machine names, and create an entry for each
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    Set objCommand.ActiveConnection = objConnection
    
    objCommand.CommandText = "Select Name, ADsPath from '" & DomainLDAPPath & "' where objectClass='computer'"
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    objCommand.Properties("Cache Results") = False
    
    Set objRecordset = objCommand.Execute
    
    If Err.Number = 0 Then
        objRecordset.MoveFirst
        
        Do Until (objRecordset.EOF)
            Computername = objRecordset.Fields("Name").Value
            lblInfo.Caption = "Found " & Computername & " in " & DomainName
            Me.Refresh
            If Dir(ParentDirectory & DomainName & "\" & Computername & ".na*") = "" Then
                Open ParentDirectory & DomainName & "\" & Computername & ".na1" For Output As #1
                Close 1
            End If
            objRecordset.MoveNext
        Loop
    End If
    
    Set objCommand = Nothing
    Set objConnection = Nothing
    End
End Sub
