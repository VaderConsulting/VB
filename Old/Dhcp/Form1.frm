VERSION 5.00
Begin VB.Form frmDHCP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DHCP Clients"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportScope 
      Caption         =   "Export Scope"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdExtract2 
      Caption         =   "Extract Clients"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstClients 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   4695
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract Scopes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstScopes 
      Height          =   5910
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   1695
   End
   Begin VB.CommandButton cmdScopes 
      Caption         =   "Get Scopes"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnum 
      Caption         =   "Enumerate clients"
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblDesc 
      Caption         =   "MAC Address"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblDesc 
      Caption         =   "Host name"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDesc 
      Caption         =   "IP Address"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   10680
      X2              =   0
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label lblClient 
      Alignment       =   2  'Center
      Caption         =   "Client information"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblScope 
      Alignment       =   2  'Center
      Caption         =   "Scopes"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblClients 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Label lblScopes 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
         Begin VB.Menu mnuExportAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnuExportSelected 
            Caption         =   "Selected"
         End
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
         Begin VB.Menu mnuScopes 
            Caption         =   "Scopes"
         End
         Begin VB.Menu mnuClients 
            Caption         =   "Clients"
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
Attribute VB_Name = "frmDHCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This program relies on DHCPCMD.exe being installed AND in the users path
' ****
' ****
' It is ESSENTIAL that the version of DHCPCMD.exe used is at least 4.00
' ****
' ****

Dim DHCPServer As String

Private Sub Form_Load()
    mnuFile.Enabled = False
    frmDHCP.Show
    frmDHCP.Refresh
    RefreshScopes                  ' Initiate DHCP scope extraction
    mnuFile.Enabled = True
End Sub

Private Sub cmdEnum_Click()
    ' Create the batch file used to enumerate the information
    cmd = "dhcpcmd " & DHCPServer & " enumclients " & lstScopes.List(lstScopes.ListIndex) & " -h >c:\temp\dhcpclients.txt"
    Open "c:\temp\getclients.bat" For Output As #1
        Print #1, cmd
    Close 1
    Shell "c:\temp\getclients.bat", vbHide
    cmdExtract2.Enabled = True
    ' Wait 4 seconds after running the command for the results to be returned.
    ' Increase this number if the DHCP server is on a slow link
    a = Time
    Do Until Time > DateAdd("s", 4, a)
        DoEvents
    Loop
    ' Parse the txt file for information required.  In this case it is the DHCP Scopes.
    cmdExtract2_Click
    lblClients = lstClients.ListCount & " Clients"
    lblClients.Refresh
End Sub

Private Sub cmdExportScope_Click()
    ' Export the information contained in the list box to a text file
    lblClients = "Exporting... please wait"
    lblClients.Refresh
    FileName = "c:\temp\" & lstScopes.Text & ".txt"
    Open FileName For Output As #3
        For lp = 0 To lstClients.ListCount - 1
            Print #3, lstClients.List(lp)
        Next lp
    Close 3
    MsgBox "Extract complete.  File saved to c:\temp\" & lstScopes.Text & ".txt"
    lblClients = lstClients.ListCount & " Clients"
    lblClients.Refresh
End Sub

Private Sub cmdExtract_Click()
    ' Get scope data and put in list box
    Open "c:\temp\scopes.txt" For Input As #1
        Do Until EOF(1)
            Line Input #1, varData
            If Left(varData, 17) = "Subnet address = " Then
                lstScopes.AddItem Right(varData, Len(varData) - 17)
            End If
        Loop
    Close 1
End Sub

Private Sub cmdExtract2_Click()
    ' Parse information in extracted text file
    lstClients.Clear
    Open "c:\temp\dhcpclients.txt" For Input As #1
    ' CHANGE:  removed version info from DHCP Server info line to allow future versions
    junk1 = "DHCP Server version" ' 4.1"
    junk2 = "Command successfully completed."
        Do Until EOF(1)
            Line Input #1, varData
            ' 19 chars is the length of the above two 'junk' lines
            If Left(varData, Len(junk1)) <> junk1 And Left(varData, Len(junk2)) <> junk2 Then
                varTab = InStr(1, varData, Chr(9))
                num = Mid(varData, 1, varTab - 1)
                varRemainder = Right(varData, Len(varData) - varTab - 1)
                IP = Trim(Left(varRemainder, 17))
                varRemainder = Right(varData, Len(varData) - varTab - 18)
                varSpace = InStr(1, varRemainder, " ")
                varHost = Trim(Left(varRemainder, varSpace))
                varMAC = Trim(Right(varRemainder, Len(varRemainder) - varSpace))
                varNewData = IP & "," & varHost & "," & varMAC
                lstClients.AddItem varNewData
            End If
        Loop
    Close 1
    
End Sub

Private Sub cmdScopes_Click()
    lstScopes.Clear
    ' Create the batch file used to enumerate the information
    cmd = "dhcpcmd " & DHCPServer & " getsuperscopetable >c:\temp\scopes.txt"
    Open "c:\temp\getscopes.bat" For Output As #1
        Print #1, cmd
    Close 1
    ' Run the batch file
    Shell "c:\temp\getscopes.bat", vbHide
    cmdExtract.Enabled = True
End Sub

Sub RefreshScopes()
    ' Ask the user for the DHCP Servers IP Address.  This is a modification that was added to
    ' allow this program to be used in a non WAPS environment.
    If DHCPServer = "" Then
        DHCPServer = InputBox("Enter your DHCP Server's IP Address", "Input Required", "10.1.1.21")
    End If
    lblScopes = "Enumerating..."
    lblScopes.Refresh
    cmdScopes_Click
    frmDHCP.Refresh
    ' Wait 4 seconds after running the command for the results to be returned.
    ' Increase this number if the DHCP server is on a slow link
    a = Time
    Do Until Time > DateAdd("s", 4, a)
        DoEvents
    Loop
    cmdExtract_Click
    lblScopes = lstScopes.ListCount & " Scopes"
    lblScopes.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmExport
    Unload Me
    End
End Sub

Private Sub lstScopes_Click()
    RefreshClients
End Sub

Sub RefreshClients()
    lblClients = "Enumerating... please wait"
    lblClients.Refresh
    cmdEnum_Click
End Sub

Private Sub mnuClients_Click()
    RefreshClients
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuExportAll_Click()
    frmExport.Show vbModal
End Sub

Private Sub mnuExportSelected_Click()
    cmdExportScope_Click
End Sub

Private Sub mnuScopes_Click()
    RefreshScopes
End Sub
