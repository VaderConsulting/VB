VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComputers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Management"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8160
   Icon            =   "frmComputers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fmeTasks 
      Caption         =   "Tasks"
      Height          =   1095
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkTask 
         Caption         =   "Add to Group1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkTask 
         Caption         =   "Move to WEL Managed Computers"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtComputers 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdAcceptList 
      Caption         =   "Cleanup this list"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlImages 
      Left            =   2880
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":08CA
            Key             =   "User"
            Object.Tag             =   "User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":0BE4
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":0EFE
            Key             =   "DisabledUser"
            Object.Tag             =   "DisabledUser"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":1218
            Key             =   "DisabledComputer"
            Object.Tag             =   "DisabledComputer"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":1532
            Key             =   "UnknownComputer"
            Object.Tag             =   "UnknownComputer"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":1E0C
            Key             =   "WaitingComputer"
            Object.Tag             =   "WaitingComputer"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":2126
            Key             =   "DownComputer"
            Object.Tag             =   "DownComputer"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":2440
            Key             =   "UpComputer"
            Object.Tag             =   "UpComputer"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":275A
            Key             =   "Stop"
            Object.Tag             =   "Stop"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":2BAC
            Key             =   "Question"
            Object.Tag             =   "Question"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":2FFE
            Key             =   "Exclamation"
            Object.Tag             =   "Exclamation"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComputers.frx":3450
            Key             =   "Info"
            Object.Tag             =   "Info"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   4320
      Width           =   7455
   End
   Begin VB.Image imgInfo 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   8400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   8400
      Y1              =   20
      Y2              =   20
   End
   Begin VB.Label lblComputers 
      Alignment       =   2  'Center
      Caption         =   "Computers"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import List"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmComputers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAcceptList_Click()
    txtComputers.Text = Replace(txtComputers.Text, " ", vbCrLf)
    txtComputers.Text = Replace(txtComputers.Text, ";", vbCrLf)
    txtComputers.Text = Replace(txtComputers.Text, ",", vbCrLf)
    txtComputers.Text = Replace(txtComputers.Text, vbTab, vbCrLf)
    txtComputers.Text = Replace(txtComputers.Text, vbCrLf & vbCrLf, vbCrLf)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intCount As Integer
    Dim strComputers() As String
    Dim strComputername As String
    Dim oUser As IADsUser
    Dim oComputer As IADsComputer
    Dim intTaskNumber As Integer
    Dim oGroup As IADsGroup
    Dim strDNSDomainName As String
    Dim strADName As String
    Dim iadsCont As IADsContainer
    Dim strComputerPath As String
    
    strDNSDomainName = Environ$("USERDNSDOMAIN")
    
    strADName = "DC=" & Replace(strDNSDomainName, ".", ",DC=")
    strComputers() = Split(txtComputers.Text, vbCrLf)
    
    
    For intCount = 0 To UBound(strComputers())
        strComputername = strComputers(intCount) & "$"
        On Error Resume Next
            Set oUser = GetObject("WinNT://" & strDomainName & "/" & strComputername & ",User")
            If Err.Number = 0 Then
                For intTaskNumber = 0 To 1
                    If chkTask(intTaskNumber).Value = vbChecked Then
                        Select Case intTaskNumber
                            Case 0 ' Add to Group1
                                Set oGroup = GetObject("WinNT://" & strDomainName & "/" & strComputerGroups(0))
                                    On Error Resume Next
                                        oGroup.Add "WinNT://" & strDomainName & "/" & strComputername
                                        If Err.Number = 0 Then
                                            AddtoAuditTrail strComputername & ": Added to Group: " & strComputerGroups(0)
                                            Open strApp_Path & "Groups-" & strComputername & ".txt" For Append As #1
                                                Print #1, strComputerGroups(0)
                                            Close 1
                                        ElseIf Err.Number = -2147022660 Then
                                            AddtoAuditTrail strComputername & ": is already a member of Group: " & strComputerGroups(0)
                                        Else
                                            AddtoAuditTrail strComputername & ": Error " & Err.Number & " whilst attempting to add this Computer to group " & strComputerGroups(0) & "!"
                                        End If
                                    On Error GoTo 0
                                Set oGroup = Nothing
                            Case 1 ' Move to OU
                                If Right(strComputername, 1) = "$" Then strComputername = Left(strComputername, Len(strComputername) - 1)
                                strComputerPath = GetObjectPath(strComputername, Computer)
                                AddtoAuditTrail strComputername & ": Previous OU = " & strComputerPath
                                On Error Resume Next
                                Set iadsCont = GetObject("LDAP://" & strComputerOU & "," & strADName)
                                If Err.Number = 0 Then
                                    iadsCont.MoveHere strComputerPath, vbNullString
                                    Set iadsCont = Nothing
                                    AddtoAuditTrail strComputername & ": Moved to OU (" & strComputerOU & "," & strADName & ")."
                                Else
                                    Logevent "Error binding to " & "LDAP://" & strComputerOU & "," & strADName, EventStop
                                End If
                        End Select
                    End If
                Next
            Else
                Logevent "Error binding to object " & strComputername, EventStop
            End If
            Set oComputer = Nothing
        On Error GoTo 0
    Next
    cmdOK.Enabled = False

End Sub

Private Sub Form_Load()
    Dim strKey As String
    
    AddtoAuditTrail "*********************************************************************************************"
    AddtoAuditTrail "Application started. User: " & Environ$("USERNAME") & " Computer: " & Environ$("COMPUTERNAME")
    
    ' Generate an App_Path variable that is the path to this application, similar to App.Path, though it will always end with a backslash.
    If Right(App.Path, 1) <> "\" Then
        strApp_Path = App.Path & "\"
    Else
        strApp_Path = App.Path
    End If
    
    GetConfig "Computer Groups"
    
    strDomainName = Environ$("USERDOMAIN")
    
    chkTask(0).Caption = "Add to " & strComputerGroups(0)
    strComputerOU = GetIniValue("Standard OU", "Computer", "", strApp_Path & "mgecomp.ini")
    
    chkTask(1).Caption = "Move to " & strComputerOU
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_QueryUnload
' DateTime  : 05-05-2003 20:20
' Author    : Dave Robinson
' Purpose   :
'
'  V    Date        Author            History
' 1.0   05-05-2003      Dave Robinson          Initial Version
'---------------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    AddtoAuditTrail "Application finished. User: " & Environ$("USERNAME") & " Computer: " & Environ$("COMPUTERNAME")
    AddtoAuditTrail "*********************************************************************************************"
End Sub


Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuImport_Click()
    Dim strFilename As String
    Dim strData As String
    
    strFilename = InputBox("Enter the filename to import", "Input", Environ$("TEMP") & "\Computers.txt")
    
    If strFilename <> "" Then
        Open strFilename For Input As #1
            Do Until EOF(1)
                Line Input #1, strData
                If strData <> "" Then
                    txtComputers.Text = txtComputers.Text & strData & vbCrLf
                End If
            Loop
        Close 1
        cmdAcceptList_Click
    End If
    
End Sub
