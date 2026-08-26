VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DCD PC Info"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   3300
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   3300
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   3300
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   3300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.ComboBox cmbSites 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3300
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Memory:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Speed:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Manufacturer:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Model Number:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lblSite 
      Alignment       =   1  'Right Justify
      Caption         =   "Site Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim App_Path As String
Dim FoundDb As Boolean
Dim Serial As String
Dim Manufacturer As String
Dim Model As String
Dim Speed As String
Dim Memory As String
Dim DSN As String

Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long

Public Function ChangeComputerName(sNewComputerName As String) As Boolean
    On Error Resume Next
    Dim nReturn As Long
    
    nReturn = SetComputerName(sNewComputerName)
    If Err.Number = 0 Then
       ChangeComputerName = nReturn <> 0
    End If

End Function


Private Function AccessDateTime(DateTime As Date) As String
    AccessDateTime = Day(DateTime) & "-" & Month(DateTime) & "-" & Year(DateTime)
End Function

Private Sub Wait(Period As Integer)
    Dim t1 As Date
    Dim t2 As Date
    t1 = Now
    t2 = DateAdd("s", CDbl(Period), Now)
    Do
        DoEvents
    Loop Until Now >= t2
End Sub

Private Sub cmbSites_Change()
    If cmbSites.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub cmbSites_Click()
    If cmbSites.Text <> "" Then
        ' Pre april 5
        'lblStatus = "Computername: " & UCase(cmbSites.Text) & "-" & UCase(Serial)
        ' Post April 5
        lblStatus = "Computername: " & UCase(Serial)
        lblStatus.Refresh
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    Dim SQL As String
    Dim StringIn As String
    Dim StringOut As String
    Dim Reply As Long
    Dim NoOfApps As Integer
    Dim INIFilename As String
    Dim result As Boolean
    Dim i As Integer
    Dim NewComputerName As String
    
    If cmdOK.Caption = "Close" Then
        End
    Else
        cmdOK.Enabled = False
        If FoundDb Then
            Dim adoConn As ADODB.Connection
            Dim adoRS As ADODB.Recordset
            
            lblStatus = "Opening Database"
            lblStatus.Refresh
            Set adoConn = CreateObject("ADODB.Connection")
            Set adoRS = CreateObject("ADODB.Recordset")
            SQL = "SELECT Serial FROM tblPCInfo WHERE Serial = '" & Serial & "' AND Model = '" & Model & "'"
            adoConn.Open DSN
            adoRS.Open SQL, adoConn
            If (adoRS.BOF) Or (adoRS.EOF) Then
                SQL = "INSERT INTO tblPCInfo (Site, Serial, Manufacturer, Model, Memory, Speed, BuildDate) VALUES ('" & cmbSites.Text & "','" & Serial & "','" & Manufacturer & "','" & Model & "','" & Memory & "'," & CInt(Speed) & ",'" & AccessDateTime(Now) & "')"
            Else
                SQL = "UPDATE tblPCInfo Set Site = '" & cmbSites.Text & "', Serial = '" & Serial & "', Manufacturer = '" & Manufacturer & "', Model = '" & Model & "', Memory = '" & Memory & "', Speed = " & CInt(Speed) & ", BuildDate = '" & AccessDateTime(Now) & "' WHERE Serial = '" & Serial & "' AND Model = '" & Model & "'"
            End If
            adoRS.Close
            lblStatus = "Updating database"
            lblStatus.Refresh
            adoConn.Execute SQL
            adoConn.Close
            Set adoRS = Nothing
            Set adoConn = Nothing
        End If
        
        lblStatus = "Creating Output file"
        lblStatus.Refresh
        Open App_Path & "template.txt" For Input As #1
            Open App_Path & Serial & ".kix" For Output As #2
                Do Until EOF(1)
                    Line Input #1, StringIn
                    Select Case UCase(Left(StringIn, 9))
                        Case "$SITENAME"
                            StringOut = "$SITENAME = " & Chr(34) & UCase(cmbSites.Text) & Chr(34)
                        Case "$SERIALNU"
                            StringOut = "$SERIALNUMBER = " & Chr(34) & UCase(Serial) & Chr(34)
                        Case Else
                            StringOut = StringIn
                    End Select
                    Print #2, StringOut
                Loop
            Close 2
        Close 1
        
        ' Old (pre April 5)
        'NewComputerName = UCase(cmbSites.Text) & "-" & Serial
        
        ' New (Post April 5)
        NewComputerName = Serial
        
        Reply = MsgBox("Are you certain you want to:" & vbCrLf & vbCrLf & _
                       "  1.  Change the computer name to " & NewComputerName & vbCrLf & vbCrLf & _
                       "  2.  Set the SITE environment variable to " & UCase(cmbSites.Text) & "NW", vbYesNo, "Please confirm")
        If Reply = vbYes Then
            lblStatus = "Applying changes"
            lblStatus.Refresh
            Shell App_Path & "wkix32.exe " & App_Path & Serial & ".kix"
            Wait 5
            ChangeComputerName NewComputerName
            MsgBox "Complete." & vbCrLf & "You must reboot the computer for changes to take effect", vbInformation + vbOKOnly
        Else
            lblStatus = "No changes made.  Click Close to exit."
            lblStatus.Refresh
        End If
        cmdOK.Caption = "Close"
        cmdOK.Enabled = True
        Wait 2
        Kill App_Path & Serial & ".kix"
        
        INIFilename = App_Path & "deploy.ini"
        StringIn = Get_INI_Key(INIFilename, "PostApps", "No", result)
        If result = False Then
            MsgBox INIFilename & " not found!", vbCritical, "Error"
        Else
            NoOfApps = CInt(StringIn)
            For i = 1 To NoOfApps
                StringIn = Get_INI_Key(INIFilename, "AppPaths", CStr(i), result)
                If StringIn <> "" Then
                    Shell StringIn
                End If
            Next i
        End If
    End If
End Sub


Private Sub Form_Load()
    Dim INIFilename As String
    Dim DbFilename As String
    Dim DefManufacturer As String
    Dim DefModel As String
    Dim DefSpeed As String
    Dim SQL As String
    Dim OutPath As String
    Dim i As Integer
    Dim StringIn As String
    Dim StringOut As String
    Dim result As Boolean
    
    Me.Show
    Me.Refresh
    cmdOK.Enabled = False
    Screen.MousePointer = vbHourglass
    
    lblStatus = "Retrieving .ini information"
    lblStatus.Refresh
    App_Path = App.Path
    If Right(App_Path, 1) <> "\" Then App_Path = App_Path & "\"
    INIFilename = App_Path & "deploy.ini"
    
    DbFilename = Get_INI_Key(INIFilename, "Database", "Path", result)
    If result = False Then
        MsgBox INIFilename & " not found!", vbCritical, "Error"
        End
    End If
    DefManufacturer = Get_INI_Key(INIFilename, "Defaults", "Manufacturer", result)
    If result = False Then
        MsgBox INIFilename & " not found!", vbCritical, "Error"
        End
    End If
    DefModel = Get_INI_Key(INIFilename, "Defaults", "Model", result)
    If result = False Then
        MsgBox INIFilename & " not found!", vbCritical, "Error"
        End
    End If
    DefSpeed = Get_INI_Key(INIFilename, "Defaults", "Speed", result)
    If result = False Then
        MsgBox INIFilename & " not found!", vbCritical, "Error"
        End
    End If
    If DbFilename = "" Then
        MsgBox "Database path not known." & vbCrLf & vbCrLf & "Defaulting to " & App_Path & "deploy.mdb", vbExclamation + vbOKOnly
        FoundDb = False
    Else
        On Error Resume Next
        If Dir(DbFilename, vbNormal) <> "" Then
            If Err.Number = 0 Then
                FoundDb = True
                lblStatus = "Found Database"
                lblStatus.Refresh
            Else
                If Dir(App_Path & "deploy.mdb") <> "" Then
                    MsgBox "Database '" & DbFilename & "' not found." & vbCrLf & vbCrLf & "Defaulting to " & App_Path & "deploy.mdb", vbExclamation + vbOKOnly
                    DbFilename = App_Path & "deploy.mdb"
                    FoundDb = True
                Else
                    MsgBox "Database '" & DbFilename & "' not found." & vbCrLf & vbCrLf & "Database updates not possible", vbExclamation + vbOKOnly
                    FoundDb = False
                End If
            End If
        Else
            If Dir(App_Path & "deploy.mdb") <> "" Then
                MsgBox "Database '" & DbFilename & "' not found." & vbCrLf & vbCrLf & "Defaulting to " & App_Path & "deploy.mdb", vbExclamation + vbOKOnly
                DbFilename = App_Path & "deploy.mdb"
                FoundDb = True
            Else
                MsgBox "Database '" & DbFilename & "' not found." & vbCrLf & vbCrLf & "Database updates not possible", vbExclamation + vbOKOnly
                FoundDb = False
            End If
        End If
        On Error GoTo 0
    End If

    lblStatus = "Retrieving Hardware info"
    lblStatus.Refresh
    
    If Serial = "" Then
        'Use a WMI query to retrieve hardware info
        Open App_Path & "localhost.vbs" For Output As #1
        Print #1, "Set fso = CreateObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ")"
        Print #1, "winmgmt1 = " & Chr(34) & "winmgmts:{impersonationLevel=impersonate}!//localhost" & Chr(34)
        Print #1, "Set SNSet = GetObject(winmgmt1).InstancesOf(" & Chr(34) & "Win32_BIOS" & Chr(34) & ")"
        Print #1, "Set OutFile = fso.CreateTextFile(" & Chr(34) & App_Path & "stats.ini" & Chr(34) & ")"
        Print #1, "OutFile.Write (" & Chr(34) & "[Statistics]" & Chr(34) & ") & Chr(13) & Chr(10)"
        Print #1, "For Each SN In SNSet"
        Print #1, "    OutFile.Write (" & Chr(34) & "serial=" & Chr(34) & " & SN.SerialNumber) & Chr(13) & Chr(10)"
        Print #1, "    OutFile.Write (" & Chr(34) & "manufacturer=" & Chr(34) & " & SN.Manufacturer) & Chr(13) & Chr(10)"
        Print #1, "Next"
        Print #1, "Set System = GetObject(winmgmt1).InstancesOf(" & Chr(34) & "Win32_ComputerSystem" & Chr(34) & ")"
        Print #1, "For Each S In System"
        Print #1, "    OutFile.Write (" & Chr(34) & "model=" & Chr(34) & " & S.Model) & Chr(13) & Chr(10)"
        Print #1, "    OutFile.Write (" & Chr(34) & "memory=" & Chr(34) & " & S.TotalPhysicalMemory) & Chr(13) & Chr(10)"
        Print #1, "Next"
        Print #1, "Set Processor = GetObject(winmgmt1).InstancesOf(" & Chr(34) & "Win32_Processor" & Chr(34) & ")"
        Print #1, "For Each CPU In Processor"
        Print #1, "    OutFile.Write (" & Chr(34) & "speed=" & Chr(34) & " & CPU.CurrentClockSpeed) & Chr(13) & Chr(10)"
        Print #1, "Next"
        Print #1, "OutFile.Close"
        Close 1
        Shell "cscript.exe " & App_Path & "localhost.vbs"
        
        Wait 5
        INIFilename = App_Path & "stats.ini"
        If Dir(INIFilename) <> "" Then
            Serial = Get_INI_Key(INIFilename, "Statistics", "Serial", result)
            If result = False Then
                MsgBox INIFilename & " not found!", vbCritical, "Error"
                End
            End If
           Manufacturer = Get_INI_Key(INIFilename, "Statistics", "Manufacturer", result)
           If result = False Then
                MsgBox INIFilename & " not found!", vbCritical, "Error"
                End
            End If
           Model = Get_INI_Key(INIFilename, "Statistics", "Model", result)
           If result = False Then
                MsgBox INIFilename & " not found!", vbCritical, "Error"
                End
            End If
           Speed = Get_INI_Key(INIFilename, "Statistics", "Speed", result)
           If result = False Then
                MsgBox INIFilename & " not found!", vbCritical, "Error"
                End
            End If
           Memory = Get_INI_Key(INIFilename, "Statistics", "Memory", result)
           If result = False Then
                MsgBox INIFilename & " not found!", vbCritical, "Error"
                End
            End If
            If Memory <> "" Then
                Memory = CStr(CLng(Memory / 1024))
                Memory = CStr(Int(Memory / 1000))
            End If
            ' Erase any sign of us being here
            Kill App_Path & "localhost.vbs"
            Kill INIFilename
        End If
    End If
    
    ' Last option - ask the user
    If Serial = "" Then
        Serial = UCase(InputBox("Please enter the Serial number of this PC", "Input Required", ""))
    End If
    If Manufacturer = "" Then
        Manufacturer = UCase(InputBox("Please enter the Manufacturer of this PC", "Input Required", DefManufacturer))
    End If
    If Model = "" Then
        Model = UCase(InputBox("Please enter the Model number of this PC", "Input Required", DefModel))
    End If
    If (Speed = "") Or (Not IsNumeric(Speed)) Then
        Speed = UCase(InputBox("Please enter the Speed of this PC", "Input Required", DefSpeed))
    End If
    
    Serial = UCase(Serial)
    Manufacturer = UCase(Manufacturer)
    Model = UCase(Model)
    
    ' Display the results found to the user
    txtInfo(0).Text = Serial
    txtInfo(1).Text = Model
    txtInfo(2).Text = Manufacturer
    txtInfo(3).Text = Speed & " Mhz"
    txtInfo(4).Text = Memory & " Mb"
    
    If FoundDb Then
        Dim adoConn As ADODB.Connection
        Dim adoRS As ADODB.Recordset
        
        lblStatus = "Opening Database"
        lblStatus.Refresh
        Set adoConn = CreateObject("ADODB.Connection")
        Set adoRS = CreateObject("ADODB.Recordset")
        DSN = "Driver={Microsoft Access Driver (*.mdb)};" & "Dbq=" & DbFilename
        SQL = "SELECT Code FROM tblSites ORDER BY Code"
        adoConn.Open DSN
        adoRS.Open SQL, adoConn
        
        lblStatus = "Retrieving Sitecodes"
        lblStatus.Refresh
        
        Do Until adoRS.EOF
            cmbSites.AddItem UCase(adoRS("Code") & "")
            adoRS.MoveNext
        Loop
        adoRS.Close
        adoConn.Close
        Set adoRS = Nothing
        Set adoConn = Nothing
        Screen.MousePointer = vbDefault
        lblStatus = "Waiting for user input - select a site code"
    Else
        Screen.MousePointer = vbDefault
        lblStatus = "Waiting for user input - enter a site code"
    End If
    
    cmdOK.Enabled = True
    lblStatus.Refresh
    
End Sub

' Get_INI_Key
' Opens the nominated .ini file, and retrieves the 1st instance of the specified key in the
' 1st instance of the specified section
' Author: D. Robinson
'
' Change History:
'V1.0  06 Mar 2002 D. Robinson       Initial concept
'
Function Get_INI_Key(INIFilename As String, Section As String, Key As String, result As Boolean) As String
    Dim StringIn As String
    Dim RPos  As Integer
    
    result = False
    
    If Dir(INIFilename) = "" Then Exit Function
    Section = LCase(Section)
    Key = LCase(Key)
    ' Open ini file to get database path
    Open INIFilename For Input As #1
        Do Until EOF(1)
            Line Input #1, StringIn
            StringIn = LCase(Trim(StringIn))
            If Len(StringIn) > 0 Then
                If Left(StringIn, 1) <> ";" And Left(StringIn, 1) <> "'" And Left(StringIn, 3) <> "rem" Then
                    If StringIn = "[" & Section & "]" Then
                        Do Until EOF(1) Or Get_INI_Key <> ""
                            Line Input #1, StringIn
                            StringIn = LCase(Trim(StringIn))
                            If Left(StringIn, Len(Key)) = Key Then
                                RPos = InStr(Len(Key), StringIn, "=")
                                Get_INI_Key = Trim(Right(StringIn, Len(StringIn) - RPos))
                                result = True
                                Close 1
                                Exit Function
                            End If
                        Loop
                    End If
                End If
            End If
        Loop
    Close 1
End Function
