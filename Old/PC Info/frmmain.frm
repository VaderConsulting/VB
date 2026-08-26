VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC Info"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   8265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8265
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      Default         =   -1  'True
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtComputername 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer tmrInfo 
      Interval        =   2000
      Left            =   120
      Top             =   6120
   End
   Begin VB.ListBox lstInfo 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8055
   End
   Begin VB.Label lblComputername 
      Caption         =   "Computer name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Width           =   8295
   End
   Begin VB.Image imgInfo 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image imgImage 
      Height          =   240
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image imgImage 
      Height          =   255
      Index           =   1
      Left            =   480
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image imgImage 
      Height          =   255
      Index           =   2
      Left            =   840
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image imgImage 
      Height          =   255
      Index           =   3
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image imgImage 
      Height          =   255
      Index           =   4
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image imgImage 
      Height          =   240
      Index           =   5
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   240
   End
   Begin VB.Image imgClear 
      Height          =   255
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image imgOldInfo 
      Height          =   255
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label lblOldInfo 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   5280
      Width           =   8175
   End
   Begin VB.Image imgImage 
      Height          =   255
      Index           =   6
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public MyOS As String

Private Sub GetInfo(strComputer As String)
    Dim Group As IADsGroup, Member As Object, objComputer As IADs
    Dim ADDescription As String, objOS As Object, i As Integer, bIsUp As Boolean
    Dim objWMIService As Object, colChassis As Object, Chassis As String, o As Object
    Dim Memory As String, Processor As String, FormFactor As String, MemoryType As String
    Dim MainboardManufacturer As String, SerialNumber As String, DomainRole As String
    Dim WakeupType As String, DiskSize As Double, oUser As IADsUser, strTempInfo As String
    Dim objCollection As Object, objEnum As Object
    Dim strTemp(9) As String, dtmStartDate As Object, dtmEndDate As Object, DateToCheck As Date
    
    Const CONVERT_TO_LOCAL_TIME = True
    
    lstInfo.Clear
    
    On Error GoTo Hell
    
    Screen.MousePointer = vbHourglass
    
    LogMessage "Pinging " & strComputer, 3
    
    ' If the host is Windows XP, then ping the selected computer
    If (InStr(1, MyOS, "XP") > 0) And (strComputer <> Environ("computername")) Then
        If isUP(strComputer) Then
            LogMessage strComputer & " responded.", 4
            bIsUp = True
        Else
            LogMessage strComputer & " did not respond.", 5
            bIsUp = False
        End If
    Else
        ' Not XP or same machine, so assume the computer is up.
        bIsUp = True
    End If
    
    If bIsUp Then
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        
        Set objCollection = objWMIService.ExecQuery("Select Description from Win32_QuickFixEngineering")
        
        LogMessage "Listing hotfixes on " & strComputer, 0

        lstInfo.AddItem "Hotfixes:"
        For Each objEnum In objCollection
            If objEnum.Description <> "" Then
                lstInfo.AddItem objEnum.Description
            End If
            DoEvents
        Next
        
'        lstInfo.AddItem ""
'        lstInfo.AddItem "Software:"
'
'        Set objCollection = objWMIService.ExecQuery("Select Caption from Win32_Product")
'
'        LogMessage "Listing software on " & strComputer, 0
'
'        For Each objEnum In objCollection
'            If objEnum.Caption <> "" Then
'                lstInfo.AddItem objEnum.Caption
'            End If
'            DoEvents
'        Next
'
'        LogMessage "Listing services on " & strComputer, 0
'
'        lstInfo.AddItem ""
'        lstInfo.AddItem "Services:"
'
'        Set objCollection = objWMIService.ExecQuery("Select Caption, Description, State from Win32_Service")
'
'        For Each objEnum In objCollection
'            If objEnum.Description <> "" Then
'                lstInfo.AddItem objEnum.Caption & vbTab & objEnum.Description & " (" & objEnum.State & ")"
'            End If
'            DoEvents
'        Next
    
        Set objCollection = objWMIService.ExecQuery("Select Name, Path from Win32_Share")
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Network Shares:"
        
        LogMessage "Listing shared folders on " & strComputer, 0
      
        For Each objEnum In objCollection
            If objEnum.Name <> "" Then
                'strTemp(1) = objEnum.Caption & ""
                strTemp(2) = objEnum.Name & ""
                strTemp(3) = objEnum.Path & ""
                lstInfo.AddItem strTemp(2) & " = " & strTemp(3)
                Erase strTemp()
            End If
            DoEvents
        Next
        
'        lstInfo.AddItem ""
'        lstInfo.AddItem "Event Log:"
'
'        Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
'        Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
'
'        DateToCheck = CDate(Now)
'        dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
'        dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME
'
'        Set objCollection = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where TimeWritten >= '" & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'")
'
'        For Each objEnum In objCollection
'            lstInfo.AddItem objEnum.Category & vbTab & objEnum.EventCode & vbTab & objEnum.Message & vbTab & objEnum.SourceName & vbTab & objEnum.Type & vbTab & objEnum.User
'            DoEvents
'        Next
        
        LogMessage "Collecting Chassis Info ", 3
        
        ADDescription = ""
        
        ' --------------------------------------------------------
        Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
        Set colChassis = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
        For Each o In colChassis
            For i = LBound(o.ChassisTypes) To UBound(o.ChassisTypes)
                Select Case o.ChassisTypes(i)
                    Case 1:  Chassis = "Other"
                    Case 2:  Chassis = "Unknown"
                    Case 3:  Chassis = "Desktop"
                    Case 4:  Chassis = "Low Profile Desktop"
                    Case 5:  Chassis = "Pizza Box"
                    Case 6:  Chassis = "Mini Tower"
                    Case 7:  Chassis = "Tower"
                    Case 8:  Chassis = "Portable"
                    Case 9:  Chassis = "Laptop"
                    Case 10: Chassis = "Notebook"
                    Case 11: Chassis = "Hand Held"
                    Case 12: Chassis = "Docking Station"
                    Case 13: Chassis = "All in One"
                    Case 14: Chassis = "Sub Notebook"
                    Case 15: Chassis = "Space-Saving"
                    Case 16: Chassis = "Lunch Box"
                    Case 17: Chassis = "Main System Chassis"
                    Case 18: Chassis = "Expansion Chassis"
                    Case 19: Chassis = "SubChassis"
                    Case 20: Chassis = "Bus Expansion Chassis"
                    Case 21: Chassis = "Peripheral Chassis"
                    Case 22: Chassis = "Storage Chassis"
                    Case 23: Chassis = "Rack Mount Chassis"
                    Case 24: Chassis = "Sealed-Case PC"
                End Select
            Next
        Next
        lstInfo.AddItem ""
        lstInfo.AddItem "Chassis: "
        lstInfo.AddItem Chassis
        DoEvents
        ' --------------------------------------------------------
        
        If ADDescription = "" Then
            ADDescription = "[No Description]"
        Else
            lstInfo.AddItem "AD Description: " & ADDescription
        End If
        
        DoEvents
        
        LogMessage "Collecting Memory Info ", 3
        lstInfo.AddItem ""
        FormFactor = ""
        lstInfo.AddItem "Memory:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_PhysicalMemory")
            Select Case o.FormFactor
                Case 0:  FormFactor = "Unknown"
                Case 1:  FormFactor = "Other"
                Case 2:  FormFactor = "SIP"
                Case 3:  FormFactor = "DIP"
                Case 4:  FormFactor = "ZIP"
                Case 5:  FormFactor = "SOJ"
                Case 6:  FormFactor = "Proprietary"
                Case 7:  FormFactor = "SIMM"
                Case 8:  FormFactor = "DIMM"
                Case 9:  FormFactor = "TSOP"
                Case 10: FormFactor = "TGA"
                Case 11: FormFactor = "RIMM"
                Case 12: FormFactor = "SODIMM"
            End Select
            
            Select Case o.MemoryType
                Case 1:  MemoryType = "Unknown"
                Case 2:  MemoryType = "Other"
                Case 3:  MemoryType = "DRAM"
                Case 4:  MemoryType = "Synchronous DRAM"
                Case 5:  MemoryType = "Cache DRAM"
                Case 6:  MemoryType = "EDO"
                Case 7:  MemoryType = "EDRAM"
                Case 8:  MemoryType = "VRAM"
                Case 9:  MemoryType = "SRAM"
                Case 10: MemoryType = "RAM"
                Case 11: MemoryType = "ROM"
                Case 12: MemoryType = "Flash"
                Case 13: MemoryType = "EEPROM"
                Case 14: MemoryType = "FEPROM"
                Case 15: MemoryType = "EPROM"
                Case 16: MemoryType = "CDRAM"
                Case 17: MemoryType = "3DRAM"
                Case 18: MemoryType = "SDRAM"
                Case 19: MemoryType = "SGRAM"
            End Select
            
            lstInfo.AddItem o.DeviceLocator & ": " & "Size: " & o.Capacity & " (" & o.Capacity / 1024 / 1024 & " Mb) " & FormFactor & " " & MemoryType
        Next
        
        DoEvents
        
        LogMessage "Collecting Processor Info ", 3
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Processor:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Processor")
            lstInfo.AddItem o.CurrentClockSpeed & " MHz " & o.Name
        Next
        
        LogMessage "Collecting Drive Info ", 3
        
        lstInfo.AddItem ""
        i = 0
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_FloppyDrive")
            i = i + 1
        Next
        lstInfo.AddItem i & " Floppy Drive(s)"
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Disks:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_DiskDrive")
            DiskSize = 0
            DiskSize = o.Size / 1024 / 1024 / 1024
            DiskSize = Int(DiskSize * 10) / 10
            lstInfo.AddItem o.Model & " " & " (" & DiskSize & " Gb)"
        Next
        
        DoEvents
        
        LogMessage "Collecting Motherboard Info ", 3
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Motherboard:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_BaseBoard")
            lstInfo.AddItem o.Manufacturer & " " & o.Product
        Next
        
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_BIOS")
            lstInfo.AddItem "Serial Number: " & o.SerialNumber
            lstInfo.AddItem "BIOS Version: " & o.SMBIOSBIOSVersion
        Next
        
        DoEvents
        
        LogMessage "Collecting System Info ", 3
        
        lstInfo.AddItem ""
        lstInfo.AddItem "System:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_ComputerSystem")
            Select Case o.DomainRole
                Case 0: DomainRole = "Standalone Workstation"
                Case 1: DomainRole = "Member Workstation"
                Case 2: DomainRole = "Standalone Server"
                Case 3: DomainRole = "Member Server"
                Case 4: DomainRole = "Backup Domain Controller"
                Case 5: DomainRole = "Primary Domain Controller"
            End Select
            lstInfo.AddItem "Domain Role: " & DomainRole
            lstInfo.AddItem "Logged on user: " & o.UserName
            
            Select Case o.WakeupType
                Case 0: WakeupType = "Reserved"
                Case 1: WakeupType = "Other"
                Case 2: WakeupType = "Unknown"
                Case 3: WakeupType = "APM Timer"
                Case 4: WakeupType = "Modem Ring"
                Case 5: WakeupType = "LAN Remote"
                Case 6: WakeupType = "Power Switch"
                Case 7: WakeupType = "PCI PME#"
                Case 8: WakeupType = "AC Power Restored"
            End Select
            lstInfo.AddItem "Power-up due to: " & WakeupType
        Next
        
        DoEvents
        
        LogMessage "Collecting Operating System Info ", 3
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Operating System:"
        ' Connect to WMI and obtain instances of Win32_OperatingSystem
        For Each objOS In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_OperatingSystem")
            lstInfo.AddItem objOS.Caption & " (" & objOS.Version & ") "
            lstInfo.AddItem "Registered User = " & objOS.RegisteredUser
        Next
        
        If DomainRole <> "Primary Domain Controller" And DomainRole <> "Backup Domain Controller" Then
            For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_NTDomain")
                If o.Status = "OK" Then
                    lstInfo.AddItem "AD Site name: " & o.ClientSiteName
                End If
            Next
        End If
        
        LogMessage "Collecting Network Info ", 3
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Network Connections:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_NetworkConnection")
            strTempInfo = UCase(o.LocalName & "")
            If strTempInfo = "" Then strTempInfo = "  "
            lstInfo.AddItem strTempInfo & " " & o.RemoteName & " (" & o.ConnectionState & " - " & o.Status & ") as " & o.UserName
        Next
        
        lstInfo.AddItem ""
        lstInfo.AddItem "Printers:"
        For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Printer")
            lstInfo.AddItem o.DeviceID & " (" & o.DriverName & ") on " & o.PortName
        Next
        
        Me.Refresh
        
        LogMessage "Idle", 6
        
    End If
    Set objComputer = Nothing
    Set objCollection = Nothing
    Set objWMIService = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
Hell:
    LogMessage "Error " & Err.Number & " (" & Err.Description & ")", 1
    Err.Clear
    Resume Next
End Sub

Public Sub LogMessage(strMessage As String, Category As Integer)
    tmrInfo.Enabled = False
    tmrInfo.Interval = 2000
    lblOldInfo.Caption = lblInfo.Caption
    imgOldInfo.Picture = imgInfo.Picture
    lblInfo = strMessage
    Select Case Category
        Case -1
            imgInfo.Picture = imgClear.Picture
        Case Else
            imgInfo.Picture = imgImage(Category).Picture
    End Select
    
    tmrInfo.Enabled = True
    
    frmMain.Refresh
End Sub

Private Sub cmdInfo_Click()
    If txtComputername <> "" Then
            GetInfo txtComputername.Text
    End If
End Sub

Private Sub Form_Load()
    Dim objOS As Object
    ' Get the OS of the host
    'For Each objOS In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
    '    MyOS = objOS.Caption
    'Next
    
    frmMain.Show
    frmMain.Refresh
    
    LogMessage "Idle", 6
    
    txtComputername = Environ("computername")
    
End Sub

Private Sub tmrInfo_Timer()
    lblOldInfo.Caption = ""
    imgOldInfo.Picture = imgClear.Picture
    tmrInfo.Enabled = False
End Sub
