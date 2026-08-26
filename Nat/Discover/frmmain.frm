VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NAT Discover Process"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstInfo 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NA1Filename As String
Public NA2Filename As String
Public NA3Filename As String
Public NA4Filename As String
Public ErrFilename As String

Private Sub Form_Load()
    Dim s As String
    Dim Computername As String
    Dim ParentDirectory As String
    Dim Group As IADsGroup, Member As Object, objComputer As IADs, strComputer As String
    Dim ADDescription As String, objOS As Object, i As Integer, bIsUp As Boolean
    Dim objWMIService As Object, colChassis As Object, Chassis As String, o As Object
    Dim Memory As String, Processor As String, FormFactor As String, MemoryType As String
    Dim MainboardManufacturer As String, SerialNumber As String, DomainRole As String
    Dim WakeupType As String, DiskSize As Double, oUser As IADsUser
    Dim m_objApiClass As Object, p_vntRtn As Variant, p_vntRtn2 As Variant
    Dim lngError As Long, Transparency As Integer, NoOfMemorySlots As Integer
    Dim NoOfErrors As Integer, CurrentStep As String
    
    CurrentStep = "Initialising"
    
    Me.Show
    Me.Refresh
    If Command$ = "" Then End
    
    On Error GoTo Hell
    
    Transparency = 255
    
    ParentDirectory = App.Path
    
    If Dir(ParentDirectory & "\" & Command$ & "\*.na2") = "" Then End
    s = Dir(ParentDirectory & "\" & Command$ & "\*.na2")
    Computername = Left(s, InStr(1, s, ".") - 1)
    
    NA1Filename = ParentDirectory & "\" & Command$ & "\" & Computername & ".na1"
    NA2Filename = ParentDirectory & "\" & Command$ & "\" & s
    NA3Filename = ParentDirectory & "\" & Command$ & "\" & Computername & ".na3"
    NA4Filename = ParentDirectory & "\" & Command$ & "\" & Computername & ".na4"
    ErrFilename = ParentDirectory & "\" & Command$ & "\" & Computername & ".err"
    
    Name NA2Filename As NA3Filename
    
    Me.Refresh
    
    strComputer = Computername
    
    Me.Caption = Me.Caption & " [" & Computername & "]"
    
    LogMessage "Connecting to " & strComputer
    
    CurrentStep = "Connecting to Computer"

    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
    Me.Refresh
    
    CurrentStep = "Memory"
    
    LogMessage "Retrieving details: Memory"
    
    FormFactor = ""
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
        
        lstInfo.AddItem "INSERT INTO tblMemory (HostID,DeviceLocator,[Size],FormFactor,MemoryType) VALUES (HOSTID,'" & o.DeviceLocator & "'," & o.Capacity / 1024 / 1024 & ",'" & FormFactor & "','" & MemoryType & "')"
        
    Next
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_PhysicalMemoryArray")
        NoOfMemorySlots = o.MemoryDevices
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Processor(s)"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Processor")
        lstInfo.AddItem "INSERT INTO tblProcessor (HostID,CurrentClockSpeed,Name) VALUES (HOSTID," & o.CurrentClockSpeed & ",'" & o.Name & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Operating System"
    
    For Each objOS In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_OperatingSystem")
        lstInfo.AddItem "INSERT INTO tblOperatingSystem (HostID,Caption,Version,RegisteredUser,SerialNumber,CSDVersion) VALUES (HOSTID,'" & objOS.Caption & "','" & objOS.Version & "','" & objOS.RegisteredUser & "','" & objOS.SerialNumber & "','" & objOS.CSDVersion & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Disk Drives"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_DiskDrive")
        DiskSize = 0
        DiskSize = o.Size / 1024 / 1024 / 1024
        DiskSize = Int(DiskSize * 10) / 10
        lstInfo.AddItem "INSERT INTO tblDiskDrives (HostID,Model,DiskSize) VALUES (HOSTID,'" & o.Model & "'," & DiskSize & ")"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: System"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_ComputerSystem")
        Select Case o.DomainRole
            Case 0: DomainRole = "Standalone Workstation"
            Case 1: DomainRole = "Member Workstation"
            Case 2: DomainRole = "Standalone Server"
            Case 3: DomainRole = "Member Server"
            Case 4: DomainRole = "Backup Domain Controller"
            Case 5: DomainRole = "Primary Domain Controller"
            Case Else
                DomainRole = ""
        End Select
        
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
            Case Else
                WakeupType = ""
        End Select
        lstInfo.AddItem "INSERT INTO tblSystem (HostID,DomainRole,Username,WakeupType,MemorySlots) VALUES (HOSTID,'" & DomainRole & "','" & o.UserName & "" & "','" & WakeupType & "" & "'," & NoOfMemorySlots & ")"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Chassis"

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
                Case Else
                    Chassis = ""
            End Select
        Next
    Next
    
    lstInfo.AddItem "INSERT INTO tblChassis (HostID,ChassisType) VALUES (HOSTID,'" & Chassis & "')"
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Keyboard"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Keyboard")
        lstInfo.AddItem "INSERT INTO tblKeyboard (HostID,Description) VALUES (HOSTID,'" & o.Description & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Mouse"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_PointingDevice")
        lstInfo.AddItem "INSERT INTO tblMouse (HostID,Name) VALUES (HOSTID,'" & o.Name & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Monitor"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_DesktopMonitor")
        lstInfo.AddItem "INSERT INTO tblMonitor (HostID,MonitorManufacturer) VALUES (HOSTID,'" & o.MonitorManufacturer & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Video"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_VideoController")
        lstInfo.AddItem "INSERT INTO tblVideo (HostID,Caption,AdapterRAM) VALUES (HOSTID,'" & o.Caption & "'," & o.AdapterRAM / 1024 / 1024 & ")"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Floppy Drive"

    i = 0
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_FloppyDrive")
        i = i + 1
        DoEvents
    Next
    lstInfo.AddItem "INSERT INTO tblFloppyDrives (HostID,FloppyDrives) VALUES (HOSTID," & i & ")"
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: CD ROM Drives"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_CDROMDrive")
        lstInfo.AddItem "INSERT INTO tblCDROMDrives (HostID,Caption,Drive) VALUES (HOSTID,'" & o.Caption & "','" & o.Drive & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Tape Drives"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_TapeDrive")
        lstInfo.AddItem "INSERT INTO tblTapeDrives (HostID,Caption,Description) VALUES (HOSTID,'" & o.Caption & "','" & o.Description & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Motherboard"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_BaseBoard")
        lstInfo.AddItem "INSERT INTO tblMotherboard (HostID,Manufacturer,Product) VALUES (HOSTID,'" & o.Manufacturer & "" & "','" & o.Product & "" & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: BIOS"

    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_BIOS")
        lstInfo.AddItem "INSERT INTO tblBIOS (HostID,SerialNumber,SMBIOSBIOSVersion) VALUES (HOSTID,'" & o.SerialNumber & "','" & o.SMBIOSBIOSVersion & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Slots"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_SystemSlot")
        lstInfo.AddItem "INSERT INTO tblSlots (HostID,Tag,Designation) VALUES (HOSTID,'" & o.Tag & "','" & o.SlotDesignation & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Sound"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_SoundDevice")
        lstInfo.AddItem "INSERT INTO tblSound (HostID,Caption,Manufacturer) VALUES (HOSTID,'" & o.Caption & "','" & o.Manufacturer & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Modems"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_POTSModem")
        lstInfo.AddItem "INSERT INTO tblModems (HostID,Description,DeviceType) VALUES (HOSTID,'" & o.Description & "','" & o.DeviceType & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Shares"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Share")
        lstInfo.AddItem "INSERT INTO tblShares (HostID,Name,Caption,Path) VALUES (HOSTID,'" & o.Name & "','" & o.Caption & "','" & o.Path & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
    LogMessage "Retrieving details: Printers"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Printer")
        lstInfo.AddItem "INSERT INTO tblPrinters (HostID,DeviceID,DriverName,PortName) VALUES (HOSTID,'" & o.DeviceID & "','" & o.DriverName & "','" & o.PortName & "')"
        DoEvents
    Next
            
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
        
    LogMessage "Retrieving details: Software"
    
    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_Product")
        lstInfo.AddItem "INSERT INTO tblSoftware (HostID,Caption) VALUES (HOSTID,'" & o.Caption & "')"
        DoEvents
    Next
    
    MakeTransparent Me.hWnd, Transparency
    Transparency = Transparency - 12
    Me.Refresh
    
'    LogMessage "Retrieving details: Hotfixes"
'
'    For Each o In GetObject("winmgmts:\\" & strComputer).InstancesOf("Win32_QuickFixEngineering")
'        lstInfo.AddItem "2E: " & o.Description
'    Next
'
'    Me.Refresh
    
    LogMessage "Idle"
    
    Set objComputer = Nothing
    
    Name NA3Filename As NA4Filename
    Open NA4Filename For Output As #1
    For i = 0 To frmMain.lstInfo.ListCount - 1
        Print #1, frmMain.lstInfo.List(i)
    Next i
    Close 1
    
    If Dir(ErrFilename) <> "" Then
        ' Delete error File
        Kill ErrFilename
    End If
    
    End
Hell:
    LogMessage "Error " & Err.Number & " (" & Err.Description & ")"
    lstInfo.AddItem "//Error: " & Err.Number & " (" & Err.Description & ")"
    Err.Clear
    
    NoOfErrors = NoOfErrors + 1
    
    On Error Resume Next
    
    Set objComputer = Nothing
    
    ' Delete pre-existing Error file
    If Dir(ErrFilename) <> "" Then
        Kill ErrFilename
    End If
    
    If Dir(ErrFilename) <> "" Then
        ' Could not delete Error file, so rename to NA1 file
        Name NA3Filename As NA1Filename
        'Kill NA3Filename
    Else
        ' No Error file, so rename NA3 file to Error file
        Name NA3Filename As ErrFilename
    End If
    
    Open ErrFilename For Output As #1
    For i = 0 To frmMain.lstInfo.ListCount - 1
        Print #1, frmMain.lstInfo.List(i)
    Next i
    Close 1
    If NoOfErrors >= 5 Then
        End
    End If
    Resume Next
End Sub

Sub LogMessage(strMessage As String)
    lblInfo = strMessage
    Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Dir(NA3Filename) <> "" Then
        Name NA3Filename As NA2Filename
    End If
End Sub
