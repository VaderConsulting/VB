VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Logical Drive Information"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Drive Information"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.ListBox List2 
         Height          =   2010
         ItemData        =   "Form1.frx":0442
         Left            =   2640
         List            =   "Form1.frx":0444
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   460
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Get Info"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1620
         ItemData        =   "Form1.frx":0446
         Left            =   480
         List            =   "Form1.frx":0448
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Drive Information:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Select Domain/Workgroup:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Information about the owner transition process will appear here"
      Top             =   3465
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6297
            Text            =   "System messages will display here"
            TextSave        =   "System messages will display here"
            Key             =   "msg"
            Object.ToolTipText     =   "System messages will display here"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:05"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "22/05/2000"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    ' SHWDRV.EXE
    ' Enumerating Drive Information on Networked NT Servers & Workstations
    '
    ' Richard Puckett
    ' rpuckett@snl.com
    '
    ' Caveats for use of this program:
    '
    ' 1.) you must be an Administrator on the NT system you are attempting
    '     to view (no surprises here) - vb*admin*code
    ' 2.) the server service must be running on the system you are attempting
    '     to enumerate.
    ' 3.) Since NT, by default, creates administrative shares at the root
    '     of each physical drive installed (C$, D$, etc.), these shares
    '     cannot be disabled to use this program.
    '
    '   They are disabled through this Registry edit:
    '
    '   HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\LanmanServer\Parameters
    '   "AutoShareServer" = dword:00000000 ; to disable NT Server/ 00000001 to enable
    '   "AutoShareWks" = dword:00000000 ; to disable NT Workstation/ 00000001 to enable
    '
    '   NOTE: A reboot is required to reset default Administrative shares
    '
    ' 4.) After selecting the Domain/Workgroup, click into either Listbox to
    '     fill List1 with server names. Combo1 Lost_Focus calls the enumeration
    '     function NetServerEnum.
    '
    ' How the program works:
    '
    ' This project collects network information using the unicode NET APIs (NETAPI32.DLL)
    ' then walks through a static collection of default share names (C$, D$, etc.). A
    ' UNC path is created \\<servername>\<default share> that is in turn passed to the
    ' GetDiskFreeSpaceEx API (which accepts UNC paths as well as conventional drive info).
    ' The output is checked and if the return is not 0 MB (no drive), then it adds it
    ' to a viewable list. Note that the Currency Data type is replacing Long Data types
    ' in the GetDiskFreeSpaceEx Declare (See code for more info)
    '
    ' NOTE: Thanks to all of the folks in the newsgroups and on the lists for all the
    '       assistance with this and other NT projects.
    
    'Initialize Public Colllection
    Public Drives As New Collection
    
    ' Define GB and MB values
    Const lGByte As Long = 1073741824
    Const lMByte As Long = 1048576

    
    Private Sub Combo2_LostFocus()
        ' Once the Domain/Workgroup is selected, fill
        ' List1 with any located NT Servers or Workstations
        ListServers (SV_TYPE_NT)
    End Sub
    
    Private Sub Command2_Click()
        Unload Me
        End
    End Sub
    
    Private Sub Form_Load()
        ' get the Domain/Workgroup info, fill Combo2 list
        GetDomain (SV_TYPE_DOMAIN_ENUM)
                           
        ' Retrieve stored default Domain/Workgroup setting, if it exists
        Combo2.Text = GetSetting("ShowDrives", "Configuration", "DefaultDomain", "Pick One")
        
        ' Add the names of possible default Administrative
        ' shares to the Public Collection 'Drives'
                
        Drives.Add "C$"
        Drives.Add "D$"
        Drives.Add "E$"
        Drives.Add "F$"
        Drives.Add "G$"
        Drives.Add "H$"
        Drives.Add "I$"
        Drives.Add "J$"
        Drives.Add "K$"
        Drives.Add "L$"
        Drives.Add "M$"
        Drives.Add "N$"
        Drives.Add "O$"
        Drives.Add "P$"
        ' If you have more physical drives in your system than this, and
        ' they are not in a RAID configuration, then I think you need help ;-)
        Drives.Add "Q$"
        Drives.Add "R$"
        Drives.Add "S$"
        Drives.Add "T$"
        Drives.Add "U$"
        Drives.Add "V$"
        ' For the real sickos
        Drives.Add "W$"
        Drives.Add "X$"
        Drives.Add "Y$"
        Drives.Add "Z$"
        
    End Sub

    Public Sub Command1_Click()
    
        Dim vDrive As Variant
    
        List2.Clear
        
        ' Add the name of the Server you are checking to
        ' the top of the list of drives
        
        List2.AddItem "Server Name: " & List1.Text
        List2.AddItem "-----"
        
        ' Walk through the collection of possible default shares
        
        For Each vDrive In Drives
            FindDriveInfo (CStr(vDrive))
        Next


    End Sub
    
    Public Sub FindDriveInfo(sDrive As String)
            
        Dim lReturn As Long
        Dim cBytesToCall As Currency
        Dim cBytesOnDrive As Currency
        Dim cFreeBytes As Currency
        Dim cUsedBytes As Currency
        Dim sServer As String
        Dim Output1 As String
        Dim Output2 As String
        Dim Output3 As String

        sServer = "\\" & List1.Text & "\" & sDrive
         
        ' Use in place of GetDiskFreeSpaceEx to bypass the
        ' 2 GB limit.  Use the Currency Data Type multiplied by
        ' 10000 so that 9,223,372,036,854,775,807 is the limit
        ' of the return value (see "Harcore Visual Basic", Chpt. 2
        ' for more information
                
        lReturn = GetDiskFreeSpaceEx(sServer, cBytesToCall, cBytesOnDrive, cFreeBytes)

        If lReturn = 0 Then
            Exit Sub
        End If

        If lReturn <> 0 Then
        
        ' Format the output based on the returned value so that
        ' the string output is in MBs or GBs.
    
            ' Total Free Bytes on Drive
            Output1 = Format((cBytesOnDrive * 10000) / IIf(cBytesOnDrive * 10000 >= lGByte, lGByte, lMByte), IIf(cBytesOnDrive * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                      
            ' Total Available Bytes on Drive
            Output2 = Format((cFreeBytes * 10000) / IIf(cFreeBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cFreeBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
                      
            ' Calculate Used Bytes
            cUsedBytes = cBytesOnDrive - cFreeBytes
            
            ' Total Used Bytes on Drive
            Output3 = Format((cUsedBytes * 10000) / IIf(cUsedBytes * 10000 >= lGByte, lGByte, lMByte), IIf(cUsedBytes * 10000 >= lGByte, "##0.### GB", "##0.### MB"))
        
        End If
        
        ' If Output1 (Total Free Bytes on Drive) is 0, then
        ' you know the drive does not exist.  Otherwise, port
        ' the output to List2
        
        If Output1 <> "0.MB" Then
            List2.AddItem "Drive Letter: " & sDrive
            List2.AddItem " Total Bytes on Drive: " & Output1
            List2.AddItem " Total Bytes Free: " & Output2
            List2.AddItem " Total Bytes Used: " & Output3
            List2.AddItem " "
        End If
        
        
    End Sub
    
    Private Sub ListServers(lType As Long)
    
    Dim lReturn As Long
    Dim Server_Info As Long
    Dim lEntries As Long
    Dim lTotal As Long
    Dim lMax As Long
    Dim vResume As Variant
    Dim tServer_info_101 As SERVER_INFO_101
    Dim sServer As String
    Dim sDomain As String
    Dim lServerInfo101StructPtr As Long
    Dim X As Long, i As Long
    Dim bBuffer(512) As Byte
    
    ' Get the selected Domain/Workgroup name, convert to Unicode
    sDomain = StrConv(Combo2.Text, vbUnicode)
    
        List1.Clear
        
        ' Enumerate servers
        lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    
        If lReturn <> 0 Then
            StatusBar1.Panels("msg").Text = "Error: " & Err.LastDllError & "  has occurred"
            Exit Sub
        End If
    
        X = 1
        lServerInfo101StructPtr = Server_Info
        
        Do While X <= lTotal
            
            ' Copy the returned data into the SERVER_INFO_101 struct
            RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
            
            ' Copy the strings from SERVER_INFO_101 (ptr_name) into bBuffer
            lstrcpyW bBuffer(0), tServer_info_101.ptr_name

            i = 0
            Do While bBuffer(i) <> 0
            
                ' Strip NULLS from the server name
                sServer = sServer & Chr$(bBuffer(i))
                i = i + 2
            Loop
            
            ' Add the name to the list
            List1.AddItem sServer
            DoEvents
            
            X = X + 1
                
            sServer = ""
            
            ' Goto next record
            lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
    
        Loop
        
        'Free up memory buffer
        lReturn = NetApiBufferFree(Server_Info)
    
    
    End Sub

    Private Sub GetDomain(lType As Long)
    
    Dim lReturn As Long
    Dim Server_Info As Long
    Dim lEntries As Long
    Dim lTotal As Long
    Dim lMax As Long
    Dim vResume As Variant
    Dim tServer_info_101 As SERVER_INFO_101
    Dim sServer As String
    Dim sDomain As String
    Dim lServerInfo101StructPtr As Long
    Dim X As Long, i As Long
    Dim bBuffer(512) As Byte
    
    ' See the above Sub for info on this Sub, it's the same
    
        Combo2.Clear
    
        lReturn = NetServerEnum(ByVal 0&, 101, Server_Info, lMax, lEntries, lTotal, ByVal lType, sDomain, vResume)
    
        If lReturn <> 0 Then
            StatusBar1.Panels("msg").Text = "Error: " & Err.LastDllError & "  has occurred"
            Exit Sub
        End If
    
        X = 1
        lServerInfo101StructPtr = Server_Info
    
        Do While X <= lTotal
            RtlMoveMemory tServer_info_101, ByVal lServerInfo101StructPtr, Len(tServer_info_101)
            lstrcpyW bBuffer(0), tServer_info_101.ptr_name
            i = 0
            Do While bBuffer(i) <> 0
                sServer = sServer & Chr$(bBuffer(i))
                i = i + 2
            Loop
            Combo2.AddItem sServer
            DoEvents
            X = X + 1
                sServer = ""
            lServerInfo101StructPtr = lServerInfo101StructPtr + Len(tServer_info_101)
        Loop
        lReturn = NetApiBufferFree(Server_Info)
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' Save the default Domain/Workgroup name
    SaveSetting "ShowDrives", "Configuration", "DefaultDomain", Combo2.Text
    
End Sub
