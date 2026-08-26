VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   14940
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   9360
      TabIndex        =   20
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   9360
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   1485
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Form1.frx":0000
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Internet"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Network"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Info"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "RAS"
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Online"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Proxy"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Modem"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "LAN"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "AOL"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "RAS"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "LAN"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Type"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Alive"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const NETWORK_ALIVE_LAN = &H1  'net card connection
Private Const NETWORK_ALIVE_WAN = &H2  'RAS connection
Private Const NETWORK_ALIVE_AOL = &H4  'AOL
         
Private Const MAX_INTERFACE_NAME_LEN               As Long = 256
Private Const ERROR_SUCCESS                        As Long = 0
Private Const MAXLEN_IFDESCR                       As Long = 256
Private Const MAXLEN_PHYSADDR                      As Long = 8

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL   As Long = 0
Private Const MIB_IF_OPER_STATUS_UNREACHABLE       As Long = 1
Private Const MIB_IF_OPER_STATUS_DISCONNECTED      As Long = 2
Private Const MIB_IF_OPER_STATUS_CONNECTING        As Long = 3
Private Const MIB_IF_OPER_STATUS_CONNECTED         As Long = 4
Private Const MIB_IF_OPER_STATUS_OPERATIONAL       As Long = 5

Private Const MIB_IF_TYPE_OTHER                    As Long = 1
Private Const MIB_IF_TYPE_ETHERNET                 As Long = 6
Private Const MIB_IF_TYPE_TOKENRING                As Long = 9
Private Const MIB_IF_TYPE_FDDI                     As Long = 15
Private Const MIB_IF_TYPE_PPP                      As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK                 As Long = 24
Private Const MIB_IF_TYPE_SLIP                     As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP               As Long = 1
Private Const MIB_IF_ADMIN_STATUS_DOWN             As Long = 2
Private Const MIB_IF_ADMIN_STATUS_TESTING          As Long = 3
    
Private Type MIB_IFROW
    wszName(0 To (MAX_INTERFACE_NAME_LEN - 1) * 2) As Byte
    dwIndex                                        As Long
    dwType                                         As Long
    dwMtu                                          As Long
    dwSpeed                                        As Long
    dwPhysAddrLen                                  As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1)            As Byte
    dwAdminStatus                                  As Long
    dwOperStatus                                   As Long
    dwLastChange                                   As Long
    dwInOctets                                     As Long
    dwInUcastPkts                                  As Long
    dwInNUcastPkts                                 As Long
    dwInDiscards                                   As Long
    dwInErrors                                     As Long
    dwInUnknownProtos                              As Long
    dwOutOctets                                    As Long
    dwOutUcastPkts                                 As Long
    dwOutNUcastPkts                                As Long
    dwOutDiscards                                  As Long
    dwOutErrors                                    As Long
    dwOutQLen                                      As Long
    dwDescrLen                                     As Long
    bDescr(0 To MAXLEN_IFDESCR - 1)                As Byte

End Type
    
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function inet_ntoa Lib "wsock32" (ByVal addr As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long
Private Declare Function GetFriendlyIfIndex Lib "iphlpapi" (ByVal IfIndex As Long) As Long
Private Declare Function IsNetworkAlive Lib "Sensapi" (lpdwFlags As Long) As Long

Private Sub Command1_Click()
    Text1.Text = IsNetConnectionAlive()
    Text2.Text = GetNetConnectionType()
    
    Text3.Text = IsNetConnectionLAN()
    Text4.Text = IsNetConnectionRAS()
    Text5.Text = IsNetConnectionAOL()
End Sub


Private Function IsNetConnectionAlive() As Boolean
    Dim tmp As Long
    
    IsNetConnectionAlive = IsNetworkAlive(tmp) = 1
End Function


Private Function IsNetConnectionLAN() As Boolean
    Dim tmp As Long
    
    If IsNetworkAlive(tmp) = 1 Then
        IsNetConnectionLAN = tmp = NETWORK_ALIVE_LAN
    End If
End Function


Private Function IsNetConnectionRAS() As Boolean
    Dim tmp As Long
    
    If IsNetworkAlive(tmp) = 1 Then
        IsNetConnectionRAS = tmp = NETWORK_ALIVE_WAN
    End If
End Function


Private Function IsNetConnectionAOL() As Boolean
    Dim tmp As Long
    
    If IsNetworkAlive(tmp) = 1 Then
        IsNetConnectionAOL = tmp = NETWORK_ALIVE_AOL
    End If
End Function


Private Function GetNetConnectionType() As String
    Dim tmp As Long

    If IsNetworkAlive(tmp) = 1 Then
        Select Case tmp
            Case NETWORK_ALIVE_LAN
                GetNetConnectionType = "The system has one or more active LAN cards"
            Case NETWORK_ALIVE_WAN
                GetNetConnectionType = "The system has one or more active RAS connections"
            Case NETWORK_ALIVE_AOL
                GetNetConnectionType = "The system is connected to America Online"
            Case Else
        End Select
    Else
        GetNetConnectionType = "The system has no connection or an error occurred"
    End If
End Function


Private Sub Command2_Click()
    Text1.Text = IsNetConnectViaLAN()
    Text2.Text = IsNetConnectViaModem()
    Text3.Text = IsNetConnectViaProxy()
    Text4.Text = IsNetConnectOnline()
    Text5.Text = IsNetRASInstalled()
    Text6.Text = GetNetConnectString()
End Sub

Private Sub Command3_Click()
    Dim IPInterfaceRow As MIB_IFROW
    Dim buff() As Byte
    Dim cbRequired As Long
    Dim nStructSize As Long
    Dim nRows As Long
    Dim cnt As Long
    Dim n As Long
    Dim itmx As ListItem
    Dim tmp As String
    
    Call GetIfTable(ByVal 0&, cbRequired, 1)

    If cbRequired > 0 Then
     
        ReDim buff(0 To cbRequired - 1) As Byte
        
        If GetIfTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
            'saves using LenB in the CopyMemory calls below
            nStructSize = LenB(IPInterfaceRow)
            'first 4 bytes is a long indicating the number of entries in the table
            CopyMemory nRows, buff(0), 4
            For cnt = 1 To nRows
                'moving past the four bytes obtained above, get one chunk of data and cast into an IPInterfaceRow type
                CopyMemory IPInterfaceRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
                
                With ListView1
                
                    .ColumnHeaders.Add , , "Adapter " & CStr(cnt)
                    
                    Set itmx = .ListItems(1)
                    itmx.SubItems(cnt) = TrimNull(StrConv(IPInterfaceRow.bDescr, vbUnicode))

                    Set itmx = .ListItems(2)
                    itmx.SubItems(cnt) = GetFriendlyIfIndex(IPInterfaceRow.dwIndex)
                    
                    Select Case IPInterfaceRow.dwType
                        Case MIB_IF_TYPE_ETHERNET:     tmp = "Ethernet"
                        Case MIB_IF_TYPE_TOKENRING:    tmp = "TokenRing"
                        Case MIB_IF_TYPE_FDDI:          tmp = "FDDI"
                        Case MIB_IF_TYPE_PPP:            tmp = "Point-to-Point"
                        Case MIB_IF_TYPE_LOOPBACK:     tmp = "Loopback"
                        Case MIB_IF_TYPE_SLIP:          tmp = "Slip"
                        Case MIB_IF_TYPE_OTHER:         tmp = "Other"
                    End Select
                        
                    Set itmx = .ListItems(3)
                    itmx.SubItems(cnt) = IPInterfaceRow.dwType & " " & tmp
                    tmp = ""
                    
                    Set itmx = .ListItems(4)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwMtu, 0)
                    
                    Set itmx = .ListItems(5)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwSpeed, 0)
                    
                    
                    For n = 0 To IPInterfaceRow.dwPhysAddrLen - 1
                        tmp = tmp & IPInterfaceRow.bPhysAddr(n) & " "
                        Next
                    Print
                    Set itmx = .ListItems(6)
                    itmx.SubItems(cnt) = tmp
                    tmp = ""
                    
                    For n = 0 To IPInterfaceRow.dwPhysAddrLen - 1
                        tmp = tmp & Hex(IPInterfaceRow.bPhysAddr(n)) & " "
                        Next
                    Print
                    
                    Set itmx = .ListItems(7)
                    itmx.SubItems(cnt) = tmp
                    tmp = ""
                    
                    Select Case IPInterfaceRow.dwAdminStatus
                  
                        Case MIB_IF_ADMIN_STATUS_UP:        tmp = "Enabled"
                        Case MIB_IF_ADMIN_STATUS_DOWN:     tmp = "Disabled"
                        Case MIB_IF_ADMIN_STATUS_TESTING: tmp = "Testing"
    
                    End Select
                
                    Set itmx = .ListItems(8)
                    itmx.SubItems(cnt) = IPInterfaceRow.dwAdminStatus & " " & tmp
                    tmp = ""
                    
                    Select Case IPInterfaceRow.dwOperStatus
                    
                        Case MIB_IF_OPER_STATUS_NON_OPERATIONAL:  tmp = "Non-operational"
                        Case MIB_IF_OPER_STATUS_UNREACHABLE:        tmp = "Unreachable"
                        Case MIB_IF_OPER_STATUS_DISCONNECTED:      tmp = "Disconnected"
                        Case MIB_IF_OPER_STATUS_CONNECTING:         tmp = "Connecting"
                        Case MIB_IF_OPER_STATUS_CONNECTED:          tmp = "Connected"
                        Case MIB_IF_OPER_STATUS_OPERATIONAL:        tmp = "Operational"
                    End Select
                    
                    Set itmx = .ListItems(9)
                    itmx.SubItems(cnt) = IPInterfaceRow.dwOperStatus & " " & tmp
                    tmp = ""
                    
                    Set itmx = .ListItems(10)
                    itmx.SubItems(cnt) = IPInterfaceRow.dwLastChange
                    
                    Set itmx = .ListItems(11)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInOctets, 0)
                    
                    Set itmx = .ListItems(12)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInUcastPkts, 0)
                    
                    Set itmx = .ListItems(13)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInNUcastPkts, 0)
                    
                    Set itmx = .ListItems(14)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInDiscards, 0)
                    
                    Set itmx = .ListItems(15)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInErrors, 0)
                    
                    Set itmx = .ListItems(16)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwInUnknownProtos, 0)
                    
                    Set itmx = .ListItems(17)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutOctets, 0)
                    
                    Set itmx = .ListItems(18)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutUcastPkts, 0)
                    
                    Set itmx = .ListItems(19)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutNUcastPkts, 0)
                    
                    Set itmx = .ListItems(20)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutDiscards, 0)
                    
                    Set itmx = .ListItems(21)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutErrors, 0)
                    
                    Set itmx = .ListItems(22)
                    itmx.SubItems(cnt) = FormatNumber(IPInterfaceRow.dwOutQLen, 0)
                    
                End With  'Listview1

             Next cnt
             
        End If  'If GetIfTable( ...
        
    End If  'If cbRequired > 0
End Sub

' ----------------------------------------------------------------------------------

Private Sub Form_Load()
    Dim itmx As ListItem
     
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add , , "Information"
     
    With ListView1.ListItems
         Set itmx = .Add(, "bDescr", "description of interface")
          
         Set itmx = .Add(, , "interface index")
         Set itmx = .Add(, , "interface type")
         Set itmx = .Add(, , "Maximum Transmission Unit")
         Set itmx = .Add(, , "interface speed (bps)")
         Set itmx = .Add(, , "physical address (decimal)")
         Set itmx = .Add(, , "physical address (hex)")
         Set itmx = .Add(, , "admin enabled or disabled")
         Set itmx = .Add(, , "interface operational status")
         Set itmx = .Add(, , "last time op status changed")
          
         Set itmx = .Add(, , "data received (octets)")
         Set itmx = .Add(, , "packets received (unicast)")
         Set itmx = .Add(, , "packets received (non-unicast)")
         Set itmx = .Add(, , "packets discarded")
         Set itmx = .Add(, , "discarded with errors")
         Set itmx = .Add(, , "discarded, unknown protocol")
          
         Set itmx = .Add(, , "data sent (octets)")
         Set itmx = .Add(, , "packets sent (unicast)")
         Set itmx = .Add(, , "packets sent (non-unicast)")
         Set itmx = .Add(, , "packets discarded, no errors")
         Set itmx = .Add(, , "packets discarded with errors")
         Set itmx = .Add(, , "output queue length")
    End With
End Sub


Public Function GetInetStrFromPtr(ByVal Address As Long) As String
    GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))
End Function


Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function


Private Function TrimNull(item As String)
    Dim pos As Integer
     
    pos = InStr(item, Chr$(0))
     
    If pos Then
         TrimNull = Left$(item, pos - 1)
    Else
         TrimNull = item
    End If
End Function

