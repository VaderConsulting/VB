VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Monitor"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   2460
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrRefresh 
      Interval        =   1000
      Left            =   3600
      Top             =   120
   End
   Begin VB.Label lblOutRAS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblOutLAN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblInRAS 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblInLAN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblRASSpeed 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblLANSpeed 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Image imgDisconnected 
      Height          =   360
      Left            =   4800
      Picture         =   "frmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgConnected 
      Height          =   360
      Left            =   4320
      Picture         =   "frmMain.frx":0894
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgConn 
      Height          =   375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image imgModem 
      Height          =   375
      Left            =   1800
      Picture         =   "frmMain.frx":09DE
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgLAN 
      Height          =   375
      Left            =   840
      Picture         =   "frmMain.frx":0CE8
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape shpConnection 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub tmrRefresh_Timer()
    GetInterfaceInfo
End Sub

Private Sub GetInterfaceInfo()
    Dim IPInterfaceRow As MIB_IFROW
    Dim buff() As Byte
    Dim cbRequired As Long
    Dim nStructSize As Long
    Dim nRows As Long
    Dim cnt As Long
    Dim n As Long
    Dim tmp As String
    Dim bIsLANConnected As Boolean
    Dim bIsRASConnected As Boolean
    Dim bIsLANConnecting As Boolean
    Dim bIsRASConnecting As Boolean
    Dim bIsConnected As Boolean
    Dim bIsAlive As Boolean
    Dim lLANSpeed As Long
    Dim lRASSpeed As Long
    Dim lTemp As Long
    Dim lDiff As Long
    
    Static lOutRAS As Long
    Static lOutLAN As Long
    Static lInRAS As Long
    Static lInLAN As Long
    
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
                    
                Select Case IPInterfaceRow.dwType
                    Case MIB_IF_TYPE_ETHERNET, MIB_IF_TYPE_TOKENRING, MIB_IF_TYPE_FDDI, MIB_IF_TYPE_OTHER
                        Select Case IPInterfaceRow.dwOperStatus
                            Case MIB_IF_OPER_STATUS_NON_OPERATIONAL, MIB_IF_OPER_STATUS_UNREACHABLE, MIB_IF_OPER_STATUS_DISCONNECTED
                                'bIsLANConnected = False
                                'lblInLAN = ""
                                'lblOutLAN = ""
                            Case MIB_IF_OPER_STATUS_CONNECTING
                                bIsLANConnecting = True
                                bIsLANConnected = False
                                bIsAlive = False
                                'lblInLAN = ""
                                'lblOutLAN = ""
                            Case MIB_IF_OPER_STATUS_CONNECTED, MIB_IF_OPER_STATUS_OPERATIONAL
                                bIsLANConnecting = False
                                bIsLANConnected = True
                                If IPInterfaceRow.dwSpeed > lLANSpeed Then
                                    lLANSpeed = IPInterfaceRow.dwSpeed
                                End If
                                lDiff = 0
                                lTemp = IPInterfaceRow.dwInOctets
                                If lTemp > lInLAN Then
                                    lDiff = lTemp - lInLAN
                                    lInLAN = lTemp
                                End If
                                
                                lblInLAN = lDiff & " Bps In"
                                lDiff = 0
                                lTemp = IPInterfaceRow.dwOutOctets
                                If lTemp > lOutLAN Then
                                    lDiff = lTemp - lOutLAN
                                    lOutLAN = lTemp
                                End If
                                
                                lblOutLAN = lDiff & " Bps Out"
                        End Select
                    Case MIB_IF_TYPE_PPP, MIB_IF_TYPE_SLIP
                        Select Case IPInterfaceRow.dwOperStatus
                            Case MIB_IF_OPER_STATUS_NON_OPERATIONAL, MIB_IF_OPER_STATUS_UNREACHABLE, MIB_IF_OPER_STATUS_DISCONNECTED
                                'bIsRASConnected = False
                                'lblInRAS = ""
                                'lblOutRAS = ""
                            Case MIB_IF_OPER_STATUS_CONNECTING
                                bIsRASConnecting = True
                                'lblInRAS = ""
                                'lblOutRAS = ""
                            Case MIB_IF_OPER_STATUS_CONNECTED, MIB_IF_OPER_STATUS_OPERATIONAL
                                bIsRASConnecting = False
                                bIsRASConnected = True
                                If IPInterfaceRow.dwSpeed > lRASSpeed Then
                                    lRASSpeed = IPInterfaceRow.dwSpeed
                                End If
                                
                                lTemp = IPInterfaceRow.dwInOctets
                                If lTemp > lInLAN Then
                                    lDiff = lTemp - lInRAS
                                    lInRAS = lTemp
                                    lblInRAS = lDiff & " Bps In"
                                End If
                                lTemp = IPInterfaceRow.dwOutOctets
                                If lTemp > lOutLAN Then
                                    lDiff = lTemp - lOutRAS
                                    lOutRAS = lTemp
                                    lblOutRAS = lDiff & " Bps Out"
                                End If
                        End Select
                    Case MIB_IF_TYPE_LOOPBACK
                        ' Should always have a loopback adapter
                End Select
                
             Next cnt
             bIsAlive = bIsLANConnected Or bIsRASConnected
             If bIsAlive Then
                imgConn.Picture = imgConnected.Picture
                shpConnection.FillColor = &H80FF80 ' Green
                imgLAN.Visible = bIsLANConnected
                imgModem.Visible = bIsRASConnected
                If lLANSpeed > 0 Then
                    lblLANSpeed = CStr((CLng(lLANSpeed) / 1000000)) & " Mbps"
                Else
                    lblLANSpeed = ""
                End If
                If lRASSpeed > 0 Then
                    lblRASSpeed = CStr((CLng(lRASSpeed) / 1000)) & " Kbps"
                Else
                    lblRASSpeed = ""
                End If
            Else
                shpConnection.FillColor = &H8080FF ' Red
                imgConn.Picture = imgDisconnected.Picture
                imgLAN.Visible = False
                imgModem.Visible = False
                lblLANSpeed = ""
                lblRASSpeed = ""
            End If
        End If  'If GetIfTable( ...
        
    End If  'If cbRequired > 0
End Sub

