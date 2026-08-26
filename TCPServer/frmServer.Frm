VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   2430
   ClientLeft      =   210
   ClientTop       =   2700
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2430
   ScaleWidth      =   8520
   Begin VB.ListBox lstConnections 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Index           =   0
      Left            =   8520
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngCurrentPortIndex As Long
Private lngMaxPorts As Long
Private bAuthorised As Boolean
Private iState As Integer

' State Machine
'   0: Connected
'   1: Waiting for Authorisation string
'   2: Waiting for Identification string

Private Sub Form_Load()
    lngCurrentPortIndex = 0
    lngMaxPorts = 5
    tcpServer(lngCurrentPortIndex).LocalPort = 1001
    tcpServer(lngCurrentPortIndex).Listen
End Sub

Private Sub tcpServer_Connect(Index As Integer)
    SendData Index, "***AUTHORISATION***"
End Sub

Private Sub tcpServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim IndexToAssign As Integer
    If Index = 0 Then
        lngCurrentPortIndex = lngCurrentPortIndex + 1
        If lngCurrentPortIndex > lngMaxPorts Then
            IndexToAssign = NextPort ' Find the first available port
        Else
            IndexToAssign = lngCurrentPortIndex
        End If
        lstConnections.AddItem "-!- (" & IndexToAssign & ") Connected", 0
        
        If IndexToAssign <> 0 Then
            Load tcpServer(IndexToAssign)
            tcpServer(IndexToAssign).LocalPort = 0
            tcpServer(IndexToAssign).Accept requestID
            lstConnections.AddItem "--> (" & IndexToAssign & ") ***CONNECT***", 0
        Else ' no free port found
            lstConnections.AddItem "-X- Cannot accept additional connection requests", 0
            Disconnect 0
        End If
    End If
End Sub

Private Sub tcpServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String, bResult As Boolean
    Dim Path As String
    
    Static StateMachine  '  <---- this is why a class module is required.. to keep states between connections separate
    tcpServer(Index).GetData strData
    
    lstConnections.AddItem "<-- (" & Index & ") " & strData, 0
    
    Debug.Print "<-- (" & Index & ") " & strData, 0
    
    bResult = strData Like "CON:LOGONDATA*" ' <-- Mod this.. looks strange
    If bResult Then
        ' Compose drive paths for the connecting user, send back to this connection
        Path = "P:+\\ServerA\ShareA+T:+\\ServerB\ShareB+N:+\\ServerC\ShareC"
        SendData Index, "PATH:" & Path
    ElseIf strData = "***DISCONNECT***" Then
        SendData Index, "***DISCONNECT***"
        Disconnect Index
    End If
    
    If Not bAuthorised Then
        If strData = "AUTH:LOGON" Then
            bAuthorised = True
            Exit Sub
        Else
            lstConnections.AddItem "--> (" & IndexToAssign & ") ***NO AUTH***", 0
        End If
    End If
End Sub

Private Sub Disconnect(Index As Integer)
    Unload tcpServer(Index)
    lstConnections.AddItem "-X- (" & Index & ") Disconnected", 0
End Sub

Function NextPort() As Integer
    Dim IndexToAssign As Long, lngState As Long
    
    For IndexToAssign = 1 To lngMaxPorts
        'lstConnections.AddItem "Is " & IndexToAssign & " free ?", 0
        If tcpServer(IndexToAssign) Is Nothing Then
            NextPort = IndexToAssign
            'lstConnections.AddItem IndexToAssign & " is free", 0
            Exit Function
        Else
            On Error Resume Next
            lngState = tcpServer(IndexToAssign).State
            If Err.Number <> 0 Then
                NextPort = IndexToAssign
                Exit Function
            End If
            On Error GoTo 0
            'If tcpServer(IndexToAssign).State = sckClosed Then
            '    NextPort = IndexToAssign
            '    lstConnections.AddItem IndexToAssign & " is closed", 0
            'Else
            '    NextPort = 0
            '    lstConnections.AddItem "Could not find a free port", 0
            'End If
        End If
    Next IndexToAssign
End Function

Sub SendData(Index As Integer, strData As String)
    lstConnections.AddItem "--> (" & Index & ") " & strData, 0
    tcpServer(Index).SendData strData
End Sub
