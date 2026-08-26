VERSION 5.00
Begin VB.Form frmmslot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launcher Client"
   ClientHeight    =   2145
   ClientLeft      =   1080
   ClientTop       =   1500
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4155
   End
End
Attribute VB_Name = "frmmslot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Declare Function CreateMailslot Lib "kernel32" Alias "CreateMailslotA" (ByVal lpName As String, ByVal nMaxMessageSize As Long, ByVal lReadTimeout As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFileAsync Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Const OPEN_EXISTING = 3
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Dim serverhandle As Long    ' Server Mailslot handle
Dim clienthandle As Long    ' Client Mailslot handle

Private Sub cmdConnect_Click()
    Create
    SendGUID
End Sub

Private Sub cmdDisconnect_Click()
    Destroy
End Sub

Sub SendGUID()
    Dim res As Long, byteswritten As Long, t As Integer, portion As Integer
    Dim GUID As String, lp As Integer, r As Integer
    GUID = ""
    For lp = 1 To 16
        r = Rnd(1) * 255
        GUID = GUID & Hex(r)
    Next lp
    GUID = GUID & vbCr
    For t = 1 To Len(GUID)
        portion = Asc(Mid(GUID, t, 1))
        res = WriteFile(serverhandle, portion, 2, byteswritten, 0)
    Next t
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    Randomize Timer
    serverhandle = INVALID_HANDLE_VALUE
    clienthandle = INVALID_HANDLE_VALUE
End Sub

Private Sub Create()
    If serverhandle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(serverhandle)
        serverhandle = 0
    End If
    If clienthandle <> INVALID_HANDLE_VALUE Then
        Call CloseHandle(clienthandle)
        clienthandle = 0
    End If
    
    txtInfo.Text = ""
    clienthandle = CreateMailslot("\\.\mailslot\launcher\client", 0, 0, 0)
    If clienthandle = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to create client mailbox"
    End If
    PostOwnName
    ' Open as a client
    serverhandle = CreateFile("\\warnbro\mailslot\launcher\server", GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If serverhandle = INVALID_HANDLE_VALUE Then
        MsgBox "Unable to find Launcher Server"
    End If
End Sub

Sub PostOwnName()
    Dim res As Long, byteswritten As Long, t As Integer, portion As Integer
    Dim Computername As String, lp As Integer, r As Integer
    Computername = UCase(Environ("computername"))
    
    For t = 1 To Len(Computername)
        portion = Asc(Mid(Computername, t, 1))
        res = WriteFile(clienthandle, portion, 2, byteswritten, 0)
    Next t
End Sub

Private Sub Destroy()
    Call CloseHandle(serverhandle)
    Call CloseHandle(clienthandle)
    serverhandle = 0
    clienthandle = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If serverhandle <> INVALID_HANDLE_VALUE Then
        Destroy
    End If
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
    Dim res As Long, byteswritten As Long
    If serverhandle <> INVALID_HANDLE_VALUE Then
        ' Send the message
        res = WriteFile(serverhandle, KeyAscii, 2, byteswritten, 0)
        KeyAscii = 0
    End If
End Sub
