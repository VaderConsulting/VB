VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Registry"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSiteloc 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtHostname 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "C01961890"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblSiteloc 
      Caption         =   "Siteloc"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
    If txtHostname = "" Then Exit Sub
    'retVal = RegQueryValueEx(hRegistry, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\SiteLoc", DummyVar, lpType, Value, ValueLength)
    dummy = 1
    s = Siteloc
End Sub

Function Siteloc() As String
    Dim hRegistry As Long, varResult As Long, Value As String, RetVal As Long
    RetVal = RegConnectRegistry(txtHostname, HKEY_LOCAL_MACHINE, hRegistry)
    Siteloc = RegGetString(hRegistry, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Siteloc")
End Function

Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
    Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long
    Dim R As Long
    
    RetVal$ = ""
    R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
    If R <> ERROR_SUCCESS Then GoTo Quit_Now
    
    SZ = 256: v$ = String$(SZ, 0)
    R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If R = ERROR_SUCCESS And dwType = REG_SZ Then
        RetVal$ = Left$(v$, SZ)
    Else
        RetVal$ = "--Not String--"
    End If
    If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
    RegGetString$ = RetVal$
End Function

