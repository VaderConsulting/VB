VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim OsInfo As OSVERSIONINFO
    Dim retval As Integer
    'Get the Operating System Version
    OsInfo.dwOSVersionInfoSize = 148
    OsInfo.szCSDVersion = Space$(128)
    retval = GetVersionExA(OsInfo)
    Os = OsInfo.dwMajorVersion
    
    'Default Values
    Text1.Text = Environ("USERNAME")
    Text3.Text = Environ("USERDOMAIN")
    Command1.Caption = "(...)"
    Command2.Caption = "&Install"
    Command3.Caption = "&Exit"

End Sub

Private Sub Command1_Click()
    'Select a File to Run.
    With CommonDialog1
        .Filter = "Msi Files|*.msi|All Files|*.*"
        .ShowOpen
        Text4.Text = .FileName
    End With
End Sub

Private Sub Command2_Click()
    Dim retval As Long
    
    If UCase(Mid(Trim(Text4.Text), (Len(Text4.Text) - 3), 4)) = ".MSI" Then
        If Left(Text4.Text, 1) <> Space(1) Then Text4.Text = " /i " & Text4.Text
        TheApplication = "C:\Winnt\System32\MsiExec.exe"
    Else
        TheApplication = Trim(Text4.Text)
        Text4.Text = ""
    End If
    
    Select Case Os
        Case 4
            'NT4.0
            retval = WinNTSetup(Trim(Text1.Text), Trim(Text2.Text), _
            Trim(Text3.Text), Text4.Text, CurDir)
        Case 5
            'Windows2K
            retval = Win2KSetup(Trim(Text1.Text), Trim(Text2.Text), _
            Trim(Text3.Text), Text4.Text)
    End Select

End Sub

Private Sub Command3_Click()
    Unload Me
End Sub


