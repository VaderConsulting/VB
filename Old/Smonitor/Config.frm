VERSION 5.00
Begin VB.Form Config 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration Options"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "Config.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "View Security Events:"
      Height          =   855
      Left            =   3360
      TabIndex        =   26
      Top             =   3480
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "Oldest Records First"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Newest Records First"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Polling Interval:"
      Height          =   615
      Left            =   3360
      TabIndex        =   23
      Top             =   2760
      Width           =   2895
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Time (In Minutes):"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close && Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notification Options:"
      Height          =   2535
      Left            =   3360
      TabIndex        =   15
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Text            =   "Enter system name here"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox Check17 
         Caption         =   " Alert Someone"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Text            =   "Enter e-mail address here"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CheckBox Check16 
         Caption         =   " E-Mail Someone"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Text            =   "Enter pager address here"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox Check15 
         Caption         =   " Page Someone"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Security Events to Monitor:"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox ChkCustom 
         Caption         =   "Custom Event:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   3720
         Width           =   615
      End
      Begin VB.CheckBox Chk512 
         Caption         =   "512 - System Restart"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox Chk513 
         Caption         =   "513 - System Shutdown"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox Chk517 
         Caption         =   "517 - Audit Log Cleared"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox Chk528 
         Caption         =   "528 - Successful Logon"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox Chk529 
         Caption         =   "529 - Bad Username/Password"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox Chk530 
         Caption         =   "530 - Logon Restriction Violation"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox Chk531 
         Caption         =   "531 - Account Disabled"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox Chk532 
         Caption         =   "532 - Account Expired"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox Chk533 
         Caption         =   "533 - User Not Allowed to Logon"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CheckBox Chk534 
         Caption         =   "534 - Logon Type Restricted"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox Chk535 
         Caption         =   "535 - Password Expired"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox Chk537 
         Caption         =   "537 - Unsuccessful Logon"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CheckBox Chk538 
         Caption         =   "538 - User Logoff"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox Chk539 
         Caption         =   "539 - Account Locked Out"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   3480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Private Sub Check15_Click()
    
        If Check15 = 1 Then Text1.Enabled = True
        If Check15 = 0 Then Text1.Enabled = False
    
    End Sub
    
    Private Sub Check16_Click()
    
        If Check16 = 1 Then Text2.Enabled = True
        If Check16 = 0 Then Text2.Enabled = False
    
    End Sub
    
    Private Sub Check17_Click()
    
        If Check17 = 1 Then Text3.Enabled = True
        If Check17 = 0 Then Text3.Enabled = False
        
    End Sub
    
    Private Sub ChkCustom_Click()
    
        If ChkCustom.Value = 1 Then Text5.Enabled = True
        If ChkCustom.Value = 0 Then Text5.Enabled = False
    
    End Sub
    
    Private Sub Command1_Click()
    
        SaveSetting "SMonitor", "Configuration", "Chk512", Chk512.Value
        SaveSetting "SMonitor", "Configuration", "Chk513", Chk513.Value
        SaveSetting "SMonitor", "Configuration", "Chk517", Chk517.Value
        SaveSetting "SMonitor", "Configuration", "Chk528", Chk528.Value
        SaveSetting "SMonitor", "Configuration", "Chk529", Chk529.Value
        SaveSetting "SMonitor", "Configuration", "Chk530", Chk530.Value
        SaveSetting "SMonitor", "Configuration", "Chk531", Chk531.Value
        SaveSetting "SMonitor", "Configuration", "Chk532", Chk532.Value
        SaveSetting "SMonitor", "Configuration", "Chk533", Chk533.Value
        SaveSetting "SMonitor", "Configuration", "Chk534", Chk534.Value
        SaveSetting "SMonitor", "Configuration", "Chk535", Chk535.Value
        SaveSetting "SMonitor", "Configuration", "Chk537", Chk537.Value
        SaveSetting "SMonitor", "Configuration", "Chk538", Chk538.Value
        SaveSetting "SMonitor", "Configuration", "Chk539", Chk539.Value
        SaveSetting "SMonitor", "Configuration", "ChkCustom", ChkCustom.Value
        SaveSetting "SMonitor", "Configuration", "PageChk", Check15.Value
        SaveSetting "SMonitor", "Configuration", "MailChk", Check16.Value
        SaveSetting "SMonitor", "Configuration", "AlertChk", Check17.Value
        SaveSetting "SMonitor", "Configuration", "PageAddress", Text1.Text
        SaveSetting "SMonitor", "Configuration", "MailAddress", Text2.Text
        SaveSetting "SMonitor", "Configuration", "AlertSrv", Text3.Text
        SaveSetting "SMonitor", "Configuration", "Timeout", Text4.Text
        SaveSetting "SMonitor", "Configuration", "CustomEvt", Text5.Text
        SaveSetting "SMonitor", "Configuration", "NewestRecords", Option1.Value
        SaveSetting "SMonitor", "Configuration", "OldestRecords", Option2.Value
        Config.Hide
    
    End Sub
    
    Private Sub Form_Load()
    
        Chk512.Value = Val(GetSetting("SMonitor", "Configuration", "Chk512", "0"))
        Chk513.Value = Val(GetSetting("SMonitor", "Configuration", "Chk513", "0"))
        Chk517.Value = Val(GetSetting("SMonitor", "Configuration", "Chk517", "0"))
        Chk528.Value = Val(GetSetting("SMonitor", "Configuration", "Chk528", "0"))
        Chk529.Value = Val(GetSetting("SMonitor", "Configuration", "Chk529", "1"))
        Chk530.Value = Val(GetSetting("SMonitor", "Configuration", "Chk530", "0"))
        Chk531.Value = Val(GetSetting("SMonitor", "Configuration", "Chk531", "0"))
        Chk532.Value = Val(GetSetting("SMonitor", "Configuration", "Chk532", "0"))
        Chk533.Value = Val(GetSetting("SMonitor", "Configuration", "Chk533", "0"))
        Chk534.Value = Val(GetSetting("SMonitor", "Configuration", "Chk534", "0"))
        Chk535.Value = Val(GetSetting("SMonitor", "Configuration", "Chk535", "0"))
        Chk537.Value = Val(GetSetting("SMonitor", "Configuration", "Chk537", "0"))
        Chk538.Value = Val(GetSetting("SMonitor", "Configuration", "Chk538", "0"))
        Chk539.Value = Val(GetSetting("SMonitor", "Configuration", "Chk539", "1"))
        ChkCustom.Value = Val(GetSetting("SMonitor", "Configuration", "ChkCustom", "0"))
        Check15.Value = Val(GetSetting("SMonitor", "Configuration", "PageChk", "0"))
        Text1.Text = GetSetting("SMonitor", "Configuration", "PageAddress", "")
        Check16.Value = Val(GetSetting("SMonitor", "Configuration", "MailChk", "0"))
        Text2.Text = GetSetting("SMonitor", "Configuration", "MailAddress", "")
        Check17.Value = Val(GetSetting("SMonitor", "Configuration", "AlertChk", "0"))
        Text3.Text = GetSetting("SMonitor", "Configuration", "AlertSrv", GetLocalSystemName)
        Text4.Text = GetSetting("SMonitor", "Configuration", "Timeout", "5")
        Text5.Text = GetSetting("SMonitor", "Configuration", "CustomEvt", "")
        Option1.Value = GetSetting("SMonitor", "Configuration", "NewestRecords", "True")
        Option2.Value = GetSetting("SMonitor", "Configuration", "OldestRecords", "False")
        
        If Check15 = 1 Then Text1.Enabled = True
        If Check15 = 0 Then Text1.Enabled = False
        If Check16 = 1 Then Text2.Enabled = True
        If Check16 = 0 Then Text2.Enabled = False
        If Check17 = 1 Then Text3.Enabled = True
        If Check17 = 0 Then Text3.Enabled = False
        If ChkCustom.Value = 1 Then Text5.Enabled = True
        If ChkCustom.Value = 0 Then Text5.Enabled = False
        
    End Sub
    
