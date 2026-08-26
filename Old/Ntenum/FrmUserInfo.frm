VERSION 5.00
Begin VB.Form FrmUserInfo 
   Caption         =   "UserInfo"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbGlobalGroups 
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Press button to view memberships"
      Top             =   5265
      Width           =   3000
   End
   Begin VB.TextBox TxtWorkstationsAllowed 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4935
      Width           =   3000
   End
   Begin VB.TextBox TxtHomedirDrive 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4575
      Width           =   3000
   End
   Begin VB.TextBox TxtHomedir 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4335
      Width           =   3000
   End
   Begin VB.TextBox TxtLoginScriptName 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4095
      Width           =   3000
   End
   Begin VB.TextBox TxtUserProfilePath 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3855
      Width           =   3000
   End
   Begin VB.TextBox TxtBadPwCount 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3495
      Width           =   3000
   End
   Begin VB.TextBox TxtPasswordNeverExpires 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3255
      Width           =   3000
   End
   Begin VB.TextBox TxtPasswordChangeable 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3015
      Width           =   3000
   End
   Begin VB.TextBox TxtPasswordExpired 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2775
      Width           =   3000
   End
   Begin VB.TextBox TxtPasswordLastSet 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2535
      Width           =   3000
   End
   Begin VB.TextBox TxtNumLogons 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2175
      Width           =   3000
   End
   Begin VB.TextBox TxtLastLogon 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1935
      Width           =   3000
   End
   Begin VB.TextBox TxtAcctExpires 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1575
      Width           =   3000
   End
   Begin VB.TextBox TxtAcctLockedOut 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1335
      Width           =   3000
   End
   Begin VB.TextBox TxtAcctActive 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1095
      Width           =   3000
   End
   Begin VB.TextBox TxtPrivilege 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   735
      Width           =   3000
   End
   Begin VB.TextBox TxtDescription 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   495
      Width           =   3000
   End
   Begin VB.TextBox TxtFullName 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   255
      Width           =   3000
   End
   Begin VB.TextBox TxtUserName 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label30 
      Caption         =   "Global Groups ......................"
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   5295
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Workstations allowed ..........."
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   4950
      Width           =   2055
   End
   Begin VB.Label Label28 
      Caption         =   "Home dir drive ......................"
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   4590
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Home dir .............................."
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   4350
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Login script name ................."
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   4110
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "User profile path ..................."
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   3870
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Last logon ............................"
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   1950
      Width           =   2055
   End
   Begin VB.Label Label21 
      Caption         =   "Num logons .........................."
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   2190
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Bad pw count ......................"
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   3510
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Acct locked out ..................."
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   1350
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Password expires ................."
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   3270
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Password changeable ........."
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   3030
      Width           =   2055
   End
   Begin VB.Label Label29 
      Caption         =   "Password expired ................."
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   2790
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Password last set ................."
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   2550
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Acct active .........................."
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   1110
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Privilege ..............................."
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   750
      Width           =   2055
   End
   Begin VB.Label Label16 
      Caption         =   "Acct expires ........................."
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   1590
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Description ..........................."
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   510
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Full name ..........................."
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   270
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Username .........................."
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   15
      Width           =   1935
   End
End
Attribute VB_Name = "FrmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

