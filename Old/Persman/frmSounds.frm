VERSION 5.00
Begin VB.Form frmSounds 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sounds"
   ClientHeight    =   1888
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   8320
   Icon            =   "frmSounds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1888
   ScaleWidth      =   8320
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSound 
      Caption         =   "16"
      Height          =   528
      Index           =   15
      Left            =   7296
      TabIndex        =   16
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "15"
      Height          =   528
      Index           =   14
      Left            =   6272
      TabIndex        =   15
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   7424
      TabIndex        =   14
      Top             =   1408
      Width           =   784
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "14"
      Height          =   528
      Index           =   13
      Left            =   5248
      TabIndex        =   13
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "13"
      Height          =   528
      Index           =   12
      Left            =   4224
      TabIndex        =   12
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "12"
      Height          =   528
      Index           =   11
      Left            =   3200
      TabIndex        =   11
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "11"
      Height          =   528
      Index           =   10
      Left            =   2176
      TabIndex        =   10
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "10"
      Height          =   528
      Index           =   9
      Left            =   1152
      TabIndex        =   9
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "9"
      Height          =   528
      Index           =   8
      Left            =   128
      TabIndex        =   8
      Top             =   768
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "8"
      Height          =   528
      Index           =   7
      Left            =   7296
      TabIndex        =   7
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "7"
      Height          =   528
      Index           =   6
      Left            =   6272
      TabIndex        =   6
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "6"
      Height          =   528
      Index           =   5
      Left            =   5248
      TabIndex        =   5
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "5"
      Height          =   528
      Index           =   4
      Left            =   4224
      TabIndex        =   4
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "4"
      Height          =   528
      Index           =   3
      Left            =   3200
      TabIndex        =   3
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "3"
      Height          =   528
      Index           =   2
      Left            =   2176
      TabIndex        =   2
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "2"
      Height          =   528
      Index           =   1
      Left            =   1152
      TabIndex        =   1
      Top             =   128
      Width           =   912
   End
   Begin VB.CommandButton cmdSound 
      Caption         =   "1"
      Height          =   528
      Index           =   0
      Left            =   128
      TabIndex        =   0
      Top             =   128
      Width           =   912
   End
End
Attribute VB_Name = "frmSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
