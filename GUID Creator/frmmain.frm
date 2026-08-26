VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GUID Creator"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrCountdown 
      Interval        =   5000
      Left            =   1320
      Top             =   480
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtGUID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CheckBox chkBraces 
      Caption         =   "Braces"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblGUID 
      Caption         =   "GUID:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Function CreateGUID(Optional bWithBraces As Boolean = False) As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    lLen = 40
    bGUID = String(lLen, 0)
    
    CoCreateGuid uGUID
    
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    
    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then
        RetVal = RetVal - 1
    End If
    
    CreateGUID = LCase(Left$(sGUID, RetVal))
    
    If Not bWithBraces Then
        CreateGUID = Replace(CreateGUID, "{", "")
        CreateGUID = Replace(CreateGUID, "}", "")
    End If
    
End Function

Private Sub cmdCreate_Click()
    tmrCountdown.Enabled = False
    tmrCountdown.Interval = 5000
    If chkBraces.Value = vbChecked Then
        txtGUID.Text = CreateGUID(True)
    Else
        txtGUID.Text = CreateGUID(False)
    End If
    Clipboard.SetText txtGUID.Text
    lblInfo.Caption = "This GUID has been added to the clipboard.  You may now paste it using CTRL-V."
    tmrCountdown.Enabled = True
End Sub

Private Sub cmdHelp_Click()
    MsgBox "Contact Dave Robinson via Email (vader@iinet.net.au) or on 0417 927 828 for assistance.", vbOKOnly + vbInformation, "Help"
End Sub

Private Sub tmrCountdown_Timer()
    lblInfo.Caption = "Idle"
    tmrCountdown.Interval = "5000"
    tmrCountdown.Enabled = False
End Sub
