VERSION 5.00
Begin VB.Form frmHotKeyDemo 
   Caption         =   "Completely Pointless Demo of HotKey Response..."
   ClientHeight    =   3990
   ClientLeft      =   2745
   ClientTop       =   2955
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6795
   Begin VB.CommandButton cmdConfigure 
      Caption         =   "&Configure.."
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.Timer tmrEarthquake 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   420
      Top             =   180
   End
End
Attribute VB_Name = "frmHotKeyDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cHK As cRegHotKey
Attribute m_cHK.VB_VarHelpID = -1
Private m_iType As Long
Private m_iKey(1 To 2) As KeyCodeConstants
Private m_iMod(1 To 2) As EHKModifiers

Private Sub pAttachKeys()
   m_cHK.RegisterKey "Earthquake", m_iKey(1), m_iMod(1)
   m_cHK.RegisterKey "GoodbyeAssistant", m_iKey(2), m_iMod(2)
End Sub
Private Sub pDetachKeys()
   m_cHK.UnregisterKey "Earthquake"
   m_cHK.UnregisterKey "GoodbyeAssistant"
End Sub

Private Sub cmdConfigure_Click()
   ' remove the hotkeys whilst we let them be set up!
   pDetachKeys
   
   Dim fc As New frmOptions
   fc.KeyForItem(1) = m_iKey(1)
   fc.ModifierForItem(1) = m_iMod(1)
   fc.KeyForItem(2) = m_iKey(2)
   fc.ModifierForItem(2) = m_iMod(2)
   fc.Show vbModal, Me
   If Not fc.Cancelled Then
      m_iKey(1) = fc.KeyForItem(1)
      m_iMod(1) = fc.ModifierForItem(1)
      m_iKey(2) = fc.KeyForItem(2)
      m_iMod(2) = fc.ModifierForItem(2)
   End If
   
   ' Attach the hotkeys again:
   pAttachKeys
   
End Sub

Private Sub Form_Load()
   ' Set default:
   Set m_cHK = New cRegHotKey
   m_cHK.Attach Me.hwnd
   m_iKey(1) = vbKeyC
   m_iMod(1) = MOD_CONTROL Or MOD_ALT
   m_iKey(2) = vbKeyX
   m_iMod(2) = MOD_CONTROL Or MOD_ALT
   m_iType = 2
   pAttachKeys
End Sub

Private Sub m_cHK_HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)
   If (sName = "Earthquake") Then
      tmrEarthquake.Enabled = Not (tmrEarthquake.Enabled)
      If (tmrEarthquake.Enabled) Then
         Me.Visible = False
      Else
         Me.Visible = True
         m_cHK.RestoreAndActivate Me.hwnd
      End If
   ElseIf (sName = "GoodbyeAssistant") Then
      KillAssistant
   End If
End Sub

Private Sub tmrEarthquake_Timer()
   DoEarthQuake 1
End Sub
