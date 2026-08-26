VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Earthquake Options"
   ClientHeight    =   3330
   ClientLeft      =   4725
   ClientTop       =   3015
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSep 
      Height          =   75
      Left            =   -180
      TabIndex        =   4
      Top             =   2760
      Width           =   6915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4500
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin HotKeyDemo.HotKey hkyEarthquake 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotKeyModifier  =   0
   End
   Begin HotKeyDemo.HotKey hkyAssistant 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotKeyModifier  =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Office Assistant Killer HotKey:"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   6675
   End
   Begin VB.Label lblHotKey 
      Caption         =   "Earthquake HotKey:"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6675
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_eMod(1 To 2) As EHKModifiers
Private m_eKey(1 To 2) As KeyCodeConstants
Private m_bCancel As Boolean

Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property
Public Property Get KeyForItem(ByVal iItem As Long) As KeyCodeConstants
   KeyForItem = m_eKey(iItem)
End Property
Public Property Let KeyForItem(ByVal iItem As Long, ByVal iKey As KeyCodeConstants)
   m_eKey(iItem) = iKey
End Property
Public Property Get ModifierForItem(ByVal iItem As Long) As EHKModifiers
   ModifierForItem = m_eMod(iItem)
End Property
Public Property Let ModifierForItem(ByVal iItem As Long, ByVal eMod As EHKModifiers)
   m_eMod(iItem) = eMod
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   m_bCancel = False
   m_eKey(1) = hkyEarthquake.HotKey
   m_eMod(1) = peInterpret(hkyEarthquake.HotKeyModifier)
   m_eKey(2) = hkyAssistant.HotKey
   m_eMod(2) = peInterpret(hkyEarthquake.HotKeyModifier)
   Unload Me
End Sub
Private Function peInterpret(eModifier As echkModifierKeys) As EHKModifiers
Dim eR As EHKModifiers
   
   ' Rather annoyingly, the key codes in the HotKey control don't
   ' map onto the ones in the RegisterHotKey call!
   If (eModifier And HOTKEYF_ALT) = HOTKEYF_ALT Then
      eR = MOD_ALT
   End If
   If (eModifier And HOTKEYF_SHIFT) = HOTKEYF_SHIFT Then
      eR = eR Or MOD_SHIFT
   End If
   If (eModifier And HOTKEYF_CONTROL) = HOTKEYF_CONTROL Then
      eR = eR Or MOD_CONTROL
   End If
   If (eModifier And HOTKEYF_EXT) = HOTKEYF_EXT Then
      eR = eR Or MOD_WIN
   End If
   peInterpret = eR
End Function
Private Function peDeterpret(eModifier As EHKModifiers) As echkModifierKeys
Dim eR As echkModifierKeys

   ' Rather annoyingly, the key codes in the HotKey control don't
   ' map onto the ones in the RegisterHotKey call!
   If (eModifier And MOD_ALT) = MOD_ALT Then
      eR = HOTKEYF_ALT
   End If
   If (eModifier And MOD_SHIFT) = MOD_SHIFT Then
      eR = eR Or HOTKEYF_SHIFT
   End If
   If (eModifier And MOD_CONTROL) = MOD_CONTROL Then
      eR = eR Or HOTKEYF_CONTROL
   End If
   If (eModifier And MOD_WIN) = MOD_WIN Then
      eR = eR Or HOTKEYF_EXT
   End If
   peDeterpret = eR
   
End Function

Private Sub Form_Load()
   With hkyEarthquake
      .HotKeyModifier = peDeterpret(m_eMod(1))
      .HotKey = m_eKey(1)
      .InvalidHotKeyOperation(HKCOMB_NONE, HOTKEYF_CONTROLALT) = True
      .InvalidHotKeyOperation(HKCOMB_C, HOTKEYF_CONTROLALT) = True
      .InvalidHotKeyOperation(HKCOMB_S, HOTKEYF_SHIFTCONTROL) = True
   End With
   With hkyAssistant
      .HotKeyModifier = peDeterpret(m_eMod(2))
      .HotKey = m_eKey(2)
      .InvalidHotKeyOperation(HKCOMB_NONE, HOTKEYF_CONTROLALT) = True
      .InvalidHotKeyOperation(HKCOMB_C, HOTKEYF_CONTROLALT) = True
      .InvalidHotKeyOperation(HKCOMB_S, HOTKEYF_SHIFTCONTROL) = True
   End With
   m_bCancel = True
End Sub
