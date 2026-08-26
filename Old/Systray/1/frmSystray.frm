VERSION 5.00
Begin VB.Form frmSystray 
   Caption         =   "Sample systray application"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   Icon            =   "FRMSYS~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   3240
      Top             =   450
   End
   Begin VB.Frame fraToolTip 
      Height          =   1095
      Left            =   1260
      TabIndex        =   5
      Top             =   810
      Width           =   2625
      Begin VB.CommandButton cmdChangeTip 
         Caption         =   "Change Tip"
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtToolTip 
         Height          =   285
         Left            =   90
         MaxLength       =   64
         TabIndex        =   1
         Text            =   "Cypher's Systray Example"
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label lblInfo2 
         Caption         =   "Change ToolTip Text"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   1635
      End
   End
   Begin VB.Frame fraFlashIcon 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   810
      Width           =   1275
      Begin VB.CheckBox chkFlash 
         Caption         =   "Flash Icon"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image imgFlashIcon 
         Height          =   480
         Left            =   450
         Picture         =   "FRMSYS~1.frx":0442
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   $"FRMSYS~1.frx":0884
      Height          =   645
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   3795
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "System Tray Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuFlash 
         Caption         =   "&Flash Icon"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////
'//Project:         Sample Code Library
'//Date:            May 2nd, 1999  8:30AM EST
'//Programmer:      Robert J. Reich  (cypher@tir.com)
'//Company:         CypherSolutions
'//
'//Name:            System Tray Application
'//Description:     Demonstrates how to create an application that
'//                 resides in the systray instead of on the taskbar.
'//                 This is done with the WIN32 API.
'//
'//Note:            Usually the two APIs, the constants, and UDT will
'//                 all be declared in a stadard module and not in a
'//                 form.  This is only done here to make the example
'//                 as simple as possible.
'//////////////////////////////////////////////////////////////////
'//
'//WARNING:         If you run this in the IDE, do NOT use the VCR-STOP
'//                 button to end this program.  That will bypass the
'//                 Form_Unload event which is neccisary to restore
'//                 the system tray back to it's original state.
'//////////////////////////////////////////////////////////////////


'//These are the two API functions we'll need to use here.  The first
'//is the one that really does the work.  The second function is used
'//to take action when the user clicks on the mouse icon (restores the
'//program and brings it to front of all other windows.)
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
          (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32" _
          (ByVal hwnd As Long) As Long


'//UDT required by Shell_NotifyIcon API call
Private Type NOTIFYICONDATA
 cbSize As Long             '//size of this UDT
 hwnd As Long               '//handle of the app
 uId As Long                '//unused (set to vbNull)
 uFlags As Long             '//Flags needed for actions
 uCallBackMessage As Long   '//WM we are going to subclass
 hIcon As Long              '//Icon we're going to use for the systray
 szTip As String * 64       '//ToolTip for the mouse_over of the icon.
End Type


'//Constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0             '//Flag : "ALL NEW nid"
Private Const NIM_MODIFY = &H1          '//Flag : "ONLY MODIFYING nid"
Private Const NIM_DELETE = &H2          '//Flag : "DELETE THE CURRENT nid"
Private Const NIF_MESSAGE = &H1         '//Flag : "Message in nid is valid"
Private Const NIF_ICON = &H2            '//Flag : "Icon in nid is valid"
Private Const NIF_TIP = &H4             '//Flag : "Tip in nid is valid"
Private Const WM_MOUSEMOVE = &H200      '//This is our CallBack Message
Private Const WM_LBUTTONDOWN = &H201    '//LButton down
Private Const WM_LBUTTONUP = &H202      '//LButton up
Private Const WM_LBUTTONDBLCLK = &H203  '//LDouble-click
Private Const WM_RBUTTONDOWN = &H204    '//RButton down
Private Const WM_RBUTTONUP = &H205      '//RButton up
Private Const WM_RBUTTONDBLCLK = &H206  '//RDouble-click

Private nid As NOTIFYICONDATA       '//global UDT for the systray function


Private Sub Form_Activate()
'//////////////////////////////////////////////////////////////////
'//Purpose:         Load up the UDT for the Systray Function.  This
'//                 must be done after the form is fully visable.
'//                 The Form_Activate is a perfect place for that.
'//////////////////////////////////////////////////////////////////
 
  With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Cypher's Systray Example" & vbNullChar
  End With
 
  Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                                X As Single, Y As Single)
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is the callback function of icon in the
'//                 system tray.  This is where will will process
'//                 what the application will do when Mouse Input
'//                 is given to the icon.
'//
'//Inputs:          What Button was clicked (this is button & shift),
'//                 also, the X & Y coordinates of the mouse.
'//////////////////////////////////////////////////////////////////

  Dim msg As Long     '//The callback value
  
  '//The value of X will vary depending
  '//upon the ScaleMode setting.  Here
  '//we are using that fact to determine
  '//what the value of 'msg' should really be
  If (Me.ScaleMode = vbPixels) Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If

  Select Case msg
    Case WM_LBUTTONDBLCLK    '515 restore form window
      Me.WindowState = vbNormal
      Call SetForegroundWindow(Me.hwnd)
      Me.Show
      
    Case WM_RBUTTONUP        '517 display popup menu
      Call SetForegroundWindow(Me.hwnd)
      Me.PopupMenu Me.mnuSystray
    
    Case WM_LBUTTONUP        '514 restore form window
      '//commonly an application on the
      '//systray will do nothing on a
      '//single mouse_click, so nothing
  End Select

  '//small note:  I just learned that when using a Select Case
  '//structure you always want to place the most commonly anticipated
  '//action highest. Saves CPU cycles becuase of less evaluations.
End Sub

Private Sub Form_Resize()
'//////////////////////////////////////////////////////////////////
'//Purpose:         This is just to check to make sure that, if
'//                 indeed the application is minimized (hence on
'//                 the systray) to also hide the form.
'//////////////////////////////////////////////////////////////////
  If (Me.WindowState = vbMinimized) Then Me.Hide
End Sub

Private Sub mnuFlash_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         To do EXACTLY what the chkFlash does.  In fact
'//                 it actually calls that Click event and lets it
'//                 do all the work.
'//////////////////////////////////////////////////////////////////
  Call chkFlash_Click     '//let it do the work for us.
End Sub

Private Sub mnuRestore_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will restore it.
'//////////////////////////////////////////////////////////////////
  Me.WindowState = vbNormal
  Call SetForegroundWindow(Me.hwnd)
  Me.Show
End Sub

Private Sub mnuExit_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         When the application is minimized on the systray
'//                 this will close the application.
'//////////////////////////////////////////////////////////////////
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'//////////////////////////////////////////////////////////////////
'//Purpose:         Deletes the systray icon, and makes the application
'//                 "safe" to unload.
'//////////////////////////////////////////////////////////////////
   Shell_NotifyIcon NIM_DELETE, nid
   Set frmSystray = Nothing
End Sub

Private Sub chkFlash_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         Change the "Enabled" state of the Timer to
'//                 it's inverse.  While the Timer is enabled it will
'//                 cause the icon to flash, else it will stop.
'//////////////////////////////////////////////////////////////////
  Dim nidflash As NOTIFYICONDATA

  If (Timer1.Enabled) Then
    chkFlash.Value = vbUnchecked
    mnuFlash.Checked = False
    Timer1.Enabled = False
    With nidflash
    '//we've disabled the flash, we've
    '//got to make sure that we've got
    '//our default icon back.
      .cbSize = Len(nidflash)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON
      .hIcon = Me.Icon
    End With
    Shell_NotifyIcon NIM_MODIFY, nidflash
  Else
    chkFlash.Value = vbChecked
    mnuFlash.Checked = True
    Timer1.Enabled = True
  End If
End Sub

Private Sub cmdChangeTip_Click()
'//////////////////////////////////////////////////////////////////
'//Purpose:         Change the ToolTip of the System Tray Icon
'//////////////////////////////////////////////////////////////////
  Dim nidNewTip As NOTIFYICONDATA     '//New ToolTip nid
    
    
  With nidNewTip
    .cbSize = Len(nidNewTip)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_TIP       '//Here the Tip is the only valid "new data"
    .szTip = txtToolTip.Text & vbNullChar
  End With
    
  Shell_NotifyIcon NIM_MODIFY, nidNewTip
End Sub

Private Sub Timer1_Timer()
'//////////////////////////////////////////////////////////////////
'//Purpose:         Creates a new NOTIFYICONDATA UDT and use that to
'//                 modify the current options.  This is done to
'//                 "flash" the icon in the system tray.
'//////////////////////////////////////////////////////////////////
  Dim nidflash As NOTIFYICONDATA     '//New "Flashing" nid
  Static SwitchIcons As Boolean      '//Flag to decide whether to
                                     '//use the form icon, or the
                                     '//flash icon.
                                     '//(True=flash icon)
                                     '//(False=form icon)
  
  With nidflash
  '//only thing we're really changing from
  '//the original structure is the icon.
  '//Hence the NIF_ICON flag.
    .cbSize = Len(nidflash)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON
    
    If (SwitchIcons) Then
    '//we need to change it to the
    '//"non-app" icon.
      .hIcon = imgFlashIcon.Picture
      SwitchIcons = False
    Else
      .hIcon = Me.Icon
      SwitchIcons = True
    End If
  End With
  
  Shell_NotifyIcon NIM_MODIFY, nidflash
End Sub
