VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmColors 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Color Settings"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "FrmColors.frx":0000
   ScaleHeight     =   5370
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2400
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Image Image14 
      Height          =   180
      Left            =   3000
      Picture         =   "FrmColors.frx":35B8
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Apply"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3240
      TabIndex        =   17
      Tag             =   "TitleColor"
      Top             =   2985
      Width           =   1005
   End
   Begin VB.Image Image13 
      Height          =   180
      Left            =   2160
      Picture         =   "FrmColors.frx":3973
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image12 
      Height          =   180
      Left            =   2160
      Picture         =   "FrmColors.frx":3D2E
      Top             =   4920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2880
      TabIndex        =   16
      Tag             =   "TitleColor"
      Top             =   3525
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2880
      TabIndex        =   15
      Tag             =   "TitleColor"
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Image imgSelected2 
      Height          =   180
      Index           =   2
      Left            =   480
      Picture         =   "FrmColors.frx":4107
      Top             =   3480
      Width           =   180
   End
   Begin VB.Image imgSelected2 
      Height          =   180
      Index           =   1
      Left            =   480
      Picture         =   "FrmColors.frx":44F4
      Top             =   3120
      Width           =   180
   End
   Begin VB.Image imgSelected2 
      Height          =   180
      Index           =   0
      Left            =   480
      Picture         =   "FrmColors.frx":48E1
      Top             =   2760
      Width           =   180
   End
   Begin VB.Image imgSelected1 
      Height          =   180
      Index           =   1
      Left            =   2640
      Picture         =   "FrmColors.frx":4CCE
      Top             =   1440
      Width           =   180
   End
   Begin VB.Image imgSelected1 
      Height          =   180
      Index           =   0
      Left            =   2640
      Picture         =   "FrmColors.frx":50BB
      Top             =   960
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   2
      Left            =   480
      Picture         =   "FrmColors.frx":54A8
      Top             =   1560
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   1
      Left            =   480
      Picture         =   "FrmColors.frx":5895
      Top             =   1200
      Width           =   180
   End
   Begin VB.Image imgSelected 
      Height          =   180
      Index           =   0
      Left            =   480
      Picture         =   "FrmColors.frx":5C82
      Top             =   840
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   180
      Left            =   1920
      Picture         =   "FrmColors.frx":606F
      Top             =   4560
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image9 
      Height          =   180
      Left            =   1920
      Picture         =   "FrmColors.frx":645C
      Top             =   4920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   4200
      Picture         =   "FrmColors.frx":6854
      Tag             =   "Close"
      ToolTipText     =   "Close"
      Top             =   50
      Width           =   315
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   14
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   3120
      Width           =   825
   End
   Begin VB.Shape Shape3 
      Height          =   1575
      Left            =   240
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   13
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   3480
      Width           =   750
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   12
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   2760
      Width           =   840
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   11
      Tag             =   "TitleColor"
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Form"
      Height          =   225
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3000
      TabIndex        =   9
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   1440
      Width           =   825
   End
   Begin VB.Shape Shape2 
      Height          =   1575
      Left            =   2400
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3000
      TabIndex        =   8
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2640
      TabIndex        =   7
      Tag             =   "TitleColor"
      Top             =   1965
      Width           =   1485
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats Form"
      Height          =   225
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Main Form"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   4
      Tag             =   "TitleColor"
      Top             =   1965
      Width           =   1485
   End
   Begin VB.Image Image4 
      Height          =   330
      Left            =   120
      Picture         =   "FrmColors.frx":6D76
      Top             =   4560
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image5 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmColors.frx":7476
      Top             =   4560
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image Image7 
      Height          =   330
      Left            =   120
      Picture         =   "FrmColors.frx":796B
      Top             =   4920
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image8 
      Height          =   315
      Left            =   1560
      Picture         =   "FrmColors.frx":807A
      Top             =   4920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   3
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bar Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   2
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   1560
      Width           =   750
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   240
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Tag             =   "TitleColor"
      Top             =   0
      Width           =   885
   End
   Begin VB.Label lblLabelColor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Color"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   840
      TabIndex        =   0
      Tag             =   "Label"
      ToolTipText     =   "Click here to set the color of all labels"
      Top             =   1200
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   480
      Picture         =   "FrmColors.frx":859C
      ToolTipText     =   "Settings"
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   2640
      Picture         =   "FrmColors.frx":8CAB
      ToolTipText     =   "Settings"
      Top             =   1920
      Width           =   1425
   End
   Begin VB.Image Image3 
      Height          =   330
      Left            =   480
      Picture         =   "FrmColors.frx":93BA
      ToolTipText     =   "Settings"
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Image Image10 
      Height          =   330
      Left            =   2880
      Picture         =   "FrmColors.frx":9AC9
      ToolTipText     =   "Settings"
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Image Image11 
      Height          =   330
      Left            =   2880
      Picture         =   "FrmColors.frx":A1D8
      ToolTipText     =   "Settings"
      Top             =   3480
      Width           =   1425
   End
End
Attribute VB_Name = "FrmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For Dragging Borderless Forms...
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim sControlSelected As String
Dim sControlSelected1 As String
Dim sControlSelected2 As String
Private Sub Form_Load()
Dim AutoApply As String
AutoApply = ReadINI("colors", "autoapply", App.Path & "\settings.ini")

If AutoApply = "unchecked" Then
Image14.Picture = Image13.Picture
End If
If AutoApply = "checked" Then
Image14.Picture = Image12.Picture
End If
Me.Height = 4485
Me.Width = 4590

sControlSelected = "text"
sControlSelected1 = "text1"
sControlSelected2 = "text2"
Call imgSelected_Click(0)
Call imgSelected1_Click(0)
Call imgSelected2_Click(0)
Call LoadColors
Call SetColors(Me)

Me.Top = Screen.Height / 2 - Me.Height / 2 + 350
Me.Left = Screen.Width / 2 - Me.Width / 2
DoEvents
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
    DoEvents
End If
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image5.Picture
End If

End Sub
Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    imgClose.Picture = Image8.Picture
End If

End Sub

Private Sub imgSelected_Click(Index As Integer)
Dim X As Byte

'Clear the radio buttons...
For X = 0 To 2
    imgSelected(X).Picture = Image6.Picture
Next X

'Update the radio buttons...
imgSelected(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected = "title"
    Case 1
        sControlSelected = "text"
    Case 2
        sControlSelected = "bar"
End Select
End Sub

Private Sub imgSelected1_Click(Index As Integer)
Dim X As Byte

'Clear the radio buttons...
For X = 0 To 1
    imgSelected1(X).Picture = Image6.Picture
Next X

'Update the radio buttons...
imgSelected1(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected1 = "title1"
    Case 1
        sControlSelected1 = "text1"
End Select
End Sub

Private Sub imgSelected2_Click(Index As Integer)
Dim X As Byte

'Clear the radio buttons...
For X = 0 To 2
    imgSelected2(X).Picture = Image6.Picture
Next X

'Update the radio buttons...
imgSelected2(Index).Picture = Image9.Picture

'Remember the control selected...
Select Case Index
    Case 0
        sControlSelected2 = "title2"
    Case 1
        sControlSelected2 = "text2"
    Case 2
        sControlSelected2 = "bar2"
End Select
End Sub

Private Sub Label10_Click()
On Local Error Resume Next

'Set the Dialog boxes color to the color of the control that is selected...
If sControlSelected2 = "title2" Then
    Dialog.Color = Label11.ForeColor
ElseIf sControlSelected2 = "text2" Then
    Dialog.Color = Label14.ForeColor
ElseIf sControlSelected2 = "bar2" Then
    Dialog.Color = Label13.ForeColor
End If

'Show the color Dialog box...
Dialog.Flags = cdlCCFullOpen Or cdlCCRGBInit
Dialog.ShowColor
If Err.Number > 0 Then Exit Sub
DoEvents
'Set the color to the control currently selected...
If sControlSelected2 = "title2" Then
    Label11.ForeColor = Dialog.Color
End If
If sControlSelected2 = "text2" Then
    Label14.ForeColor = Dialog.Color
End If
If sControlSelected2 = "bar2" Then
    Label13.ForeColor = Dialog.Color
End If
DoEvents
    If Image14.Picture = Image12.Picture Then
    Call Label15_Click
    End If
    
End Sub

Private Sub Label12_Click()
On Local Error Resume Next

'Set the Dialog boxes color to the color of the control that is selected...
If sControlSelected = "title" Then
    Dialog.Color = Label2.ForeColor
ElseIf sControlSelected = "text" Then
    Dialog.Color = lblLabelColor.ForeColor
ElseIf sControlSelected = "bar" Then
    Dialog.Color = Label1.ForeColor
End If

'Show the color Dialog box...
Dialog.Flags = cdlCCFullOpen Or cdlCCRGBInit
Dialog.ShowColor
If Err.Number > 0 Then Exit Sub
DoEvents
'Set the color to the control currently selected...
If sControlSelected = "title" Then
    Label2.ForeColor = Dialog.Color
End If
If sControlSelected = "text" Then
    lblLabelColor.ForeColor = Dialog.Color
End If
If sControlSelected = "bar" Then
    Label1.ForeColor = Dialog.Color
End If
DoEvents
    If Image14.Picture = Image12.Picture Then
    Call Label15_Click
    End If
    
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image2.Picture = Image4.Picture
End If
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image2.Picture = Image7.Picture
End If
End Sub

Private Sub Label15_Click()
Call Me.SaveChanges
DoEvents
Call LoadColors
Call SetColors(Me)
Call SetColors(FrmMain)
Call SetColors(FrmStats)
Call SetColors(FrmDesktop)
End Sub
Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image11.Picture = Image4.Picture
End If
End Sub
Private Sub Label15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Image11.Picture = Image7.Picture
End If
End Sub
Private Sub Label5_Click()
On Local Error Resume Next

'Set the Dialog boxes color to the color of the control that is selected...
If sControlSelected1 = "title1" Then
    Dialog.Color = Label6.ForeColor
ElseIf sControlSelected1 = "text1" Then
    Dialog.Color = Label8.ForeColor
End If

'Show the color Dialog box...
Dialog.Flags = cdlCCFullOpen Or cdlCCRGBInit
Dialog.ShowColor
If Err.Number > 0 Then Exit Sub
DoEvents
'Set the color to the control currently selected...
If sControlSelected1 = "title1" Then
    Label6.ForeColor = Dialog.Color
End If
If sControlSelected1 = "text1" Then
    Label8.ForeColor = Dialog.Color
End If
DoEvents
    If Image14.Picture = Image12.Picture Then
    Call Label15_Click
    End If
    
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image1.Picture = Image4.Picture
End If
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image1.Picture = Image7.Picture
End If
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image4.Picture
End If
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image3.Picture = Image7.Picture
End If
End Sub
Sub SaveChanges()

'On Local Error Resume Next

'Save the color settings to the skin scheme ini file...
Call WriteINI("colors", "maintextcolors", lblLabelColor.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "maintitlecolors", Label2.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "mainbarcolors", Label1.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "desktoptextcolors", Label14.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "desktoptitlecolors", Label11.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "desktopbarcolors", Label13.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "statstitlecolors", Label6.ForeColor, App.Path & "\settings.ini")
Call WriteINI("colors", "statstextcolors", Label8.ForeColor, App.Path & "\settings.ini")
    If Image14.Picture = Image13.Picture Then
    Call WriteINI("colors", "autoapply", "unchecked", App.Path & "\settings.ini")
    End If
    If Image14.Picture = Image12.Picture Then
    Call WriteINI("colors", "autoapply", "checked", App.Path & "\settings.ini")
    End If
End Sub

Private Sub Label7_Click()
Unload Me
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Image10.Picture = Image4.Picture
End If
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Image10.Picture = Image7.Picture
End If
End Sub
Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Image14.Picture = Image13.Picture Then
    Image14.Picture = Image12.Picture
    Exit Sub
    End If
    If Image14.Picture = Image12.Picture Then
    Image14.Picture = Image13.Picture
    End If
End If

End Sub
Public Sub DragForm(Frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
