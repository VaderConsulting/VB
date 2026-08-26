VERSION 5.00
Begin VB.Form frmOnOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On / Off"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton optChoice 
      Caption         =   "Surname"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.OptionButton optChoice 
      Caption         =   "Personal Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox txtLocationCode 
      Height          =   285
      Left            =   4440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtPersonalNumber 
      Height          =   285
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblName 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblLocation 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label lblLocationCode 
      Caption         =   "Location Code"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmOnOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Len(txtPersonalNumber) > 1 Then txtLocationCode.SetFocus
End Sub

Private Sub optChoice_Click(Index As Integer)
    If Index = 1 Then
        txtPersonalNumber.Enabled = False
        txtName.Enabled = True
        txtName.SetFocus
    Else
        txtPersonalNumber.Enabled = True
        txtPersonalNumber.SetFocus
        txtName.Enabled = False
    End If
End Sub

Private Sub txtLocationCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub ' backspace
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtLocationCode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And txtLocationCode <> "" And frmLeave.tmrDelay.Interval = 250 Then
        KeyCode = 0
        txtLocationCode.Refresh
        lblLocation = gLocationText(txtLocationCode)
        lblLocation.Refresh
        a = 0
        Do Until (gPersonalNumber(a) = txtPersonalNumber) Or (a = 232)
            a = a + 1
        Loop
        If a = 232 Then
            MsgBox "Unknown personal number"
            frmLeave.tmrDelay.Enabled = True
            frmLeave.KeyPreview = False
            Unload Me
        End If
        gLocation(a) = Val(txtLocationCode)
        Select Case gLocation(a)
            Case 1
                frmLeave.imgLocation(a).Picture = frmLeave.imgLeave.Picture
            Case 2
                frmLeave.imgLocation(a).Picture = frmLeave.imgHome.Picture
            Case 3
                frmLeave.imgLocation(a).Picture = frmLeave.imgOut.Picture
            Case 4
                frmLeave.imgLocation(a).Picture = frmLeave.imgOut.Picture
        End Select
        gTime(a) = Time
        frmLeave.tmrDelay.Enabled = True
        frmLeave.KeyPreview = False
        Unload Me
    End If
End Sub

Private Sub txtPersonalNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub ' backspace
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtPersonalNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        MsgBox "CR character by itself is an error"
        a = 0
        Do Until (gPersonalNumber(a) = txtPersonalNumber) Or (a = 232)
            a = a + 1
        Loop
        KeyCode = 0
        ' check if the person is already out
        If gLocation(a) <> 0 Then
            frmLeave.imgLocation(a).Picture = frmLeave.imgIn.Picture
            gLocation(a) = 0
            frmLeave.tmrDelay.Enabled = True
            frmLeave.KeyPreview = False
            Unload Me
        Else
            ' not out
            txtLocationCode.SetFocus
            ' txtPersonalNumber = Left$(txtPersonalNumber.Text, Len(txtPersonalNumber) - 2)
            For a = 0 To 232
                If txtPersonalNumber = gPersonalNumber(a) Then lblName = gName(a)
            Next a
        End If
    End If
End Sub
