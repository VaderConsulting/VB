VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9090
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPhoto 
      Height          =   3088
      Left            =   5888
      ScaleHeight     =   3030
      ScaleWidth      =   3030
      TabIndex        =   20
      ToolTipText     =   "Photo"
      Top             =   3456
      Width           =   3088
   End
   Begin VB.TextBox txtLocation 
      Height          =   304
      Left            =   5888
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   896
      Width           =   3088
   End
   Begin VB.TextBox txtMessages 
      Height          =   4240
      Left            =   128
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2816
      Width           =   5648
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   8192
      TabIndex        =   6
      Top             =   6656
      Width           =   784
   End
   Begin VB.ListBox lstMember 
      Enabled         =   0   'False
      Height          =   1230
      ItemData        =   "frmProperties.frx":0000
      Left            =   5888
      List            =   "frmProperties.frx":0002
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1664
      Width           =   3056
   End
   Begin VB.CheckBox chkOnboard 
      Caption         =   "Onboard"
      Height          =   255
      Left            =   5888
      TabIndex        =   7
      Top             =   128
      Width           =   1335
   End
   Begin VB.TextBox txtSurname 
      Height          =   285
      Left            =   1536
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   128
      Width           =   4215
   End
   Begin VB.TextBox txtRank 
      Height          =   288
      Left            =   1536
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   896
      Width           =   4215
   End
   Begin VB.TextBox txtBillet 
      Height          =   285
      Left            =   1536
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1280
      Width           =   4215
   End
   Begin VB.TextBox txtBilletDesc 
      Height          =   285
      Left            =   1536
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1664
      Width           =   4215
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1536
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2048
      Width           =   4215
   End
   Begin VB.TextBox txtChristian 
      Height          =   285
      Left            =   1536
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   496
      Width           =   4215
   End
   Begin VB.Label lblPhoto 
      Alignment       =   2  'Center
      Caption         =   "Photo"
      Height          =   272
      Left            =   5888
      TabIndex        =   21
      Top             =   3200
      Width           =   3088
   End
   Begin VB.Label lblLocation 
      Alignment       =   2  'Center
      Caption         =   "Location"
      Height          =   272
      Left            =   5888
      TabIndex        =   19
      Top             =   512
      Width           =   3088
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      Caption         =   "Messages"
      Height          =   272
      Left            =   128
      TabIndex        =   17
      Top             =   2560
      Width           =   5648
   End
   Begin VB.Label lblMember 
      Alignment       =   2  'Center
      Caption         =   "Member of"
      Height          =   256
      Left            =   5888
      TabIndex        =   15
      Top             =   1408
      Width           =   3056
   End
   Begin VB.Label lblSurname 
      Caption         =   "Surname"
      Height          =   256
      Left            =   128
      TabIndex        =   13
      Top             =   128
      Width           =   1328
   End
   Begin VB.Label lblRank 
      Caption         =   "Rank"
      Height          =   256
      Left            =   128
      TabIndex        =   12
      Top             =   896
      Width           =   1328
   End
   Begin VB.Label lblBillet 
      Caption         =   "Billet"
      Height          =   256
      Left            =   128
      TabIndex        =   11
      Top             =   1280
      Width           =   1328
   End
   Begin VB.Label lblBilletDesc 
      Caption         =   "Billet Description"
      Height          =   256
      Left            =   128
      TabIndex        =   10
      Top             =   1664
      Width           =   1328
   End
   Begin VB.Label lblComment 
      Caption         =   "Comment"
      Height          =   256
      Left            =   128
      TabIndex        =   9
      Top             =   2048
      Width           =   1328
   End
   Begin VB.Label lblChristian 
      Caption         =   "Christian Name"
      Height          =   256
      Left            =   128
      TabIndex        =   8
      Top             =   512
      Width           =   1328
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If txtLocation = "ONBOARD" Then
        chkOnboard.Value = vbChecked
        txtLocation.Visible = False
    End If
    If txtLocation <> "ONBOARD" Then
        chkOnboard.Value = vbUnchecked
        txtLocation.Visible = True
    End If
End Sub

Public Sub Form_Load()
    If gClickOrigin = "In" Then ' from lstIn
        PersString = frmPersonnel.lstIn.List(frmPersonnel.lstIn.ListIndex)
    End If
    If gClickOrigin = "Out" Then ' from lstOut
        PersString = frmPersonnel.lstOut.List(frmPersonnel.lstOut.ListIndex)
    End If
    If gClickOrigin = "Menu" Then ' select source
        'if lstin.
        reply = MsgBox("Select 'Yes' for In data, Select 'No' for Out data, Select 'Cancel' to cancel", vbYesNoCancel + vbDefaultButton1, "Select data source")
        If reply = vbYes Then PersString = frmPersonnel.lstIn.List(frmPersonnel.lstIn.ListIndex)
        If reply = vbNo Then PersString = frmPersonnel.lstOut.List(frmPersonnel.lstOut.ListIndex)
        If reply = vbCancel Then Exit Sub
    End If
    PersLen = Len(PersString) - 1
    Start = InStr(1, PersString, "(")
    PersID = Mid$(PersString, Start + 1, PersLen)
    PersID = Val(Left$(PersID, Len(PersID) - 1))
    'frmPersonnel.fctPopulateTextBoxes (PersID)
End Sub
