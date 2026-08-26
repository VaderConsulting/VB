VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optMode 
      Caption         =   "Strict search"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Ends with"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Starts with"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Find anywhere"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   4080
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox lstFound 
      Height          =   3180
      ItemData        =   "frmSearch.frx":6DC2
      Left            =   120
      List            =   "frmSearch.frx":6DC4
      TabIndex        =   4
      Top             =   720
      Width           =   9015
   End
   Begin VB.CheckBox chkMidi 
      Caption         =   "MIDI"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox chkText 
      Caption         =   "Lyrics (*.txt)"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblHelp 
      Caption         =   "% and _ are wildcards.  % represents any characters, _ represents any single character."
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label lblSearch 
      Caption         =   "Search options:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   9015
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    txtSearch = UCase(txtSearch)
    
    lblStatus = "Searching for " & txtSearch & "..."
    Screen.MousePointer = vbHourglass
    lstFound.Clear
    
    If optMode(0).Value = True Then
        SQL = "SELECT * FROM tblSongs WHERE Name LIKE '%" & txtSearch & "%' ORDER BY Name"
    End If
    If optMode(1).Value = True Then
        SQL = "SELECT * FROM tblSongs WHERE Name LIKE '" & txtSearch & "%' ORDER BY Name"
    End If
    If optMode(2).Value = True Then
        SQL = "SELECT * FROM tblSongs WHERE Name LIKE '%" & txtSearch & "' ORDER BY Name"
    End If
    If optMode(3).Value = True Then
        SQL = "SELECT * FROM tblSongs WHERE Name LIKE '" & txtSearch & "' ORDER BY Name"
    End If
    adoConn.Open DSN
        adoRS.Open SQL, adoConn
            Do Until adoRS.EOF
                If (chkMidi.Value = vbChecked And adoRS("isMIDI")) Or (chkText.Value = vbChecked And adoRS("isLyrics")) Then
                    lstFound.AddItem adoRS("name") & " (" & adoRS("Filename") & ")"
                End If
                adoRS.MoveNext
            Loop
        adoRS.Close
    adoConn.Close
    Screen.MousePointer = vbDefault
    lblStatus = "Search complete."
    txtSearch.SelStart = 0
    txtSearch.SelLength = Len(txtSearch)
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    If txtSearch.Text <> "" Then
        cmdSearch.Enabled = True
    Else
        cmdSearch.Enabled = False
    End If
End Sub
