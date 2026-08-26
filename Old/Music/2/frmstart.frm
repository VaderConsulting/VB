VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6660
   ClientLeft      =   1770
   ClientTop       =   1965
   ClientWidth     =   9600
   Icon            =   "frmstart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6660
   ScaleWidth      =   9600
   Begin VB.CheckBox ssFormat 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9240
      TabIndex        =   71
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox ssClearMatch 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3960
      TabIndex        =   67
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox chkText 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9240
      TabIndex        =   66
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdTextView 
      Caption         =   "View Song"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7650
      TabIndex        =   63
      Top             =   5220
      Width           =   1185
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Height          =   285
      Left            =   4860
      TabIndex        =   62
      Top             =   450
      Width           =   555
   End
   Begin VB.TextBox txtPicture 
      Height          =   285
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   450
      Width           =   3525
   End
   Begin VB.CommandButton cmdText 
      Caption         =   "Text Options..."
      Height          =   315
      Left            =   7020
      TabIndex        =   59
      Top             =   5940
      Width           =   1185
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   330
      Left            =   4860
      TabIndex        =   58
      Top             =   90
      Width           =   555
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   90
      Width           =   3525
   End
   Begin VB.CommandButton cmdAV 
      Caption         =   "Rec Options..."
      Height          =   315
      Left            =   8370
      TabIndex        =   55
      Top             =   5940
      Width           =   1185
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Song"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7020
      TabIndex        =   54
      Top             =   4860
      Width           =   1185
   End
   Begin VB.Timer timDelay 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5490
      Top             =   450
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2880
      TabIndex        =   53
      Top             =   5760
      Width           =   915
   End
   Begin VB.CommandButton cmdAny 
      Caption         =   "Any"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4230
      TabIndex        =   52
      Top             =   5760
      Width           =   915
   End
   Begin VB.ListBox lstMatch 
      Height          =   1035
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   50
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Boxes"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8370
      TabIndex        =   33
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Details"
      Enabled         =   0   'False
      Height          =   315
      Left            =   7020
      TabIndex        =   34
      Top             =   5580
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Song"
      Enabled         =   0   'False
      Height          =   315
      Left            =   8370
      TabIndex        =   32
      Top             =   4860
      Width           =   1185
   End
   Begin VB.Frame Frame5 
      Caption         =   "Navigation Mode"
      Height          =   1395
      Left            =   5400
      TabIndex        =   47
      Top             =   4860
      Width           =   1545
      Begin VB.OptionButton optEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   810
         WhatsThisHelpID =   2
         Width           =   915
      End
      Begin VB.OptionButton optAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   29
         Top             =   270
         WhatsThisHelpID =   1
         Width           =   615
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   540
         WhatsThisHelpID =   3
         Width           =   830
      End
      Begin VB.OptionButton optBrowse 
         Caption         =   "Browse"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   30
         Top             =   1080
         Value           =   -1  'True
         WhatsThisHelpID =   2
         Width           =   830
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sex"
      Height          =   465
      Left            =   7560
      TabIndex        =   46
      Top             =   3420
      Width           =   1995
      Begin VB.CheckBox chkSex 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   70
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   180
         Width           =   885
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Era"
      Height          =   1155
      Left            =   7560
      TabIndex        =   45
      Top             =   2190
      Width           =   2025
      Begin VB.CheckBox chkEra 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   65
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton opt00 
         Caption         =   "00's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   810
         TabIndex        =   16
         Top             =   810
         Width           =   615
      End
      Begin VB.OptionButton opt90 
         Caption         =   "90's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   810
         TabIndex        =   15
         Top             =   540
         Width           =   615
      End
      Begin VB.OptionButton opt80 
         Caption         =   "80's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   810
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton opt70 
         Caption         =   "70's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   810
         Width           =   615
      End
      Begin VB.OptionButton opt60 
         Caption         =   "60's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   615
      End
      Begin VB.OptionButton opt50 
         Caption         =   "50's"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "Digital"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   8460
      TabIndex        =   28
      Top             =   4320
      Width           =   735
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "Picture"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   8460
      TabIndex        =   27
      Top             =   4050
      Width           =   915
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "Sheet"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   7650
      TabIndex        =   25
      Top             =   4590
      Width           =   735
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "CD"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   8460
      TabIndex        =   26
      Top             =   4590
      Width           =   615
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "Vinyl"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   7650
      TabIndex        =   24
      Top             =   4320
      Width           =   735
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "Tape"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   7650
      TabIndex        =   23
      Top             =   4050
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nationality"
      Height          =   1365
      Left            =   5400
      TabIndex        =   44
      Top             =   3420
      Width           =   2085
      Begin VB.CheckBox chkNationality 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   69
         Top             =   1080
         Width           =   255
      End
      Begin VB.OptionButton optOther 
         Caption         =   "Other"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optAmerican 
         Caption         =   "American"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   810
         Width           =   1095
      End
      Begin VB.OptionButton optEnglish 
         Caption         =   "English"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton optAustralian 
         Caption         =   "Australian"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type"
      Height          =   1155
      Left            =   5400
      TabIndex        =   43
      Top             =   2190
      Width           =   2085
      Begin VB.CheckBox chkType 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1680
         TabIndex        =   68
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton optDance 
         Caption         =   "Dance Music"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton optListen 
         Caption         =   "Listening Music"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   810
         Width           =   1455
      End
      Begin VB.OptionButton optRock 
         Caption         =   "Rock n Roll"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAlbumName 
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   450
      Width           =   2535
   End
   Begin VB.TextBox txtSong 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   810
      Width           =   2535
   End
   Begin VB.TextBox txtMagazinePage 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8550
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox txtMagazineDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtMagazine 
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ListBox lstSongs 
      Height          =   2985
      Left            =   2790
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtBand 
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   2535
   End
   Begin VB.ListBox lstBands 
      Height          =   5130
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin MCI.MMControl mmWave 
      Height          =   330
      Left            =   5400
      TabIndex        =   64
      Top             =   1710
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      _Version        =   327681
      PlayEnabled     =   -1  'True
      RecordEnabled   =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label lblPicture 
      Caption         =   "Picture Filename"
      Height          =   285
      Left            =   90
      TabIndex        =   61
      Top             =   450
      Width           =   1185
   End
   Begin VB.Label lblFilename 
      Alignment       =   1  'Right Justify
      Caption         =   "Audio Filename"
      Height          =   285
      Left            =   0
      TabIndex        =   56
      Top             =   90
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      Height          =   1995
      Left            =   2790
      Top             =   4230
      Width           =   2535
   End
   Begin VB.Label lblSearch 
      Alignment       =   2  'Center
      Caption         =   "Search Results"
      Height          =   255
      Left            =   2880
      TabIndex        =   51
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label lblHelp 
      Height          =   255
      Left            =   90
      TabIndex        =   49
      Top             =   6360
      Width           =   9465
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Album Name"
      Height          =   255
      Left            =   5940
      TabIndex        =   42
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Page"
      Height          =   255
      Left            =   8010
      TabIndex        =   41
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Date "
      Height          =   255
      Left            =   6300
      TabIndex        =   40
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Magazine Name"
      Height          =   255
      Left            =   5640
      TabIndex        =   39
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Song Title"
      Height          =   255
      Left            =   6120
      TabIndex        =   38
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Artist / Band Name"
      Height          =   255
      Left            =   5580
      TabIndex        =   37
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Songs"
      Height          =   285
      Left            =   2760
      TabIndex        =   36
      Top             =   810
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Artists / Bands"
      Height          =   285
      Left            =   90
      TabIndex        =   35
      Top             =   810
      Width           =   2535
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Database"
      End
      Begin VB.Menu mnuHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuModifyUser 
         Caption         =   "Modify"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuHyphen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim datmarty As Database
    Dim tabSongs As Recordset
    Dim tabBands As Recordset
Private Sub cmdNew_Click()
    txtBand = ""
    txtSong = ""
    txtMagazine = ""
    txtMagazineDate = ""
    txtMagazinePage = ""
    txtBand.SetFocus
    txtBand.Locked = 0
    txtSong.Locked = 0
    txtMagazine.Locked = 0
    txtMagazineDate.Locked = 0
    txtMagazinePage.Locked = 0
    txtFilename.Locked = 0
End Sub


Private Sub chkEra_Click() ' value as integer
    chkEra = 1
    opt50 = 0
    opt60 = 0
    opt70 = 0
    opt80 = 0
    opt90 = 0
    opt00 = 0

End Sub

Private Sub chkEra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the era."
End Sub

Private Sub chkFormat_Click(Index As Integer)
    If Index = 5 And chkFormat(5) = 1 Then
        If optAdd = True Or optEdit = True Then
            Rem cmdAV_Click
        End If
    End If
End Sub

Private Sub chkFormat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then lblHelp.Caption = "A tick in this box indicates that you have this song on cassette."
    If Index = 1 Then lblHelp.Caption = "A tick in this box indicates that you have this song on record."
    If Index = 2 Then lblHelp.Caption = "A tick in this box indicates that you have this song on paper."
    If Index = 3 Then lblHelp.Caption = "A tick in this box indicates that you have this song on audio CD."
    If Index = 4 Then lblHelp.Caption = "A tick in this box indicates that you have this song on video."
    If Index = 5 Then lblHelp.Caption = "A tick in this box indicates that you have this song on computer."
End Sub

Private Sub chkNationality_Click() ' value as integer
    chkNationality = 1
    optAustralian = 0
    optEnglish = 0
    optAmerican = 0
    optOther = 0
End Sub

Private Sub chkNationality_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the nationality."
End Sub

Private Sub chkSex_Click() ' value as integer
    chkSex.Value = 1
    optMale = 0
    optFemale = 0
End Sub

Private Sub chkSex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the sex."
End Sub

Private Sub chkText_Click() ' value as integer
    chkText = 1
    txtBand = ""
    txtSong = ""
    txtMagazine = ""
    txtMagazineDate = ""
    txtMagazinePage = ""
    txtMagazinePage = ""
    txtAlbumName = ""
End Sub

Private Sub chkText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the text boxes."
End Sub

Private Sub chkType_Click() 'Value As Integer
    chkType = 1
    optRock = 0
    optListen = 0
    optDance = 0
End Sub

Private Sub chkType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the type."
End Sub

Private Sub cmdAll_Click()
    'this routine is used to ensure that ALL of the selected options match 1 or more songs
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    If tabSongs.RecordCount = 0 Then Exit Sub
    tabSongs.MoveFirst
    varMatch = 0
    lstMatch.Clear
    Do Until tabSongs.EOF
        If UCase(txtSong) <> "" And UCase(tabSongs(1)) = UCase(txtSong) Then varMatch = 1
        If UCase(txtSong) <> "" And UCase(tabSongs(1)) <> UCase(txtSong) Then varMatch = 0
        If UCase(txtBand) <> "" And UCase(tabSongs(2)) = UCase(txtBand) Then varMatch = 1
        If UCase(txtBand) <> "" And UCase(tabSongs(2)) <> UCase(txtBand) Then varMatch = 0
        If UCase(txtMagazine) <> "" And UCase(tabSongs(3)) = UCase(txtMagazine) Then varMatch = 1
        If UCase(txtMagazine) <> "" And UCase(tabSongs(3)) <> UCase(txtMagazine) Then varMatch = 0
        If UCase(txtMagazineDate) <> "" And UCase(tabSongs(4)) = UCase(txtMagazineDate) Then varMatch = 1
        If UCase(txtMagazineDate) <> "" And UCase(tabSongs(4)) <> UCase(txtMagazineDate) Then varMatch = 0
        If UCase(txtMagazinePage) <> "" And UCase(tabSongs(5)) = UCase(txtMagazinePage) Then varMatch = 1
        If UCase(txtMagazinePage) <> "" And UCase(tabSongs(5)) <> UCase(txtMagazinePage) Then varMatch = 0
        If UCase(txtAlbumName) <> "" And UCase(tabSongs(6)) = UCase(txtAlbumName) Then varMatch = 1
        If UCase(txtAlbumName) <> "" And UCase(tabSongs(6)) <> UCase(txtAlbumName) Then varMatch = 0
        If optRock = True And tabSongs(7) = 1 Then varMatch = 1
        If optRock = True And tabSongs(7) <> 1 Then varMatch = 0
        If optListen = True And tabSongs(7) = 2 Then varMatch = 1
        If optListen = True And tabSongs(7) <> 2 Then varMatch = 0
        If optDance = True And tabSongs(7) = 3 Then varMatch = 1
        If optDance = True And tabSongs(7) <> 3 Then varMatch = 0
        If optMale = True And tabSongs(8) = "M" Then varMatch = 1
        If optMale = True And tabSongs(8) <> "M" Then varMatch = 0
        If optFemale = True And tabSongs(8) = "F" Then varMatch = 1
        If optFemale = True And tabSongs(8) <> "F" Then varMatch = 0
        If optAustralian = True And tabSongs(9) = 1 Then varMatch = 1
        If optAustralian = True And tabSongs(9) <> 1 Then varMatch = 0
        If optEnglish = True And tabSongs(9) = 2 Then varMatch = 1
        If optEnglish = True And tabSongs(9) <> 2 Then varMatch = 0
        If optAmerican = True And tabSongs(9) = 3 Then varMatch = 1
        If optAmerican = True And tabSongs(9) <> 3 Then varMatch = 0
        If optOther = True And tabSongs(9) = 4 Then varMatch = 1
        If optOther = True And tabSongs(9) <> 4 Then varMatch = 0
        If opt50 = True And tabSongs(10) = 5 Then varMatch = 1
        If opt50 = True And tabSongs(10) <> 5 Then varMatch = 0
        If opt60 = True And tabSongs(10) = 6 Then varMatch = 1
        If opt60 = True And tabSongs(10) <> 6 Then varMatch = 0
        If opt70 = True And tabSongs(10) = 7 Then varMatch = 1
        If opt70 = True And tabSongs(10) <> 7 Then varMatch = 0
        If opt80 = True And tabSongs(10) = 8 Then varMatch = 1
        If opt80 = True And tabSongs(10) <> 8 Then varMatch = 0
        If opt90 = True And tabSongs(10) = 9 Then varMatch = 1
        If opt90 = True And tabSongs(10) <> 9 Then varMatch = 0
        If opt00 = True And tabSongs(10) = 0 Then varMatch = 1
        If opt00 = True And tabSongs(10) <> 0 Then varMatch = 0
        For varLoop = 0 To 5
            If chkFormat(varLoop) = 1 And tabSongs(11 + varLoop) = -1 Then varMatch = 1
            If chkFormat(varLoop) = 1 And tabSongs(11 + varLoop) <> -1 Then varMatch = 0
        Next varLoop
        If varMatch = 1 Then lstMatch.AddItem tabSongs(1)
        varMatch = 0
        tabSongs.MoveNext
        lblSearch.Caption = "Search Results - " & lstMatch.ListCount & " found"
        lblSearch.Refresh
    Loop
End Sub

Private Sub cmdAV_Click()
    If optEdit = True Or optAdd = True Then
        cmdSave_Click
    End If
    If txtSong <> "" Then
        frmStart.Hide
        frmAV.Show
    Else
        msg = "Hang on... select a song first"
        MsgBox msg, vbOKOnly
    End If
End Sub

Private Sub cmdAV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click this button to advance to the Audio and Video(Still) options screen."
End Sub

Private Sub cmdClear_Click()
    txtSong = ""
    subClearAll
End Sub
Private Sub subClearAll()
    txtBand = ""
    '
    txtMagazine = ""
    txtMagazineDate = ""
    txtMagazinePage = ""
    txtMagazinePage = ""
    txtAlbumName = ""
    txtFilename = ""
    txtPicture = ""
    txtText = ""
    optRock = 0
    optListen = 0
    optDance = 0
    optMale = 0
    optFemale = 0
    optAustralian = 0
    optEnglish = 0
    optAmerican = 0
    optOther = 0
    opt50 = 0
    opt60 = 0
    opt70 = 0
    opt80 = 0
    opt90 = 0
    opt00 = 0
    For varLoop = 0 To 5
        chkFormat(varLoop).Value = 0
    Next varLoop
End Sub
Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click this button to clear the boxes."
End Sub

Private Sub cmdDelete_Click()
    subDelete
End Sub
Private Sub subDelete()
    If tabSongs.RecordCount = 0 Then Exit Sub
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    Dim msg As String
    Dim varStyle As Integer
    Dim response As Integer
    tabSongs.MoveFirst
    Do Until tabSongs.EOF
        If UCase(tabSongs(1)) = UCase(txtSong) And UCase(txtSong) <> "" Then
            msg = "Delete this song"
            varStyle = vbYesNo + vbQuestion + vbDefaultButton2
            response = MsgBox(msg, varStyle, "Delete " & txtSong)
            If response = IDYES Then
                tabSongs.Delete
                Rem now delete the .wav file
                Rem kill "songname.wav"
                subUpdateSongs
                subBandListUpdate
                subClearAll
                txtSong = ""
            End If
            Exit Sub
        End If
        DoEvents
        tabSongs.MoveNext
    Loop
End Sub
Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click on this button to delete the currently selected song."
End Sub


Private Sub cmdAny_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    If tabSongs.RecordCount = 0 Then Exit Sub
    tabSongs.MoveFirst
    varMatch = 0
    lstMatch.Clear
    Do Until tabSongs.EOF
        If UCase(tabSongs(1)) = UCase(txtSong) And UCase(txtSong) <> "" Then varMatch = 1
        If UCase(tabSongs(2)) = UCase(txtBand) And UCase(txtBand) <> "" Then varMatch = 1
        If UCase(tabSongs(3)) = UCase(txtMagazine) And UCase(txtMagazine) <> "" Then varMatch = 1
        If UCase(tabSongs(4)) = UCase(txtMagazineDate) And UCase(txtMagazineDate) <> "" Then varMatch = 1
        If UCase(tabSongs(5)) = UCase(txtMagazinePage) And UCase(txtMagazinePage) <> "" Then varMatch = 1
        If UCase(tabSongs(6)) = UCase(txtAlbumName) And UCase(txtAlbumName) <> "" Then varMatch = 1
        If optRock = True And tabSongs(7) = 1 Then varMatch = 1
        If optListen = True And tabSongs(7) = 2 Then varMatch = 1
        If optDance = True And tabSongs(7) = 3 Then varMatch = 1
        If optMale = True And tabSongs(8) = "M" Then varMatch = 1
        If optFemale = True And tabSongs(8) = "F" Then varMatch = 1
        If optAustralian = True And tabSongs(9) = 1 Then varMatch = 1
        If optEnglish = True And tabSongs(9) = 2 Then varMatch = 1
        If optAmerican = True And tabSongs(9) = 3 Then varMatch = 1
        If optOther = True And tabSongs(9) = 4 Then varMatch = 1
        If opt50 = True And tabSongs(10) = 5 Then varMatch = 1
        If opt60 = True And tabSongs(10) = 6 Then varMatch = 1
        If opt70 = True And tabSongs(10) = 7 Then varMatch = 1
        If opt80 = True And tabSongs(10) = 8 Then varMatch = 1
        If opt90 = True And tabSongs(10) = 9 Then varMatch = 1
        If opt00 = True And tabSongs(10) = 0 Then varMatch = 1
        For varLoop = 0 To 5
            If chkFormat(varLoop) = 1 And tabSongs(11 + varLoop) = -1 Then varMatch = 1
        Next varLoop
        If varMatch = 1 Then lstMatch.AddItem tabSongs(1)
        varMatch = 0
        tabSongs.MoveNext
        lblSearch.Caption = "Search Results - " & lstMatch.ListCount & " found"
        lblSearch.Refresh
    Loop
End Sub

Private Sub cmdAny_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Search for any song with any of the current parameters."
End Sub

Private Sub cmdFind_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    If cmdFind.Caption = "Find Song" Then
        subClearAll
        shpPlay.Visible = 1
        imgPlay.Visible = 1
        txtSong = ""
        txtSong.Locked = 0
        txtSong.SetFocus
        cmdFind.Caption = "Start Search"
        Exit Sub
    End If
    If tabSongs.RecordCount = 0 Then
        lblHelp.Caption = "No Records to search"
        txtSong = ""
        txtSong.Locked = 1
        cmdFind.Caption = "Find Song"
        Exit Sub
    End If
    If txtSong = "" Then
        cmdFind.Caption = "Find Song"
        Exit Sub
    End If
    subLook
    If txtBand = "" Or (optMale = 0 And optFemale = 0) Then txtSong = ""
    txtSong.Locked = 1
    cmdFind.Caption = "Find Song"
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Search for a song by name."
End Sub

Private Sub cmdPlay_Click()
    mmWave.Command = "Back"
    mmWave.Command = "Close"
    If txtFilename <> "" Then
        soundname$ = txtFilename
        mmWave.filename = soundname$
        mmWave.Command = "Open"
        mmWave.Command = "Play"
    End If
End Sub

Private Sub cmdPreview_Click()
    soundname$ = lstPlay.List(lstPlay.ListIndex)
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Rem X% = sndPlaySound(soundname$, wFlags%)
End Sub

Private Sub cmdPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to play the sample of the currently selected song."
End Sub

Private Sub cmdSave_Click()
    If txtSong = "" Then
        Beep
        lblHelp.Caption = "You must supply a song name"
        txtSong.BackColor = &HFFFF&
        timDelay.Enabled = 1
        varerror = 1
    End If
    If txtBand = "" Then
        Beep
        lblHelp.Caption = "You must supply a band name"
        txtBand.BackColor = &HFFFF&
        timDelay.Enabled = 1
        varerror = 1
    End If
    If varerror = 1 Then Exit Sub
    subAdd
End Sub
Private Sub subAdd()
    If optAdd = True Then
        tabSongs.AddNew
    End If
    If optEdit = True Then
        tabSongs.MoveFirst
        Do Until tabSongs.EOF
            If UCase$(tabSongs(1)) = UCase$(txtSong) Then
                'Found it!
                Exit Do
            End If
            DoEvents
            tabSongs.MoveNext
        Loop
        'the next 3 lines are for if there is no match, which should not happen..
        If UCase$(tabSongs(1)) <> UCase$(txtSong) Then
            Beep
            frmHelp.lblHelp.Caption = "Song not found!"
            frmHelp.Show 1
            Exit Sub
        End If
    End If
    If optEdit = True Then tabSongs.Edit
    tabSongs(1) = txtSong
    tabSongs(2) = txtBand
    tabSongs(3) = txtMagazine
    If txtMagazineDate <> "" Then tabSongs(4) = txtMagazineDate
    tabSongs(5) = txtMagazinePage
    tabSongs(6) = txtAlbumName
    If optRock = True Then tabSongs(7) = 1
    If optListen = True Then tabSongs(7) = 2
    If optDance = True Then tabSongs(7) = 3
    If optMale = True Then tabSongs(8) = "M"
    If optFemale = True Then tabSongs(8) = "F"
    If optAustralian = True Then tabSongs(9) = 1
    If optEnglish = True Then tabSongs(9) = 2
    If optAmerican = True Then tabSongs(9) = 3
    If optOther = True Then tabSongs(9) = 4
    If opt50 = True Then tabSongs(10) = 5
    If opt60 = True Then tabSongs(10) = 6
    If opt70 = True Then tabSongs(10) = 7
    If opt80 = True Then tabSongs(10) = 8
    If opt90 = True Then tabSongs(10) = 9
    If opt00 = True Then tabSongs(10) = 0
    For varLoop = 0 To 5
        If chkFormat(varLoop).Value = 1 Then tabSongs(11 + varLoop) = -1
    Next varLoop
    For varLoop = 0 To 5
        If chkFormat(varLoop).Value <> 1 Then tabSongs(11 + varLoop) = 0
    Next varLoop
    If chkFormat(5).Value <> 1 Then txtFilename = ""   ' clear filename if digital is not checked
    tabSongs(17) = txtFilename
    tabSongs(18) = txtPicture
    tabSongs.Update
    subUpdateSongs
    subBandListUpdate
    ' subClearAll
    For varLoop = 0 To lstBands.ListCount - 1
        If lstBands.List(varLoop) = txtBand Then lstBands.ListIndex = varLoop
    Next varLoop
    For varLoop = 0 To lstSongs.ListCount - 1
        If lstBands.List(varLoop) = txtSong Then lstSongs.ListIndex = varLoop
    Next varLoop
    'txtSong = ""
    txtBand.SetFocus
End Sub
Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Save the current song details."
End Sub
Private Sub cmdSelect_Click()
    If lstPlay.ListIndex <> -1 Then
        tabSongs.MoveFirst
        Do Until tabSongs.EOF
            If tabSongs(1) = txtSong Then
                'Found it!
                Exit Do
            End If
            DoEvents
            tabSongs.MoveNext
        Loop
        'the next 3 lines are for if there is no match, which should not happen..
        If tabSongs(1) <> txtSong Then
            Beep
            frmHelp.lblHelp.Caption = "Song not found!"
            frmHelp.Show 1
            Exit Sub
        End If
        tabSongs.Edit
        tabSongs(17) = lstPlay.List(lstPlay.ListIndex)
        tabSongs.Update
        cmdPlay.Enabled = 1
    End If
    
    
End Sub


Private Sub cmdText_Click()
    If txtSong <> "" Then
        frmStart.Hide
        frmText.Show
    Else
        msg = "Hang on... select a song first"
        MsgBox msg, vbOKOnly
    End If
End Sub


Private Sub cmdTextView_Click()
    frmView.Show
    frmStart.Hide
End Sub

Private Sub cmdView_Click()
    If txtPicture <> "" Then frmPicture.Show
End Sub

Private Sub cmdView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to view the picture of the currently selected Artist."
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the type of song."
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the Nationality of the Artist / Band."
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the era that the song is from."
End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the sex of the Artist / lead singer of the Band."
End Sub


Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Select your Navigation Mode."
End Sub

Private Sub imgPlay_Click()
    soundname$ = txtFilename
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Rem X% = sndPlaySound(soundname$, wFlags%)
End Sub

Private Sub imgPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Clicking this button will play the selected song sample."
End Sub

Public Sub lstBands_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    txtBand = lstBands.List(lstBands.ListIndex)
    txtSong = ""
    txtMagazine = ""
    txtMagazineDate = ""
    txtMagazinePage = ""
    txtMagazinePage = ""
    txtAlbumName = ""
    txtFilename = ""
    txtPicture = ""
    txtText = ""
    optRock = 0
    optListen = 0
    optDance = 0
    optMale = 0
    optFemale = 0
    optAustralian = 0
    optEnglish = 0
    optAmerican = 0
    optOther = 0
    opt50 = 0
    opt60 = 0
    opt70 = 0
    opt80 = 0
    opt90 = 0
    opt00 = 0
    For varLoop = 0 To 5
        chkFormat(varLoop).Value = 0
    Next varLoop
    subUpdateSongs
End Sub
Public Sub subUpdateSongs()
    tabSongs.MoveFirst
    lstSongs.Clear
    Do Until tabSongs.EOF
        If tabSongs(2) = txtBand Then
            lstSongs.AddItem tabSongs(1)
        End If
        tabSongs.MoveNext
    Loop
    lstSongs.Refresh
End Sub
Private Sub lstBands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the list of Artists / Bands to choose from or add to."
End Sub

Private Sub lstMatch_Click()
    chkText = 1
    txtBand = ""
    txtSong = ""
    txtMagazine = ""
    txtMagazineDate = ""
    txtMagazinePage = ""
    txtMagazinePage = ""
    txtAlbumName = ""
    txtSong = lstMatch.List(lstMatch.ListIndex)
    txtFilename = ""
    subLook
    cmdPlay_Click
End Sub

Private Sub lstMatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is where the result of your last search is displayed."
End Sub

Private Sub lstSongs_Click()
    If optBrowse.Value = -1 Or optEdit.Value = -1 Then
        txtSong = lstSongs.List(lstSongs.ListIndex)
        subLook
    End If
    If gTimeCripple = 1 And gRegistered = 0 Then
        'Give the non registered user 5 minutes
        If Timer > gTime + 300 Then
            frmTimesUp.Show
            frmStart.Hide
        End If
    End If
    If chkFormat(2) = 1 Then
        cmdTextView.Enabled = 1
    Else
        cmdTextView.Enabled = 0
    End If
End Sub

Private Sub lstSongs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "This is the list of Songs for the chosen Artist / Band."
End Sub

Private Sub mmPlay_Done(NotifyCode As Integer)
    mmPlay.Command = "Close"
End Sub

Private Sub mmPlay_PlayClick(Cancel As Integer)
    mmPlay.Enabled = 1
    mmPlay.filename = txtFilename
    mmPlay.Command = "Open"
    If Dir(txtFilename) = "" Then
        msg = "Sorry, no song to play"
        MsgBox msg, vbOKOnly
    End If
    mmPlay.PlayEnabled = 1
End Sub

Private Sub mmWave_PlayCompleted(Errorcode As Long)
    mmWave.Command = "Back"
    mmWave.Command = "Close"
End Sub

Private Sub mnuAbout_Click()
    subAbout
End Sub
Private Sub mnuDeleteUser_Click()
    subDelete
End Sub
Private Sub mnuExit_Click()
    subEnd
End Sub

Private Sub mnuLoad_Click()
    Rem subBandListUpdate
    optAdd.Enabled = 1
    optEdit.Enabled = 1
    optBrowse.Enabled = 1
    optSearch.Enabled = 1
    cmdDelete.Enabled = 1
    cmdFind.Enabled = 1
    lblHelp.Caption = ""
End Sub

Private Sub mnuLogout_Click()
    frmLogin.Show
    frmStart.Hide
End Sub
Private Sub mnuSetup_Click()
    frmSetup.Show
    frmStart.Hide
End Sub
Private Sub mnuShow_Click()
    frmStart.Hide
    frmMe.Show
End Sub
Public Sub Form_Load()
    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    gWindowsDir = GetWindowsDir
    varDir = Dir(gWindowsDir & gINIFile & ".INI")
    Caption = gProgramName
    gDatabasePath = GetSetting(gINIFile, "Setup", "Database", "")
    If gDatabasePath = "" Then
        frmDatabase.Show 1
    End If
    subBandListUpdate
    optAdd.Enabled = 1
    optEdit.Enabled = 1
    optBrowse.Enabled = 1
    optSearch.Enabled = 1
    Rem cmdDelete.Enabled = 1
    cmdFind.Enabled = 1
    lblHelp.Caption = ""
    mmWave.Notify = False
    mmWave.Wait = True
    mmWave.Shareable = False
    mmWave.DeviceType = "WaveAudio"
End Sub

Private Sub optAdd_Click()
    If optAdd.Value = -1 Then
        cmdSave.Enabled = 1
        lstSongs.Enabled = 0
    End If
    If optAdd.Value = 0 Then
        cmdSave.Enabled = 0
        lstSongs.Enabled = 1
    End If
    subClearAll
    txtSong = ""
    cmdClear.Enabled = 1
    txtBand.Locked = 0
    txtSong.Locked = 0
    txtAlbumName.Locked = 0
    txtMagazine.Locked = 0
    txtMagazineDate.Locked = 0
    txtMagazinePage.Locked = 0
    txtFilename.Locked = 0
    cmdAny.Enabled = 0
    cmdAll.Enabled = 0
    cmdClear.Enabled = 1
    optDance.Enabled = 1
    optRock.Enabled = 1
    optListen.Enabled = 1
    opt50.Enabled = 1
    opt60.Enabled = 1
    opt70.Enabled = 1
    opt80.Enabled = 1
    opt90.Enabled = 1
    opt00.Enabled = 1
    optAustralian.Enabled = 1
    optEnglish.Enabled = 1
    optAmerican.Enabled = 1
    optOther.Enabled = 1
    optMale.Enabled = 1
    optFemale.Enabled = 1
    chkSex.Enabled = 1
    chkType.Enabled = 1
    chkNationality.Enabled = 1
    chkEra.Enabled = 1
    ssFormat.Enabled = 1
    txtBand.SetFocus
    For varLoop = 0 To 5
        chkFormat(varLoop).Enabled = 1
    Next varLoop
End Sub

Private Sub optBrowse_Click()
    If optAdd.Value = -1 Then
        cmdSave.Enabled = 1
        lstSongs.Enabled = 0
    End If
    If optAdd.Value = 0 Then
        cmdSave.Enabled = 0
        lstSongs.Enabled = 1
    End If
    txtBand.Locked = 1
    txtSong.Locked = 1
    txtAlbumName.Locked = 1
    txtMagazine.Locked = 1
    txtMagazineDate.Locked = 1
    txtMagazinePage.Locked = 1
    txtFilename.Locked = 1
    cmdClear.Enabled = 0
    cmdAny.Enabled = 0
    cmdAll.Enabled = 0
    optDance.Enabled = 0
    optRock.Enabled = 0
    optListen.Enabled = 0
    opt50.Enabled = 0
    opt60.Enabled = 0
    opt70.Enabled = 0
    opt80.Enabled = 0
    opt90.Enabled = 0
    opt00.Enabled = 0
    optAustralian.Enabled = 0
    optEnglish.Enabled = 0
    optAmerican.Enabled = 0
    optOther.Enabled = 0
    optMale.Enabled = 0
    optFemale.Enabled = 0
    chkSex.Enabled = 0
    chkType.Enabled = 0
    chkNationality.Enabled = 0
    chkEra.Enabled = 0
    ssFormat.Enabled = 0
    For varLoop = 0 To 5
        chkFormat(varLoop).Enabled = 0
    Next varLoop
End Sub

Private Sub optEdit_Click()
    lstSongs.Enabled = 1
    If optEdit.Value = -1 Then
        cmdSave.Enabled = 1
    End If
    If optEdit.Value = 0 Then
        cmdSave.Enabled = 0
    End If
    cmdClear.Enabled = 1
    txtBand.Locked = 0
    txtSong.Locked = 0
    txtAlbumName.Locked = 0
    txtMagazine.Locked = 0
    txtMagazineDate.Locked = 0
    txtMagazinePage.Locked = 0
    txtFilename.Locked = 0
    cmdAny.Enabled = 0
    cmdAll.Enabled = 0
    cmdClear.Enabled = 1
    optDance.Enabled = 1
    optRock.Enabled = 1
    optListen.Enabled = 1
    opt50.Enabled = 1
    opt60.Enabled = 1
    opt70.Enabled = 1
    opt80.Enabled = 1
    opt90.Enabled = 1
    opt00.Enabled = 1
    optAustralian.Enabled = 1
    optEnglish.Enabled = 1
    optAmerican.Enabled = 1
    optOther.Enabled = 1
    optMale.Enabled = 1
    optFemale.Enabled = 1
    chkSex.Enabled = 1
    chkType.Enabled = 1
    chkNationality.Enabled = 1
    chkEra.Enabled = 1
    ssFormat.Enabled = 1
    txtBand.SetFocus
    For varLoop = 0 To 5
        chkFormat(varLoop).Enabled = 1
    Next varLoop
End Sub

Private Sub optSearch_Click()
    If optSearch.Value = -1 Then
        cmdSave.Enabled = 1
        lstSongs.Enabled = 0
    End If
    If optAdd.Value = 0 Then
        cmdSave.Enabled = 0
    End If
    cmdClear.Enabled = 1
    txtBand.Locked = 0
    txtSong.Locked = 0
    txtAlbumName.Locked = 0
    txtMagazine.Locked = 0
    txtMagazineDate.Locked = 0
    txtMagazinePage.Locked = 0
    txtFilename.Locked = 0
    cmdAny.Enabled = 1
    cmdAll.Enabled = 0
    cmdClear.Enabled = 1
    optDance.Enabled = 1
    optRock.Enabled = 1
    optListen.Enabled = 1
    opt50.Enabled = 1
    opt60.Enabled = 1
    opt70.Enabled = 1
    opt80.Enabled = 1
    opt90.Enabled = 1
    opt00.Enabled = 1
    optAustralian.Enabled = 1
    optEnglish.Enabled = 1
    optAmerican.Enabled = 1
    optOther.Enabled = 1
    optMale.Enabled = 1
    optFemale.Enabled = 1
    chkSex.Enabled = 1
    chkType.Enabled = 1
    chkNationality.Enabled = 1
    chkEra.Enabled = 1
    ssFormat.Enabled = 1
    txtBand.SetFocus
    For varLoop = 0 To 5
        chkFormat(varLoop).Enabled = 1
    Next varLoop
    cmdClear_Click
End Sub
Public Sub subBandListUpdate()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    lstBands.Clear
    Do Until tabSongs.EOF
        For varLoop = 0 To lstBands.ListCount
            If lstBands.List(varLoop) = tabSongs(2) Then varPresent = 1
        Next varLoop
        If varPresent = 0 Then lstBands.AddItem tabSongs(2)
        varPresent = 0
        tabSongs.MoveNext
    Loop
    lstBands.Refresh
End Sub
Public Sub subLook()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    subClearAll
    tabSongs.MoveFirst
    Do Until tabSongs.EOF
        If UCase$(tabSongs(1)) = UCase$(txtSong) Then
            txtSong = tabSongs(1)
            txtBand = tabSongs(2)
            txtMagazine = tabSongs(3)
            If tabSongs(4) <> "" Then txtMagazineDate = tabSongs(4)
            If tabSongs(5) <> "" Then txtMagazinePage = tabSongs(5)
            If tabSongs(6) <> "" Then txtAlbumName = tabSongs(6)
            If tabSongs(7) = 1 Then optRock = 1
            If tabSongs(7) = 2 Then optListen = 1
            If tabSongs(7) = 3 Then optDance = 1
            If tabSongs(8) = "M" Then optMale = 1
            If tabSongs(8) = "F" Then optFemale = 1
            If tabSongs(9) = 1 Then optAustralian = 1
            If tabSongs(9) = 2 Then optEnglish = 1
            If tabSongs(9) = 3 Then optAmerican = 1
            If tabSongs(9) = 4 Then optOther = 1
            If tabSongs(10) = 5 Then opt50 = 1
            If tabSongs(10) = 6 Then opt60 = 1
            If tabSongs(10) = 7 Then opt70 = 1
            If tabSongs(10) = 8 Then opt80 = 1
            If tabSongs(10) = 9 Then opt90 = 1
            If tabSongs(10) = 0 Then opt00 = 1
            For varLoop = 0 To 5
                chkFormat(varLoop).Value = 0
            Next varLoop
            For varLoop = 0 To 5
                If tabSongs(11 + varLoop) = -1 Then chkFormat(varLoop).Value = 1
            Next varLoop
            If tabSongs(17) <> "" Then
                txtFilename = tabSongs(17)
            End If
            If tabSongs(18) <> "" Then
                txtPicture = tabSongs(18)
            End If
            If tabSongs(19) <> "" Then
                txtText = tabSongs(19)
            End If
        End If
        tabSongs.MoveNext
    Loop
End Sub

Private Sub ssClearMatch_Click() 'Value As Integer
    ssClearMatch = 1
    lstMatch.Clear
    lblSearch.Caption = "Search Results"
End Sub

Private Sub ssClearMatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the search results."
End Sub


Private Sub ssFormat_Click() 'Value As Integer
    ssFormat = 1
    For varLoop = 0 To 5
        chkFormat(varLoop).Value = 0
    Next varLoop
End Sub

Private Sub ssFormat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHelp.Caption = "Click here to clear the storage types."
End Sub

Private Sub timDelay_Timer()
    txtSong.BackColor = &HFFFFFF
    txtBand.BackColor = &HFFFFFF
    txtMagazineDate.BackColor = &HFFFFFF
    timDelay.Enabled = 0
End Sub

Private Sub txtBand_Change()
    If txtBand = "" Or Len(txtSong) = 1 Then cmdDelete.Enabled = 0
    If txtBand <> "" Or Len(txtSong) > 1 Then cmdDelete.Enabled = 1
End Sub

Private Sub txtMagazineDate_Change()
    If InStr(txtMagazineDate, ".") <> 0 Then
        Beep
        lblHelp.Caption = "Please enter either a / or a - as the separating character"
        txtMagazineDate.BackColor = &HFFFF&
        timDelay.Enabled = 1
    End If
End Sub

Private Sub txtSong_Change()
    If txtSong = "" Or Len(txtSong) = 1 Then cmdDelete.Enabled = 0
    If txtSong <> "" Or Len(txtSong) > 1 Then cmdDelete.Enabled = 1
End Sub
