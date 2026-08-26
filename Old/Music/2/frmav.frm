VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form frmAV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio and Video"
   ClientHeight    =   6795
   ClientLeft      =   2580
   ClientTop       =   1860
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6795
   ScaleWidth      =   5850
   Begin VB.TextBox txtReturnValue 
      Height          =   285
      Left            =   1350
      TabIndex        =   55
      Top             =   5940
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton cmdSync 
      Caption         =   "Syncronised Recording"
      Height          =   645
      Left            =   4410
      TabIndex        =   54
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   29
      Left            =   3960
      TabIndex        =   51
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   28
      Left            =   3690
      TabIndex        =   50
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   27
      Left            =   3420
      TabIndex        =   49
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   26
      Left            =   3150
      TabIndex        =   48
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   2880
      TabIndex        =   47
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   2610
      TabIndex        =   46
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   2340
      TabIndex        =   45
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   2070
      TabIndex        =   44
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   1800
      TabIndex        =   43
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   1530
      TabIndex        =   42
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   1260
      TabIndex        =   41
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   990
      TabIndex        =   40
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   720
      TabIndex        =   39
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   450
      TabIndex        =   38
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   180
      TabIndex        =   37
      Top             =   3240
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   3960
      TabIndex        =   36
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   3690
      TabIndex        =   35
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   3420
      TabIndex        =   34
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3150
      TabIndex        =   33
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   2880
      TabIndex        =   32
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   2610
      TabIndex        =   31
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   2340
      TabIndex        =   30
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2070
      TabIndex        =   29
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   28
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1530
      TabIndex        =   27
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   26
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   990
      TabIndex        =   25
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   24
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   450
      TabIndex        =   23
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdTrack 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   2970
      Width           =   285
   End
   Begin VB.CommandButton cmdCD 
      Caption         =   "Enable CD"
      Height          =   285
      Left            =   4410
      TabIndex        =   14
      Top             =   3600
      Width           =   1275
   End
   Begin VB.ListBox lstVideo 
      Height          =   1620
      Left            =   180
      TabIndex        =   7
      Top             =   4140
      Width           =   4065
   End
   Begin VB.VScrollBar scrLength 
      Height          =   285
      Left            =   5400
      Max             =   5
      Min             =   300
      TabIndex        =   5
      Top             =   630
      Value           =   5
      Width           =   195
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Text            =   "5"
      Top             =   630
      Width           =   375
   End
   Begin VB.ListBox lstPlay 
      Height          =   1620
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   180
      Width           =   4065
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2430
      TabIndex        =   0
      Top             =   6390
      Width           =   1095
   End
   Begin MCI.MMControl mmWave 
      Height          =   330
      Left            =   3780
      TabIndex        =   57
      Top             =   1980
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
   End
   Begin VB.PictureBox ssRemove 
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   2520
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   56
      Top             =   1980
      Width           =   1095
   End
   Begin VB.PictureBox ssDelete 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1350
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   17
      Top             =   1980
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdBrowse 
      Left            =   90
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin MCI.MMControl mmCD 
      Height          =   330
      Left            =   180
      TabIndex        =   12
      Top             =   3600
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   582
      _Version        =   327681
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
      RecordVisible   =   0   'False
      DeviceType      =   "CDAudio"
   End
   Begin VB.PictureBox cmdSelect2 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4590
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   5940
      Width           =   1095
   End
   Begin VB.PictureBox ssBrowse2 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   5940
      Width           =   1095
   End
   Begin VB.PictureBox ssBrowse1 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.PictureBox cmdSelect1 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4590
      ScaleHeight     =   255
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Image imgPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1365
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   1365
   End
   Begin VB.Label lblTrackLength 
      Caption         =   "Track Length"
      Height          =   195
      Left            =   2340
      TabIndex        =   53
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label lblLength 
      Height          =   195
      Left            =   3420
      TabIndex        =   52
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblTracks 
      Height          =   285
      Left            =   1710
      TabIndex        =   21
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label lblOf 
      Caption         =   "of"
      Height          =   285
      Left            =   1440
      TabIndex        =   20
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblTrack 
      Caption         =   "Track No"
      Height          =   285
      Left            =   180
      TabIndex        =   19
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblTrackNo 
      Height          =   285
      Left            =   990
      TabIndex        =   18
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblApprox 
      Alignment       =   2  'Center
      Caption         =   "52.5 Kbytes"
      Height          =   195
      Left            =   4410
      TabIndex        =   16
      Top             =   1260
      Width           =   1185
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      Caption         =   "Approx Size:"
      Height          =   195
      Left            =   4410
      TabIndex        =   15
      Top             =   990
      Width           =   1185
   End
   Begin VB.Label lblCD 
      Alignment       =   2  'Center
      Caption         =   "CD Audio"
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Shape Shape4 
      Height          =   1545
      Left            =   90
      Top             =   2430
      Width           =   5685
   End
   Begin VB.Shape Shape3 
      Height          =   2265
      Left            =   90
      Top             =   4050
      Width           =   5685
   End
   Begin VB.Label lblVideo 
      Alignment       =   2  'Center
      Caption         =   "Picture"
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   4140
      Width           =   1365
   End
   Begin VB.Label lblAudio 
      Alignment       =   2  'Center
      Caption         =   ".WAV Audio"
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   180
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      Height          =   2265
      Left            =   90
      Top             =   90
      Width           =   5685
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   4320
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label lblSample 
      Caption         =   "Sample Length"
      Height          =   465
      Left            =   4410
      TabIndex        =   4
      Top             =   540
      Width           =   645
   End
End
Attribute VB_Name = "frmAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim datmarty As Database
    Dim tabSongs As Recordset
    Dim tabBands As Recordset

Private Sub cmdCD_Click()
    mmCD.Command = "Open"
    lblTrackNo = mmCD.Track
    lblTracks = mmCD.Tracks
    mmCD.Wait = False
    subUpdateTrack
End Sub

Private Sub cmdClose_Click()
    mmWave.Command = "Stop"
    mmWave.Command = "Back"
    mmWave.Command = "Close"
    frmStart.Show
    frmAV.Hide
End Sub

Private Sub cmdPreview_Click()
    soundname$ = lstPlay.List(lstPlay.ListIndex)
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Rem X% = sndPlaySound(soundname$, wFlags%)
End Sub

Private Sub cmdSelect1_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    frmStart.txtFilename = lstPlay.List(lstPlay.ListIndex)
    varerror = 1
    If lstPlay.ListIndex <> -1 Then
        tabSongs.MoveFirst
        Do Until tabSongs.EOF
            If tabSongs(1) = frmStart.txtSong Then
                'Found it!
                varerror = 0
                Exit Do
            End If
            DoEvents
            tabSongs.MoveNext
        Loop
        'the next 3 lines are for if there is no match, which should not happen..
        Rem If tabSongs(1) <> frmStart.txtSong Then
        Rem     Beep
        Rem     frmHelp.lblHelp.Caption = "Song not found!"
        Rem     frmHelp.Show 1
        Rem     Exit Sub
        Rem End If
        If varerror = 0 Then
            tabSongs.Edit
            tabSongs(17) = lstPlay.List(lstPlay.ListIndex)
            frmStart.chkFormat(5) = 1
            tabSongs(16) = -1   ' update Digital check box
            tabSongs.Update
        Else
            Beep
            frmHelp.Caption = "Error"
            frmHelp.lblHelp.Caption = "Song not found! - check with " & gAuthor
            frmHelp.Show 1
            Exit Sub
        End If
    End If
    cmdClose_Click
End Sub

Private Sub cmdSelect2_Click()
    Set datmarty = OpenDatabase(gDatabasePath)
    Set tabSongs = datmarty.OpenRecordset("Songs")
    frmStart.txtPicture = lstVideo.List(lstVideo.ListIndex)
    varerror = 1
    If lstVideo.ListIndex <> -1 Then
        tabSongs.MoveFirst
        Do Until tabSongs.EOF
            If tabSongs(1) = frmStart.txtSong Then
                'Found it!
                varerror = 0
                Exit Do
            End If
            DoEvents
            tabSongs.MoveNext
        Loop
        'the next 3 lines are for if there is no match, which should not happen..
        Rem If tabSongs(1) <> frmStart.txtSong Then
        Rem     Beep
        Rem     frmHelp.lblHelp.Caption = "Song not found!"
        Rem     frmHelp.Show 1
        Rem     Exit Sub
        Rem End If
        If varerror = 0 Then
            tabSongs.Edit
            tabSongs(18) = lstVideo.List(lstVideo.ListIndex)
            frmStart.chkFormat(4) = 1
            tabSongs(15) = -1   ' update Picture check box
            tabSongs.Update
        Else
            Beep
            frmHelp.Caption = "Error"
            frmHelp.lblHelp.Caption = "Song not found! - check with " & gAuthor
            frmHelp.Show 1
            Exit Sub
        End If
    End If
    cmdClose_Click
End Sub

Private Sub cmdSync_Click()
    Message = "Enter a valid long filename for " & frmStart.txtSong   ' Set prompt.
    Title = "Enter filename" ' Set title.
    Default = frmStart.txtSong & ".WAV"  ' Set default.
    ' Display message, title, and default value.
    txtReturnValue = InputBox(Message, Title, Default)
    mmWave.Command = "Stop"
    mmWave.Command = "Back"
    mmWave.Command = "Close"
    If Right$(CurDir, 1) <> "\" Then
        mmWave.filename = CurDir & "\" & txtReturnValue
    Else
        mmWave.filename = CurDir & txtReturnValue
    End If
    mmWave.Command = "Open"
    mmCD.Command = "Play"
    mmWave.Command = "Record"
    Wait (txtLength)
    mmWave.Command = "Stop"
    mmCD.Command = "Stop"
    mmWave.Command = "Save"
    mmWave.Command = "Back"
    mmWave.Command = "Close"
    Rem subRecord
    subSongListUpdate
    For varLoop = 0 To lstPlay.ListIndex
        If Right$(lstPlay.List(varLoop), Len(txtReturnValue)) = txtReturnValue Then
            lstPlay.ListIndex = varLoop
            lstPlay.Refresh
        End If
    Next varLoop
    mmCD.Command = "Stop"
End Sub

Private Sub cmdTrack_Click(Index As Integer)
    mmCD.Command = "Close"
    mmCD.Command = "Open"
    mmCD.Silent = True
    mmCD.Track = Index + 1
    For varLoop = 1 To 30
        mmCD.Command = "Prev"
    Next varLoop
    For varLoop = 0 To Index - 1
        mmCD.Command = "Next"
    Next varLoop
    mmCD.Silent = False
    mmCD.Command = "Play"
    mmCD.Track = Index + 1
    subUpdateTrack
    lblTrackNo = Index + 1
    lblTrackNo.Refresh
End Sub


Private Sub Form_Activate()
    frmAV.Caption = "Audio and Video - " & frmStart.txtSong
    subSongListUpdate
    subPictureListUpdate
    mmWave.Notify = False
    mmWave.Wait = True
    mmWave.Shareable = False
    mmWave.DeviceType = "WaveAudio"
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    For varLoop = 0 To 29
        cmdTrack(varLoop).Enabled = 0
    Next varLoop
End Sub

Private Sub Form_Terminate()
    mmWave.Command = "Close"
End Sub

Private Sub lstPlay_Click()
    soundname$ = lstPlay.List(lstPlay.ListIndex)
    mmWave.Command = "Back"
    mmWave.Command = "Close"
    mmWave.filename = soundname$
    mmWave.Command = "Open"
    mmWave.Command = "Play"
End Sub

Private Sub lstVideo_Click()
    imgPicture = LoadPicture(lstVideo.List(lstVideo.ListIndex))
End Sub

Private Sub mmCD_EjectClick(Cancel As Integer)
    mmCD.Command = "Stop"
End Sub

Private Sub mmCD_EjectCompleted(Errorcode As Long)
    mmCD.Track = 1
    lblTrackNo = mmCD.Track
    lblTracks = mmCD.Tracks
    subUpdateTrack
End Sub

Private Sub mmCD_NextCompleted(Errorcode As Long)
    lblTrackNo = mmCD.Track
    lblLength = mmCD.TrackLength
    subUpdateTrack
End Sub

Private Sub mmCD_PlayClick(Cancel As Integer)
    lblTrackNo = mmCD.Track
    lblLength = mmCD.TrackLength
    subUpdateTrack
End Sub

Private Sub mmCD_PrevCompleted(Errorcode As Long)
    lblTrackNo = mmCD.Track
    lblLength = mmCD.TrackLength
    subUpdateTrack
End Sub

Private Sub mmWave_PlayCompleted(Errorcode As Long)
    mmWave.Command = "Back"
    mmWave.Command = "Close"
End Sub

Private Sub scrLength_Change()
    txtLength = scrLength.Value
End Sub
Private Sub subSongListUpdate()
    lstPlay.Clear
    Songname = Dir(CurDir & "\*.WAV")
    If Songname <> "" Then lstPlay.AddItem CurDir & "\" & Songname
    Do Until Songname = ""
        Songname = Dir
        lstPlay.AddItem CurDir & "\" & Songname
    Loop
    On Error Resume Next
    lstPlay.RemoveItem 0
End Sub
Private Sub subPictureListUpdate()
    lstVideo.Clear
    Picturename = Dir(CurDir & "\*.BMP")
    If Picturename <> "" Then lstVideo.AddItem CurDir & "\" & Picturename
    Do Until Picturename = ""
        Picturename = Dir
        lstPlay.AddItem CurDir & "\" & Picturename
    Loop
    On Error Resume Next
    lstPicture.RemoveItem 0
End Sub
Private Sub subRecord()
    Shell (gWindowsDir & "SNDREC32.EXE"), vbNormalFocus
    Rem Start recording
    SendKeys " "
    Rem wait the allotted time in seconds
    Wait (Val(txtLength))
    Rem then stop recording
    SendKeys " "
    Wait (1)
    Rem now send 'save As'
    SendKeys "%fa"
    Wait (1)
    Rem now send the name of the song, character by character
    For varLoop = 1 To Len(txtReturnValue)
        SendKeys Mid$(txtReturnValue, varLoop, 1)
    Next varLoop
    ' Rem Add '.wav' to the file name
    ' SendKeys ".WAV"
    Rem change the format to radio quality
    SendKeys "%c{DOWN}{DOWN}{ENTER}"
    Wait (1)
    Rem save the .wav file
    SendKeys "%s"
    Rem wait a few secs for the hard disk drive to save the info
    Wait (3)
    Rem now exit the sndrec32 program
    SendKeys "%fx"
End Sub

Private Sub ssBrowse1_Click()
    cdBrowse.filename = ""
    Do Until UCase$(Right$(cdBrowse.filename, 3)) = "WAV"
        cdBrowse.Filter = "Audio files (*.WAV)|*.WAV"
        cdBrowse.ShowOpen
    Loop
    lstPlay.AddItem cdBrowse.filename
End Sub

Private Sub ssBrowse2_Click()
    cdBrowse.filename = ""
    Do Until UCase$(Right$(cdBrowse.filename, 3)) = "BMP"
        cdBrowse.Filter = "Bitmap files (*.BMP)|*.BMP"
        cdBrowse.ShowOpen
    Loop
    lstVideo.AddItem cdBrowse.filename
End Sub

Private Sub ssDelete_Click()
    msg = "Are you sure you wish to delete " & lstPlay.List(lstPlay.ListIndex)
    response = MsgBox(msg, vbYesNo)
    If response = IDYES Then
        Kill lstPlay.List(lstPlay.ListIndex)
        subSongListUpdate
    End If
End Sub

Private Sub ssRemove_Click()
    msg = "Are you sure you wish to remove " & lstPlay.List(lstPlay.ListIndex)
    response = MsgBox(msg, vbYesNo)
    If response = IDYES Then
        lstPlay.RemoveItem (lstPlay.ListIndex)
        subSongListUpdate
    End If
End Sub

Private Sub txtLength_Change()
    lblApprox = 10.5 * Val(txtLength) & " Kbytes"
End Sub

Private Sub subUpdateTrack()
    mmCD.TimeFormat = mciFormatMilliseconds
    msec = (CDbl(mmCD.TrackLength) / 1000)
    varPosition = InStr(msec, ".")
    If varPosition > 0 Then
        varLeft = Left$(msec, varPosition - 1)
        varLeft = Val(varLeft)
        varMin = Int(varLeft / 60)
        varSec = varLeft - (varMin * 60)
        If Len(varSec) = 1 Then varSec = "0" & varSec
        If Len(varMin) = 1 Then varMin = "0" & varMin
        varTime = varMin & varSec
    End If
    lblLength.Caption = Format$(varTime, "0:00")
    For varLoop = 0 To 29
        If (Val(lblTracks) - 1) >= varLoop Then
            cmdTrack(varLoop).Enabled = 1
        Else
            cmdTrack(varLoop).Enabled = 0
        End If
    Next varLoop
End Sub
