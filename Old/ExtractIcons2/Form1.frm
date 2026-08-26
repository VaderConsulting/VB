VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract Icons"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Internal"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Associated"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.DirListBox DirListBox 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox DriveListBox 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.FileListBox FileListBox 
      Height          =   2040
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myPath As String
Dim myMaxIcon As Long, myIcon As Long
Dim hInstance As Long

Private Sub cmdSave_Click()
    Dim strType As String
    If Option1.Value Then
        strType = " Associated Icon"
    Else
        strType = " Internal Icon " & HScroll.Value
    End If
    SavePicture Picture1.Image, App.Path & "\" & FileListBox.FileName & strType & ".ico"
End Sub

Private Sub DirListBox_Change()
    FileListBox.Path = DirListBox.Path
End Sub

Private Sub DriveListBox_Change()
    On Error GoTo errorhandle
        DirListBox.Path = DriveListBox.Drive
        Exit Sub
errorhandle:
    DriveListBox.Drive = "c:"
    Err.Number = 0
    Exit Sub
End Sub

Private Sub FileListBox_Click()
    Picture1.Cls
    If Option1 Then
        myIcon = 0
        myPath = FileListBox.Path + "\" + FileListBox.FileName
        myIcon = ExtractAssociatedIconA(hInstance, myPath, 0)
        DrawIcon Picture1.hdc, 0, 0, myIcon
        DestroyIcon myIcon
    Else
        HScroll.Enabled = False
        If Len(DirListBox.Path) > 3 Then
            myPath = DirListBox.Path + "\" + FileListBox.FileName
        Else
            myPath = Left(DriveListBox.Drive, 2) + "\" + FileListBox.FileName
        End If

        myMaxIcon = ExtractIcon(hInstance, myPath, -1)
        Form1.Caption = "Extract Icon (total icons:" + Str(myMaxIcon) + ")"
        HScroll.Value = 0
        Label1.Caption = "Icon num. " + Str(HScroll.Value)
        If myMaxIcon > 0 Then
            HScroll.Enabled = True
            HScroll.Max = myMaxIcon - 1
        End If
        myIcon = ExtractIcon(hInstance, myPath, HScroll.Value)
        DrawIcon Picture1.hdc, 0, 0, myIcon
        DestroyIcon myIcon
    End If
End Sub

Private Sub HScroll_Change()
    Picture1.Cls
    myIcon = ExtractIcon(hInstance, myPath, HScroll.Value)
    Label1.Caption = "Icon num. " + Str(HScroll.Value)
    DrawIcon Picture1.hdc, 10, 2, myIcon
    DestroyIcon myIcon
End Sub

Private Sub Option1_Click()
    HScroll.Visible = False
End Sub

Private Sub Option2_Click()
    HScroll.Visible = True
End Sub
