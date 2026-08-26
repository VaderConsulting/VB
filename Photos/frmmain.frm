VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{50F16B18-467E-11D1-8271-00C04FC3183B}#1.0#0"; "shimgvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Photos"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin PREVIEWLibCtl.Preview Preview 
      Height          =   5775
      Left            =   2400
      TabIndex        =   2
      Top             =   3360
      Width           =   6255
   End
   Begin MSComctlLib.ImageList imlPhotos 
      Left            =   8760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "Folder"
            Object.Tag             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   "File"
            Object.Tag             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   3135
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "imlPhotos"
      SmallIcons      =   "imlPhotos"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DirListBox DirFolders 
      Height          =   8865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DirFolders_Change()
    Dim strPath As String
    Dim strDir As String
    
    If Right(DirFolders.Path, 1) = "\" Then
        strPath = DirFolders.Path
    Else
        strPath = DirFolders.Path & "\"
    End If
    
    lvwFiles.ListItems.Clear
    
    strDir = Dir(strPath & "*.jpg")
    Do Until strDir = ""
        lvwFiles.ListItems.Add , strDir, strDir, "File"
        strDir = Dir()
    Loop
    
End Sub

Private Sub Form_Load()
    frmSplash.Show vbModal
    frmSplash.Refresh
    
    DirFolders.Path = App.Path
    'DirFolders.Path = "\\server\dave\my pictures\stuff"
End Sub

Private Sub lvwFiles_Click()
    Dim strPath As String
    
    If Right(DirFolders.Path, 1) = "\" Then
        strPath = DirFolders.Path
    Else
        strPath = DirFolders.Path & "\"
    End If
    Preview.ShowFile strPath & lvwFiles.ListItems(lvwFiles.SelectedItem.Index), 1
    Preview.BestFit
End Sub
