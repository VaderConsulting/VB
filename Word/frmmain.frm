VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Property Search"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtPropertyValue 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   9375
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   19
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   20
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Text            =   "*.doc"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblInfo 
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   6360
      Width           =   7455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblFilename 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   31
      Top             =   3360
      Width           =   8295
   End
   Begin VB.Label lblSelected 
      Caption         =   "Selected File :"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   29
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   28
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   4920
      TabIndex        =   27
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4920
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4920
      TabIndex        =   25
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4920
      TabIndex        =   24
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblPropertyName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   23
      Top             =   480
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   15000
      Y1              =   6255
      Y2              =   6255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   15000
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label lblFilter 
      Caption         =   "Filter :"
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblPath 
      Caption         =   "Path :"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FSO As Scripting.FileSystemObject

Sub doSearch(oFolder As Scripting.Folder)
    Dim varFile As Variant
    Dim varSubfolder As Scripting.Folder
    
    For Each varFile In oFolder.Files
        If LCase(varFile.Name) Like LCase(txtFilter.Text) Then
            lstFiles.AddItem oFolder.Path & "\" & varFile.Name
            lblInfo.Caption = "Searching... found " & lstFiles.ListCount & " files matching " & txtFilter.Text
            lblInfo.Refresh
        End If
    Next
    For Each varSubfolder In oFolder.SubFolders
        doSearch varSubfolder
    Next
End Sub

Private Sub cmdClose_Click()
    Set FSO = Nothing
    End
End Sub

Private Sub cmdSearch_Click()
    Dim intLoop As Integer
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    cmdSearch.Enabled = False
    cmdClose.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    txtPath.Enabled = False
    txtFilter.Enabled = False
    lblFilename.Caption = ""
    lblInfo.Caption = "Searching..."
    lstFiles.Clear
    
    For intLoop = 0 To txtPropertyValue().UBound
        txtPropertyValue(intLoop).Text = ""
    Next intLoop
    
    doSearch FSO.GetFolder(txtPath)
    
    Screen.MousePointer = vbDefault
    lblInfo.Caption = "Idle.  Found " & frmMain.lstFiles.ListCount & " files found matching " & txtFilter.Text
    txtPath.Enabled = True
    txtFilter.Enabled = True
    cmdClose.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub Form_Load()
    lblPropertyName(0).Caption = "Title"
    lblPropertyName(1).Caption = "Subject"
    lblPropertyName(2).Caption = "Author"
    lblPropertyName(3).Caption = "Manager"
    lblPropertyName(4).Caption = "Company"
    lblPropertyName(5).Caption = "Category"
    lblPropertyName(6).Caption = "Keywords"
    lblPropertyName(7).Caption = "Comments"
    
    lblPropertyName(8).Caption = "Document number"
    lblPropertyName(9).Caption = "Release Date"
    lblPropertyName(10).Caption = "Version"
    lblPropertyName(11).Caption = "CDC #"
    lblPropertyName(12).Caption = "Release Statement"
    lblPropertyName(13).Caption = "Requisition #"
    lblPropertyName(14).Caption = "Contract #"
End Sub

Private Sub lstFiles_Click()
    Dim strFilename As String
    Dim intLoop As Integer
    
    cmdClose.Enabled = False
    cmdSearch.Enabled = False
    strFilename = lstFiles.List(lstFiles.ListIndex)
    
    lblInfo.Caption = "Loading properties..."
    
    lblFilename.Caption = strFilename
    
    For intLoop = 0 To txtPropertyValue().UBound
        txtPropertyValue(intLoop).Text = ""
    Next intLoop
    
    Screen.MousePointer = vbHourglass
    GetDocInfo strFilename
    Screen.MousePointer = vbDefault
    
    lblInfo.Caption = "Idle"
    
    cmdClose.Enabled = True
    cmdSearch.Enabled = True
End Sub

Private Sub GetDocInfo(strWordDoc As String)
    'Dim oWord As Word.Application
    Dim oWord As Object
    Dim oWordDoc As Object
    'Dim oWordDoc As Document
    
    Dim vProperty As Variant
    
    Set oWord = CreateObject("Word.Application")
    
    On Error Resume Next
    
    Set oWordDoc = oWord.Documents.Open(strWordDoc)
    txtPropertyValue(0) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Title")
    txtPropertyValue(1) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Subject")
    txtPropertyValue(2) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Author")
    txtPropertyValue(3) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Manager")
    txtPropertyValue(4) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Company")
    txtPropertyValue(5) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Category")
    txtPropertyValue(6) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Keywords")
    txtPropertyValue(7) = oWordDoc.ActiveWindow.Document.BuiltInDocumentProperties("Comments")

    txtPropertyValue(8) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(1)
    txtPropertyValue(9) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(2)
    txtPropertyValue(10) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(3)
    txtPropertyValue(11) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(4)
    txtPropertyValue(12) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(5)
    txtPropertyValue(13) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(6)
    txtPropertyValue(14) = oWordDoc.ActiveWindow.Document.CustomDocumentProperties(7)
        
    oWordDoc.Application.Documents(1).Close False
    oWord.Application.Quit False
    
    Set oWordDoc = Nothing
End Sub

Private Sub txtPath_Change()
    If Trim(txtPath.Text) = "" Then
        cmdSearch.Enabled = False
    Else
        cmdSearch.Enabled = True
    End If
End Sub
