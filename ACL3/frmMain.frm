VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton command2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim sUserName As String
    Dim sFolderName As String
    
    sUserName = Trim$(CStr(text2.Text))
    sFolderName = Trim$(CStr(text1.Text))
    SetAccess sUserName, sFolderName, GENERIC_READ Or GENERIC_EXECUTE Or DELETE Or GENERIC_WRITE
End Sub

Private Sub Command2_Click()
    Dim sUserName As String
    Dim sFolderName As String
    
    sUserName = Trim$(text2.Text)
    sFolderName = Trim$(text1.Text)
    SetAccess sUserName, sFolderName, GENERIC_EXECUTE Or GENERIC_READ
End Sub

Private Sub Command3_Click()
    Dim sUserName As String
    Dim sFolderName As String
    
    sUserName = Trim$(text2.Text)
    sFolderName = Trim$(text1.Text)
    SetAccess sUserName, sFolderName, GENERIC_ALL
End Sub

Private Sub Command4_Click()
    Dim sUserName As String
    Dim sFolderName As String
    
    sUserName = Trim$(text2.Text)
    sFolderName = Trim$(text1.Text)
    SetAccess sUserName, sFolderName, FILE_GENERIC_LIST
End Sub

Private Sub Form_Load()
    text1.Text = "enter folder name"
    text2.Text = "enter username"
    Command1.Caption = "Change"
    command2.Caption = "Read && Add"
    Command3.Caption = "Full"
    Command4.Caption = "List"
    
    STANDARD_RIGHTS_READ = READ_CONTROL
    STANDARD_RIGHTS_WRITE = READ_CONTROL
    STANDARD_RIGHTS_EXECUTE = READ_CONTROL
    
    FILE_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF

    FILE_GENERIC_READ = STANDARD_RIGHTS_READ Or _
                        FILE_READ_DATA Or _
                        FILE_READ_ATTRIBUTES Or _
                        FILE_READ_EA Or _
                        SYNCHRONIZE
    
    FILE_GENERIC_WRITE = STANDARD_RIGHTS_WRITE Or _
                         FILE_WRITE_DATA Or _
                         FILE_WRITE_ATTRIBUTES Or _
                         FILE_WRITE_EA Or _
                         FILE_APPEND_DATA Or _
                         SYNCHRONIZE
    
    FILE_GENERIC_EXECUTE = STANDARD_RIGHTS_EXECUTE Or _
                           FILE_READ_ATTRIBUTES Or _
                           FILE_EXECUTE Or _
                           SYNCHRONIZE
    
    FILE_GENERIC_MODIFY = STANDARD_RIGHTS_READ Or _
                          STANDARD_RIGHTS_WRITE Or _
                          FILE_TRAVERSE Or _
                          FILE_LIST_DIRECTORY Or _
                          FILE_READ_ATTRIBUTES Or _
                          FILE_READ_EA Or _
                          FILE_ADD_FILE Or _
                          FILE_ADD_SUBDIRECTORY Or _
                          FILE_WRITE_ATTRIBUTES Or _
                          FILE_WRITE_EA Or _
                          DELETE Or _
                          READ_CONTROL Or _
                          SYNCHRONIZE
    
    FILE_GENERIC_LIST = FILE_TRAVERSE Or _
                        FILE_LIST_DIRECTORY Or _
                        FILE_READ_ATTRIBUTES Or _
                        FILE_READ_EA Or _
                        READ_CONTROL
End Sub

