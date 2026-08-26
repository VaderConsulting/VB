VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Write & Read from/to an ini file"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Remove Value"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add Value"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save ini file"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load ini file"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Value"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Retrieve Value"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "ini reader.frx":0000
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program contains functions and examples for editing a ini-file
'Created by Lennert Van Damme
'the textbox contains an example ini-file
'Feel free to mail any bugs/suggestions/questions
'vdlennert@hotmail.com

Private Sub Command1_Click()
    'Get the amount of variables stored
    MsgBox GetNumberOfChars(Text1.Text, "=")
End Sub

Private Sub Command2_Click()
    'Retrieve the value of a certain variable
    MsgBox RetrieveValue(Text1.Text, InputBox("Enter name of variable:"))
End Sub

Function GetNumberOfChars(Text As String, Char As String) As Long
    If Len(Char) <> 1 Then Exit Function
    Dim InstrPos As Long, TextLen As Long, a As Variant
    Dim Count As Long
    
    TextLen = 1
    Count = 0
    
    For i% = 1 To Len(Text)
        a = Right(Text, TextLen)
        If Left(a, 1) = Char Then
            Count = Count + 1
        End If
        TextLen = TextLen + 1
    Next i%
    
    GetNumberOfChars = Count
    
End Function

Function RetrieveValue(Text As String, Value)
    Dim SearchText, tBuf
    
    If InStr(Text, Value & "=") = 0 Then Exit Function
    
    
    SearchText = InStr(Text, Value & "=")
    SearchText = SearchText - 1
    tBuf = Right(Text, Len(Text) - SearchText)
    SearchText = InStr(tBuf, Chr(13))
    SearchText = SearchText - 1
    tBuf = Left(tBuf, SearchText)
    tBuf = Right(tBuf, Len(tBuf) - Len(Value) - 1)
    RetrieveValue = tBuf

End Function

Function ChangeValue(Text As String, OldValue As String, NewValue As String)
    If Text = "" Or OldValue = "" Or NewValue = "" Then Exit Function
    ChangeValue = Replace(Text, OldValue, NewValue)
End Function

Private Sub Command3_Click()
    'Change the value of a variable
    a = InputBox("Enter name of variable to change:")
    Text1.Text = ChangeValue(Text1.Text, a & "=" & RetrieveValue(Text1.Text, a), a & "=" & InputBox("Enter new value:"))
End Sub

Private Sub Command4_Click()
    'Load an ini file into the textbox
    CommonDialog1.DialogTitle = "Load ini file"
    CommonDialog1.Filter = "ini files (*.ini)|*.ini"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    
    Open CommonDialog1.FileName For Input As #1
        Text1.Text = Input$(LOF(1), #1)
    Close #1
End Sub

Private Sub Command5_Click()
    CommonDialog1.DialogTitle = "Save ini file"
    CommonDialog1.Filter = "ini files (*.ini)|*.ini"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    
    Open CommonDialog1.FileName For Output As #1
        Print #1, Text1.Text
    Close #1
End Sub

Private Sub Command6_Click()
    a = InputBox("Enter variable name:")
    b = InputBox("Enter value:")
    Text1.Text = Text1.Text & a & "=" & b & vbCrLf
End Sub

Private Sub Command7_Click()
    a = InputBox("Enter name of variable to delete:")
    'Text1.Text = ChangeValue(Text1.Text, a & "=" & RetrieveValue(Text1.Text, a), a & "=" & InputBox("Enter new value:"))
    Text1.Text = Replace(Text1.Text, a & "=" & RetrieveValue(Text1.Text, a) & vbCrLf, "")
    'Text1.Text = Left(Text1.Text, Len(Text1.Text) - 2)
    
End Sub

Function GetNextReturn(Text As String, StartPoint As Long)
    GetNextReturn = InStr(StartPoint, Text, vbCrLf)
End Function

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = Len(Text1.Text)
End Sub
