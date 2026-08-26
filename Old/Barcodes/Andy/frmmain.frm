VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barcode Input Application"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.ListBox lstInvalid 
      Height          =   4155
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ListBox lstCodes 
      Height          =   4155
      Left            =   4560
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtCodes 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblCodes 
      Alignment       =   2  'Center
      Caption         =   "Valid Codes"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblInvalid 
      Alignment       =   2  'Center
      Caption         =   "Invalid codes"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblRaw 
      Alignment       =   2  'Center
      Caption         =   "Raw text to parse"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcess_Click()
    Dim RawCodes As String
    Dim Codes() As String
    
    ' If no data to process, exit.
    If txtCodes.Text = "" Then Exit Sub
    
    ' Get a copy of the inputted text.
    RawCodes = txtCodes.Text
    
    ' Split each line of text into individual items.
    Codes() = Split(RawCodes, vbCrLf)
    
    ' Go through each item, and determine what it should actually be.
    For Each Code In Codes()
        ' If the item is a valid length, process it further.
        'If Len(Code) = 8 Or Len(Code) = 16 Then
            Select Case Left(Code, 2)
                Case "S9"
                    AddifNotPresent lstCodes, Mid(Code, 2, 7)
                Case "1S"
                    If Left(Code, 6) = "1S6266" Then
                        ' This is a Model Number and Serial Number combined
                        AddifNotPresent lstCodes, Mid(Code, 10, 7)
                        ' Model Number = Mid(Code, 3, 7)
                    Else
                        lstInvalid.AddItem Code
                    End If
                Case "1P"
                    If Left(Code, 6) = "1P6339" Then
                        ' This is a model number ONLY
                        ' Model Number = Mid(Code, 3, 7)
                    Else
                        lstInvalid.AddItem Code
                    End If
                Case "EO"
                    ' Do nothing, this is an EOF marker
                Case Else
                    ' An unknown item type.
                    lstInvalid.AddItem Code
            End Select
        'End If
    Next
    txtCodes.Text = ""
    cmdNext.Enabled = True
End Sub

' Adds an item to a listbox if it doesn't already exist
Sub AddifNotPresent(L As ListBox, str As String)
    Dim Present As Boolean
    For a = 0 To L.ListCount - 1
        If L.List(a) = str Then Present = True
    Next a
    If Not Present Then
        L.AddItem str
    End If
End Sub

