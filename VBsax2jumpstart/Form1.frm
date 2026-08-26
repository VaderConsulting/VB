VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SAX Parser Example"
   ClientHeight    =   5250
   ClientLeft      =   7515
   ClientTop       =   5580
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   5850
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parse"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Test.xml"
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "File name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
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

    Dim reader As New SAXXMLReader                  ' This one will do the work
    Dim contentHandler As New ContentHandlerImpl    ' This one will receive parsing events
    Dim errorHandler As New ErrorHandlerImpl        ' and this one will receive errors
    
    Text2.text = ""
    Set reader.contentHandler = contentHandler      ' And they should work together
    Set reader.errorHandler = errorHandler          ' the same here
    On Error GoTo 10
    reader.parseURL (Text1.text)                    '  Parse it
    
    Exit Sub                                        ' That's all, folks!
    
10  Text2.text = Text2.text & "*** Error *** " & Err.Number & " : " & Err.Description
End Sub

Private Sub Command2_Click()
    End
End Sub

