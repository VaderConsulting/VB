VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Context Sensitive Menu's"
   ClientHeight    =   3255
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView tvwNodes 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuCS 
         Caption         =   "Context Sensitive Type 1"
         Index           =   1
         Visible         =   0   'False
         Begin VB.Menu mnuTest1 
            Caption         =   "Test Case 1"
            Index           =   1
         End
         Begin VB.Menu mnuTest1 
            Caption         =   "Test Case 2"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCS 
         Caption         =   "Context Sensitive Type 2"
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu mnuTest2 
            Caption         =   "CS2 - Test 1"
            Index           =   1
         End
         Begin VB.Menu mnuTest2 
            Caption         =   "CS2 - Test 2"
            Index           =   2
         End
         Begin VB.Menu mnuTest2 
            Caption         =   "CS2 - Test 3"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mButton As Integer

Private Sub Form_Load()
    Dim n As Node
    Set n = tvwNodes.Nodes.Add(, tvwparent, "P1", "Parent 1")
    Set n = tvwNodes.Nodes.Add("P1", tvwChild, "C1", "Child 1") '  Bit 0
    Set n = tvwNodes.Nodes.Add("P1", tvwChild, "C2", "Child 2") '  Bit 1
    Set n = tvwNodes.Nodes.Add("P1", tvwChild, "C4", "Child 3") '  Bit 2
    
    Set n = tvwNodes.Nodes.Add(, tvwparent, "P2", "Parent 2")
    Set n = tvwNodes.Nodes.Add("P2", tvwChild, "C8", "Child 4") '  Bit 3
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub tvwNodes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then ' Right mouse
        mButton = Button
    Else
        mButton = 0
    End If
End Sub

Private Sub tvwNodes_NodeClick(ByVal Node As MSComctlLib.Node)
    If mButton = 2 Then                 ' Right mouse clicked
        If Left(Node.Key, 1) = "C" Then ' Only interested in children
            ' Don't have to do it using bit comparisons, but makes for interesting possibilities
            Bit = Val(Right(Node.Key, 1))
            If Bit And 1 Then
                Debug.Print "Case 1"
                frmMain.PopupMenu mnuCS(1) ' You could use 'Bit' here instead of '1'
            End If
            If Bit And 2 Then
                Debug.Print "Case 2"
                mnuTest2(1).Visible = True
                mnuTest2(2).Visible = False
                mnuTest2(3).Visible = False
                frmMain.PopupMenu mnuCS(2)
            End If
            If Bit And 4 Then
                Debug.Print "Case 3"
                mnuTest2(1).Visible = True
                mnuTest2(2).Visible = True
                mnuTest2(3).Visible = False
                frmMain.PopupMenu mnuCS(2)
            End If
            If Bit And 8 Then
                Debug.Print "Case 4"
                mnuTest2(1).Visible = True
                mnuTest2(2).Visible = True
                mnuTest2(3).Visible = True
                frmMain.PopupMenu mnuCS(2)
            End If
        End If
    End If
End Sub

