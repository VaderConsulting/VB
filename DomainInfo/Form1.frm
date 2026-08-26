VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Domain Info"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportGroups 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdRemoveObject 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstGroups 
      Enabled         =   0   'False
      Height          =   4350
      Left            =   2640
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdGroups 
      Caption         =   "Groups"
      Enabled         =   0   'False
      Height          =   975
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtDomain 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Enabled         =   0   'False
      Height          =   975
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Computers"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Groups"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optClass 
      Caption         =   "Users"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox lstObjects 
      Height          =   4350
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   5415
   End
   Begin VB.Label lblCount 
      Caption         =   "0 listed"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblGroupInfo 
      Caption         =   "Clicking the 'Groups' button will retrieve the list of groups that all objects listed are a member of."
      Height          =   975
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblDomain 
      Caption         =   "Domain name"
      Height          =   255
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
Private Sub cmdExportGroups_Click()
    Dim AppDir As String, i As Integer
    
    If Right(App.Path, 1) = "\" Then
        AppDir = App.Path
    Else
        AppDir = App.Path & "\"
    End If
    Open AppDir & "grouplist.txt" For Output As #1
        For i = 0 To Form1.lstGroups.ListCount - 1
            Print #1, Form1.lstGroups.List(i)
        Next i
    Close 1
    MsgBox "Exported to " & AppDir & "grouplist.txt", vbOKOnly + vbInformation
End Sub

Private Sub cmdExtract_Click()
    Dim sDomainName As String
    
    sDomainName = Form1.txtDomain
    
    If Len(Trim(sDomainName)) > 0 Then
        If optClass(0).Value Then GetDomainObjects sDomainName, True, False, False
        If optClass(1).Value Then GetDomainObjects sDomainName, False, True, False
        If optClass(2).Value Then GetDomainObjects sDomainName, False, False, True
    End If
End Sub

Private Sub cmdGroups_Click()
    Dim i As Integer
    
    Form1.lstGroups.Clear
    Form1.lstGroups.Enabled = True
    
    For i = 0 To Form1.lstObjects.ListCount - 1
        If Form1.lstObjects.Selected(i) = True Then
            GetGroupMembership Form1.lstObjects.List(i)
        End If
    Next i
End Sub

Private Sub lstObjects_Click()
    If lstObjects.SelCount > 0 Then
        cmdRemoveObject.Enabled = True
        cmdGroups.Enabled = True
    Else
        cmdRemoveObject.Enabled = False
        cmdGroups.Enabled = True
    End If
End Sub

Private Sub optClass_Click(Index As Integer)
    If Len(txtDomain.Text) > 0 Then
        cmdExtract.Enabled = True
        Form1.lstGroups.Clear
    Else
        cmdExtract.Enabled = False
    End If
    
    If Index = 1 Then cmdGroups.Enabled = False
    
End Sub

Private Sub txtDomain_Change()
    If optClass(0).Value Or optClass(1).Value Or optClass(2).Value Then
        If Len(txtDomain.Text) > 0 Then
            cmdExtract.Enabled = True
        Else
            cmdExtract.Enabled = False
        End If
    Else
        cmdExtract.Enabled = False
    End If
End Sub
