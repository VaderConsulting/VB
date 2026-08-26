VERSION 5.00
Object = "{BCA4E9CD-1DCA-11D3-BAC1-00104B0FDC01}#1.0#0"; "DirMon.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS Dir Watcher"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   855
   End
   Begin VB.ListBox List5 
      Height          =   3765
      Left            =   4560
      TabIndex        =   10
      Top             =   600
      Width           =   4335
   End
   Begin VB.ListBox List3 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   4335
   End
   Begin VB.ListBox List4 
      Height          =   3375
      Left            =   4560
      TabIndex        =   6
      Top             =   4440
      Width           =   4335
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin DIRMONLibCtl.DirMonCtl DirMonCtl1 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   7920
      Width           =   1575
   End
   Begin DIRMONLibCtl.DirMonCtl DirMonCtl2 
      Height          =   495
      Left            =   1800
      OleObjectBlob   =   "Form1.frx":0024
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin DIRMONLibCtl.DirMonCtl DirMonCtl3 
      Height          =   495
      Left            =   3480
      OleObjectBlob   =   "Form1.frx":0048
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin DIRMONLibCtl.DirMonCtl DirMonCtl4 
      Height          =   495
      Left            =   5160
      OleObjectBlob   =   "Form1.frx":006C
      TabIndex        =   7
      Top             =   7920
      Width           =   1575
   End
   Begin DIRMONLibCtl.DirMonCtl DirMonCtl5 
      Height          =   495
      Left            =   6840
      OleObjectBlob   =   "Form1.frx":0090
      TabIndex        =   11
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CAP_S01\sinv.box, SMS\sinv.box"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label lblCAP 
      Alignment       =   2  'Center
      Caption         =   "CAP_S01\inventry.box, SMS\inventry.box, SMS\dataldr.box"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRoot1 As String
Dim strRoot2 As String
Dim strRoot3 As String
Dim strRoot4 As String
Dim strRoot5 As String


Private Sub cmdClear_Click()
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
End Sub

Private Sub DirMonCtl1_DirChanged()
    Dim fName As String, fSize As Long, fPath As String
    
    fPath = Dir(strRoot1 & "\*.*")
    If fPath <> "" Then
        List1.AddItem "** Dir change.  Files in this dir are:"
        Do Until fPath = ""
            List1.AddItem fPath & vbTab & FileLen(strRoot1 & "\" & fPath) & vbTab & FileDateTime(strRoot1 & "\" & fPath)
            fPath = Dir
        Loop
        List1.AddItem "** ------------------------------ **"
    Else
        List1.AddItem "** Dir Change:  No files exist in this directory."
    End If
End Sub

Private Sub DirMonCtl2_DirChanged()
    Dim fName As String, fSize As Long, fPath As String
    
    fPath = Dir(strRoot2 & "\*.*")
    If fPath <> "" Then
        List2.AddItem "** Dir change.  Files in this dir are:"
        Do Until fPath = ""
            List2.AddItem fPath & vbTab & FileLen(strRoot2 & "\" & fPath) & vbTab & FileDateTime(strRoot2 & "\" & fPath)
            fPath = Dir
        Loop
        List2.AddItem "** ------------------------------ **"
    Else
        List2.AddItem "** Dir Change:  No files exist in this directory."
    End If
End Sub

Private Sub DirMonCtl3_DirChanged()
    Dim fName As String, fSize As Long, fPath As String
    
    fPath = Dir(strRoot3 & "\*.*")
    If fPath <> "" Then
        List3.AddItem "** Dir change.  Files in this dir are:"
        Do Until fPath = ""
            List3.AddItem fPath & vbTab & FileLen(strRoot3 & "\" & fPath) & vbTab & FileDateTime(strRoot3 & "\" & fPath)
            fPath = Dir
        Loop
        List3.AddItem "** ------------------------------ **"
    Else
        List3.AddItem "** Dir Change:  No files exist in this directory."
    End If
End Sub

Private Sub DirMonCtl4_DirChanged()
    Dim fName As String, fSize As Long, fPath As String
    
    fPath = Dir(strRoot4 & "\*.*")
    If fPath <> "" Then
        List4.AddItem "** Dir change.  Files in this dir are:"
        Do Until fPath = ""
            List4.AddItem fPath & vbTab & FileLen(strRoot4 & "\" & fPath) & vbTab & FileDateTime(strRoot4 & "\" & fPath)
            fPath = Dir
        Loop
        List4.AddItem "** ------------------------------ **"
    Else
        List4.AddItem "** Dir Change:  No files exist in this directory."
    End If
End Sub

Private Sub DirMonCtl5_DirChanged()
    Dim fName As String, fSize As Long, fPath As String
    
    fPath = Dir(strRoot5 & "\*.*")
    If fPath <> "" Then
        List5.AddItem "** Dir change.  Files in this dir are:"
        Do Until fPath = ""
            List5.AddItem fPath & vbTab & FileLen(strRoot5 & "\" & fPath) & vbTab & FileDateTime(strRoot5 & "\" & fPath)
            fPath = Dir
        Loop
        List5.AddItem "** ------------------------------ **"
    Else
        List5.AddItem "** Dir Change:  No files exist in this directory."
    End If
End Sub

Private Sub Form_Load()
    strRoot1 = "D:\CAP_S01\inventry.box"
    strRoot2 = "D:\SMS\inboxes\inventry.box"
    strRoot3 = "D:\SMS\inboxes\dataldr.box"
    strRoot4 = "D:\SMS\inboxes\sinv.box"
    strRoot5 = "D:\CAP_S01\sinv.box"
    
    DirMonCtl1.MonitorDir strRoot1
    DirMonCtl2.MonitorDir strRoot2
    DirMonCtl3.MonitorDir strRoot3
    DirMonCtl4.MonitorDir strRoot4
    DirMonCtl5.MonitorDir strRoot5
End Sub

