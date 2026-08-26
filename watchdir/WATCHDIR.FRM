VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   1680
   ClientTop       =   1545
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   6630
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lbMsg 
      Caption         =   "Label2"
      Height          =   1455
      Left            =   3960
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hChangeHandle As Long
Dim hWatched As Long
Dim terminateFlag As Long

Private Sub Form_Load()

   lbMsg = "Press 'Begin Watch'"
   
End Sub


Private Sub cmdEnd_Click()

   If hWatched > 0 Then Call WatchDelete(hWatched)
      
   Unload Me
   Set Form1 = Nothing

End Sub


Private Sub cmdStop_Click()

  'clean up by deleting the handle to the watched directory
   Call WatchDelete(hWatched)
   hWatched = 0
      
   cmdBegin.Enabled = True
   lbMsg = "Press 'Begin Watch'"

End Sub
Private Sub cmdBegin_Click()

   Dim r As Long
   Dim watchPath As String
   Dim watchStatus As Long
   
   watchPath = "c:\bat"
   
   terminateFlag = False
   cmdBegin.Enabled = False
   
   lbMsg = "Using Explorer and Notepad, create, modify, rename, delete or "
   lbMsg = lbMsg & "change the attributes of a text file in the watched directory."""

  'get the first file text attributes to the listbox (if any)
   WatchChangeAction watchPath
   
  'show a msgbox to indicate the watch is starting
   MsgBox "Beginning watching of folder " & watchPath & " .. press OK"
   
  'create a watched directory
   hWatched = WatchCreate(watchPath, FILE_NOTIFY_FLAGS)
   
  'poll the watched folder
   watchStatus = WatchDirectory(hWatched, 100)
  
  'if WatchDirectory exited with watchStatus = 0,
  'then there was a change in the folder.
   If watchStatus = 0 Then
   
      'update the listbox for the first file found in the
      'folder and indicate a change took place.
       WatchChangeAction watchPath
       
       MsgBox "The watched directory has been changed.  Resuming watch..."
       
      '(perform actions)
      'this is where you'd actually put code to perform a
      'task based on the folder changing.
      'FindNextChangeNotification API, again exiting if
      'watchStatus indicates the terminate flag was set
       Do
         watchStatus = WatchResume(hWatched, 100)
         
         If watchStatus = -1 Then
              'watchStatus must have exited with the terminate flag
               MsgBox "Watching has been terminated for " & watchPath
         
         Else: WatchChangeAction watchPath
               MsgBox "The watched directory has been changed again."
              
              '(perform actions)
              'this is where you'd actually put code to perform a
              'task based on the folder changing.
               
         End If
         
       Loop While watchStatus = 0
   
   
   Else
     'watchStatus must have exited with the terminate flag
      MsgBox "Watching has been terminated for " & watchPath
   
   End If
   
End Sub



Private Function WatchCreate(lpPathName As String, flags As Long) As Long

  'FindFirstChangeNotification members:
  '
  '  lpPathName: folder to watch
  '  bWatchSubtree:
  '     True = watch specified folder and its sub folders
  '     False = watch the specified folder only
  '  flags: OR'd combination of the FILE_NOTIFY_ flags to apply
  
   WatchCreate = FindFirstChangeNotification(lpPathName, False, flags)

End Function


Private Sub WatchDelete(hWatched As Long)
    
   Dim r As Long
   
   terminateFlag = True
   DoEvents

   r = FindCloseChangeNotification(hWatched)
 
End Sub


Private Function WatchDirectory(hWatched As Long, interval As Long) As Long

  'Poll the watched folder.
  'The Do..Loop will exit when:
  '   r = 0, indicating a change has occurred
  '   terminateFlag = True, set by the WatchDelete routine
  
   Dim r As Long
   
   Do
   
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchDirectory = r
   
End Function


Private Function WatchResume(hWatched As Long, interval) As Boolean

   Dim r As Long
   
   r = FindNextChangeNotification(hWatched)
   
   Do
      
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchResume = r
   
End Function


Private Sub WatchChangeAction(fPath As String)

   Dim fName As String
   
   List1.Clear

   fName = Dir(fPath & "\" & "*.txt")
   
   If fName > "" Then
   
      List1.AddItem "path: " & vbTab & fPath
      List1.AddItem "file: " & vbTab & fName
      List1.AddItem "size: " & vbTab & FileLen(fPath & "\" & fName)
      List1.AddItem "attr: " & vbTab & GetAttr(fPath & "\" & fName)
   
   End If

End Sub


