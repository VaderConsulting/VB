Attribute VB_Name = "Module1"
Option Explicit

Public Const INFINITE = &HFFFF

Public Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Public Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Public Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Public Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                 FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                 FILE_NOTIFY_CHANGE_LAST_WRITE

Declare Function FindFirstChangeNotification Lib "kernel32" _
    Alias "FindFirstChangeNotificationA" _
   (ByVal lpPathName As String, _
    ByVal bWatchSubtree As Long, _
    ByVal dwNotifyFilter As Long) As Long

Declare Function FindCloseChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Declare Function FindNextChangeNotification Lib "kernel32" _
   (ByVal hChangeHandle As Long) As Long

Declare Function WaitForSingleObject Lib "kernel32" _
   (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_ABANDONED = &H80
Public Const WAIT_IO_COMPLETION = &HC0
Public Const WAIT_TIMEOUT = &H102
Public Const STATUS_PENDING = &H103

