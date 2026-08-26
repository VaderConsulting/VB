Attribute VB_Name = "modDatabase"
Option Explicit

Public DBConn As ADODB.Connection


Public Function OpenDatabase() As Boolean
  '****************
  '* Debug Mode
  If DebugMode Then
    OpenDatabase = True
    Exit Function
  End If
  '****************
  
  On Error GoTo OpenDatabaseError
  OpenDatabase = False
  DBConn.Mode = adModeShareDenyNone
  DBConn.Open SQLDSN
  
  OpenDatabase = True
  Exit Function
OpenDatabaseError:
  Exit Function
End Function

Public Function CloseDatabase() As Boolean
  '****************
  '* Debug Mode
  If DebugMode Then
    CloseDatabase = True
    Exit Function
  End If
  '****************
  
  On Error GoTo CloseDatabaseError
  CloseDatabase = False
  If (DBConn.State <> adStateOpen) Then Exit Function
  DBConn.Close
  CloseDatabase = True
  Exit Function
CloseDatabaseError:
  Exit Function
End Function
