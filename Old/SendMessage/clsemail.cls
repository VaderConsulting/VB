VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SendEmail(strFrom As String, strTo As String, strSubject As String, strMessage As String, strServer As String) As Boolean
    Dim objEmail As Object
    
    If Trim(strFrom) = "" Or Trim(strTo) = "" Or Trim(strSubject) = "" Or Trim(strMessage) = "" Or Trim(strServer) = "" Then
        Exit Function
    End If
    
    SendEmail = True
    
    Set objEmail = CreateObject("CDO.Message")
    
    objEmail.From = strFrom
    objEmail.To = strTo
    objEmail.Subject = strSubject
    objEmail.Textbody = strMessage
    
    ' sendusing:
    ' 1 use local SMTP service (local to the specified smtpserver)
    ' 2 use port
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objEmail.Configuration.Fields.Update
    
    On Error Resume Next
        objEmail.Send
        'Return True if it worked
        If Err <> 0 Then SendEmail = False
    On Error GoTo 0
    
    Set objEmail = Nothing
End Function
