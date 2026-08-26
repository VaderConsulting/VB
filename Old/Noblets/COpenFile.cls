VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COpenFile"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Attribute VB_Description = "A COpenFile object which is a file returned by NetFileEnum - resources opened at this server or share point.  A COpenFile is a dependent object of a CServer - you can't have one without a CServer object."
Option Explicit

Public ID            As Long
Public Permissions   As Long
Public NumLocks      As Long
Public PathName      As String
Public UserName      As String   ' Will be computername if server is share level security






