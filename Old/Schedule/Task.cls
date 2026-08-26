VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Job"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Schedule Control
'
' Job Class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 05/27/1999
' Updates: 12/12/1999. * The class name was changed to Job to avoid
'                      problems with the Task object from the typelib.
'                    
'
'*********************************************************************************************
Option Explicit

Dim m_Task As Task
Dim m_Name As String
Dim m_FileName As String
Dim m_Trgs As Triggers

Public Property Get ApplicationName() As String
    
    ApplicationName = StrFromPtrW(m_Task.GetApplicationName)
    
End Property

Public Property Let ApplicationName(ByVal New_App As String)
    
    m_Task.SetApplicationName New_App
    
End Property

Public Property Get CommandLine() As String
    
    CommandLine = StrFromPtrW(m_Task.GetParameters)
    
End Property

Public Property Let CommandLine(ByVal New_CommandLine As String)
    
    m_Task.SetParameters New_CommandLine
    
End Property

Public Property Get Comment() As String
    
    Comment = StrFromPtrW(m_Task.GetComment)
    
End Property

Public Property Let Comment(ByVal New_Comment As String)
    
    m_Task.SetComment New_Comment
    
End Property

Public Property Get Creator() As String
    
    Creator = StrFromPtrW(m_Task.GetCreator)
    
End Property

Public Property Let Creator(ByVal New_Creator As String)
    
    m_Task.SetCreator New_Creator
    
End Property

Public Property Get ErrorRetryCount() As Integer
    
    ErrorRetryCount = m_Task.GetErrorRetryCount
    
End Property

Public Property Let ErrorRetryCount(ByVal New_ErrorRetryCount As Integer)
    
    m_Task.SetErrorRetryCount New_ErrorRetryCount
    
End Property

Public Property Get ErrorRetryInterval() As Integer
    
    ErrorRetryInterval = m_Task.GetErrorRetryInterval
    
End Property

Public Property Let ErrorRetryInterval(ByVal New_ErrorRetryInterval As Integer)
    
    m_Task.SetErrorRetryInterval New_ErrorRetryInterval
    
End Property

Public Property Get ExitCode() As Long
    
    ExitCode = m_Task.GetExitCode
    
End Property

Public Property Get Flags() As saTaskFlags
    
    Flags = m_Task.GetFlags
    
End Property

Public Property Let Flags(ByVal New_Flags As saTaskFlags)
    
    m_Task.SetFlags New_Flags
    
End Property

Public Property Get IdleWait() As Integer
    
    IdleWait = m_Task.GetIdleWait
    
End Property

Public Property Let IdleWait(ByVal New_IdleWait As Integer)
    
    m_Task.SetIdleWait New_IdleWait
    
End Property

Public Property Get MaxRunTime() As Long
    
    MaxRunTime = m_Task.GetMaxRunTime
    
End Property

Public Property Let MaxRunTime(ByVal New_MaxRunTimeMS As Long)
    
    m_Task.SetMaxRunTime New_MaxRunTimeMS
    
End Property

Public Property Get Priority() As Long
    
    Priority = m_Task.GetPriority
    
End Property

Public Property Let Priority(ByVal New_Priority As Long)
    
    m_Task.SetPriority New_Priority
    
End Property

Public Property Get Status() As Long
    
    Status = m_Task.GetStatus
    
End Property

Public Sub Edit(ByVal Parent As Object)

    m_Task.EditWorkItem Parent.hWnd, 0
    
End Sub

Public Sub Run()
    
    m_Task.Run
    
End Sub

Public Sub Terminate()
    
    m_Task.Terminate
    
End Sub

Public Property Get LastRunTime() As Date
Dim ST As SYSTEMTIME, S As Double

    m_Task.GetMostRecentRunTime ST
    SystemTimeToVariantTime ST, S
    
    LastRunTime = S
    
End Property

Public Property Get NextRunTime() As Date
Dim ST As SYSTEMTIME, S As Double

    m_Task.GetNextRunTime ST
    SystemTimeToVariantTime ST, S
    
    NextRunTime = S
    
End Property

Public Property Get FileName() As String
Dim PF As IPersistFile, Ptr As Long

    Set PF = m_Task
   
    Ptr = PF.GetCurFile
    FileName = StrFromPtrW(Ptr)

End Property

Public Property Get Name() As String

    Name = m_Name
    
End Property

Friend Property Let Name(ByVal NewName As String)

    m_Name = NewName
    
End Property

Friend Property Let FileName(ByVal NewName As String)

    m_FileName = NewName
    
End Property

Public Sub Save(Optional ByVal FileName As String)
Dim PF As IPersistFile, T As Trigger

    Set PF = m_Task
    
    For Each T In m_Trgs
        T.Update
    Next
    
    If FileName <> "" Then
        PF.Save FileName, 0
    Else
        PF.Save vbNullString, 0
    End If
    
End Sub

'Friend Sub Load(ByVal FileName As String)
'Dim PF As IPersistFile
'
'    Set PF = m_Task
'    PF.Load FileName, 0
'
'    m_FileName = FileName
'
'End Sub

Public Sub ShowProperties(Optional ParenthWnd As Long)

    m_Task.EditWorkItem ParenthWnd, 0
    
End Sub

Friend Property Get TaskObject() As Task

    Set TaskObject = m_Task
    
End Property

Friend Property Set TaskObject(ByVal New_Job As Task)

    Set m_Task = New_Job
    Set m_Trgs = New Triggers
    Set m_Trgs.Task = m_Task
    
End Property

Public Property Get WorkingDirectory() As String
    
    WorkingDirectory = StrFromPtrW(m_Task.GetWorkingDirectory)
    
End Property

Public Property Let WorkingDirectory(ByVal New_WorkingDirectory As String)
    
    m_Task.SetWorkingDirectory New_WorkingDirectory
    
End Property

Public Property Get Triggers() As Triggers

    Set Triggers = m_Trgs
    
End Property

Private Sub Class_Terminate()

    Set m_Task = Nothing
    Set m_Trgs = Nothing
    
End Sub