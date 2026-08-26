VERSION 5.00
Begin VB.UserControl Schedule 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Schedule.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Schedule.ctx":047A
End
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Schedule Control
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 05/27/1999
'
'*********************************************************************************************
Option Explicit

Dim m_SA As SchedulingAgent
Dim m_Jobs As Collection

Enum saTaskFlags
    tfInteractive = &H1
    tfDeleteWhenDone = &H2
    tfDisabled = &H4
    tfStartOnlyIfIdle = &H10
    tfKillOnIdleEnd = &H20
    tfDontStartIfOnBatteries = &H40
    tfKillIfGoingOnBatteries = &H80
    tfRunOnlyIfDocked = &H100
    tfHidden = &H200
    tfRunIfConnectedToInternet = &H400
    tfRestartOnIdleResume = &H800
End Enum

Enum saTriggerTypes
    ttAtLogon = TASK_EVENT_TRIGGER_AT_LOGON
    ttAtSystemStart = TASK_EVENT_TRIGGER_AT_SYSTEMSTART
    ttOnIdle = TASK_EVENT_TRIGGER_ON_IDLE
    ttDaily = TASK_TIME_TRIGGER_DAILY
    ttMonthlyDate = TASK_TIME_TRIGGER_MONTHLYDATE
    ttMonthlyWeekDay = TASK_TIME_TRIGGER_MONTHLYDOW
    ttOnce = TASK_TIME_TRIGGER_ONCE
    ttWeekly = TASK_TIME_TRIGGER_WEEKLY
End Enum

Enum saTriggerMonths
    tmJanuary = TASK_JANUARY
    tmFebruary = TASK_FEBRUARY
    tmMarch = TASK_MARCH
    tmApril = TASK_APRIL
    tmMay = TASK_MAY
    tmJune = TASK_JUNE
    tmJuly = TASK_JULY
    tmAugust = TASK_AUGUST
    tmSeptember = TASK_SEPTEMBER
    tmOctober = TASK_OCTOBER
    tmNovember = TASK_NOVEMBER
    tmDecember = TASK_DECEMBER
End Enum

Enum saTriggerDays
    tdSunday = TASK_SUNDAY
    tdMonday = TASK_MONDAY
    tdTuesday = TASK_TUESDAY
    tdWednesday = TASK_WEDNESDAY
    tdThursday = TASK_THURSDAY
    tdFriday = TASK_FRIDAY
    tdSaturday = TASK_SATURDAY
End Enum

Enum saTriggerWeeks
    twFirst = TASK_FIRST_WEEK
    twLast = TASK_LAST_WEEK
    twSecond = TASK_SECOND_WEEK
    twThird = TASK_THIRD_WEEK
    twFourth = TASK_FOURTH_WEEK
End Enum

Enum saTriggerFlags
    tfTriggerDisabled = TASK_TRIGGER_FLAG_DISABLED
    tfHasEndDate = TASK_TRIGGER_FLAG_HAS_END_DATE
    tfKillAtDurationEnd = TASK_TRIGGER_FLAG_KILL_AT_DURATION_END
End Enum
'Default Property Values:
Const m_def_xxx = 0
Const m_def_Count = 0
'Const m_def_xxx = 0
'Property Variables:
Dim m_xxx As Variant
Dim m_Count As Long
'Dim m_xxx As Variant



'*********************************************************************************************
' CreateTask: Creates a new Task object
'*********************************************************************************************
Public Function CreateTask(ByVal Name As String) As Job
Dim Jb As Job, IJb As Task
Dim cls As IID, riid As IID

    ' Create a new Task class object
    Set Jb = New Job
    
    With riid
        ' IID of ITask
        ' 148BD524-A2AB-11CE-B11F-00AA00530503
        .Data1 = &H148BD524
        .Data2 = &HA2AB
        .Data3 = &H11CE
        .Data4(0) = &HB1
        .Data4(1) = &H1F
        .Data4(3) = &HAA
        .Data4(5) = &H53
        .Data4(6) = &H5
        .Data4(7) = &H3
    End With
    
    With cls
        ' CLSID of Task
        ' 148BD520-A2AB-11CE-B11F-00AA00530503
        .Data1 = &H148BD520
        .Data2 = &HA2AB
        .Data3 = &H11CE
        .Data4(0) = &HB1
        .Data4(1) = &H1F
        .Data4(3) = &HAA
        .Data4(5) = &H53
        .Data4(6) = &H5
        .Data4(7) = &H3
    End With

    ' Create the new Task object
    Set IJb = m_SA.NewWorkItem(Name, cls, riid)

    ' Pass the Task object to
    ' the Job object
    Set Jb.TaskObject = IJb
    
    ' Set the default creator name
    Jb.Creator = App.Title
    
    ' Save the task to disk
    Jb.Save
    
    ' Add this task to the Jobs collection
    m_Jobs.Add Jb
    
    ' Return the created task object
    Set CreateTask = Jb
    
End Function

'*********************************************************************************************
' NewEnum: Returns a collection enumerator for "For Each...Next" statement.
'*********************************************************************************************
Public Function NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
        
    Set NewEnum = m_Jobs.[_NewEnum]
    
End Function

'*********************************************************************************************
' Refresh: Updates the jobs collection
'*********************************************************************************************
Public Sub Refresh()
Dim IEnum As IEnumWorkItems, Cnt As Long
Dim Ptr As Long, Ptr2 As Long, Jb As Job
Dim FileName As String, Tsk As Task
Dim riid As IID, I As Long
    
    On Error Resume Next
    
    ' Create a new Jobs collection
    ' and destroy the previous one
    Set m_Jobs = New Collection
    
    ' Get the jobs enumerator class
    Set IEnum = m_SA.Enum
    
    With riid
        ' 148BD524-A2AB-11CE-B11F-00AA00530503
        .Data1 = &H148BD524
        .Data2 = &HA2AB
        .Data3 = &H11CE
        .Data4(0) = &HB1
        .Data4(1) = &H1F
        .Data4(3) = &HAA
        .Data4(5) = &H53
        .Data4(6) = &H5
        .Data4(7) = &H3
    End With
        
    I = 0
    
    ' Get all the jobs
    Do While IEnum.Next(, Ptr) = 1
        
        ' Get the string pointer
        ' from the array
        MoveMemory Ptr2, ByVal Ptr, 4
    
        If Ptr2 <> 0 Then
            
            ' Get the job filename
            FileName = StrFromPtrW(Ptr2)
            
            ' Create a new Job object
            Set Jb = New Job
        
            ' Obtain a Task object
            ' from the file
            Set Tsk = m_SA.Activate(FileName, riid)
        
            ' Set the TaskObject
            Set Jb.TaskObject = Tsk
            
            ' Update the Task object name and
            ' filename properties
            Jb.Name = Mid$(FileName, 1, Len(FileName) - 4)
            Jb.FileName = FileName
            
            ' Add the task to the jobs collection
            m_Jobs.Add Jb
            
            Set Jb = Nothing
            
        End If

        ' Free the memory used
        ' by the array
        CoTaskMemFree Ptr

    Loop
        
End Sub

'*********************************************************************************************
' Delete: Removes a task from the Scheduler
'*********************************************************************************************
Public Sub Delete(ByVal Index As Variant)

    ' Remove the scheduler item
    m_SA.Delete m_Jobs(Index).Name & ".job"
    
    ' Remove the item from Jobs collection
    m_Jobs.Remove Index
    
End Sub
'
''*********************************************************************************************
'' Count: Returns the Jobs count
''*********************************************************************************************
'Public Property Get Count() As Long
'
'    Count = m_Jobs.Count
'
'End Property

'*********************************************************************************************
' Task: Returns a Task object given its index/key
'*********************************************************************************************
Public Property Get Job(ByVal Index As Variant) As Job
Attribute Job.VB_UserMemId = 0
Attribute Job.VB_MemberFlags = "200"

    Set Job = m_Jobs(Index)
    
End Property

'*********************************************************************************************
' TargetComputer: Returns/Sets the target computer where the task will run
'*********************************************************************************************
Public Property Get TargetComputer() As String
    
    TargetComputer = StrFromPtrW(m_SA.GetTargetComputer)
    
End Property

'*********************************************************************************************
' TargetComputer: Returns/Sets the target computer where the task will run
'*********************************************************************************************
Public Property Let TargetComputer(ByVal New_Computer As String)
    
    If Ambient.UserMode = False Then Err.Raise 394
    
    m_SA.SetTargetComputer New_Computer
    
End Property

Private Sub UserControl_Initialize()
    
    ' Create a SchedulingAgent object
    Set m_SA = New SchedulingAgent
        
End Sub


Private Sub UserControl_Paint()
Dim R As RECT

    R.Bottom = ScaleHeight
    R.Right = ScaleWidth
    
    DrawEdge UserControl.hdc, R, 5, &HF
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    If Ambient.UserMode Then
        
        Set m_Jobs = New Collection

        Refresh
        
    End If

    m_Count = PropBag.ReadProperty("Count", m_def_Count)
'    m_xxx = PropBag.ReadProperty("xxx", m_def_xxx)
    m_xxx = PropBag.ReadProperty("xxx", m_def_xxx)
End Sub

Private Sub UserControl_Resize()
Static OnResize As Boolean

    If Not OnResize Then
        OnResize = True
        
        Width = 41 * Screen.TwipsPerPixelX
        Height = 41 * Screen.TwipsPerPixelY
        
        OnResize = False
    End If
    
End Sub


Public Property Get Count() As Long
    Count = m_Count
End Property

