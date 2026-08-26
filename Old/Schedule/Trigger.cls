VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Trigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Schedule Control
'
' Trigger Class
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

Dim m_Trg As ITaskTrigger
Dim m_Task As Task
Dim m_Idx As Integer
Dim m_Trigger As TASK_TRIGGER

Public Property Get DaysInterval() As Integer

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_DAILY Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    DaysInterval = m_Trigger.Type.vb_compat_union(0)
    
End Property

Public Property Let DaysInterval(ByVal New_DaysInterval As Integer)

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_DAILY Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    m_Trigger.Type.vb_compat_union(0) = New_DaysInterval
    
End Property

Public Property Get Day() As Integer

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_MONTHLYDATE Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    Day = m_Trigger.Type.vb_compat_union(0)
    
End Property

Public Property Let Day(ByVal New_Days As Integer)

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_MONTHLYDATE Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    m_Trigger.Type.vb_compat_union(0) = New_Days
    
End Property

Public Property Get Months() As saTriggerMonths

    Select Case m_Trigger.TriggerType
        Case TASK_TIME_TRIGGER_MONTHLYDATE
            Months = m_Trigger.Type.vb_compat_union(2)
        Case TASK_TIME_TRIGGER_MONTHLYDOW
            Months = m_Trigger.Type.vb_compat_union(1)
        Case Else
            Err.Raise vbObjectError + 1, , STR_NOTPROP
    End Select
    
End Property

Public Property Let Months(ByVal New_Months As saTriggerMonths)

    Select Case m_Trigger.TriggerType
        Case TASK_TIME_TRIGGER_MONTHLYDATE
            m_Trigger.Type.vb_compat_union(2) = New_Months
        Case TASK_TIME_TRIGGER_MONTHLYDOW
            m_Trigger.Type.vb_compat_union(1) = New_Months
        Case Else
            Err.Raise vbObjectError + 1, , STR_NOTPROP
    End Select
    
End Property

Public Property Get DaysOfTheWeek() As saTriggerDays

    Select Case m_Trigger.TriggerType
        Case TASK_TIME_TRIGGER_WEEKLY, TASK_TIME_TRIGGER_MONTHLYDOW
            DaysOfTheWeek = m_Trigger.Type.vb_compat_union(0)
        Case Else
            Err.Raise vbObjectError + 1, , STR_NOTPROP
    End Select
    
End Property

Public Property Let DaysOfTheWeek(ByVal New_DaysOfTheWeek As saTriggerDays)

    Select Case m_Trigger.TriggerType
        Case TASK_TIME_TRIGGER_WEEKLY, TASK_TIME_TRIGGER_MONTHLYDOW
            m_Trigger.Type.vb_compat_union(1) = New_DaysOfTheWeek
        Case Else
            Err.Raise vbObjectError + 1, , STR_NOTPROP
    End Select
    
End Property

Public Property Get Week() As saTriggerWeeks

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_MONTHLYDOW Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    Week = m_Trigger.Type.vb_compat_union(2)
    
End Property

Public Property Let Week(ByVal New_Week As saTriggerWeeks)

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_MONTHLYDOW Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    m_Trigger.Type.vb_compat_union(2) = New_Week
    
End Property

Public Property Get WeeksInterval() As Integer

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_WEEKLY Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    Week = m_Trigger.Type.vb_compat_union(1)
    
End Property

Public Property Let WeeksInterval(ByVal New_WeeksInterval As Integer)

    If m_Trigger.TriggerType <> TASK_TIME_TRIGGER_WEEKLY Then Err.Raise vbObjectError + 1, , STR_NOTPROP
    
    m_Trigger.Type.vb_compat_union(0) = New_WeeksInterval
    
End Property

Friend Sub Init(ByVal IT As ITaskTrigger, ByVal Tsk As Task, ByVal Idx As Integer)

    Set m_Trg = IT
    Set m_Task = Tsk
    m_Idx = Idx
    
    m_Trigger.cbTriggerSize = Len(m_Trigger)
    m_Trg.GetTrigger m_Trigger
    
End Sub

Public Sub Refresh()
    
    m_Trg.GetTrigger m_Trigger

End Sub

Public Sub Update()

    m_Trg.SetTrigger m_Trigger
    m_Trg.GetTrigger m_Trigger
    
End Sub

Public Property Get Text() As String
    
    Text = StrFromPtrW(m_Task.GetTriggerString(m_Idx))

End Property

Public Property Get TriggerType() As saTriggerTypes
    
    TriggerType = m_Trigger.TriggerType
    
End Property

Public Property Let TriggerType(ByVal New_TriggerType As saTriggerTypes)
    
    m_Trigger.TriggerType = New_TriggerType
    
End Property

Private Sub Class_Terminate()

    Set m_Trg = Nothing
    Set m_Task = Nothing

End Sub

Public Property Get Duration() As Long

    Duration = m_Trigger.MinutesDuration
    
End Property


Public Property Let Duration(ByVal New_MinutesDuration As Long)

    m_Trigger.MinutesDuration = New_MinutesDuration
    
End Property

Public Property Get Interval() As Long

    Duration = m_Trigger.MinutesInterval
    
End Property

Public Property Let Interval(ByVal New_MinutesInterval As Long)

    m_Trigger.MinutesInterval = New_MinutesInterval
    
End Property

Public Property Get Flags() As saTriggerFlags

    Duration = m_Trigger.rgFlags
    
End Property

Public Property Let Flags(ByVal New_Flags As saTriggerFlags)

    m_Trigger.rgFlags = New_Flags
    
End Property

Public Property Get BeginDay() As Date

    BeginDay = DateSerial(m_Trigger.wBeginYear, m_Trigger.wBeginMonth, m_Trigger.wBeginDay)
    
End Property

Public Property Let BeginDay(ByVal New_BeginDay As Date)

    m_Trigger.wBeginDay = VBA.Day(New_BeginDay)
    m_Trigger.wBeginMonth = VBA.Month(New_BeginDay)
    m_Trigger.wBeginYear = VBA.Year(New_BeginDay)
    
End Property

Public Property Get EndDay() As Date

    EndDay = DateSerial(m_Trigger.wEndYear, m_Trigger.wEndMonth, m_Trigger.wEndDay)
    
End Property

Public Property Let EndDay(ByVal New_EndDay As Date)

    m_Trigger.wEndDay = VBA.Day(New_EndDay)
    m_Trigger.wEndMonth = VBA.Month(New_EndDay)
    m_Trigger.wEndYear = VBA.Year(New_EndDay)
    
End Property

Public Property Get RandomInterval() As Integer

    RandomInterval = m_Trigger.wRandomMinutesInterval
    
End Property

Public Property Let RandomInterval(ByVal New_RandomInterval As Integer)

    m_Trigger.wRandomMinutesInterval = New_RandomInterval
    
End Property

Public Property Get StartTime() As Date

    StartTime = TimeSerial(m_Trigger.wStartHour, m_Trigger.wStartMinute, 0)
    
End Property

Public Property Let StartTime(ByVal New_StartTime As Date)

    m_Trigger.wStartHour = VBA.Hour(New_StartTime)
    m_Trigger.wStartMinute = VBA.Minute(New_StartTime)
    
End Property

