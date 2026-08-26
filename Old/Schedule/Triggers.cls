VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Triggers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
' Schedule Control
'
' Triggers collection class
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 05/27/1999
' Updates: 12/21/1999. * When creating a new trigger the
'                        trigger flags are set to 0.
'
'*********************************************************************************************
Option Explicit

Dim m_Collection As Collection
Dim m_Task As Task

'*********************************************************************************************
' Adds a new trigger to the collection
'*********************************************************************************************
Public Function Add() As Trigger
Dim IT As ITaskTrigger, Trgr As Trigger
Dim Idx As Integer

    ' Create the new ITaskTrigger object
    Set IT = m_Task.CreateTrigger(Idx)
    
    ' Create a new Trigger object
    Set Trgr = New Trigger
    
    ' Set the ITaskTrigger and
    ' parent task in the Trigger
    ' object
    Trgr.Init IT, m_Task, Idx
    
    ' Set the trigger flags to 0.
    Trgr.Flags = 0
    
    ' Add the trigger to the collection
    m_Collection.Add Trgr
    
    ' Return the object
    Set Add = Trgr
    
End Function

'*********************************************************************************************
' Returns a new enumerator object
'*********************************************************************************************
Public Function NewEnum() As Variant
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = m_Collection.[_NewEnum]
    
End Function

'*********************************************************************************************
' Deletes a trigger from the collection and task
'*********************************************************************************************
Public Sub Remove(ByVal Index As Integer)

    m_Collection.Remove Index
    m_Task.DeleteTrigger Index - 1
        
End Sub

'*********************************************************************************************
' Gets the trigger count
'*********************************************************************************************
Public Property Get Count() As Long

    Count = m_Collection.Count
    
End Property

'*********************************************************************************************
' Returns a trigger from the collection
'*********************************************************************************************
Public Property Get Item(ByVal Index As Variant) As Trigger
Attribute Item.VB_UserMemId = 0

    Set Item = m_Collection(Index)
    
End Property

'*********************************************************************************************
' Sets the parent Task object, and updates the collection
'*********************************************************************************************
Friend Property Set Task(ByVal Tsk As Task)

    Set m_Task = Tsk
    
    Update
    
End Property

'*********************************************************************************************
' Updates the trigger collection
'*********************************************************************************************
Public Sub Update()
Dim I As Integer, Tgr As Trigger
Dim IT As ITaskTrigger
    
    Set m_Collection = New Collection
    
    I = m_Task.GetTriggerCount()
    
    For I = 0 To I - 1
        
        Set IT = m_Task.GetTrigger(I)
          
        Set Tgr = New Trigger
        
        Tgr.Init IT, m_Task, I
        
        m_Collection.Add Tgr
        
    Next

End Sub

