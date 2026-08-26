Attribute VB_Name = "mdlSchedule"
'*********************************************************************************************
'
' Schedule Control
'
' API declaration & support functions module
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

Public Const STR_NOTPROP = "Invalid trigger type"
Public Function StrFromPtrW(lpszA As Long) As String
    
   ' Create a BSTR from the pointer
   StrFromPtrW = SysAllocString(lpszA)
   
   ' Release the pointed memory
   CoTaskMemFree lpszA
    
End Function


