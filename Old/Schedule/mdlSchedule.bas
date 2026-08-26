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

Public Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Public Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function SystemTimeToVariantTime Lib "oleaut32.dll" (lpSystemTime As SYSTEMTIME, vtime As Double) As Long

Public Declare Sub DrawEdge Lib "user32" (ByVal hdc As Long, R As RECT, ByVal edge As Integer, ByVal grfFlags As Integer)

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const STR_NOTPROP = "Invalid trigger type"
Public Function StrFromPtrW(lpszA As Long) As String
    
    StrFromPtrW = Space$(lstrlenW(ByVal lpszA))
    MoveMemory ByVal StrPtr(StrFromPtrW), ByVal lpszA, Len(StrFromPtrW) * 2
    
    CoTaskMemFree lpszA
    
End Function


