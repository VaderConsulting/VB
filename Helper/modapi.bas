Attribute VB_Name = "modAPI"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'---------------------------------------------------------------------------------------
' Procedure : CreateGUID
' DateTime  : 12/03/2003 22:52
' Author    : Dave Robinson
' Purpose   : Used to create a GUID for filenames
'
'  V    History            Author
' 1.0   Initial Version    Dave Robinson
'---------------------------------------------------------------------------------------
'
Public Function CreateGUID() As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    
    On Error GoTo CreateGUIDError
    
    lLen = 40
    bGUID = String(lLen, 0)
    
    CoCreateGuid uGUID
    
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    
    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then
        RetVal = RetVal - 1
    End If
    
    sGUID = Replace(sGUID, "{", "")
    sGUID = Replace(sGUID, "}", "")
    
    CreateGUID = Left$(sGUID, RetVal)
    CreateGUID = Replace(CreateGUID, Chr(0), "")
    Exit Function
CreateGUIDError:
    Err.Clear
End Function

