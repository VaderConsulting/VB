VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Function Create(Optional bWithBraces As Boolean = False) As String

  Dim uGUID As GUID
  Dim sGUID As String
  Dim bGUID() As Byte
  Dim lLen As Long
  Dim RetVal As Long

    lLen = 40
    bGUID = String$(lLen, 0)

    CoCreateGuid uGUID

    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)

    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then
        RetVal = RetVal - 1
    End If

    Create = LCase$(Left$(sGUID, RetVal))

    If Not bWithBraces Then
        Create = Replace(Create, "{", "")
        Create = Replace(Create, "}", "")
    End If

End Function

Private Sub Form_Load()

    UpdateTrendGUID False
    End

End Sub

Public Sub UpdateTrendGUID(Optional bForceUpdate As Boolean = False)

  Dim oRegistry As REGTool5.Registry
  Dim strOldGUID As String
  Dim strNewGUID As String

    Set oRegistry = CreateObject("RegTool5.Registry")

    strNewGUID = Create(False)

    oRegistry.GetKeyValue REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\TRENDMICRO\PC-CILLINNTCORP\CurrentVersion", "GUID", strOldGUID

    If strOldGUID = "" Then
        oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\TRENDMICRO\PC-CILLINNTCORP\CurrentVersion", "GUID", strNewGUID
      Else 'NOT STROLDGUID...
        If bForceUpdate Then
            oRegistry.UpdateKey REGTool5.HKEY_LOCAL_MACHINE, "SOFTWARE\TRENDMICRO\PC-CILLINNTCORP\CurrentVersion", "GUID", strNewGUID
        End If
    End If

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2003-Jul-28 12:43) 11 + 60 = 71 Lines
