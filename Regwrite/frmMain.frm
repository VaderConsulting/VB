VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegWrite"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' DateTime  : 15-04-2003 10:45
' Author    : Dave Robinson
' Purpose   : Insert registry keys/strings/values into the current machine's registry
'
'  V    Date        Author          History
' 1.0   15-04-2003  Dave Robinson   Initial Version
' 1.1   15-04-2003  Dave Robinson   Added DOS ErrorLevel
' 1.2   20-07-2003  Dave Robinson   Added REG_EXPAND_SZ capability, cleaned up and sorted.
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub Form_Load()

  Dim strCommandlineArguments() As String
  Dim strDateNow As String
  Dim strTimeNow As String
  Dim strRegType As String
  Dim intRegType As Integer
  Dim strRegRoot As String
  Dim lngRegRoot As Long
  Dim strRegKey As String
  Dim strRegValue As String
  Dim strRegData As String
  Dim intParameterCount As Integer
  Dim intLoop As Integer

    On Error GoTo Load_Error
    strDateNow = FormatDateTime(Date, vbShortDate)
    strTimeNow = FormatDateTime(Time, vbShortTime)

    strDateNow = Replace(strDateNow, "/", "-")

    intParameterCount = GetParamCount
    For intLoop = 1 To intParameterCount
        ReDim Preserve strCommandlineArguments(intLoop)
        strCommandlineArguments(intLoop) = GetParam(intLoop)

        Select Case LCase$(GetParamName(strCommandlineArguments(intLoop)))
          Case "root"
            strRegRoot = GetparamValue(strCommandlineArguments(intLoop))
            lngRegRoot = GetRootValue(strRegRoot)
          Case "key"
            strRegKey = GetparamValue(strCommandlineArguments(intLoop))
          Case "value"
            strRegValue = GetparamValue(strCommandlineArguments(intLoop))
            strRegValue = Replace(LCase$(strRegValue), "(default)", "", 1, -1, vbTextCompare)
            strRegValue = Replace(LCase$(strRegValue), "default", "", 1, -1, vbTextCompare)
          Case "data"
            strRegData = GetparamValue(strCommandlineArguments(intLoop))
            strRegData = Replace(LCase$(strRegData), "$time$", strTimeNow, 1, -1, vbTextCompare)
            strRegData = Replace(LCase$(strRegData), "$date$", strDateNow, 1, -1, vbTextCompare)
            strRegData = Replace(LCase$(strRegData), "$datetime$", strDateNow & " " & strTimeNow, 1, -1, vbTextCompare)
            strRegData = Replace(LCase$(strRegData), "$computername$", Environ$("COMPUTERNAME"), 1, -1, vbTextCompare)
            strRegData = Replace(LCase$(strRegData), "$username$", Environ$("USERNAME"), 1, -1, vbTextCompare)
            strRegData = Replace(LCase$(strRegData), "/%", "%", 1, -1, vbTextCompare)
          Case "type"
            strRegType = GetparamValue(strCommandlineArguments(intLoop))
            intRegType = GetRegType(strRegType)
        End Select

    Next intLoop

    If strRegType = "" Then
        strRegType = "REG_SZ"
        intRegType = GetRegType(strRegType)
    End If

    If strRegRoot <> "" And strRegKey <> "" And strRegValue <> "" Then
        Select Case intRegType
          Case 1 ' REG_SZ
            SaveRegString lngRegRoot, strRegKey, strRegValue, strRegData
            ExitProcess 0 ' All OK
          Case 2 ' REG_EXPAND_SZ
            SaveRegExpandString lngRegRoot, strRegKey, strRegValue, strRegData
            ExitProcess 0 ' All OK
        End Select
    End If
    
    Exit Sub
Load_Error:
    ' or unsupported registry type
    ExitProcess 1 ' NOT OK

End Sub
