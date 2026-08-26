VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADM Export"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlFiles 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4560
      Y1              =   980
      Y2              =   980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   4560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblADMPath 
      Caption         =   "Path to ADM File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strOutputFilename As String

Private Sub cmdBrowse_Click()
    Dim strDefault As String
    Dim intOne As Integer
    
    cdlFiles.DialogTitle = "Select .ADM file to extract"
    cdlFiles.Filter = "Policy Templates (*.ADM)|*.adm"
    cdlFiles.InitDir = "c:\"
    cdlFiles.ShowOpen
    txtPath.Text = cdlFiles.FileName
    frmMain.Refresh
    
    intOne = InStrRev(txtPath.Text, "\")
    
    strDefault = "C:\" & Right(txtPath.Text, Len(txtPath.Text) - intOne)
    
    strDefault = Left(strDefault, Len(strDefault) - 4) & ".txt"
    
    strOutputFilename = InputBox("Please enter a filename for the output file", "Input required", strDefault)
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdStart_Click()
    cmdClose.Enabled = False
    If txtPath.Text <> "" Then
        Screen.MousePointer = vbHourglass
        ConvertADM
        Screen.MousePointer = vbDefault
    End If
    cmdClose.Enabled = True
End Sub

Private Sub ConvertADM()
    Dim strFileData As String
    Dim strPolicyName As String
    Dim strCaseSelect As String
    Dim strPart As String
    Dim strValueName As String
    Dim strValueOn As String
    Dim strValueOff As String
    Dim strItem As String
    Dim strHelp As String
    Dim strInfo As String
    Dim strValueType As String
    Dim intOne As Integer
    Dim intTwo As Integer
    Dim intClassLevel As Integer
    Dim intCategoryLevel As Integer
    Dim strClass() As String
    Dim strCategory() As String
    Dim strKeyname As String
    
    Open txtPath.Text For Input As #1
        Do Until EOF(1)
            DoEvents
            Line Input #1, strFileData
            strFileData = Trim(strFileData)
            
            If strFileData <> "" Then
                strFileData = Trim(Replace(strFileData, Chr(9), " "))
                If InStr(1, strFileData, " ") <> 0 Then
                    strCaseSelect = LCase(Left(strFileData, InStr(1, strFileData, " ") - 1))
                Else
                    strCaseSelect = LCase(strFileData)
                End If
                Select Case strCaseSelect
                    Case "class"
                        intClassLevel = intClassLevel + 1
                        ReDim Preserve strClass(intClassLevel)
                        strClass(intClassLevel) = Replace(Trim(Right(strFileData, Len(strFileData) - 5)), Chr(34), "")
                    Case "category"
                        intCategoryLevel = intCategoryLevel + 1
                        ReDim Preserve strCategory(intCategoryLevel)
                        strCategory(intCategoryLevel) = Replace(Trim(Right(strFileData, Len(strFileData) - 9)), Chr(34), "")
                    Case "policy"
                        strPolicyName = Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ") - 1), Chr(34), "")
                    Case "keyname"
                        strKeyname = Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ")), Chr(34), "")
                    Case "part"
                        Select Case LCase(Right(strFileData, 4))
                            Case "text"
                                strHelp = strHelp & " " & Mid(strFileData, 6, Len(strFileData) - 6 - 4)
                                If InStr(1, LCase(strFileData), "edittext") = 0 Then
                                    strValueType = "TEXT"
                                    If InStr(1, LCase(strFileData), "numeric") <> 0 Then
                                        Debug.Print "value"
                                    End If
                                Else
                                    strValueType = "EDITABLE TEXT"
                                    strInfo = strInfo & Mid(strFileData, 6, Len(strFileData) - 8 - 8)
                                End If
                            Case "kbox"
                                strInfo = strInfo & Mid(strFileData, 6, Len(strFileData) - 6 - 8)
                                 strValueType = "CHECKBOX"
                            Case "list"
                                strValueType = "DROPDOWNLIST"
                                strInfo = strInfo & Mid(strFileData, 6, Len(strFileData) - 11 - 8)
                            Case Else
                                strPart = strPart & "," & Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ")), Chr(34), "")
                        End Select
                        strHelp = Replace(strHelp, Chr(34), "")
                        strInfo = Replace(strInfo, Chr(34), "")
                    Case "valuename"
                        strValueName = Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ")), Chr(34), "")
                        strValueName = Replace(strValueName, Chr(34), "")
                    Case "valueon"
                        strValueOn = Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ")), Chr(34), "")
                        strValueOn = Replace(LCase(strValueOn), "numeric", "")
                        strValueOn = Replace(LCase(strValueOn), ";reverse", "")
                    Case "valueoff"
                        strValueOff = Replace(Right(strFileData, Len(strFileData) - InStr(1, strFileData, " ")), Chr(34), "")
                        strValueOff = Replace(LCase(strValueOff), "numeric", "")
                        strValueOff = Replace(LCase(strValueOff), ";reverse", "")
                    Case "name"
                        intOne = InStr(1, strFileData, Chr(34)) + 1
                        intTwo = InStr(intOne, strFileData, Chr(34))
                        If intTwo > 0 Then
                            strItem = strItem & " OR " & Replace(Mid(strFileData, intOne, intTwo - intOne), Chr(34), "")
                        Else
                            intOne = InStr(1, strFileData, " ") + 1
                            intTwo = InStr(intOne, strFileData, " ")
                            If intTwo = 0 Then intTwo = Len(strFileData) + 1
                            strItem = strItem & " OR " & Replace(Mid(strFileData, intOne, intTwo - intOne), Chr(34), "")
                        End If
                        If InStr(1, LCase(strFileData), "numeric") <> 0 Then
                            intOne = InStr(1, LCase(strFileData), "numeric") + 7
                            strItem = strItem & " [" & Trim(Mid(strFileData, intOne, 255)) & "]"
                        End If
                    Case "end"
                        If LCase(strFileData) = "end itemlist" Then
                            'strItem = ""
                        End If
                        If LCase(strFileData) = "end part" Then
                            'strPart = ""
                        End If
                        If LCase(strFileData) = "end policy" Then
                            If Left(strPart, 1) = "," Then strPart = Right(strPart, Len(strPart) - 1)
                            If Left(strItem, 4) = " OR " Then strItem = Right(strItem, Len(strItem) - 4)
                            strHelp = Trim(strHelp)
                            strInfo = Trim(strInfo)
                            LogOutput "Class: " & strClass(intClassLevel)
                            LogOutput "Category: " & strCategory(intCategoryLevel)
                            LogOutput "POLICY:  " & strPolicyName
                            LogOutput "Key: " & strKeyname
                            If strValueOn <> "" Then
                                LogOutput "VALUE: " & strPart & " " & Trim(strValueName) & " = " & Trim(strValueOn) & " OR " & Trim(strValueOff) & " (" & strValueType & ")"
                            Else
                                LogOutput "VALUE: " & strPart & " " & Trim(strValueName) & " = " & Trim(strItem) & " (" & strValueType & ")"
                            End If
                            LogOutput strInfo
                            If strHelp <> "" Then
                                LogOutput "DESCRIPTION: " & strHelp
                            End If
                            LogOutput "--------------------------------------------------------------"
                            ' Cleanup
                            strPolicyName = ""
                            strPart = ""
                            strValueName = ""
                            strValueOn = ""
                            strValueOff = ""
                            strHelp = ""
                            strInfo = ""
                            strItem = ""
                            strValueType = ""
                            strKeyname = ""
                        End If
                        If LCase(strFileData) = "end category" Then
                            intCategoryLevel = intCategoryLevel - 1
                        End If
                        If LCase(strFileData) = "end class" Then
                            intClassLevel = intClassLevel - 1
                        End If
                End Select
            End If
        Loop
    Close 1
End Sub

Private Sub LogOutput(strMessage As String)
    Open strOutputFilename For Append As #2
        Print #2, strMessage
    Close 2
End Sub
