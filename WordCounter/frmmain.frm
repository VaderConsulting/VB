VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Word Counter"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstKix 
      Height          =   2400
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox lstWords 
      Height          =   2400
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSourceDir 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\temp\kix"
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    Dim s As String, Data As String
    Dim i As Integer
    
    If Dir(txtSourceDir & "\wordlist.txt") <> "" Then
        Kill txtSourceDir & "\wordlist.txt"
    End If
    
    txtSourceDir = Trim(txtSourceDir)
    If Right(txtSourceDir, 1) = "\" Then txtSourceDir = Left(txtSourceDir, Len(txtSourceDir) - 1)
    s = Dir(txtSourceDir & "\*.kix", vbNormal)
    Do Until s = ""
        Me.Caption = "Word Counter - " & s
        Open txtSourceDir & "\" & s For Input As #1
            Do Until EOF(1)
                Line Input #1, Data
                GetWords LCase(Data)
            Loop
        Close 1

        Open txtSourceDir & "\wordlist.txt" For Append As #1
            For i = 0 To lstWords.ListCount - 1
                Print #1, lstWords.List(i)
            Next i
        Close 1
        lstWords.Clear
        DoEvents
        s = Dir
    Loop
    Me.Caption = "Word Counter"
End Sub

Sub GetWords(strInput As String)
    Dim SplitString() As String, s2 As Variant, s3 As String, s4 As String, i As Integer
    Dim bMatch As Boolean, strWord As String, strWordCount As String, intCount As Integer
    
    If Trim(strInput) = "" Then Exit Sub
    If Trim(Left(strInput, 1)) = ";" Then Exit Sub
    strInput = Replace(strInput, Chr(9), " ")

    strInput = Trim(strInput)
    SplitString = Split(strInput, " ")
    For Each s2 In SplitString()
        If Left(s2, 1) <> "$" And Left(s2, 1) <> Chr(34) And Left(s2, 1) <> "=" And Left(s2, 1) <> "+" Then
            bMatch = False
            For i = 0 To lstKix.ListCount - 1
                If LCase(lstKix.List(i)) = s2 Then
                    bMatch = True
                    Exit For
                End If
            Next i
            If bMatch = True Then ' S2 was found in lstKix
'                bMatch = False
'                For i = 0 To lstWords.ListCount - 1
'                    If InStr(1, lstWords.List(i), " ") > 0 Then
'                        strWord = Left(lstWords.List(i), InStr(1, lstWords.List(i), " ") - 1)
'                    Else
'                        strWord = lstWords.List(i)
'                    End If
'
'                    If strWord = s2 Then
'                        bMatch = True
'                        Exit For
'                    End If
'                Next i
'                If Not bMatch Then
'                    lstWords.AddItem s2
'                Else
'                    For i = 0 To lstWords.ListCount - 1
'                        s3 = lstWords.List(i)
'                        If Left(s3, Len(s2)) = s2 And InStr(1, s2, "(") > 0 Then
'                            strWordCount = Replace(s3, s2, "")
'                            strWordCount = Replace(strWordCount, "(", "")
'                            strWordCount = Replace(strWordCount, ")", "")
'                            lstWords.RemoveItem i
'                            Exit For
'                        End If
'                    Next
'
'                    If strWordCount <> "" Then
'                        intCount = CInt(strWordCount) + 1
'                    Else
'                        intCount = 2
'                    End If
                    lstWords.AddItem s2 '& " (" & intCount & ")"
                'End If
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
    lstKix.Clear
    'lstKix.AddItem "BEEP"
    'lstKix.AddItem "BIG"
    lstKix.AddItem "BREAK"
    lstKix.AddItem "CALL"
    'lstKix.AddItem "CASE"
    lstKix.AddItem "CD"
    'lstKix.AddItem "CLS"
    'lstKix.AddItem "COLOR"
    lstKix.AddItem "COOKIE1"
    lstKix.AddItem "COPY"
    lstKix.AddItem "DEBUG"
    lstKix.AddItem "DEL"
    'lstKix.AddItem "DIM"
    lstKix.AddItem "DISPLAY"
    'lstKix.AddItem "DO"
    lstKix.AddItem "EACH"
    'lstKix.AddItem "ELSE"
    'lstKix.AddItem "ENDFUNCTION"
    'lstKix.AddItem "ENDIF"
    'lstKix.AddItem "ENDSELECT"
    lstKix.AddItem "EXIT"
    lstKix.AddItem "FILE"
    lstKix.AddItem "FLUSHKB"
    'lstKix.AddItem "FOR"
    lstKix.AddItem "FUNCTION"
    lstKix.AddItem "GET"
    lstKix.AddItem "GETS"
    lstKix.AddItem "GLOBAL"
    lstKix.AddItem "GO"
    'lstKix.AddItem "GOSUB"
    'lstKix.AddItem "GOTO"
    'lstKix.AddItem "IF"
    lstKix.AddItem "IN"
    lstKix.AddItem "LIST"
    'lstKix.AddItem "LOOP"
    lstKix.AddItem "MD"
    lstKix.AddItem "MOD"
    'lstKix.AddItem "NEXT"
    lstKix.AddItem "OFF"
    lstKix.AddItem "ON"
    lstKix.AddItem "OPTIONAL"
    lstKix.AddItem "PASSWORD"
    lstKix.AddItem "PLAY"
    lstKix.AddItem "PRESERVE"
    lstKix.AddItem "QUIT"
    lstKix.AddItem "RD"
    'lstKix.AddItem "REDIM"
    'lstKix.AddItem "RETURN"
    lstKix.AddItem "RUN"
    'lstKix.AddItem "SELECT"
    lstKix.AddItem "SET"
    lstKix.AddItem "SETL"
    lstKix.AddItem "SETM"
    lstKix.AddItem "SETTIME"
    lstKix.AddItem "SHELL"
    lstKix.AddItem "SLEEP"
    'lstKix.AddItem "SMALL"
    lstKix.AddItem "STEP"
    'lstKix.AddItem "TO"
    'lstKix.AddItem "UNTIL"
    lstKix.AddItem "USE"
    'lstKix.AddItem "WHILE"
    lstKix.AddItem "Abs"
    lstKix.AddItem "AddKey"
    lstKix.AddItem "AddPrinterConnection"
    lstKix.AddItem "AddProgramGroup"
    lstKix.AddItem "AddProgramItem"
    'lstKix.AddItem "Asc"
    lstKix.AddItem "At"
    lstKix.AddItem "BackupEventLog"
    lstKix.AddItem "Box"
    lstKix.AddItem "CDbl"
    'lstKix.AddItem "Chr"
    'lstKix.AddItem "CInt"
    lstKix.AddItem "ClearEventLog"
    lstKix.AddItem "Close"
    lstKix.AddItem "CompareFileTimes"
    lstKix.AddItem "CreateObject"
    'lstKix.AddItem "CStr"
    lstKix.AddItem "DecToHex"
    lstKix.AddItem "DelKey"
    lstKix.AddItem "DelPrinterConnection"
    lstKix.AddItem "DelProgramGroup"
    lstKix.AddItem "DelProgramItem"
    lstKix.AddItem "DelTree"
    lstKix.AddItem "DelValue"
    'lstKix.AddItem "Dir"
    lstKix.AddItem "EnumGroup"
    lstKix.AddItem "EnumIPInfo"
    lstKix.AddItem "EnumKey"
    lstKix.AddItem "EnumLocalGroup"
    lstKix.AddItem "EnumValue"
    lstKix.AddItem "Execute"
    'lstKix.AddItem "Exist"
    lstKix.AddItem "ExistKey"
    lstKix.AddItem "ExpandEnvironmentVars"
    lstKix.AddItem "Fix"
    lstKix.AddItem "FormatNumber"
    lstKix.AddItem "GetDiskSpace"
    lstKix.AddItem "GetFileAttr"
    lstKix.AddItem "GetFileSize"
    lstKix.AddItem "GetFileTime"
    lstKix.AddItem "GetFileVersion"
    lstKix.AddItem "GetObject"
    lstKix.AddItem "InGroup"
    'lstKix.AddItem "InStr"
    'lstKix.AddItem "InStrRev"
    lstKix.AddItem "Int"
    lstKix.AddItem "KbHit"
    lstKix.AddItem "KeyExist"
    'lstKix.AddItem "Lcase"
    'lstKix.AddItem "Left"
    'lstKix.AddItem "Len"
    lstKix.AddItem "LoadHive"
    lstKix.AddItem "LoadKey"
    lstKix.AddItem "LogEvent"
    lstKix.AddItem "LogOff"
    'lstKix.AddItem "Ltrim"
    lstKix.AddItem "MemorySize"
    lstKix.AddItem "MessageBox"
    lstKix.AddItem "Open"
    lstKix.AddItem "ReadLine"
    lstKix.AddItem "ReadProfileString"
    lstKix.AddItem "ReadType"
    lstKix.AddItem "ReadValue"
    lstKix.AddItem "RedirectOutput"
    'lstKix.AddItem "Right"
    lstKix.AddItem "Rnd"
    lstKix.AddItem "Round"
    lstKix.AddItem "Rtrim"
    lstKix.AddItem "SaveKey"
    lstKix.AddItem "SendKeys"
    lstKix.AddItem "SendMessage"
    lstKix.AddItem "SetAscii"
    lstKix.AddItem "SetConsole"
    lstKix.AddItem "SetDefaultPrinter"
    lstKix.AddItem "SetFileAttr"
    lstKix.AddItem "SetFocus"
    lstKix.AddItem "SetOption"
    lstKix.AddItem "SetSystemState"
    lstKix.AddItem "SetTitle"
    lstKix.AddItem "SetWallPaper"
    lstKix.AddItem "ShowProgramGroup"
    lstKix.AddItem "Shutdown"
    lstKix.AddItem "SidToName"
    lstKix.AddItem "Split"
    lstKix.AddItem "Srnd"
    lstKix.AddItem "Substr"
    'lstKix.AddItem "Trim"
    'lstKix.AddItem "UBound"
    'lstKix.AddItem "Ucase"
    lstKix.AddItem "UnloadHive"
    'lstKix.AddItem "Val"
    lstKix.AddItem "VarType"
    lstKix.AddItem "VarTypeName"
    lstKix.AddItem "WriteLine"
    lstKix.AddItem "WriteProfileString"
    lstKix.AddItem "WriteValue"
    lstKix.AddItem "@address"
    lstKix.AddItem "@build"
    lstKix.AddItem "@comment"
    lstKix.AddItem "@cpu"
    lstKix.AddItem "@crlf"
    lstKix.AddItem "@csd"
    lstKix.AddItem "@curdir"
    lstKix.AddItem "@date"
    lstKix.AddItem "@day"
    lstKix.AddItem "@domain"
    lstKix.AddItem "@dos"
    lstKix.AddItem "@error"
    lstKix.AddItem "@fullname"
    lstKix.AddItem "@homedir"
    lstKix.AddItem "@homedrive"
    lstKix.AddItem "@homeshr"
    lstKix.AddItem "@hostname"
    lstKix.AddItem "@inwin"
    lstKix.AddItem "@ipaddress0"
    lstKix.AddItem "@ipaddress1"
    lstKix.AddItem "@ipaddress2"
    lstKix.AddItem "@ipaddress3"
    lstKix.AddItem "@kix"
    lstKix.AddItem "@lanroot"
    lstKix.AddItem "@ldomain"
    lstKix.AddItem "@ldrive"
    lstKix.AddItem "@lm"
    lstKix.AddItem "@logonmode"
    lstKix.AddItem "@longhomedir"
    lstKix.AddItem "@lserver"
    lstKix.AddItem "@maxpwage"
    lstKix.AddItem "@mdayno"
    lstKix.AddItem "@mhz"
    lstKix.AddItem "@month"
    lstKix.AddItem "@monthno"
    lstKix.AddItem "@msecs"
    lstKix.AddItem "@primarygroup"
    lstKix.AddItem "@priv"
    lstKix.AddItem "@productsuite"
    lstKix.AddItem "@producttype"
    lstKix.AddItem "@pwage"
    lstKix.AddItem "@ras"
    lstKix.AddItem "@rserver"
    lstKix.AddItem "@scriptdir"
    lstKix.AddItem "@scriptname"
    lstKix.AddItem "@serror"
    lstKix.AddItem "@sid"
    lstKix.AddItem "@site"
    lstKix.AddItem "@startdir"
    lstKix.AddItem "@syslang"
    lstKix.AddItem "@ticks"
    lstKix.AddItem "@time"
    lstKix.AddItem "@userid"
    lstKix.AddItem "@userlang"
    lstKix.AddItem "@wdayno"
    lstKix.AddItem "@wksta"
    lstKix.AddItem "@wuserid"
    lstKix.AddItem "@ydayno"
    lstKix.AddItem "@year"
End Sub
