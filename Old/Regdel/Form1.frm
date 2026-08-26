VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reg Delete"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Usage:
' Delete a key:     Key,HKCU\Software\VB and VBA Program Settings\SendMail\Setup
' Delete a value  Value,HKCU\Software\VB and VBA Program Settings\SendMail\Setup Test

Private Sub Form_Load()
    Dim Params() As String
    Dim Key As Long
    Dim Action As String
    Dim Root As String, i As Integer, j As Integer
    Dim KeyRoot As Long, KeyName As String, KeyValue As String
    Dim msg As String
    Dim DoError As Boolean
    
    If Command <> "" Then
        If InStr(1, Command, ",") <> 0 Then
            Params() = Split(Command, ",")
            
            If Params(0) <> "" Then
                Action = UCase(Params(0))
                
                i = InStr(1, Params(1), "\")
                If i > 0 Then
                    
                    Root = UCase(Left(Params(1), i - 1))
                    
                    ' Determine appropriate root to use
                    Select Case Root
                        Case "HKCU", "HKEY_CURRENTUSER", "HKEY_CURRENT_USER"
                            KeyRoot = EnumRegistryRootKeys.rrkHKeyCurrentUser
                            
                        Case "HKCR", "HKEY_CLASSESROOT", "HKEY_CLASSES_ROOT"
                            KeyRoot = EnumRegistryRootKeys.rrkHKeyClassesRoot
                            
                        Case "HKLM", "HKEY_LOCALMACHINE", "HKEY_LOCAL_MACHINE"
                            KeyRoot = EnumRegistryRootKeys.rrkHKeyLocalMachine
                            
                        Case "HKU", "HKEY_USERS"
                            KeyRoot = EnumRegistryRootKeys.rrkHKeyUsers
                            
                        Case Else
                            'Unsupported Root
                            DoError = True
                    End Select
                    
                    ' Determine action to perform
                    Select Case Action
                        Case "KEY"
                            i = InStr(1, Params(1), "\")
                            KeyName = Trim(Mid(Params(1), i + 1, Len(Params(1)) - i))
                            RegistryDeleteKey KeyRoot, KeyName
                        Case "VALUE"
                            i = InStr(1, Params(1), "\")
                            j = InStrRev(Params(1), " ")
                            KeyName = Trim(Mid(Params(1), i + 1, j - i))
                            KeyValue = Trim(Mid(Params(1), j + 1, Len(Params(1)) - j))
                            RegistryDeleteValue KeyRoot, KeyName, KeyValue
                        Case Else
                            ' Unsupported action
                            DoError = True
                    End Select
                Else
                    DoError = True
                End If
            Else
                DoError = True
            End If
        Else
            DoError = True
        End If
    Else
        DoError = True
    End If
    
    If DoError Then
        msg = ""
        msg = msg & " Usage:" & vbCrLf
        msg = msg & " Delete a key:       Key,ROOT\KEY" & vbCrLf
        msg = msg & " Delete a value:   Value,ROOT\KEY VALUE" & vbCrLf
        msg = msg & vbCrLf
        msg = msg & " Examples:" & vbCrLf
        msg = msg & "     Key,HKCU\Software\VB and VBA Program Settings\SendMail\Setup" & vbCrLf
        msg = msg & "   Value,HKCU\Software\VB and VBA Program Settings\SendMail\Setup Test" & vbCrLf
        MsgBox msg, vbInformation, "Help"
    End If
    Unload Me
End Sub


