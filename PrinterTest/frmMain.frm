VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Task4()
    Dim oPrinter As Printer, strDevicename As String
    Dim strCurrentPrintServername As String, strCurrentPrintername As String    ' Current
    Dim strOriginalServername As String, strOriginalPrintername As String       ' Original as specified to search for
    Dim strReplacementServername As String, strReplacementPrintername As String ' Replacement as specified
    Dim intPrinterCount As Integer
    Dim strDefaultPrinterServername As String, strDefaultPrintername As String  ' Current DEFAULTS
    Dim strNewDefaultPrintServer As String, strNewDefaultPrinter As String      ' New defaults
    Dim strNewPrintServername As String, strNewPrintername As String            ' Final that should be mapped
    Dim oPrinterList() As vbPrinter
    Dim lPrinterNo As Long
    Dim bReplacePrinter As Boolean
    Dim bRemovePrinter As Boolean
    
    On Error GoTo TaskError
    
    LogEvent "Inside Task4 subroutine", DebugInfo
    
    UpdateStatus 5
    
    OpenPrinterList oPrinterList()
    
    LogEvent UBound(oPrinterList()) & " printer(s) found in substitute list", DebugInfo
    LogEvent "---------------End Substitute list---------------", DebugInfo
    
    LogEvent "Printers defined on this system are as follows", DebugInfo
    LogEvent "---------------Start Printer list---------------", DebugInfo
    For Each oPrinter In Printers
        intPrinterCount = intPrinterCount + 1
        strDevicename = oPrinter.DeviceName
        strCurrentPrintServername = GetServerFromPath(oPrinter.DeviceName) ' From system
        strCurrentPrintername = GetPathfromUNC(strDevicename)              ' From system
        strNewPrintServername = ""
        strNewPrintername = ""
        bRemovePrinter = False
        
        If strCurrentPrintServername <> "" Then
            LogEvent "Installed Printer: " & oPrinter.DeviceName, DebugInfo
            
            Debug.Print "Substitute Printer Info follows:"
            
            For lPrinterNo = 1 To UBound(oPrinterList())
                strOriginalServername = ""
                strOriginalPrintername = ""
                strReplacementServername = ""
                strReplacementPrintername = ""
                strNewPrintServername = ""
                
                strOriginalServername = oPrinterList(lPrinterNo).OldServer
                strOriginalPrintername = oPrinterList(lPrinterNo).OldQueue
                strReplacementServername = oPrinterList(lPrinterNo).NewServer
                strReplacementPrintername = oPrinterList(lPrinterNo).NewQueue
                
                'Debug.Print "Number: " & lPrinterNo
                'Debug.Print "Old Server: " & strOriginalServername
                'Debug.Print "Old Queue : " & strOriginalPrintername
                'Debug.Print "New Server: " & strReplacementServername
                'Debug.Print "New Queue : " & strReplacementPrintername
                'Debug.Print
                
                bReplacePrinter = False
                bRemovePrinter = False
                
                Select Case LCase(strOriginalServername)
                    Case "*"
                        Select Case LCase(strOriginalPrintername)
                            Case "*"
                                LogEvent "\\*\* = ..."
                                Select Case LCase(strReplacementServername)
                                    Case "*"
                                        LogEvent "\\*\* = \\*\ ..."
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 1 - replace all printers on all servers with the same printer on the same server", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 2 - replace all printers on all servers with the same printer on the same server", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 3 - remove all printers on all servers", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 4 - replace all printers on all servers with this printer on the same server", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                    Case ""
                                        LogEvent "\\*\* = \\\ ..."
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 5 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 6 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case ""
                                                LogEvent "Case 7 - remove all printers on all servers", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 8 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                        End Select
                                    Case LCase(strCurrentPrintServername)
                                        LogEvent "\\*\* = \\[Current Server]\ ..."
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 9 - replace all printers on all servers with the same printer on this server", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 10 - replace all printers on all servers with the same printer on this server", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 11 - Remove this printer on this server.", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 12 - replace all printers on all servers with this printer on this server", DebugInfo
                                                strNewPrintServername = strReplacementPrintername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                    Case Else
                                        LogEvent "\\*\* = \\[Specified Server]\ ..."
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 13 - replace all printers on all servers with the same printer on the specified server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 14 - replace all printers on all servers with the same printer on the same server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 15 - replace all printers on all servers with no printer on this server", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 16 - replace all printers on all servers with this printer on this server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                End Select
                            Case LCase(strCurrentPrintername) ' Original printer name = current printer name
                                LogEvent "\\*\[currentprinter] = ...", DebugInfo
                                Select Case LCase(strReplacementServername)
                                    Case "*"
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 17 - replace the current printer name on any server with the current printer name on the specified server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 18 - replace the current printer name on any server with the current printer name on the specified server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 19 - remove the current printer on the specified server", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 20 - replace the current printer on the specified server with the specified printer on the specified server", DebugInfo
                                                strNewPrintServername = strOriginalServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                    Case ""
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 21 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 22 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case ""
                                                LogEvent "Case 23 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 24 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                        End Select
                                    Case LCase(strCurrentPrintServername)
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 25", DebugInfo
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 26", DebugInfo
                                            Case ""
                                                LogEvent "Case 27 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 28", DebugInfo
                                        End Select
                                    Case Else
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 29 - replace the specified printer on any server with the same printer on the specified server", DebugInfo
                                                
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 30", DebugInfo
                                            Case ""
                                                LogEvent "Case 31 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 32", DebugInfo
                                        End Select
                                End Select
                            Case Else
                                LogEvent "The specified replacement printer does not match this installed printer", DebugInfo
                        End Select
                    Case LCase(strCurrentPrintServername)
                        Select Case LCase(strOriginalPrintername)
                            Case "*"
                                LogEvent "This Server, any printer"
                                Select Case LCase(strReplacementServername)
                                    Case "*"
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 33 - replace ", DebugInfo
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 34", DebugInfo
                                            Case ""
                                                LogEvent "Case 35 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 36", DebugInfo
                                        End Select
                                    Case ""
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 37 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 38 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case ""
                                                LogEvent "Case 39 - remove all printers on this server", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 40 - Do nothing.  The replacement print server is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                        End Select
                                    Case LCase(strCurrentPrintServername)
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 41", DebugInfo
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 42", DebugInfo
                                            Case ""
                                                LogEvent "Case 43 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 44", DebugInfo
                                        End Select
                                    Case Else
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 45 - replace any printer on the specified server with the same printer name on the specified server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 46", DebugInfo
                                            Case ""
                                                LogEvent "Case 47 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 48 - replace any printer on this server with the specified printer on the specified server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                End Select
                            Case LCase(strCurrentPrintername)
                                Select Case LCase(strReplacementServername)
                                    Case "*"
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 49", DebugInfo
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 50", DebugInfo
                                            Case ""
                                                LogEvent "Case 51 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 52", DebugInfo
                                        End Select
                                    Case ""
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 53 - Do nothing.  The printer to replace is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 54 - Do nothing.  The printer to replace is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case ""
                                                LogEvent "Case 55 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 56 - Do nothing.  The printer to replace is not specified", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = False
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = False
                                                oPrinterList(lPrinterNo).bReplace = False
                                        End Select
                                    Case LCase(strCurrentPrintServername)
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 57", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 58", DebugInfo
                                                strNewPrintServername = strCurrentPrintServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 59 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 60 - replace the specified printer on the specified server with the specified printer on the same server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                    Case Else
                                        Select Case LCase(strReplacementPrintername)
                                            Case "*"
                                                LogEvent "Case 61 - replace the specified printer on the specified server with the same printer on the specified server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strCurrentPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case LCase(strCurrentPrintername)
                                                LogEvent "Case 62 - replace the specified printer on the specified server with the same printer on the specified server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                            Case ""
                                                LogEvent "Case 63 - remove this printer", DebugInfo
                                                strNewPrintServername = ""
                                                strNewPrintername = ""
                                                bRemovePrinter = True
                                                bReplacePrinter = False
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = False
                                            Case Else
                                                LogEvent "Case 64 - replace the specified printer on the specified server with the specified printer on the specified server", DebugInfo
                                                strNewPrintServername = strReplacementServername
                                                strNewPrintername = strReplacementPrintername
                                                bRemovePrinter = True
                                                bReplacePrinter = True
                                                oPrinterList(lPrinterNo).bRemove = True
                                                oPrinterList(lPrinterNo).bReplace = True
                                        End Select
                                End Select
                            Case Else
                                'LogEvent "The specified replacement printer does not match this installed printer", DebugInfo
                        End Select
                    Case Else
                        'LogEvent "This specified replacement Server does not match the installed print server", DebugInfo
                End Select
                
                ' New
                ' ****************
            Next
            
            For lPrinterNo = 1 To UBound(oPrinterList())
                If oPrinterList(lPrinterNo).bRemove Then
                    LogEvent "This printer (\\" & oPrinterList(lPrinterNo).OldServer & "\" & oPrinterList(lPrinterNo).OldQueue & ") is to be removed", DebugInfo
                    ' Check if we are removing the default printer
                    If LCase(GetServerFromPath(Printer.DeviceName)) = LCase(oPrinterList(lPrinterNo).OldServer) And GetPathfromUNC(Printer.DeviceName) = oPrinterList(lPrinterNo).OldQueue Then
                        LogEvent "The default printer (" & Printer.DeviceName & ") is to be removed", Information
                        If oPrinterList(lPrinterNo).NewServer <> "" And oPrinterList(lPrinterNo).NewQueue <> "" Then
                            LogEvent "The new default printer will be \\" & oPrinterList(lPrinterNo).NewServer & "\" & oPrinterList(lPrinterNo).NewQueue, Information
                            strNewDefaultPrintServer = oPrinterList(lPrinterNo).NewServer
                            strNewDefaultPrinter = oPrinterList(lPrinterNo).NewQueue
                        Else
                            'LogEvent "There is no printer specified to replace the default printer.", Warning
                        End If
                    End If
                    LogEvent "About to delete printer connection " & oPrinterList(lPrinterNo).OldServer & "\" & oPrinterList(lPrinterNo).OldQueue & ".", Information
                    'DeletePrinterConnection "\\" & oPrinterList(lPrinterNo).OldServer & "\" & oPrinterList(lPrinterNo).OldQueue & Chr(0) ' NOTE the use of the null character
                End If
                
                If oPrinterList(lPrinterNo).bReplace Then
                    LogEvent "This printer (\\" & oPrinterList(lPrinterNo).OldServer & "\" & oPrinterList(lPrinterNo).OldQueue & ") is to be replaced", DebugInfo
                    LogEvent "New printer: \\" & oPrinterList(lPrinterNo).NewServer & "\" & oPrinterList(lPrinterNo).NewQueue, DebugInfo
                    LogEvent "About to add printer connection " & oPrinterList(lPrinterNo).NewServer & "\" & oPrinterList(lPrinterNo).NewQueue & ".", Information
                    'AddPrinterConnection "\\" & strNewPrintServername & "\" & strNewPrintername
                    ' Is this the new default printer?
                    If LCase(oPrinterList(lPrinterNo).NewServer) = LCase(strNewDefaultPrintServer) And (oPrinterList(lPrinterNo).NewQueue = strNewDefaultPrinter) Then
                        NTSetDefaultPrinter "\\" & oPrinterList(lPrinterNo).NewServer & "\" & oPrinterList(lPrinterNo).NewQueue
                    End If
                Else
                    'LogEvent "This printer (\\" & oPrinterList(lPrinterNo).OldServer & "\" & oPrinterList(lPrinterNo).OldQueue & ") is NOT to be replaced", DebugInfo
                End If
                
                ' ****************
            Next
'            If bRemovePrinter = True Then
'                LogEvent "This printer (\\" & strCurrentPrintServername & "\" & strCurrentPrintername & ") is to be removed", DebugInfo
'                ' Check if we are removing the default printer
'                If LCase(GetServerFromPath(Printer.DeviceName)) = LCase(strCurrentPrintServername) And GetPathfromUNC(Printer.DeviceName) = strCurrentPrintername Then
'                    LogEvent "The default printer (" & Printer.DeviceName & ") is to be removed", Information
'                    If strNewPrintServername <> "" And strNewPrintername <> "" Then
'                        LogEvent "The new default printer will be \\" & strNewPrintServername & "\" & strNewPrintername, Information
'                        strNewDefaultPrintServer = strNewPrintServername
'                        strNewDefaultPrinter = strNewPrintername
'                    Else
'                        LogEvent "There is no printer specified to replace the default printer.", Warning
'                    End If
'                End If
'                LogEvent "About to delete printer connection " & strCurrentPrintServername & "\" & strCurrentPrintername & ".", Information
'                'DeletePrinterConnection "\\" & strCurrentPrintServername & "\" & strCurrentPrintername & Chr(0) ' NOTE the use of the null character
'            End If
            
'            If bReplacePrinter = True Then
'                LogEvent "This printer (\\" & strCurrentPrintServername & "\" & strCurrentPrintername & ") is to be replaced", DebugInfo
'                LogEvent "New printer: \\" & strNewPrintServername & "\" & strNewPrintername, DebugInfo
'                LogEvent "About to add printer connection " & strNewPrintServername & "\" & strNewPrintername & ".", Information
'                'AddPrinterConnection "\\" & strNewPrintServername & "\" & strNewPrintername
'                ' Is this the new default printer?
'                If LCase(strNewPrintServername) = LCase(strNewDefaultPrintServer) And (strNewPrintername = strNewDefaultPrinter) Then
'                    NTSetDefaultPrinter "\\" & strNewDefaultPrintServer & "\" & strNewPrintername
'                End If
'            Else
'                LogEvent "This printer (\\" & strCurrentPrintServername & "\" & strCurrentPrintername & ") is NOT to be replaced", DebugInfo
'            End If
         End If
    Next
    LogEvent "----------------End Printer list----------------", DebugInfo
    If intPrinterCount > 0 Then
        strDefaultPrinterServername = GetServerFromPath(Printer.DeviceName)
        strDefaultPrintername = GetPathfromUNC(Printer.DeviceName)
        LogEvent "Default Printer info- Server: " & strDefaultPrinterServername & " Name: " & strDefaultPrintername
    Else
        LogEvent "There are no printers defined on this system.", Information
    End If
    Exit Sub
TaskError:
    LogEvent "**** Logon encountered internal error " & Err.Number & " (" & Err.Description & ") in Task4.", Error
    Err.Clear
End Sub

Private Sub Form_Load()
    Set oDictionary = CreateObject("Scripting.Dictionary")
        
    strDomainname = Environ$("USERDOMAIN")
    strUsername = Environ$("USERNAME")
    strComputername = Environ$("COMPUTERNAME")
    
    ' Generate an App_Path variable that is the path to this application, similar to App.Path, though it will always end with a backslash.
    If Right(App.Path, 1) <> "\" Then
        strApp_Path = App.Path & "\"
    Else
        strApp_Path = App.Path
    End If
    
    
    Me.Show
    Me.Refresh
    
    strAppStatus = ApplicationStatus.statusOK
    
    Task4
End Sub
