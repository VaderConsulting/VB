Attribute VB_Name = "Module1"
Dim lpPrevWndProc As Long
Dim gHW As Long

Const GWL_WNDPROC = -4
Const WM_DEVICECHANGE = 537
Const DBT_DEVICEARRIVAL = 32768
Const DBT_DEVICEREMOVECOMPLETE = 32772

Private Type DEV_BROADCAST_HDR
    DBCH_Size As Long
    DBCH_DeviceType As DBT_DEVTYPE
    DBCH_Reserved As Long
End Type

Private Type DEV_BROADCAST_VOLUME
    DBCV_Size As Long
    DBCV_DeviceType As Long
    DBCV_Reserved As Long
    DBCV_UnitMask As Long
    DBCV_Flags As Integer
End Type

Private Enum DEV_BROADCAST_FLAGS
    DBTF_Other = 0      'Other Devices, including USB
    DBTF_Media = 1      'media comings and goings
    DBTF_Net = 2        'network volume
End Enum

Private Enum DBT_DEVTYPE
    DBT_DEVTYP_OEM = &H0                         ' oem-defined device type
    DBT_DEVTYP_DEVNODE = &H1                     ' devnode number
    DBT_DEVTYP_VOLUME = &H2                      ' logical volume
    DBT_DEVTYP_PORT = &H3                        ' serial, parallel
    DBT_DEVTYP_NET = &H4                         ' network resource
    DBT_DEVTYP_DEVICEINTERFACE = &H5             ' device interface class
    DBT_DEVTYP_HANDLE = &H6                      ' file system handle
End Enum

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
                       
Public Sub Hook(frm As Form)
    gHW = frm.hWnd
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHook()
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/windowprocedures/windowprocedurereference/windowprocedurefunctions/windowproc.asp
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Header As DEV_BROADCAST_HDR
    Dim Volume As DEV_BROADCAST_VOLUME
    Dim DriveLetter As String
    Dim sOut As String
    
    If uMsg = WM_DEVICECHANGE Then ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/devio/base/wm_devicechange.asp
        If (wParam = DBT_DEVICEARRIVAL) Or (wParam = DBT_DEVICEREMOVECOMPLETE) Or (wParam = DBT_CONFIGCHANGED) Or (wParam = DBT_DEVNODES_CHANGED) Then
            CopyMemory Header, ByVal lParam, Len(Header) ' Get the BASIC details of the event by coping the parameters to a type.
            Select Case Header.DBCH_DeviceType
                Case DBT_DEVTYP_VOLUME
                    CopyMemory Volume, ByVal (lParam), Len(Volume) ' Get more details of the event by coping the parameters to a type.
                    DriveLetter = Chr$(65 + (Log(Volume.DBCV_UnitMask) / Log(2)))
                    Select Case wParam
                        ' Inserted, Attached or Mapped
                        Case DBT_DEVICEARRIVAL
                            Select Case Volume.DBCV_Flags
                                Case DBTF_Media
                                    sOut = "Media inserted into drive " & DriveLetter
                                Case DBTF_Other
                                    sOut = "USB Device attached as " & DriveLetter
                                Case DBTF_Net
                                    sOut = "Network drive mapped as " & DriveLetter
                            End Select
                        ' Removed, detached or deleted
                        Case DBT_DEVICEREMOVECOMPLETE
                            Select Case Volume.DBCV_Flags
                                Case DBTF_Media
                                    sOut = "Media ejected from drive " & DriveLetter
                                Case DBTF_Other
                                    sOut = "USB device removed (" & DriveLetter & ")"
                                Case DBTF_Net
                                    sOut = "Network drive " & DriveLetter & " removed"
                            End Select
                        ' Dock or undock.  This code has not been tested.
                        Case DBT_CONFIGCHANGED
                            sOut = "Configuration change due to dock or undock"
                        ' Changes due to another device not mentioned above.  This code has not been tested.
                        Case DBT_DEVNODES_CHANGED
                            sOut = "Volume added to or removed from the system."
                    End Select
                Case Else
                    ' Not a volume.  May be an interface, handle, port, or another OEM device.
                    ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/devio/base/dev_broadcast_hdr_str.asp
                    sOut = "Non USB/Network/Media device, port, resource, interface or file system handle " & vbCrLf & "has been added to or removed from the system."
            End Select
            MsgBox sOut, vbOKOnly + vbInformation
        End If
    End If
    
    
    ' Now that we have processed the message, pass it to the window in its original format.
    WindowProc = CallWindowProc(lpPrevWndProc, hWnd, uMsg, wParam, lParam)
End Function


