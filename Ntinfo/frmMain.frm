VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Info"
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
   Begin VB.Label lblInfo 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Const PROCESSOR_INTEL_386 = 386
Const PROCESSOR_INTEL_486 = 486
Const PROCESSOR_INTEL_PENTIUM = 586
Const PROCESSOR_MIPS_R4000 = 4000
Const PROCESSOR_ALPHA_21064 = 21064

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Sub Form_Load()
    Dim msg As String         ' Status information.
    Dim nl As String          ' New-line.
    Dim ret%
    Dim ver_major$
    nl = Chr$(13) + Chr$(10)  ' New-line.

    Show
    MousePointer = 11   ' Hourglass.


        ' Get operating system and version.
        Dim verinfo As OSVERSIONINFO
        verinfo.dwOSVersionInfoSize = Len(verinfo)
        ret% = GetVersionEx(verinfo)
        If ret% = 0 Then
            MsgBox "Error Getting Version Information"
            End
        End If
        Select Case verinfo.dwPlatformId
            Case 0
                msg = msg + "Windows 32s "
            Case 1
                msg = msg + "Windows 95 "
            Case 2
                msg = msg + "Windows NT "
        End Select

        ver_major$ = verinfo.dwMajorVersion
        ver_minor$ = verinfo.dwMinorVersion
        build$ = verinfo.dwBuildNumber
        msg = msg + ver_major$ + "." + ver_minor$
        msg = msg + " (Build " + build$ + ") "
        
        msg = msg + Trim(Replace(verinfo.szCSDVersion, Chr(0), "")) + nl + nl

        ' Get CPU type and operating mode.
        Dim sysinfo As SYSTEM_INFO
        GetSystemInfo sysinfo
        msg = msg + "CPU: "
        Select Case sysinfo.dwProcessorType
            Case PROCESSOR_INTEL_386
                msg = msg + "Intel 386" + nl
            Case PROCESSOR_INTEL_486
                msg = msg + "Intel 486" + nl
            Case PROCESSOR_INTEL_PENTIUM
                msg = msg + "Intel Pentium" + nl
            Case PROCESSOR_MIPS_R4000
                msg = msg + "MIPS R4000" + nl
            Case PROCESSOR_ALPHA_21064
                msg = msg + "DEC Alpha 21064" + nl
            Case Else
                msg = msg + "(unknown)" + nl

        End Select
        msg = msg + nl
        ' Get free memory.
        Dim memsts As MEMORYSTATUS
        Dim memory&
        GlobalMemoryStatus memsts
        memory& = memsts.dwTotalPhys
        msg = msg + "Total Physical Memory: "
        msg = msg + Format$(memory& \ 1024, "###,###,###") + "K" + nl
        memory& = memsts.dwAvailPhys
        msg = msg + "Available Physical Memory: "
        msg = msg + Format$(memory& \ 1024, "###,###,###") + "K" + nl
        memory& = memsts.dwTotalVirtual
        msg = msg + "Total Virtual Memory: "
        msg = msg + Format$(memory& \ 1024, "###,###,###") + "K" + nl
        memory& = memsts.dwAvailVirtual
        msg = msg + "Available Virtual Memory: "
        msg = msg + Format$(memory& \ 1024, "###,###,###") + "K" + nl _
             + nl

        ' Get free system resources.
        ' Not applicable to 32-bit operating system (Windows NT).
    lblInfo.Caption = msg
    MousePointer = 0


End Sub
