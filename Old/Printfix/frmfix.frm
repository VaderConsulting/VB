VERSION 5.00
Begin VB.Form frmFix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printer Fix for Clients"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmFix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "This program is designed to allow the replacement of out-of-date or incorrect printer drivers with the correct drivers."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    If Dir("c:\temp\logoff.exe") = "" Then
        logonserver = Environ("logonserver")
        varPath = logonserver & "\netlogon\setup\logoff.exe"
        FileCopy varPath, "c:\temp\logoff.exe"
    End If
    If Dir("c:\temp\djrfix.txt") = "" Then
        Open "c:\temp\djrfix.txt" For Output As #1
            Print #1, "Printer fix run 1 commenced."
        Close 1
        Shell "c:\winnt\system32\con2prt.exe /f", vbHide
        Shell "NET STOP Spooler", vbHide
        On Error Resume Next
            Kill "c:\winnt\system32\spool\drivers\w32x86\2\*.*"
        On Error GoTo 0
        MsgBox "Please click OK to log off.  Please log on again and run this program again."
        Shell "c:\temp\logoff /N", vbHide
        End
    Else
        Shell "NET STOP Spooler", vbHide
        On Error Resume Next
            Kill "c:\winnt\system32\spool\drivers\w32x86\2\*.*"
            Kill "c:\temp\djrfix.txt"
        On Error GoTo 0
        Shell "NET START Spooler", vbHide
        MsgBox "Please click OK to log off.  Please log on again and your printers will be fixed."
        Shell "c:\temp\logoff /N", vbHide
        End
    End If
End Sub
