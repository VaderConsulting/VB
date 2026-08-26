VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "Named Pipe Client"
   ClientHeight    =   2040
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2040
   ScaleWidth      =   4230
   Begin VB.CheckBox chkConnect 
      Caption         =   "Connect"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   660
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   1200
   End
   Begin VB.TextBox txtPipe 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "\\.\pipe\vbpgpipe1"
      Top             =   180
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of bytes received:"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblCount 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pipe Name:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 1997 by Desaware Inc. All Rights Reserved

Dim pipehandle&
Dim totalread&

Private Sub chkConnect_Click()
    Dim res&
    If chkConnect.Value = 1 Then
        ' Connect to pipe
        res = WaitNamedPipe(txtPipe.Text, 10000)
        If res = 0 Then
            MsgBox "Pipe is not available at this time"
            chkConnect.Value = 0
            Exit Sub
        End If
        ' This could fail if server deletes pipe or other client
        ' grabs it during this time.
        pipehandle = CreateFile(txtPipe.Text, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If pipehandle = INVALID_HANDLE_VALUE Then
            MsgBox "Can't open named pipe"
            chkConnect.Value = 0
            Exit Sub
        End If
        Timer1.Enabled = True
    Else
        Timer1.Enabled = 0
        If pipehandle <> 0 Then
            Call CloseHandle(pipehandle)
            pipehandle = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        If pipehandle <> 0 Then
            Call CloseHandle(pipehandle)
            pipehandle = 0
        End If
End Sub

Private Sub Timer1_Timer()
    Dim inbuf As Byte
    Dim res&
    Dim bytesread&, bytesavail&, bytesleft&
    ' See if any data is waiting in the pipe
    res = PeekNamedPipe(pipehandle, ByVal 0&, 0, bytesread, bytesavail, bytesleft)
    If res <> 0 And bytesavail > 0 Then
        ' Read a byte off the pipe
        res = ReadFile(pipehandle, inbuf, 1, bytesread, 0)
        If res And bytesread = 1 Then
            lblCount.Caption = totalread
            totalread = totalread + 1
        End If
    End If
End Sub
