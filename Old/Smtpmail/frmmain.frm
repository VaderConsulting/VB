VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Messages"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox chkAnon 
      Caption         =   "Anonymous"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optMailType 
      Caption         =   "NET SEND"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optMailType 
      Caption         =   "SMTP"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdSend 
      Height          =   975
      Left            =   3480
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblResult 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblServer 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblFrom 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This will send approx ~70 emails per second with a non-local SMTP server, and NOT flood your SMTP server.

Option Explicit

Public AutoMode As Boolean

Private Const NERR_Success As Long = 0&

Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" (yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long

Private Sub chkAnon_Click()
    If chkAnon.Value = vbChecked Then
        lblFrom.Visible = False
        txtFrom.Visible = False
    Else
        lblFrom.Visible = True
        txtFrom.Visible = True
    End If
End Sub

Private Sub cmdHelp_Click()
    Dim msg As String
    msg = ""
    msg = msg & "Commandline usage:" & vbCrLf & vbCrLf
    msg = msg & "SendMessage.exe file=drive\path\file.txt" & vbCrLf
    msg = msg & "Each line in this file must have either of the following format:" & vbCrLf & vbCrLf
    msg = msg & "0,Email_From,Email_To,""Subject"",SMTP_SERVER_NAME,""Message"",0" & vbCrLf & vbCrLf
    msg = msg & "  -or-" & vbCrLf & vbCrLf
    msg = msg & "1,Username_From,Username_To,""This is my NET SEND Message""" & vbCrLf & vbCrLf
    msg = msg & "The first field indicates whether to send an email or a NET SEND" & vbCrLf
    MsgBox msg, vbOKOnly + vbInformation, "Help"
End Sub

Private Sub optMailType_Click(Index As Integer)
    Select Case Index
        Case 0
            lblSubject.Visible = True
            txtSubject.Visible = True
            lblServer.Visible = True
            txtServer.Visible = True
            lblFrom.Visible = True
            txtFrom.Visible = True
            
            chkAnon.Visible = False
            
            txtTo.Text = GetSetting(App.EXEName, "Setup", "SMTPTo", "")
            txtFrom.Text = GetSetting(App.EXEName, "Setup", "SMTPFrom", "")
            txtServer.Text = GetSetting(App.EXEName, "Setup", "SMTPServer", "")
        Case 1
            lblSubject.Visible = False
            txtSubject.Visible = False
            lblServer.Visible = False
            txtServer.Visible = False
            
            chkAnon.Visible = True
            
            txtTo.Text = GetSetting(App.EXEName, "Setup", "NETTo", "")
            txtFrom.Text = GetSetting(App.EXEName, "Setup", "NETFrom", Environ$("USERNAME") & " on " & Environ$("COMPUTERNAME"))
    End Select
End Sub

Private Sub cmdSend_Click()
    Dim Result As Boolean
    
    If Trim(txtMessage.Text) = "" Then
        lblResult.Caption = "Enter text to send.  Message NOT sent."
        Exit Sub
    End If
    
    If optMailType(0).Value Then
        SendEmail txtFrom.Text, txtTo.Text, txtSubject.Text, txtMessage.Text, txtServer.Text
        lblResult.Caption = "SMTP Mail sent.  Delivery is not guaranteed."
    ElseIf optMailType(1).Value Then
        If chkAnon.Value = vbChecked Then
            
        Else
            
        End If
        Result = BroadcastMessage(txtTo.Text, txtFrom.Text, txtMessage.Text)
        lblResult.Caption = "Broadcast Message returned " & Result
    Else
        
    End If
    
    If AutoMode Then End
    
End Sub

Public Function SendEmail(strFrom As String, strTo As String, strSubject As String, strMessage As String, strServer As String)
    Dim objEmail As Object
    Set objEmail = CreateObject("CDO.Message")
    objEmail.From = strFrom
    objEmail.To = strTo
    objEmail.Subject = strSubject
    objEmail.Textbody = strMessage
    ' sendusing:
    ' 1 use local SMTP service (local to the specified smtpserver)
    ' 2 use port
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strServer
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objEmail.Configuration.Fields.Update
    On Error Resume Next
        objEmail.Send
    On Error GoTo 0
    
    If Not AutoMode Then
        SaveSetting App.EXEName, "Setup", "SMTPFrom", strFrom
        SaveSetting App.EXEName, "Setup", "SMTPTo", strTo
        SaveSetting App.EXEName, "Setup", "SMTPServer", strServer
    End If
    Set objEmail = Nothing
End Function

Public Function BroadcastMessage(UserOrMachine As String, strFrom As String, Message As String) As Boolean
    Dim byteToName() As Byte
    Dim byteFromName() As Byte
    Dim MessageToSend() As Byte
    
    If chkAnon.Value = vbChecked Then
        strFrom = "Another user on this LAN"
    Else
        
    End If
    
    'Put data into byte arrays
    byteToName = UserOrMachine & vbNullChar
    byteFromName = strFrom & vbNullChar
    MessageToSend = Message & vbNullChar
    
    'Broadcast message via API
    If NetMessageBufferSend(ByVal 0&, byteToName(0), byteFromName(0), MessageToSend(0), UBound(MessageToSend)) = NERR_Success Then
        'Return True if it worked
        BroadcastMessage = True
    End If
    
    If Not AutoMode Then
        SaveSetting App.EXEName, "Setup", "NETTo", UserOrMachine
        SaveSetting App.EXEName, "Setup", "NETFrom", strFrom
    End If
End Function

Private Sub Form_Load()
    Dim Args As String, Filename As String
    Dim intMailType As Integer
    Dim strFrom As String
    Dim strTo As String
    Dim strSubject As String
    Dim strServer As String
    Dim strMessage As String
    Dim intAnon As Integer
    Dim start As Date
    Dim finish As Date
    
    start = Now
    
    If Command$ <> "" Then
        Args = Command$
        
        If LCase(Left(Args, 4)) = "file" Then
            Filename = Trim(Mid(Args, 5, Len(Args) - 4))            ' Remove 'file'
            Filename = Trim(Mid(Filename, 2, Len(Filename) - 1))    ' Remove '='
            If Dir(Filename) = "" Then
                ' File not found
                End
            End If
            Open Filename For Input As #1
                Do Until EOF(1)
                    Input #1, intMailType
                    
                    optMailType(intMailType).Value = True
                    
                    Select Case intMailType
                        Case 0
                            Input #1, strFrom, strTo, strSubject, strServer, strMessage, intAnon
                        Case 1
                            Input #1, strFrom, strTo, strMessage
                    End Select
                    
                    txtFrom.Text = strFrom: txtTo.Text = strTo: txtSubject.Text = strSubject
                    txtServer.Text = strServer: txtMessage.Text = strMessage: chkAnon.Value = intAnon
                    
                    cmdSend_Click
                    
                    intMailType = 0: strFrom = "": strTo = "": strSubject = "": strMessage = "": intAnon = 0
                Loop
            Close 1
        Else
            
        End If
        
        AutoMode = True
    Else
        txtTo.Text = GetSetting(App.EXEName, "Setup", "SMTPTo", "")
        txtFrom.Text = GetSetting(App.EXEName, "Setup", "SMTPFrom", "")
        txtServer.Text = GetSetting(App.EXEName, "Setup", "SMTPServer", "")
        AutoMode = False
    End If
    
    
    If AutoMode Then
        finish = Now
        Debug.Print Now & " Run Time: " & DateDiff("s", start, finish)
        End
    End If
End Sub

