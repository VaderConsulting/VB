VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl smtppoisctl 
   CanGetFocus     =   0   'False
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "smtppois.ctx":0000
   ScaleHeight     =   1620
   ScaleWidth      =   3315
   ToolboxBitmap   =   "smtppois.ctx":018A
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
      RemotePort      =   25
   End
End
Attribute VB_Name = "smtppoisctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private smtpCommand As String

Private tata_RCPTName(1 To 100) As String
Private NumOfRecipient As Integer
                                        
Private s_MailContent As String

Private s_SenderName As String

Private MachineName As String




''' event...
Public Event ConnectionMade()
Public Event ConnectionLost()
Public Event Error(ErrCode As Integer, ErrText As String)


''' public variable for ip address.
Private lRemoteHost As String


 
Sub ConnectionMade()

End Sub


Sub ConnectionLost()

End Sub



Public Function ConnectMailSrv(MailSrv As String) As Boolean

    If Winsock1.State = sckClosed Or Winsock1.State = sckClosing Then
        Winsock1.Close
        ConnectMailSrv = True
    Else
        ConnectMailSrv = False
        Exit Function
    End If
    


        ''' take default computer name as sender
        Dim sBuffer As String * 255
        If GetComputerName(sBuffer, 255) <> 0 Then
            MachineName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        End If
        
        With Winsock1
            .RemoteHost = MailSrv  ' "epss08"
            .RemotePort = 25  '''' smtp port
            .Protocol = sckTCPProtocol ''' tcp
            ''' randomly select port between 7001-9000
            Randomize
            .LocalPort = 7000 + Int((2000 * Rnd) + 1) ' Generate random value between 1 and 2000.
            'Debug.Print "Local Port = " & .LocalPort
        End With
    
            smtpCommand = "CONN"
            Winsock1.Connect
    
End Function

Sub Error(ErrCode As Integer, ErrText As String)
 
End Sub

Public Function SendMail(m_sender As String, m_recipients As String, m_header As String, m_priority As Byte, m_message As String) As Boolean
    
    SendMail = True
    If Winsock1.State <> sckConnected Then
        SendMail = False
        RaiseEvent Error(19999, "Mail server not connected or timeout by mail server.")
        Exit Function
    End If
        
    ''' 15/3/1999 --- Split to individual recipients
    If m_recipients = "" Then
        SendMail = False
        RaiseEvent Error(19998, "No recipients name provided in SendMail().")
       Exit Function
    End If
        
     NumOfRecipient = 0
     Dim stext As String
     Dim i As Integer
        For i = 1 To Len(m_recipients)
            If (Mid$(m_recipients, i, 1) <> ",") And (Mid$(m_recipients, i, 1) <> " ") Then
               stext = stext & Mid$(m_recipients, i, 1)
            End If
            If (Mid$(m_recipients, i, 1) = ",") Or (i = Len(m_recipients)) Then
                ''' form name
                NumOfRecipient = NumOfRecipient + 1
                tata_RCPTName(NumOfRecipient) = stext
                Debug.Print NumOfRecipient, tata_RCPTName(NumOfRecipient)
                stext = ""
            End If
        Next

    ''' 15/3/1999 --- Build mail content
    s_MailContent = ""
    s_MailContent = "Subject : " & m_header & vbCrLf
    s_MailContent = s_MailContent & "X-Priority: " & CStr(m_priority) & vbCrLf
    s_MailContent = s_MailContent & "X-Mailer: Pois Mailer Ver 1.0 " & vbCrLf & vbCrLf
    s_MailContent = s_MailContent & m_message & vbCrLf
    
    ''' 15/3/1999 --- build sender name
    If m_sender = "" Then
        s_SenderName = MachineName
    Else
        s_SenderName = m_sender & "(MachineName)"
    End If
    
    Call smtp_MAIL
    
    
End Function

Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 420
    
    
End Sub


Private Sub Winsock1_Close()
   
        Winsock1.Close
        
        RaiseEvent ConnectionLost
   
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connected (OK)"
   

        RaiseEvent ConnectionMade
        
        
        
    
    
        
    
    
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
 Dim sdata As String
        Winsock1.GetData sdata
        
            Select Case smtpCommand
                Case "CONN":
                    
                    If bytesTotal > 0 Then
                        If Left$(sdata, 4) = "220 " Then
                            Debug.Print smtpCommand & sdata
                            Call smtp_HELO
                        Else
                            Debug.Print smtpCommand & sdata
                        End If
                    End If
                        
                Case "HELO":
                    If Left$(sdata, 4) = "250 " Then
                        Debug.Print smtpCommand & sdata
                    Else
                        Debug.Print smtpCommand & sdata
                    End If
                    
                Case "MAIL":
                    If Left$(sdata, 4) = "250 " Then
                        Debug.Print smtpCommand & sdata
                        Call smtp_RCPT
                    Else
                        Debug.Print smtpCommand & sdata
                    End If
                
                Case "RCPT":
                    If Left$(sdata, 4) = "250 " Then
                        Debug.Print smtpCommand & sdata
                        NumOfRecipient = NumOfRecipient - 1
                        If NumOfRecipient > 0 Then
                            Call smtp_RCPT
                        Else
                            Call smtp_DATA
                        End If
                    Else
                        Debug.Print smtpCommand & sdata
                    End If
                
                Case "DATA":
                    If Left$(sdata, 4) = "354 " Then
                        Debug.Print smtpCommand & sdata
                        Call PutMailContent
                    Else
                        Debug.Print smtpCommand & sdata
                    End If
                    
                    
                Case "EDOT":
                    If Left$(sdata, 4) = "250 " Then
                        Debug.Print smtpCommand & sdata
                    Else
                        Debug.Print smtpCommand & sdata
                    End If
                    
                 
                    
                
                    

                Case Else
                    Debug.Print "Haha : " & sdata
            End Select
            
End Sub
Private Sub PutMailContent()
  
        Call smtp_CONTent
        
        Call smtp_EDOT
        
End Sub


Private Sub smtp_CONTent()
    smtpCommand = "CONT"
    Winsock1.SendData s_MailContent & vbCrLf

End Sub


Private Sub smtp_DATA()
    smtpCommand = "DATA"
    Winsock1.SendData "DATA " & vbCrLf
End Sub


Private Sub smtp_EDOT()
    smtpCommand = "EDOT"
    Winsock1.SendData vbCrLf & "." & vbCrLf
End Sub


Private Sub smtp_HELO()
    smtpCommand = "HELO"
    Winsock1.SendData "HELO" & vbCrLf
End Sub

Private Sub smtp_MAIL()
    smtpCommand = "MAIL"
        Winsock1.SendData "MAIL FROM:" & "<" & s_SenderName & ">" & vbCrLf
End Sub



Private Sub smtp_QUIT()
    smtpCommand = "QUIT"
    Winsock1.SendData "QUIT " & vbCrLf
  
End Sub


Private Sub smtp_RCPT()
    smtpCommand = "RCPT"
    Winsock1.SendData "RCPT TO:" & "<" & tata_RCPTName(NumOfRecipient) & ">" & vbCrLf
    
End Sub



Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
        Winsock1.Close
    RaiseEvent Error(Number, Description)
        
   
    
End Sub



Public Property Get RemoteHostIP() As String
    RemoteHostIP = Winsock1.RemoteHostIP
End Property



Public Property Get RemoteHost() As String
    RemoteHost = Winsock1.RemoteHost
End Property

 

Public Property Get Connected() As Boolean
    If Winsock1.State = sckConnected Then
        Connected = True
    Else
        Connected = False
    End If
    
End Property

 
Public Property Get LocalHostIP() As String
    LocalHostIP = Winsock1.LocalIP
End Property

 
Public Property Get LocalHost() As String
    LocalHost = Winsock1.LocalHostName
End Property

 
