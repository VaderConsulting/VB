VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PICPic"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optColour2 
      Caption         =   "Black"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optColour1 
      Caption         =   "Red"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox pic2D 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3840
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   21
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox pic2G 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3840
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   17
      Top             =   600
      Width           =   615
   End
   Begin VB.PictureBox pic2A 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3840
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox pic2H 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4440
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   22
      Top             =   1200
      Width           =   135
   End
   Begin VB.PictureBox pic2C 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   135
      TabIndex        =   20
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox pic2E 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   135
      TabIndex        =   19
      Top             =   720
      Width           =   135
   End
   Begin VB.PictureBox pic2F 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      ScaleHeight     =   375
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   240
      Width           =   135
   End
   Begin VB.PictureBox pic2B 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4320
      ScaleHeight     =   375
      ScaleWidth      =   135
      TabIndex        =   16
      Top             =   240
      Width           =   135
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox picG 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.PictureBox picF 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1440
      ScaleHeight     =   1185
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picE 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1440
      ScaleHeight     =   1185
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picD 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox picC 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2760
      ScaleHeight     =   1185
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.PictureBox picA 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   225
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picB 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2760
      ScaleHeight     =   1185
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Binary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblHex 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblBin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label lblDec 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label lblDecValueTitle 
      Alignment       =   2  'Center
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BinaryValue As String
Dim DecimalValue As Integer
Dim LEDValue As Integer

Dim Colour1 As Long
Dim Colour2 As Long
Dim CurrentColour As Long
Dim Background As Long

Private Sub cmdReset_Click()
    picA.BackColor = Background
    picB.BackColor = Background
    picC.BackColor = Background
    picD.BackColor = Background
    picE.BackColor = Background
    picF.BackColor = Background
    picG.BackColor = Background
    picH.BackColor = Background
    
    ' Copy
    pic2A.BackColor = Background
    pic2B.BackColor = Background
    pic2C.BackColor = Background
    pic2D.BackColor = Background
    pic2E.BackColor = Background
    pic2F.BackColor = Background
    pic2G.BackColor = Background
    pic2H.BackColor = Background
    
    LEDValue = 0
    UpdateValue
End Sub

Private Sub ChangeColour()
    If picA.BackColor = Colour1 Or picA.BackColor = Colour2 Then
        picA.BackColor = CurrentColour
        pic2A.BackColor = CurrentColour
    End If
    If picB.BackColor = Colour1 Or picB.BackColor = Colour2 Then
        picB.BackColor = CurrentColour
        pic2B.BackColor = CurrentColour
    End If
    If picC.BackColor = Colour1 Or picC.BackColor = Colour2 Then
        picC.BackColor = CurrentColour
        pic2C.BackColor = CurrentColour
    End If
    If picD.BackColor = Colour1 Or picD.BackColor = Colour2 Then
        picD.BackColor = CurrentColour
        pic2D.BackColor = CurrentColour
    End If
    If picE.BackColor = Colour1 Or picE.BackColor = Colour2 Then
        picE.BackColor = CurrentColour
        pic2E.BackColor = CurrentColour
    End If
    If picF.BackColor = Colour1 Or picF.BackColor = Colour2 Then
        picF.BackColor = CurrentColour
        pic2F.BackColor = CurrentColour
    End If
    If picG.BackColor = Colour1 Or picG.BackColor = Colour2 Then
        picG.BackColor = CurrentColour
        pic2G.BackColor = CurrentColour
    End If
    If picH.BackColor = Colour1 Or picH.BackColor = Colour2 Then
        picH.BackColor = CurrentColour
        pic2H.BackColor = CurrentColour
    End If
End Sub

Private Sub Form_Load()
    Colour1 = &HFF
    Colour2 = &H0&
    CurrentColour = Colour1
    Background = frmMain.BackColor
    
    UpdateValue
    ChangeColour
End Sub

Private Sub optColour1_Click()
    CurrentColour = Colour1
    ChangeColour
End Sub

Private Sub optColour2_Click()
    CurrentColour = Colour2
    ChangeColour
End Sub

Private Sub picA_Click() ' OK
    If picA.BackColor = CurrentColour Then
        picA.BackColor = Background
        pic2A.BackColor = Background
        LEDValue = LEDValue And Not &H1
    Else
        picA.BackColor = CurrentColour
        pic2A.BackColor = CurrentColour
        LEDValue = LEDValue Or &H1
    End If
    UpdateValue
End Sub

Private Sub picB_Click() ' OK
    If picB.BackColor = CurrentColour Then
        picB.BackColor = Background
        pic2B.BackColor = Background
        LEDValue = LEDValue And Not &H2
    Else
        picB.BackColor = CurrentColour
        pic2B.BackColor = CurrentColour
        LEDValue = LEDValue Or &H2
    End If
    UpdateValue
End Sub

Private Sub picC_Click() ' OK
    If picC.BackColor = CurrentColour Then
        picC.BackColor = Background
        pic2C.BackColor = Background
        LEDValue = LEDValue And Not &H4
    Else
        picC.BackColor = CurrentColour
        pic2C.BackColor = CurrentColour
        LEDValue = LEDValue Or &H4
    End If
    UpdateValue
End Sub

Private Sub picD_Click() ' OK
    If picD.BackColor = CurrentColour Then
        picD.BackColor = Background
        pic2D.BackColor = Background
        LEDValue = LEDValue And Not &H8
    Else
        picD.BackColor = CurrentColour
        pic2D.BackColor = CurrentColour
        LEDValue = LEDValue Or &H8
    End If
    UpdateValue
End Sub


Private Sub picE_Click() ' OK
    If picE.BackColor = CurrentColour Then
        picE.BackColor = Background
        pic2E.BackColor = Background
        LEDValue = LEDValue And Not &H10
    Else
        picE.BackColor = CurrentColour
        pic2E.BackColor = CurrentColour
        LEDValue = LEDValue Or &H10
    End If
    UpdateValue
End Sub

Private Sub picF_Click() ' OK
    If picF.BackColor = CurrentColour Then
        picF.BackColor = Background
        pic2F.BackColor = Background
        LEDValue = LEDValue And Not &H20
    Else
        picF.BackColor = CurrentColour
        pic2F.BackColor = CurrentColour
        LEDValue = LEDValue Or &H20
    End If
    UpdateValue
End Sub

Private Sub picG_Click()
    If picG.BackColor = CurrentColour Then
        picG.BackColor = Background
        pic2G.BackColor = Background
        LEDValue = LEDValue And Not &H40
    Else
        picG.BackColor = CurrentColour
        pic2G.BackColor = CurrentColour
        LEDValue = LEDValue Or &H40
    End If
    UpdateValue
End Sub

Private Sub picH_Click()
    If picH.BackColor = CurrentColour Then
        picH.BackColor = Background
        pic2H.BackColor = Background
        LEDValue = LEDValue And Not &H80
    Else
        picH.BackColor = CurrentColour
        pic2H.BackColor = CurrentColour
        LEDValue = LEDValue Or &H80
    End If
    UpdateValue
End Sub

Private Sub UpdateValue()
    Me.lblDec.Caption = CStr(LEDValue)
    Me.lblBin.Caption = DecimalToBinary(CLng(LEDValue))
    Me.lblHex.Caption = Hex(LEDValue)
End Sub

Public Function DecimalToBinary(DecimalNum As Long) As String
    Dim tmp As String
    Dim n As Long
    Dim i As Integer

    n = DecimalNum

    tmp = Trim(Str(n Mod 2))
    n = n \ 2

    Do While n <> 0
        tmp = Trim(Str(n Mod 2)) & tmp
        n = n \ 2
    Loop

    DecimalToBinary = tmp
    Do Until Len(DecimalToBinary) = 8
        DecimalToBinary = "0" & DecimalToBinary
    Loop
End Function

Public Function BinaryToDecimal(Binary As String) As Long
    Dim n As Long
    Dim s As Integer

    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ (s - 1)))
    Next s

    BinaryToDecimal = n
End Function


