VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imglvMainLarge 
      Left            =   1560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "File"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   "Domain"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BA4
            Key             =   "Generic"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FF6
            Key             =   "AdminShare"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1448
            Key             =   "Network"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":189A
            Key             =   "Share"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CEC
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":213E
            Key             =   "Directory"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglvMainSmall 
      Left            =   840
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2590
            Key             =   "Domain"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29E2
            Key             =   "Generic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E34
            Key             =   "AdminShare"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3286
            Key             =   "Network"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36D8
            Key             =   "Share"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B2A
            Key             =   "Server"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F7C
            Key             =   "Directory"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43CE
            Key             =   "File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   741
      ButtonWidth     =   609
      Appearance      =   1
      ImageList       =   "ImgTBMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vwDetail"
            Object.ToolTipText     =   "View Details"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "vwList"
            Object.ToolTipText     =   "View List"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Large Icons"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Small Icons"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpLevel"
            Object.ToolTipText     =   "Move up one Level"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgTBMain 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B20
            Key             =   "vwDetail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C32
            Key             =   "vwList"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D44
            Key             =   "vwLarg"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E56
            Key             =   "vwSmall"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F68
            Key             =   "UpLevel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6376
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imglvMainLarge"
      SmallIcons      =   "imglvMainSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   4830
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8520
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "30/11/98"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "17:32"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDescription 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strParent As String

Private Sub Form_Load()

Dim strDisp() As String
Dim strAttr() As String
Dim lngIcon() As Long
Dim lngReturn As Long
Dim i As Long
Dim strIcon As String
Dim strObject As String
Dim strComment As String
Dim itmX As ListItem

    With Me
    
        .Caption = "Browse Network Neighborhood"
        .lblDescription.Caption = "Network Provider"
        Call SendMessageLong(.lvMain.hwnd, _
            LVM_SETEXTENDEDLISTVIEWSTYLE, _
            LVS_EX_FULLROWSELECT, True)

    End With
    
    With Me.lvMain
    
        .ColumnHeaders.Add , , "Name", _
            (.Width / 3)
        .ColumnHeaders.Add , , "Description", _
            (.Width * 2) / 3
        .View = lvwReport
        
    End With
    
    Me.Show
    Me.MousePointer = vbHourglass
    Me.sbMain.Panels("Status").Text = "Enumerating Network..."
    DoEvents
    
    strParent = ""
    lngReturn = fnEnum("", strDisp(), strAttr(), lngIcon())
    
    If lngReturn = 0 Then
    
        For i = 1 To UBound(strDisp)
            
            lngReturn = fnGetDisplay(strDisp(i), strObject, strComment)
            strIcon = GetIcon(lngIcon(i))
            
            Set itmX = Me.lvMain.ListItems.Add(, strAttr(i), _
                strObject, strIcon, strIcon)
            itmX.SubItems(1) = strComment
            DoEvents
            
        Next i
        
    End If

    Me.sbMain.Panels("Status").Text = "Ready"
    Me.MousePointer = vbNormal
    
End Sub

Private Sub lvMain_DblClick()

Dim strKey As String
Dim strProvider As String
Dim lngProvider As Long
Dim strRemoteNames() As String
Dim strFullInfo() As String
Dim lngReturn As Long
Dim i As Long
Dim itmX As ListItem
Dim lngIcon() As Long
Dim strIcon As String
Dim strDisplay As String
Dim strObject As String
Dim strComment As String

    Me.MousePointer = vbHourglass
    strKey = Me.lvMain.SelectedItem.Key
    '
    Me.sbMain.Panels("Status").Text = "Working..."
    DoEvents
    
    Me.lblDescription.Caption = "Resources on " & Me.lvMain.SelectedItem.Text
    lngReturn = fnEnum(strKey, strRemoteNames(), strFullInfo(), lngIcon())
    If lngReturn = 0 Then
        Me.lvMain.ListItems.Clear
        For i = 1 To UBound(strRemoteNames)
            
            strIcon = GetIcon(lngIcon(i))
            lngReturn = fnGetDisplay(strRemoteNames(i), strDisplay, strComment)
            If Left$(strDisplay, 2) = "\\" Then
                strDisplay = Mid$(strDisplay, 3, Len(strDisplay) - 2)
            End If
            If InStr(strDisplay, "\") <> 0 Then
                strDisplay = Mid$(strDisplay, InStr(strDisplay, "\") + 1, Len(strDisplay) - InStr(strDisplay, "\") + 1)
            End If
            If strDisplay <> "" Then
                Set itmX = Me.lvMain.ListItems.Add(, _
                    strFullInfo(i), _
                    strDisplay, strIcon, strIcon)
                itmX.SubItems(1) = strComment
                DoEvents
            End If
        Next i
    Else
        MsgBox "Error " & str$(lngReturn)
    End If
    
    Me.sbMain.Panels("Status").Text = "Ready"
    Me.MousePointer = vbNormal
 

End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strKey As String
Dim strProvider As String
Dim lngProvider As Long
Dim strRemoteNames() As String
Dim strFullInfo() As String
Dim lngReturn As Long
Dim i As Long
Dim itmX As ListItem
Dim lngIcon() As Long
Dim strIcon As String
Dim strDisplay As String
Dim strObject As String
Dim strComment As String

    Select Case Button.Key
        
        Case "vwDetails"
            Me.lvMain.View = lvwReport
        Case "vwList"
            Me.lvMain.View = lvwList
        Case "vwLarge"
            Me.lvMain.View = lvwIcon
        Case "vwSmall"
            Me.lvMain.View = lvwSmallIcon
        Case "UpLevel"
    

            Me.MousePointer = vbHourglass
            strKey = Me.lvMain.SelectedItem.Key
            'strParent = strKey
            Me.sbMain.Panels("Status").Text = "Working..."
            DoEvents
    
            Me.lblDescription.Caption = "Resources on " & Me.lvMain.SelectedItem.Text
            lngReturn = fnEnum(strParent, strRemoteNames(), strFullInfo(), lngIcon())
    
                If lngReturn = 0 Then
                Me.lvMain.ListItems.Clear
                For i = 1 To UBound(strRemoteNames)
            
                    strIcon = GetIcon(lngIcon(i))
                    lngReturn = fnGetDisplay(strRemoteNames(i), strDisplay, strComment)
                    If Left$(strDisplay, 2) = "\\" Then
                        strDisplay = Mid$(strDisplay, 3, Len(strDisplay) - 2)
                    End If
                    If InStr(strDisplay, "\") <> 0 Then
                        strDisplay = Mid$(strDisplay, InStr(strDisplay, "\") + 1, Len(strDisplay) - InStr(strDisplay, "\") + 1)
                    End If
                    If strDisplay <> "" Then
                        Set itmX = Me.lvMain.ListItems.Add(, _
                            strFullInfo(i), _
                            strDisplay, strIcon, strIcon)
                        itmX.SubItems(1) = strComment
                        DoEvents
                    End If
                Next i
            Else
                MsgBox "Error " & str$(lngReturn)
            End If
    
            Me.sbMain.Panels("Status").Text = "Ready"
            Me.MousePointer = vbNormal
            
    End Select
 
End Sub
