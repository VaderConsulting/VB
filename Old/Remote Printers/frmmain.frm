VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote Printers"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optInfoReqd 
      Caption         =   "Domain"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optInfoReqd 
      Caption         =   "Server"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdGetPrinters 
      Caption         =   "Get' Em"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "\\CBDXAAG"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox lstPrinters 
      Height          =   5715
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Server name"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGetPrinters_Click()
    lstPrinters.Clear
    If optInfoReqd(0).Value = True Then
        GetPrinters 1, txtServer & Chr(0)                ' 1 = Specific server
                                                         ' 2 = Domain
    End If
    If optInfoReqd(1).Value = True Then
        GetPrinters 2, txtServer & Chr(0)                ' 1 = Specific server
                                                         ' 2 = Domain
    End If
End Sub

Sub GetPrinters(InfoReqd As Integer, Servername As String)
    Dim longbuffer() As Long                        ' resizable array receives information from the function
    Dim printinfo() As PRINTER_INFO_1               ' values inside longbuffer() will be put into here
    Dim numbytes As Long                            ' size in bytes of longbuffer()
'    Dim numneeded As Long                           ' receives number of bytes necessary if longbuffer() is too small
    Dim numprinters As Long                         ' receives number of printers found
    Dim c As Integer, retval As Long                ' counter variable & return value
    Me.AutoRedraw = True                            ' Set current graphic mode to persistent     ' Get information about the local printers
    numbytes = 3076                                 ' should be sufficiently big, but it may not be
    ReDim longbuffer(numbytes / 4) As Long          ' resize array -- note how 1 Long = 4 bytes
    'Dim InfoReqd As Integer                         ' 1 = Specific server
                                                     ' 2 = Domain
    'Dim ServerName As String                        ' Null terminated Servername
                                                    ' ie "\\CBDXAAG" & chr(0)
    
    Select Case InfoReqd
        Case 1
            retval = EnumPrinters(PRINTER_ENUM_NAME, Servername, 1, longbuffer(0), numbytes, numneeded, numprinters)
        Case 2
            retval = EnumPrinters(PRINTER_ENUM_REMOTE, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
    End Select
    
    If retval = 0 Then                              ' try enlarging longbuffer() to receive all necessary information
        numbytes = numneeded
        ReDim longbuffer(numbytes / 4) As Long      ' make it large enough
        Select Case InfoReqd
            Case 1
                retval = EnumPrinters(PRINTER_ENUM_NAME, Servername, 1, longbuffer(0), numbytes, numneeded, numprinters)
            Case 2
                retval = EnumPrinters(PRINTER_ENUM_REMOTE, "", 1, longbuffer(0), numbytes, numneeded, numprinters)
        End Select
        If retval = 0 Then                          ' failed again!
            'Debug.Print "Could not successfully enumerate the printers."
            MsgBox "Could not successfully enumerate the printers."
            'End                                     ' abort program
        End If
    End If
    ' Convert longbuffer() data into printinfo()
    ReDim printinfo(numprinters - 1) As PRINTER_INFO_1  ' room for each printer
    For c = 0 To numprinters - 1  ' loop, putting each set of information into each element
        ' longbuffer(4 * c) = .flags, longbuffer(4 * c + 1) = .pDescription, etc.
        ' For each string, the string is first buffered to provide enough room, and then the string is copied.
        printinfo(c).flags = longbuffer(4 * c)
        printinfo(c).pDescription = Space(StrLen(longbuffer(4 * c + 1)))
        retval = PtrToStr(printinfo(c).pDescription, longbuffer(4 * c + 1))
        printinfo(c).pName = Space(StrLen(longbuffer(4 * c + 2)))
        retval = PtrToStr(printinfo(c).pName, longbuffer(4 * c + 2))
        printinfo(c).pComment = Space(StrLen(longbuffer(4 * c + 3)))
        retval = PtrToStr(printinfo(c).pComment, longbuffer(4 * c + 3))
    Next c
    ' Display name of each printer
    For c = 0 To numprinters - 1
        'Me.Print "Name of printer"; c + 1; " is: "; printinfo(c).pName
        ShortPrinterName = Right(printinfo(c).pName, Len(printinfo(c).pName) - Len(Servername))
        Select Case InfoReqd
            Case 1
                lstPrinters.AddItem ShortPrinterName
            Case 2
                lstPrinters.AddItem printinfo(c).pName
        End Select
    Next c
End Sub

Private Sub optInfoReqd_Click(Index As Integer)
    Select Case Index
        Case 0
            txtServer.Enabled = True
        Case 1
            txtServer.Enabled = False
    End Select
End Sub
