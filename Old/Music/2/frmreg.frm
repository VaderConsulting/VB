VERSION 5.00
Begin VB.Form frmRegister 
   Caption         =   "Register"
   ClientHeight    =   6915
   ClientLeft      =   3720
   ClientTop       =   1485
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6915
   ScaleWidth      =   9600
   Begin VB.TextBox txtSolution 
      Height          =   285
      Left            =   3840
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Details"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
      Width           =   6375
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6375
   End
   Begin VB.CommandButton cmdCost 
      Caption         =   "How much? etc..."
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAuthor 
      Caption         =   "Author info"
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox cmbState 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox txtPostcode 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtPlace 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   5055
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
   Begin VB.TextBox txtSurname 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      Caption         =   "Solution (invisible control)"
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lbl5 
      Caption         =   "5. Click on 'Register'."
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   5160
      Width           =   9135
   End
   Begin VB.Label lbl3A 
      Caption         =   "    The Author will not send the Password until funds have been received.  You may continue to use the program."
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   4680
      Width           =   9135
   End
   Begin VB.Label lbl4 
      Caption         =   "4. On receipt, enter your Registration Password in the appropriate box."
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   4920
      Width           =   9135
   End
   Begin VB.Label lbl3 
      Caption         =   "3. The Author will contact you and discuss payment details.  Wait for the returned fax with your Registration Password.  "
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   4440
      Width           =   9135
   End
   Begin VB.Label lbl2 
      Caption         =   "2. Click on 'Submit' and then 'Print Details'.  Fax the printout to the Authors Fax Number. (see Author info)."
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   4200
      Width           =   9135
   End
   Begin VB.Label lbl1 
      Caption         =   "1. Fill in all of the above details."
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   3960
      Width           =   9135
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Instructions"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Registration Password"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Registration Code"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblFax 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblPhone 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      Caption         =   "State / Territory"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblPostcode 
      Alignment       =   1  'Right Justify
      Caption         =   "Postcode"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblPlace 
      Alignment       =   1  'Right Justify
      Caption         =   "City / Suburb / Town"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblSurname 
      Alignment       =   1  'Right Justify
      Caption         =   "Surname"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAuthor_Click()
    subAuthor
End Sub

Private Sub cmdClose_Click()
    frmAbout.Show
    frmRegister.Hide
End Sub

Private Sub cmdCost_Click()
    frmInfo.Show
    frmRegister.Hide
End Sub

Private Sub cmdPrint_Click()
    Printer.Print gProgramName; " Version " & App.Major & "." & App.Minor & "." & App.Revision
    Printer.Print "Please fax to this number: "; gAuthorFax
    Printer.Print "Or email to this address "; gEmail
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "First Name "; txtFirstName
    Printer.Print "Surname    "; txtSurname
    Printer.Print "Company    "; txtCompany
    Printer.Print "Address    "; txtAddress
    Printer.Print "City       "; txtPlace
    Printer.Print "Postcode   "; txtPostcode
    Printer.Print "State      "; cmbState.List(cmbState.ListIndex)
    Printer.Print "Phone      "; txtPhone
    Printer.Print "Fax        "; txtFax
    Printer.Print "Reg. Code  "; txtCode
    Printer.EndDoc
End Sub

Private Sub cmdRegister_Click()
    
    DriveInfo = fctGetDriveInfo
    Serial = Left$(DriveInfo, 9)
    Volume = Right$(DriveInfo, Len(DriveInfo) - 9)
    gCompany = txtCompany
    gPlace = txtPlace
    gSurname = txtSurname
    gAddress = txtAddress
    gPhone = txtPhone
    gFax = txtFax
    gFooter = txtFooter
    gState = cmbState.List(cmbState.ListIndex)
    gPostcode = txtPostcode
    gRegistered = 1
        
    SaveSetting gINIFile, "Register", "Registered", 1
    
    Code = fctEncode(Volume)
    SaveSetting gINIFile, "Register", "Volume", Code
    Code = fctEncode(Serial)
    SaveSetting gINIFile, "Register", "Serial", Code
    SaveSetting gINIFile, "Register", "Code", txtCode
    'SaveSetting gINIFile, "Register", "Password", txtPassword
    frmAbout.Show
    frmRegister.Hide
End Sub

Private Sub cmdSubmit_Click()
    Dim varLoop As Integer
    Dim J As String
    Dim varHex As String
    Dim varPassword As String
    
    If txtCode = "" Then
        For varLoop = 1 To 6
            Code = Code & fctRND(1, 255) ' get a random number between 1 and 255
        Next varLoop
        
        'Encode it using the always standard Business Name
        Key$ = gBusiness
        varPassword = fctCrypt(Key$, Code)
        varHex = ""
        For varLoop = 1 To Len(Code)
            J = Hex(Asc(Mid$(Code, varLoop, 1)))
            If Len(J) = 1 Then J = "0" + J
            varHex = varHex + J
        Next
        varHex = Format$(Len(varHex), "00") + varHex
        txtCode = varHex
        SaveSetting gINIFile, "Register", "Code", varHex    ' Save for future ref
    Else
        cmdSubmit.Enabled = 0
    End If
    Code = fctEncode(txtFirstName)
    SaveSetting gINIFile, "Register", "Name", Code
    Code = fctEncode(txtSurname)
    SaveSetting gINIFile, "Register", "Surname", Code
    Code = fctEncode(txtCompany)
    SaveSetting gINIFile, "Register", "Company", Code
    Code = fctEncode(txtAddress)
    SaveSetting gINIFile, "Register", "Address", Code
    Code = fctEncode(txtPlace)
    SaveSetting gINIFile, "Register", "Place", Code
    Code = fctEncode(txtPostcode)
    SaveSetting gINIFile, "Register", "Postcode", Code
    Code = fctEncode(cmbState.List(cmbState.ListIndex))
    SaveSetting gINIFile, "Register", "State", Code
    Code = fctEncode(txtPhone)
    SaveSetting gINIFile, "Register", "Phone", Code
    Code = fctEncode(txtFax)
    SaveSetting gINIFile, "Register", "Fax", Code
    
    'Convert the encryption code back and save it in a textbox
    Code = txtCode
    Code = Mid$(Code, 3, Val(Left$(Code, 2)))
    varPassword = ""
    For varLoop = 1 To Len(Code) Step 2
        J = Mid$(Code, varLoop, 2)
        varPassword = varPassword + Chr$(Val("&H" + J))
    Next
    txtSolution = fctCrypt(gBusiness, varPassword)

End Sub

Private Sub Form_Activate()
    'Convert the encryption code back and save it in a textbox
    Code = txtCode
    Code = Mid$(Code, 3, Val(Left$(Code, 2)))
    varPassword = ""
    For varLoop = 1 To Len(Code) Step 2
        J = Mid$(Code, varLoop, 2)
        varPassword = varPassword + Chr$(Val("&H" + J))
    Next
    txtSolution = fctCrypt(gBusiness, varPassword)
    If txtFirstName <> "" Then cmdSubmit.Enabled = 0
End Sub

Private Sub Form_Load()
    Dim varLoop As Integer
    Dim J As String
    Dim varPassword As String
    Dim varHex As String

    Width = gVGAWidth
    Height = gVGAHeight
    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
    Top = (Screen.Height - Height) / 2  ' Center form vertically.
    cmbState.AddItem "WESTERN AUSTRALIA"
    cmbState.AddItem "SOUTH AUSTRALIA"
    cmbState.AddItem "NORTHERN TERRITORY"
    cmbState.AddItem "ACT"
    cmbState.AddItem "VICTORIA"
    cmbState.AddItem "QUEENSLAND"
    cmbState.AddItem "NEW SOUTH WALES"
    cmbState.AddItem "TASMANIA"
    cmbState.AddItem "OTHER NON AUSTRALIAN"
    
    txtFirstName = fctDecode(GetSetting(gINIFile, "Register", "Name", ""))
    txtSurname = fctDecode(GetSetting(gINIFile, "Register", "Surname", ""))
    txtCompany = fctDecode(GetSetting(gINIFile, "Register", "Company", ""))
    txtAddress = fctDecode(GetSetting(gINIFile, "Register", "Address", ""))
    txtPlace = fctDecode(GetSetting(gINIFile, "Register", "Place", ""))
    txtPostcode = fctDecode(GetSetting(gINIFile, "Register", "Postcode", ""))
    
    state = fctDecode(GetSetting(gINIFile, "Register", "State", ""))
    If state <> "" Then
        For varLoop = 0 To cmbState.ListCount - 1
            If UCase$(state) = UCase$(cmbState.List(varLoop)) Then cmbState.ListIndex = varLoop
        Next varLoop
    End If
    txtPhone = fctDecode(GetSetting(gINIFile, "Register", "Phone", ""))
    txtFax = fctDecode(GetSetting(gINIFile, "Register", "Fax", ""))
    
    txtCode = GetSetting(gINIFile, "Register", "Code", "")
    
    If txtCode <> "" Then cmdSubmit.Enabled = 0
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtFirstName_LostFocus()
    If txtFirstName <> "" Then
        varChr = UCase(Left$(txtFirstName, 1))
        varRight = LCase(Right$(txtFirstName, Len(txtFirstName) - 1))
        txtFirstName = varChr & varRight
    End If
End Sub

Private Sub txtPassword_Change()
    If txtPassword = txtSolution Then
        cmdRegister.Enabled = 1
    Else
        cmdRegister.Enabled = 0
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtPlace_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
End Sub

Private Sub txtSurname_LostFocus()
    If txtFirstName <> "" Then
        varChr = UCase(Left$(txtSurname, 1))
        varRight = LCase(Right$(txtSurname, Len(txtSurname) - 1))
        txtSurname = varChr & varRight
    End If
End Sub
