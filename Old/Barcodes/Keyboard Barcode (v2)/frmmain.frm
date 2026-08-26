VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empired Barcode App"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrScan 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
   Begin VB.TextBox txtScan 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8467
            MinWidth        =   8467
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3069
            MinWidth        =   3069
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Sub Form_Load()
    GetDefaults
    GetPartNumbers
    GetSerials
    GetModes
    
    sbrStatus.Panels(2).Text = "Mode " & Mode & " (" & Modes(Mode) & ")"
    Me.Show
    Me.Refresh
End Sub

Sub GetDefaults()
    SerialsFilename = "serials.txt"
    PartNumbersFilename = "partnumbers.txt"
    Mode = 1
End Sub

Sub GetPartNumbers()
    Dim App_Path As String
    
    If Right(App.Path, 1) <> "\" Then
        App_Path = App.Path & "\"
    Else
        App_Path = App.Path
    End If
    
    sbrStatus.SimpleText = "Loading Part Number formats"
    
    Open App_Path & PartNumbersFilename For Input As #1
        Dim PartNumberData As String
        Dim PartNumberArray() As String
        Dim Item As Variant
        Dim ColumnNumber As Integer
        Dim ItemNumber As Integer
        Do
            Line Input #1, PartNumberData
            ColumnNumber = 0 ' Reset ColumnNumber
            If Left(PartNumberData, 2) <> "##" And Trim(PartNumberData) <> "" Then
                ItemNumber = ItemNumber + 1
                PartNumberArray() = Split(PartNumberData, ",")
                For Each Item In PartNumberArray()
                    ColumnNumber = ColumnNumber + 1
                    ReDim Preserve PartNumbers(ItemNumber)
                    Select Case ColumnNumber
                        Case 1 ' ID
                            PartNumbers(ItemNumber).ID = CInt(Item)
                        Case 2 ' Serial ID
                            PartNumbers(ItemNumber).SerialID = CInt(Item)
                        Case 3 ' Format
                            PartNumbers(ItemNumber).Format = Item
                        Case 4 ' Manufacturer
                            PartNumbers(ItemNumber).Manufacturer = Item
                        Case 5 ' Model
                            PartNumbers(ItemNumber).Model = Item
                        Case 6 ' Type
                            PartNumbers(ItemNumber).Type = CInt(Item)
                        Case 7 ' Description
                            PartNumbers(ItemNumber).Description = Item
                    End Select
                Next
                sbrStatus.Panels(1).Text = "Loading Part Number formats - " & ItemNumber & " loaded."
            End If
        Loop Until EOF(1)
    Close 1
    NumberOfPartNumbers = ItemNumber
    
    tmrScan.Enabled = True
End Sub

Sub GetSerials()
    Dim App_Path As String
    
    If Right(App.Path, 1) <> "\" Then
        App_Path = App.Path & "\"
    Else
        App_Path = App.Path
    End If
    
    sbrStatus.Panels(1).Text = "Loading Serial Number formats"
    
    Open App_Path & SerialsFilename For Input As #1
        Dim SerialData As String
        Dim SerialArray() As String
        Dim Item As Variant
        Dim ColumnNumber As Integer
        Dim ItemNumber As Integer
        Do
            Line Input #1, SerialData
            ColumnNumber = 0 ' Reset ColumnNumber
            If Left(SerialData, 2) <> "##" Then
                ItemNumber = ItemNumber + 1
                SerialArray() = Split(SerialData, ",")
                For Each Item In SerialArray()
                    ColumnNumber = ColumnNumber + 1
                    ReDim Preserve Serials(ItemNumber)
                    Select Case ColumnNumber
                        Case 1 ' ID
                            Serials(ItemNumber).ID = CInt(Item)
                        Case 2 ' PartNumber ID
                            Serials(ItemNumber).PartNumberID = Item
                        Case 3 ' Format
                            Serials(ItemNumber).Format = Item
                        Case 4 ' Example
                            
                    End Select
                Next
                sbrStatus.Panels(1).Text = "Loading Serial Number formats - " & ItemNumber & " loaded."
            End If
        Loop Until EOF(1)
    Close 1
    NumberOfSerials = ItemNumber
    
    tmrScan.Enabled = True
End Sub

Sub GetModes()

End Sub

Private Sub tmrScan_Timer()
    sbrStatus.Panels(1).Text = "Idle"
    tmrScan.Enabled = False
End Sub

Private Sub txtScan_Change()
    sbrStatus.SimpleText = "Scan - " & Right(txtScan.Text, 1)
    If Right(txtScan.Text, 2) = vbCrLf Then
        sbrStatus.Panels(1).Text = "Scan Complete"
        tmrScan.Enabled = True
        ProcessScan txtScan.Text
    End If
End Sub

Sub ProcessScan(Text As String)
    Dim lp As Integer
    Dim CompareString As String
    
    Text = Replace(Text, vbCrLf, "")
    
    If Text Like "M-###" Then ' Mode
        Dim ModeString As String
        
        ' Modes
        ' M-001,Add item to Database
        ' M-002,Delete item from Database
        ' M-003,Inventory Item
        ' M-004,Move Item to another location
        ' M-005,Decommision item
        ' M-006,Remove item (Stolen, missing etc)
        ' M-007,Lock Item
        ' M-008,Unlock item
        ' M-009,Warehouse item
        
        ModeString = Right(Text, 3)
        Select Case ModeString
            Case "001"
                Mode = 1
            Case "002"
                Mode = 2
            Case "003"
                Mode = 3
            Case "004"
                Mode = 4
            Case "005"
                Mode = 5
            Case "006"
                Mode = 6
            Case "007"
                Mode = 7
            Case "008"
                Mode = 8
            Case "009"
                Mode = 9
            Case Else
                Mode = 0
        End Select
        sbrStatus.Panels(2).Text = "Mode " & Mode & " (" & Modes(Mode) & ")"
    Else
        ' Part Numbers
        If PartNumber = "" Then
            For lp = 1 To NumberOfPartNumbers
                CompareString = PartNumbers(lp).Format
                CompareString = Replace(CompareString, "[", "")
                CompareString = Replace(CompareString, "]", "")
                CompareString = Replace(CompareString, "{", "")
                CompareString = Replace(CompareString, "}", "")
                
                If Text Like CompareString Then ' Match!
                    PartID = lp
                    GetPartAndSerial Text, PartNumbers(lp).Format, Serial, PartNumber
                End If
            Next lp
        End If
        
        ' Serial Numbers
        If Serial = "" Then
            For lp = 1 To NumberOfSerials
                CompareString = Serials(lp).Format
                CompareString = Replace(CompareString, "[", "")
                CompareString = Replace(CompareString, "]", "")
                If Text Like CompareString Then
                    Debug.Print CompareString
                End If
            Next lp
        End If
    
        Debug.Print PartNumbers(PartID).Manufacturer & " " & PartNumbers(PartID).Model & " " & PartNumbers(PartID).Description & " Partnumber: " & PartNumber & " Serial: " & Serial
    End If
    
    frmMain.txtScan.Text = ""
    
    If Mode > 0 And Serial <> "" And PartNumber <> "" Then
        Select Case Mode
            Case 1
            Case 2
            Case 3
            Case 4
            Case 5
            Case 6
            Case 7
            Case 8
            Case 9
        End Select
        
        PartID = 0
        PartNumber = ""
        Serial = ""
        
    End If
End Sub

Private Function GetPartAndSerial(ByVal Text As String, ByVal Template As String, ByRef Serial As String, ByRef PartNumber As String) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim a As Integer
    Dim b As Integer
    Dim OriginalText As String
    OriginalText = Text
    
    i = InStr(1, Template, "[")
    j = InStr(1, Template, "{")
    If i < j Then
        'Part Number first
        a = InStr(1, Template, "[")
        b = InStr(1, Template, "]")
        PartNumber = Mid(Text, a, (b - a) - 1)
        
        Text = Mid(Text, b - 1)
        Template = Mid(Template, b + 1, 4096) ' assumes template string is >= 4096 chars
        
        a = InStr(1, Template, "{")
        b = InStr(1, Template, "}")
        Serial = Mid(Text, a, (b - a) - 1)
    ElseIf j < i Then
        ' Serial Number first
        a = InStr(1, Template, "{")
        b = InStr(1, Template, "}")
        Serial = Mid(Text, a, (b - a) - 1)
        
        Text = Mid(Text, b - 1)
        Template = Mid(Template, b + 1, 4096) ' assumes template string is >= 4096 chars
        
        a = InStr(1, Template, "[")
        b = InStr(1, Template, "]")
        PartNumber = Mid(Text, a, (b - a) - 1)
    Else
        ' Either no serial, or no partnumber
    End If
End Function
