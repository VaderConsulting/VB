VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "TypeLib Info"
   ClientHeight    =   5325
   ClientLeft      =   2025
   ClientTop       =   2670
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   WindowState     =   2  'Maximized
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2025
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   3572
      _Version        =   327682
      Indentation     =   0
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1995
      Left            =   2730
      TabIndex        =   0
      Top             =   90
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   3519
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   4.56789e5
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4950
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0079
            Key             =   "class"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":03CB
            Key             =   "iface"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":071D
            Key             =   "mod"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0A6F
            Key             =   "fldr"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0DC1
            Key             =   "enum"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1113
            Key             =   "tl"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1465
            Key             =   "udt"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":17B7
            Key             =   "union"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1B09
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' ITypeLib & ITypeInfo interfaces sample
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 02/26/2000
'
'*********************************************************************************************
Option Explicit

Dim m_oTL As ITypeLib
Sub LoadTL(ByVal FileName As String)
   
   On Error Resume Next
   
   Set m_oTL = LoadTypeLibEx(FileName, REGKIND_NONE)
   
   If Err.Number <> 0 Then
      MsgBox "The file doesn't contains a type library"
   Else
   
      pvFillTree m_oTL
   
      RichTextBox1.Text = DecompileTypeLib(FileName)
      
   End If
   
End Sub

Private Sub pvFillTree(ByVal oTL As ITypeLib)
Dim lIdx As Long
Dim sName As String

   On Error Resume Next
   
   oTL.GetDocumentation -1, sName, "", 0, ""
   
   TreeView1.Visible = False
   
   With TreeView1.Nodes
      
      .Clear
      .Add , , "library", sName, "tl"
      .Add("library", tvwChild, "coclass", "coclasses", "fldr").Sorted = True
      .Add("library", tvwChild, "typedef", "typedefs", "fldr").Sorted = True
      .Add("library", tvwChild, "struct", "structs", "fldr").Sorted = True
      .Add("library", tvwChild, "union", "unions", "fldr").Sorted = True
      .Add("library", tvwChild, "interface", "interfaces", "fldr").Sorted = True
      .Add("library", tvwChild, "dispinterface", "dispinterfaces", "fldr").Sorted = True
      .Add("library", tvwChild, "module", "modules", "fldr").Sorted = True
      .Add("library", tvwChild, "enum", "enums", "fldr").Sorted = True
   
      With .Item("library")
         .Sorted = True
         .Expanded = True
      End With
      
      For lIdx = 0 To oTL.GetTypeInfoCount - 1
         
         oTL.GetDocumentation lIdx, sName, "", 0, ""
         
         Select Case oTL.GetTypeInfoType(lIdx)
         
            Case TKIND_ALIAS
               .Add("typedef", tvwChild, sName, sName, "alias").Tag = lIdx
            
            Case TKIND_COCLASS
               .Add("coclass", tvwChild, sName, sName, "class").Tag = lIdx
               
            Case TKIND_DISPATCH
               Dim oTI As ITypeInfo, tTA As TYPEATTR
               
               Set oTI = oTL.GetTypeInfo(lIdx)
               tTA = GetTypeAttr(oTI)
               
               If (tTA.wTypeFlags And TYPEFLAG_FDUAL) <> 0 Then
                  .Add("interface", tvwChild, sName, sName, "iface").Tag = lIdx
               Else
                  .Add("dispinterface", tvwChild, "DI_" & sName, sName, "iface").Tag = lIdx
               End If
            
               Set oTI = Nothing
               
            Case TKIND_ENUM
               .Add("enum", tvwChild, sName, sName, "enum").Tag = lIdx
               
            Case TKIND_INTERFACE
               .Add("interface", tvwChild, sName, sName, "iface").Tag = lIdx
            
            Case TKIND_MODULE
               .Add("module", tvwChild, sName, sName, "mod").Tag = lIdx
            
            Case TKIND_RECORD
               .Add("struct", tvwChild, sName, sName, "udt").Tag = lIdx
            
            Case TKIND_UNION
               .Add("union", tvwChild, sName, sName, "union").Tag = lIdx
         
         End Select
         
      Next
   
   End With
   
   TreeView1.Visible = True
   
End Sub
Private Sub Form_Resize()
   
   TreeView1.Move 0, 0, ScaleWidth / 3, ScaleHeight
   RichTextBox1.Move ScaleWidth / 3, 0, 2 * ScaleWidth / 3, ScaleHeight
   
   If Command$ <> "" Then
      LoadTL Command$
   End If
   
End Sub
Private Sub mnuExit_Click()
   Unload Me
End Sub

Private Sub mnuFSave_Click()
   
   On Error Resume Next
   
   With CommonDialog1
      .Filter = "ODL|*.odl"
      .DefaultExt = "odl"
      .Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
      .FileName = ""
      .ShowSave
      
      If Err.Number = 0 Then
            
         RichTextBox1.SaveFile .FileName, rtfText
          
      End If
      
   End With

End Sub

Private Sub mnuOpen_Click()

   On Error Resume Next
   
   With CommonDialog1
      .Filter = "Libraries|*.ocx;*.dll;*.exe;*.olb;*.tlb|Type Libraries|*.tlb;*.olb|OCXs|*.ocx|DLLs|*.dll|EXEs|*.exe"
      .DefaultExt = "tlb"
      .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
      .ShowOpen
      
      If Err.Number = 0 Then
            
         LoadTL .FileName
          
      End If
      
   End With
   
End Sub



Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim oTI As ITypeInfo

   If Node.Key = "library" Then
      
      RichTextBox1.Text = DecompileTypeLib(m_oTL)
      
   Else
      
      If Node.Tag = "" Then
         
         RichTextBox1.Text = ""
         
      Else
      
         Set oTI = m_oTL.GetTypeInfo(Node.Tag)
         
         Select Case m_oTL.GetTypeInfoType(Node.Tag)
         
            Case TKIND_ALIAS
               RichTextBox1.Text = DecompileTypedef(oTI)
            
            Case TKIND_COCLASS
               RichTextBox1.Text = DecompileCoclass(oTI)
               
            Case TKIND_DISPATCH, TKIND_INTERFACE
               RichTextBox1.Text = DecompileInterface(oTI)
            
            Case TKIND_ENUM
               RichTextBox1.Text = DecompileEnum(oTI)
            
            Case TKIND_MODULE
               RichTextBox1.Text = DecompileModule(oTI)
            
            Case TKIND_RECORD
               RichTextBox1.Text = DecompileStruct(oTI)
            
            Case TKIND_UNION
               RichTextBox1.Text = DecompileUnion(oTI)
         
         End Select
         
      End If
      
   End If
   
End Sub
