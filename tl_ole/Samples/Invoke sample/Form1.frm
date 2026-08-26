VERSION 5.00
Begin VB.Form frmInvokeSample 
   Caption         =   "IDispatch Invoke Sample"
   ClientHeight    =   1365
   ClientLeft      =   2790
   ClientTop       =   3285
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   315
      Left            =   2490
      TabIndex        =   7
      Top             =   975
      Width           =   1125
   End
   Begin VB.TextBox txtProp 
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Text            =   "Caption"
      Top             =   975
      Width           =   2295
   End
   Begin VB.CommandButton cmdLet 
      Caption         =   "&Let"
      Height          =   315
      Left            =   3660
      TabIndex        =   4
      Top             =   975
      Width           =   1125
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   2490
      TabIndex        =   6
      Top             =   345
      Width           =   2295
   End
   Begin VB.ComboBox cmbObj 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   345
      Width           =   2295
   End
   Begin VB.Label Labe1 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   2490
      TabIndex        =   5
      Top             =   60
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Object:"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Property:"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   735
      Width           =   630
   End
End
Attribute VB_Name = "frmInvokeSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' IDispatch Invoke Sample
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 02/28/2000
'
'*********************************************************************************************
Option Explicit

Dim Target As Object

Private Sub cmbObj_Click()

    Select Case cmbObj.ListIndex
        Case 0
            Set Target = Me
        Case 1
            Set Target = Label1(1)
        Case 2
            Set Target = txtProp
    End Select
    
    
End Sub


Private Sub cmdGet_Click()

    txtValue.Text = Invoke(Target, txtProp.Text, INVOKE_PROPERTYGET)
    
End Sub


Private Sub cmdLet_Click()
    
    Invoke Target, txtProp.Text, INVOKE_PROPERTYPUT, txtValue.Text

End Sub


Private Sub Form_Load()
    
    cmbObj.AddItem "Form"
    cmbObj.AddItem "Object Label"
    cmbObj.AddItem "Property TextBox"
    
    cmbObj.ListIndex = 0
    
End Sub


