VERSION 5.00
Begin VB.Form frmOrders 
   Caption         =   "Orders"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   10
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   9
      Left            =   1560
      TabIndex        =   20
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      Height          =   1245
      Index           =   8
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   17
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   6
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   5
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H0080FF80&
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   3855
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Commit Order"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label lblSupplierPhone 
      Caption         =   "Supplier Phone"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblPhone 
      Caption         =   "Customer phone"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblFault 
      Caption         =   "Fault description"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblCustomer 
      Caption         =   "Customer name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblOrder 
      Caption         =   "Order ID"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblWarranty 
      Caption         =   "Warranty ID"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblExpected 
      Caption         =   "Expected date"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblOrderDate 
      Caption         =   "Order date"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblSerial 
      Caption         =   "Faulty Serial No"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblSupplierID 
      Caption         =   "Supplier ID"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblProductID 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    rsSuppliers.MoveFirst
    rsSuppliers.MoveLast
    
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub
