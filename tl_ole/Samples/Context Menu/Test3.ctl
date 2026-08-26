VERSION 5.00
Begin VB.UserControl Test3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ScaleHeight     =   795
   ScaleWidth      =   1440
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This control does not implements a custom menu."
      Height          =   660
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1320
   End
End
Attribute VB_Name = "Test3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Design-Time Context Menu Sample - Test3"
'*********************************************************************************************
'
' UserControl design context menu sample
'
' Test control 3
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98 + VB5
'
' History:
'          02/09/2000 - The file was released
'
'*********************************************************************************************
Option Explicit

Private Sub UserControl_Resize()

   Width = 1440
   Height = 795
   
End Sub


