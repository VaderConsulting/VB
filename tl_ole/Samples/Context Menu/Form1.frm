VERSION 5.00
Object = "{B6AE1890-E0CC-11D4-B99A-97636EED8334}#1.0#0"; "DESCM.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2610
   ClientLeft      =   2640
   ClientTop       =   3375
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   1560
   Begin prjCtxMenus.Test3 Test31 
      Height          =   795
      Left            =   45
      TabIndex        =   2
      Top             =   1755
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1402
   End
   Begin prjCtxMenus.Test1 Test11 
      Height          =   795
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1402
   End
   Begin prjCtxMenus.Test2 Test21 
      Height          =   795
      Left            =   45
      TabIndex        =   0
      Top             =   900
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1402
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' UserControl design context menu sample
'
' Test form
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

