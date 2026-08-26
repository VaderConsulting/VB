Attribute VB_Name = "Module1"
Option Explicit

Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_LOCAL = &H2
Public Const PRINTER_ENUM_NAME = &H8
Public Const PRINTER_ENUM_NETWORK = &H40
Public Const PRINTER_ENUM_REMOTE = &H10
Public Const PRINTER_ENUM_SHARED = &H20

Public Type PRINTER_INFO_1
   flags As Long
   pDescription As String
   pName As String
   pComment As String
End Type

Public Type PRINTER_INFO_4
   pPrinterName As String
   pServerName As String
   Attributes As Long
End Type

Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function PtrToStr Lib "Kernel32" Alias "lstrcpyA" (ByVal retval As String, ByVal Ptr As Long) As Long
Public Declare Function StrLen Lib "Kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long

