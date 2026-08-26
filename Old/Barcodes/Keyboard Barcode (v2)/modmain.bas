Attribute VB_Name = "modMain"
Public Type ItemSerial
    ID As Integer
    PartNumberID As Integer
    Format As String
End Type

Public Type ItemPartNumber
    ID As Integer
    SerialID As Integer
    Format As String
    Manufacturer As String
    Model As String
    Type As Integer
    Description As String
End Type

Public Type HardwareItem
    SerialInfo As ItemSerial
    PartInfo As ItemPartNumber
End Type

Public Mode As String
Public Serial As String
Public PartNumber As String


Public Serials() As ItemSerial
Public PartNumbers() As ItemPartNumber

Public SerialsFilename As String
Public PartNumbersFilename As String
Public ModesFileNumber As String
Public TypesFilename As String

Public NumberOfSerials As Integer
Public NumberOfPartNumbers As Integer

Public ThisItem As HardwareItem
Public PartID As Integer

Public StatusString As String
Public Modes() As String

Sub main()
    ReDim Modes(9)
    Modes(0) = "Unknown"
    Modes(1) = "Add"
    Modes(2) = "Delete"
    Modes(3) = "Inventory"
    Modes(4) = "Move"
    Modes(5) = "Decommision"
    Modes(6) = "Remove"
    Modes(7) = "Lock"
    Modes(8) = "Unlock"
    Modes(9) = "Warehouse"
    Load frmMain
End Sub
