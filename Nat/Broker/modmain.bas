Attribute VB_Name = "modMain"
Public DoStop As Boolean
Public DoPause As Boolean

Sub Main()
    DoStop = False
    DoPause = False
    Load frmMain
    frmMain.Show
    frmMain.Refresh
End Sub
