Attribute VB_Name = "HufExp"
Option Explicit
'************************
' Declarations
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'************************
' User defined types
Private Type TreeNode
    Weight As Integer
    SavedWeight As Integer
    Child0 As Integer
    Child1 As Integer
End Type
Private Type BitFile
    Mask As Byte
    Rack As Byte
    ByteCount As Long      ' Unused, but could be used to indicate progress.
End Type
'**************************
' Constants
Public Const SRCCOPY = &HCC0020
Private Const EndOfStream = 256
Private Const EndWeightStream = &HFFF
'***************************
' Initializations
Dim Nodes(514) As TreeNode
Dim Bitio As BitFile

' The Function of this routine is to read the Node Weights from the
' compressed file.  EndWeightStream indicates all Weights have been
' read in.
'
Public Function InputCounts(hFileIn As Integer) As Integer
    Dim FirstNode As Integer
    Dim LastNode As Integer
    Dim Talley As Integer
    Dim w As Integer            ' Count (Weight)
    Dim i As Integer            ' Index
    
    If EOF(hFileIn) Then
        MsgBox "Fatal error reading counts", 48
        InputCounts = 0
        Exit Function
    Else
        Get #hFileIn, , FirstNode
        Talley = Talley + 2
    End If
    If EOF(hFileIn) Then
        MsgBox "Fatal error reading counts", 48
        InputCounts = 0
        Exit Function
    Else
        Get #hFileIn, , LastNode
        Talley = Talley + 2
    End If
    Do While 1
        For i = FirstNode To LastNode
            If EOF(hFileIn) Then
                MsgBox "Fatal error reading counts", 48
                InputCounts = 0
                Exit Function
            Else
                Get #hFileIn, , w
                Talley = Talley + 2
                Nodes(i).Weight = w
            End If
        Next
        If EOF(hFileIn) Then
            MsgBox "Fatal error reading counts", 48
            InputCounts = 0
            Exit Function
        Else
            Get #hFileIn, , FirstNode
            Talley = Talley + 2
        End If
        If FirstNode = EndWeightStream Then
            Exit Do
        End If
        If EOF(hFileIn) Then
            MsgBox "Fatal error reading counts", 48
            InputCounts = 0
            Exit Function
        Else
            Get #hFileIn, , LastNode
            Talley = Talley + 2
        End If
    Loop
    Nodes(EndOfStream).Weight = 1
    InputCounts = Talley
End Function

' The function of this routine is to build the decoding tree based on
' the InputCounts of the compressed file.
'
Public Function BuildTree() As Integer
    Dim iNextFree As Integer
    Dim i As Integer
    Dim Min1 As Integer
    Dim Min2 As Integer
    
    Nodes(513).Weight = &H7FFF
    For iNextFree = EndOfStream + 1 To 514
        Min1 = 513
        Min2 = 513
        For i = 0 To iNextFree - 1
            If Nodes(i).Weight <> 0 Then
                If Nodes(i).Weight < Nodes(Min1).Weight Then
                    Min2 = Min1
                    Min1 = i
                ElseIf Nodes(i).Weight < Nodes(Min2).Weight Then
                    Min2 = i
                End If
            End If
        Next
        If Min2 = 513 Then
            Exit For
        End If
        Nodes(iNextFree).Weight = Nodes(Min1).Weight + Nodes(Min2).Weight
        Nodes(Min1).SavedWeight = Nodes(Min1).Weight
        Nodes(Min1).Weight = 0
        Nodes(Min2).SavedWeight = Nodes(Min2).Weight
        Nodes(Min2).Weight = 0
        Nodes(iNextFree).Child0 = Min1
        Nodes(iNextFree).Child1 = Min2
    Next
    iNextFree = iNextFree - 1
    Nodes(iNextFree).SavedWeight = Nodes(iNextFree).Weight
    BuildTree = iNextFree
End Function

' The function of this routine is to walk the tree to a Leaf Node and then
' save the Leaf to the output File.
'
Public Sub ExpandData(hFileIn As Integer, hFileOut As Integer, _
    iRootNode As Integer, ProgressStep As Long)
    Dim Node As Integer
    Dim Leaf As Byte
    Dim R As Boolean        ' 1 bit = True
    Dim i As Integer, j As Integer
    Dim Work As Single
    Dim Task As String
    
    
    frmExpand!picVisible.Visible = True
    Task = "Expanding file "
    Bitio.Mask = &H80
    Do While 1
        Node = iRootNode
        Do
            R = InputBit(hFileIn)
            If R = True Then
                Node = Nodes(Node).Child1
            Else
                Node = Nodes(Node).Child0
            End If
        Loop While Node > EndOfStream
        If Node = EndOfStream Then
            Exit Do
        Else
            Leaf = CByte(Node)
            Put #hFileOut, , Leaf
            i = i + 1
            If i = ProgressStep Then
                Work = j / 100
                i = 0
                ProgressDisplay Work, Task
                j = j + 1
                DoEvents
            End If
        End If
    Loop
    frmExpand!picVisible.Visible = False
End Sub


' The function of this routine is to read the Huffman code bits and send
' them a bit at a time to theExpandData routine.
'
Public Function InputBit(hFileIn As Integer) As Boolean
    Dim Value As Integer
    
    If Bitio.Mask = &H80 Then
        Get #hFileIn, , Bitio.Rack
    End If
    Value = Bitio.Rack And Bitio.Mask
    Bitio.Mask = Bitio.Mask / 2
    If Bitio.Mask = 0 Then
        Bitio.Mask = &H80
    End If
    InputBit = IIf(Value <> 0, True, False)
End Function

Public Sub ProgressDisplay(Work As Single, Task As String)
    Dim OldScaleMode As Integer
    Dim R As Long
    Dim Msg As String
    
    Msg = Task & Format(Work, "0%") & " Complete"
    frmExpand!picVisible.Cls
    frmExpand!picInvisible.Cls
    
    frmExpand!picVisible.CurrentX = _
        (frmExpand!picVisible.Width - frmExpand!picVisible.TextWidth(Msg)) / 2
    frmExpand!picInvisible.CurrentX = frmExpand!picVisible.CurrentX
    
    frmExpand!picVisible.CurrentY = _
        (frmExpand!picVisible.Height - frmExpand!picVisible.TextHeight(Msg)) / 2
    frmExpand!picInvisible.CurrentY = frmExpand!picVisible.CurrentY
    
    frmExpand!picVisible.Print Msg
    frmExpand!picInvisible.Print Msg
    
    OldScaleMode = frmExpand!picVisible.Parent.ScaleMode
    frmExpand!picVisible.Parent.ScaleMode = 3
    
    R = BitBlt(frmExpand!picVisible.hDC, 0, 0, _
        frmExpand!picInvisible.Width * Work, _
        frmExpand!picInvisible.Height, _
        frmExpand!picInvisible.hDC, 0, 0, SRCCOPY)
        
    frmExpand!picVisible.Parent.ScaleMode = OldScaleMode
End Sub
