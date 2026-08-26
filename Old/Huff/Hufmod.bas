Attribute VB_Name = "HufComp"
Option Explicit
'**************
' Declarations
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'***************
' User defined types
Private Type TreeNode
    Weight As Integer       ' Goes to zero building tree.
    SavedWeight As Integer  ' Saved weight from building tree.
    Child1 As Integer
    Child0 As Integer
End Type
Private Type CharCodes
    Code As Integer         ' Huffman code for ASCII character.
    CodeBits As Integer     ' Number of bits in Huffman code.
End Type
Private Type BitFile
    Mask As Byte
    Rack As Byte
    OutputByteCount As Long
End Type
'*******************
' Constants
Public Const SRCCOPY = &HCC0020
Private Const EndOfStream = 256
Private Const EndWeightStream = &HFFF
'********************
' Initializations
Dim Counters(256) As Long
Dim Nodes(514) As TreeNode
Dim Codes(257) As CharCodes
Dim Bitio As BitFile


' The function of this routine is to count the occurrences of each
' ASCII character in the input file, load the counts list box and
' print a hard copy if the option is selected.  The print counts option
' is the only way to get a hard copy of the actual counts.
Public Sub CharacterCounts(hFileIn As Integer, ProgressStep As Long)
    Dim i As Integer, j As Integer
    Dim bChar As Byte
    Dim Work As Single
    Dim Task As String

    Task = "Counting Characters "
    frmHuff!picVisible.Visible = True
    Do While Not EOF(hFileIn)
        Get #hFileIn, , bChar
        Counters(bChar) = Counters(bChar) + 1
        i = i + 1
        If i = ProgressStep Then
            Work = j / 100
            i = 0
            DoEvents
            ProgressDisplay Work, Task
            j = j + 1
        End If
    Loop
    For i = 0 To 256
        If Counters(i) <> 0 Then
            frmHuff!lstcounts.AddItem Left(CStr(i) & "  ", 3) _
                & "  " & CStr(Counters(i))
            If frmHuff!chkCounts.Value = Checked Then
                frmHuff!picProgress.Cls
                frmHuff!picProgress.Print "Printing Counts..."
                Printer.Print Left(CStr(i) & "  ", 3) _
                    & "  " & CStr(Counters(i))
            End If
        End If
    Next
    Printer.EndDoc
    frmHuff!picVisible.Visible = False
    DoEvents
End Sub

' The function of this routine is to scale the counts of each ASCII
' character so they will fit in an integer when they are added together
' to obtain the relative node weight in the tree.
' These weights become the saved weight after building the binary tree
' and are saved with the compressed file.
Public Sub ScaleCounts()
    Dim MaxCount As Long
    Dim i As Integer
    
    MaxCount = 0
    For i = 0 To 256
        If Counters(i) > MaxCount Then
            MaxCount = Counters(i)
        End If
    Next
        
    MaxCount = MaxCount / 255
    MaxCount = MaxCount + 1
    For i = 0 To 256
        Nodes(i).Weight = CInt(Counters(i) / MaxCount)
        If Nodes(i).Weight = 0 And Counters(i) <> 0 Then
            Nodes(i).Weight = 1
        End If
        If Nodes(i).Weight <> 0 Then
            frmHuff!lstWeights.AddItem Left(CStr(i) & "  ", 3) _
                & "  " & CStr(Nodes(i).Weight)
        End If
    Next
    Nodes(EndOfStream).Weight = 1
    frmHuff!lstWeights.AddItem CStr(EndOfStream) _
        & "  " & CStr(Nodes(EndOfStream).Weight)
End Sub

' The function of this routine is to build the encoding tree for
' compression.  Huffman trees are built from the bottom up using the
' relative weights (counts) of the nodes.  As the minimum node pairs are
' found, the SavedWeight contains the relative Weight and the Weight is
' driven to 0 (zero) and will not be used again.  The SavedWeight is used
' for debugging and OutputCounts routine.
' The Nodes array contains the ASCII characters decimal value in nodes
' 0 - 255.  The EndOfStream node is 256 and is the last Symbol output in
' the compressed file to tell the decoder where to stop decoding.  The
' bottom of the tree will start with node 257 and progress upward until
' the Root node is stored.  The function then returns the Root node.
Public Function BuildTree() As Integer
    Dim NextFree As Integer
    Dim i As Integer
    Dim Min1 As Integer
    Dim Min2 As Integer
    
    Nodes(513).Weight = &H7FFF
    For NextFree = EndOfStream + 1 To 514
        Min1 = 513
        Min2 = 513
        For i = 0 To NextFree - 1
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
        Nodes(NextFree).Weight = Nodes(Min1).Weight + Nodes(Min2).Weight
        Nodes(Min1).SavedWeight = Nodes(Min1).Weight
        Nodes(Min1).Weight = 0
        Nodes(Min2).SavedWeight = Nodes(Min2).Weight
        Nodes(Min2).Weight = 0
        Nodes(NextFree).Child0 = Min1
        Nodes(NextFree).Child1 = Min2
    Next
    NextFree = NextFree - 1     ' This is the Root Node
    Nodes(NextFree).SavedWeight = Nodes(NextFree).Weight
    BuildTree = NextFree
End Function

' The function of this routine is to print the nodes of the binary tree
' if the option is selected.
Public Sub PrintNodes(RootNode As Integer)
    Dim i As Integer
    
    Printer.Print "Root Node: " & CStr(RootNode)
    For i = 0 To 514
        If Nodes(i).SavedWeight <> 0 Then
            Printer.Print "Node: " & Left(CStr(i) & "  ", 3); _
                "Node Weight: " & Left(CStr(Nodes(i).SavedWeight) & "     ", 8) _
                & " Child 0: " & Left(CStr(Nodes(i).Child0) & "  ", 3) _
                & " Child 1: " & CStr(Nodes(i).Child1)
        End If
    Next
    Printer.EndDoc
End Sub

' Searching the tree each time to create the Symbol that represents an
' input character makes little sense.  It makes more sense to convert the
' tree to the code symbols and look them up in an array.  The purpose of
' this routine is to do just that, by recursively walking the tree in preorder
' and saving Leaf information in the Codes array.
' CodeSoFar is the Symbol for the ASCII character.
' CodeBits is the number of bits in the Symbol.
Public Sub ConvertTreeToCode(CodeSoFar As Integer, Bits As Integer, node As Integer)
    If node <= EndOfStream Then
        Codes(node).Code = CodeSoFar
        Codes(node).CodeBits = Bits
        Exit Sub
    End If
    CodeSoFar = CodeSoFar * 2
    Bits = Bits + 1
    ConvertTreeToCode CodeSoFar, Bits, Nodes(node).Child0
    ConvertTreeToCode (CodeSoFar Or 1), Bits, Nodes(node).Child1
    CodeSoFar = CodeSoFar \ 2
    Bits = Bits - 1
End Sub

' The function of this routine is to save the weights in the compressed
' file so the decoder can build an identical tree.
' The routine uses Nodes.SavedWeight since Nodes.Weight is zero for all
' nodes at this point.
' It would be inefficient to save all weights with a zero value.  If there
' are more than 3 zero weights, a start, stop, and save weights takes place.
' Data saved to file is in the form:
'         FirstNode, LastNode, Weights,.........&HFFF
' The &HFFF tells the decoder it has reached the end of the weight stream.
Public Sub OutputCounts(hFileOut As Integer)
    Dim i As Integer
    Dim FirstNode As Integer
    Dim LastNode As Integer
    Dim NextNode As Integer
    
    FirstNode = 0
    Do While (FirstNode < 256)
        ' Search for first non zero weight
        If Nodes(FirstNode).SavedWeight <> 0 Then
            ' Search for last non zero weight
            Do While 1
                For LastNode = FirstNode + 1 To 255
                    If Nodes(LastNode).SavedWeight = 0 Then
                        Exit For
                    End If
                Next
                LastNode = LastNode - 1
                ' Check number of zero weights
                For NextNode = LastNode + 1 To 255
                    If Nodes(NextNode).SavedWeight = 0 Then
                        If NextNode > 255 Then
                            Exit For
                        End If
                        If NextNode - LastNode > 3 Then
                            Exit For
                        End If
                    Else
                        LastNode = NextNode
                    End If
                Next
                ' Save FirstNode, LastNode, Weights
                Put #hFileOut, , FirstNode
                Put #hFileOut, , LastNode
                For i = FirstNode To LastNode
                    Put #hFileOut, , Nodes(i).SavedWeight
                Next
                FirstNode = NextNode + 1
                Exit Do
            Loop
        Else
            FirstNode = FirstNode + 1
        End If
    Loop
    Put #hFileOut, , EndWeightStream
End Sub

' The function of this routine is to re-read the input file and replace the
' ASCII character with its coded Symbol.  It calls the BitHandler to save
' the symbols.  When the input reaches EOF, the BitHandler is called to
' output the EndOfStream symbol.  The EndOfStream symbol tells the decoder
' to quit processing symbols.
Public Sub CompressFile(hFileIn As Integer, hFileOut As Integer, _
    ProgressStep As Long)
    Dim iSymbol As Integer
    Dim iBitcount As Integer
    Dim bChar As Byte
    Dim i As Integer, j As Integer
    Dim Work As Single
    Dim Task As String
    
    Task = "Compressing File "
    Bitio.Mask = &H80
    frmHuff!picVisible.Visible = True
    Do While Not EOF(hFileIn)
        Get #hFileIn, , bChar
        i = i + 1
        If i = ProgressStep Then
            Work = j / 100
            i = 0
            ProgressDisplay Work, Task
            j = j + 1
        End If
        iSymbol = Codes(bChar).Code
        iBitcount = Codes(bChar).CodeBits
        BitHandler hFileOut, iSymbol, iBitcount
        DoEvents
    Loop
    Close #hFileIn
    iSymbol = Codes(EndOfStream).Code
    iBitcount = Codes(EndOfStream).CodeBits
    BitHandler hFileOut, iSymbol, iBitcount
    ' This next code saves the bits that are in a Rack that isn't full.
    If Bitio.Mask <> &H80 Then
        Put #hFileOut, , Bitio.Rack
    End If
    frmHuff!txtEncodeLength.Text = LOF(hFileOut)
    Close #hFileOut
    frmHuff!picVisible.Visible = False
End Sub

' The function of this routine is to save the bit patterns to disk.
' There is no function that stores bits, therefore the Rack is loaded with
' bits until it is full and then the rack is saved to disk.
Public Sub BitHandler(hFileOut As Integer, iSymbol As Integer, iBitcount As Integer)
    Dim SymbolMask As Integer
    Dim i As Integer

    SymbolMask = 1
    For i = 1 To iBitcount - 1
        SymbolMask = SymbolMask * 2
    Next
    Do While (SymbolMask <> 0)
        If (iSymbol And SymbolMask) Then
            Bitio.Rack = Bitio.Rack Or Bitio.Mask
        End If
        SymbolMask = SymbolMask / 2
        Bitio.Mask = Bitio.Mask / 2
        If Bitio.Mask = 0 Then
            Put #hFileOut, , Bitio.Rack
            Bitio.Mask = &H80
            Bitio.Rack = 0
            Bitio.OutputByteCount = Bitio.OutputByteCount + 2
        End If
    Loop
End Sub

' The function of this routine is to print the Nodes of the binary tree,
' their relative weights, and the binary code of the leaves.
Public Sub PrintModel(RootNode As Integer)
    Dim PrintBuffer As String
    Dim R As String
    Dim i As Integer
    
    Printer.Print "Huffman order 0 Compression Model"
    Printer.Print " "
    Printer.Print "Root Node: "; RootNode
    Printer.Print " "
    For i = 0 To RootNode
        If Nodes(i).SavedWeight <> 0 Then
            If i < 257 Then
                PrintBuffer = "Leaf Node: " & Left(CStr(i) & "  ", 3)
                R = PrintCharacter(i)
                PrintBuffer = PrintBuffer & Left(R & "    ", 8)
                PrintBuffer = PrintBuffer & " Weight: " & _
                    Left(CStr(Nodes(i).SavedWeight) & "               ", 8)
                PrintBuffer = PrintBuffer & " Code: "
                R = Hex2Bin(i)
                PrintBuffer = PrintBuffer & R
                Printer.Print PrintBuffer
            Else
                PrintBuffer = "Node: " & CStr(i)
                PrintBuffer = PrintBuffer & " Weight: " & _
                    Left(CStr(Nodes(i).SavedWeight) & "               ", 8)
                PrintBuffer = PrintBuffer & " Child 0: " & _
                            Left(CStr(Nodes(i).Child0) & "  ", 3)
                PrintBuffer = PrintBuffer & " Child 1: " & _
                              CStr(Nodes(i).Child1)
                Printer.Print PrintBuffer
            End If
        End If
    Next i
    Printer.EndDoc
End Sub

' The function of this routine is to provide information on the leaf
' node for the print model option.
Public Function PrintCharacter(i As Integer) As String
    Dim ReturnString As String
    
    Select Case i
        Case 0
            ReturnString = " (NULL)"
            PrintCharacter = ReturnString
        Case 1 To 9
            PrintCharacter = ""
        Case 10
            ReturnString = " (LF)"
            PrintCharacter = ReturnString
        Case 11, 12
            PrintCharacter = ""
        Case 13
            ReturnString = " (CR)"
            PrintCharacter = ReturnString
        Case 14 To 31
            PrintCharacter = ""
        Case 32
            ReturnString = " (Space)"
            PrintCharacter = ReturnString
        Case 33 To 126
            ReturnString = " (" & Chr(i) & ")"
            PrintCharacter = ReturnString
        Case 127 To 255
            PrintCharacter = ""
        Case 256
            ReturnString = " (EOS)"
            PrintCharacter = ReturnString
    End Select
End Function

' The function of this routine is to provide a binary representation of
' the Huffman code for the PrintModel option.
Public Function Hex2Bin(i As Integer) As String
    Dim j As Integer
    Dim Number As Integer
    Dim BitLength As Integer
    Dim d As Long
    Dim Bin As String
    
    Number = Codes(i).Code
    BitLength = Codes(i).CodeBits
    Bin = ""
    For j = 15 To 0 Step -1
        d = 2 ^ j
        Bin = Bin & Format(Number \ d, "0")
        If Number >= d Then
            Number = Number - d
        End If
    Next j
    Hex2Bin = Right(Bin, BitLength)
End Function

' The function of this routine is to display a progress bar for time
' consuming operations (character counts and compression of file data).
Public Sub ProgressDisplay(Work As Single, Task As String)
    Dim OldScaleMode As Integer
    Dim R As Long
    Dim Msg As String
    
    Msg = Task & Format(Work, "0%") & " Complete"
    frmHuff!picVisible.Cls
    frmHuff!picInvisible.Cls
    
    frmHuff!picVisible.CurrentX = _
        (frmHuff!picVisible.Width - frmHuff!picVisible.TextWidth(Msg)) / 2
    frmHuff!picInvisible.CurrentX = frmHuff!picVisible.CurrentX
    
    frmHuff!picVisible.CurrentY = _
        (frmHuff!picVisible.Height - frmHuff!picVisible.TextHeight(Msg)) / 2
    frmHuff!picInvisible.CurrentY = frmHuff!picVisible.CurrentY
    
    frmHuff!picVisible.Print Msg
    frmHuff!picInvisible.Print Msg
    
    OldScaleMode = frmHuff!picVisible.Parent.ScaleMode
    frmHuff!picVisible.Parent.ScaleMode = 3
    
    R = BitBlt(frmHuff!picVisible.hDC, 0, 0, _
        frmHuff!picInvisible.Width * Work, _
        frmHuff!picInvisible.Height, _
        frmHuff!picInvisible.hDC, 0, 0, SRCCOPY)
    
    frmHuff!picVisible.Parent.ScaleMode = OldScaleMode
End Sub




