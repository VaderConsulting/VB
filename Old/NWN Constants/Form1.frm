VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NWN_Path As String = "C:\temp\NWN"

Private Sub Form_Load()
    Dim FileToLoad As String
    Dim FieldName As String
    Dim NWN_String As String
    Dim TabString As String
    Dim FieldData As String
    Dim Const_Type As String
    Dim Constant As String, Constant_Value As Variant
    Dim D As String, Content_Flag As Boolean
    
    Dim i As Integer
    
    D = Dir(NWN_Path & "\*.NSS")
    Do
        FileToLoad = D
        
        Open NWN_Path & "\CONSTANTS_" & D & ".CSV" For Output As #2
        
            Open NWN_Path & "\" & FileToLoad For Input As #1
                Do
                    Line Input #1, NWN_String
                    
                    ' Remove left and right spaces
                    If NWN_String <> "" Then
                        NWN_String = Trim(NWN_String)
                    End If
                    
                    ' Some lines are commented out.  Ignore these
                    If Left(NWN_String, 2) <> "//" Then
                        
                        ' Replace instances of chr(9) with spaces
                        NWN_String = Replace(NWN_String, Chr(9), " ")
                        
                        'Get type
                        i = InStr(1, NWN_String, " ")
                        If i <> 0 Then
                            Const_Type = Left(NWN_String, i - 1)
                        End If
                        
                        ' A constant MUST have a value, but will NOT have a bracket
                        If (InStr(1, NWN_String, "=") = 0) Then
                            
                        Else
                            If (InStr(1, NWN_String, "(") = 0) And (InStr(1, NWN_String, "/") = 0) And (InStr(1, NWN_String, "*") = 0) And (InStr(1, NWN_String, "+") = 0) And (InStr(1, NWN_String, "-") = 0) Then
                                Constant_Value = "NOT SPECIFIED"
                                
                                Select Case UCase(Const_Type)
                                    Case "INT", "FLOAT"
                                        Constant = GetConstant(NWN_String, Constant_Value)
                                    Case "VOID"
                                    Case "EFFECT"
                                    Case "STRING"
                                    Case "OBJECT"
                                    Case "LOCATION"
                                    Case "EVENT"
                                    Case "TALENT"
                                    Case "VECTOR"
                                    Case "#DEFINE"
                                        
                                    Case Else
                                    
                                End Select
                                
                                ' ASSUMPTION MADE THAT ALL CONSTANTS ARE UPPERCASE
                                If UCase(Constant) = Constant Then
                                    If Constant <> "" And Constant_Value <> "NOT SPECIFIED" Then
                                        Print #2, Const_Type & ","; Constant & ","; Constant_Value
                                        Content_Flag = True
                                    End If
                                End If
                                Const_Type = ""
                                Constant = ""
                            End If
                        End If
                    End If
                Loop Until EOF(1)
            Close 1
        Close 2
        D = Dir
        Content_Flag = False
    Loop Until D = ""
End Sub

Function GetConstant(FullString As String, Value) As String
    Dim i As Integer
    Dim LeftPos As Integer, RightPos As Integer
    Dim PartString As String
    
    i = InStr(1, FullString, " ")
    LeftPos = i
    PartString = Mid(FullString, LeftPos, 255)
    PartString = Trim(PartString)
    RightPos = InStr(LeftPos, PartString, " ")
    If RightPos = 0 Then Exit Function
    GetConstant = Trim(Left(PartString, RightPos))
    PartString = Trim(Mid(PartString, RightPos, 255))
    PartString = Replace(PartString, "=", "")
    RightPos = InStr(1, PartString, ";")
    PartString = Left(PartString, RightPos - 1)
    Value = Trim(PartString)
    If IsNumeric(Value) Then
        
    Else
        GetConstant = ""
    End If
End Function
