VERSION 5.00
Begin VB.Form Tank 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Tanks"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10560
   DrawWidth       =   4
   ForeColor       =   &H00000000&
   Icon            =   "3DTank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRight 
      Caption         =   "-->"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<--"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   855
   End
   Begin VB.CheckBox Hlp 
      Caption         =   "I'm thick and need help"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CheckBox Pause 
      Caption         =   "Paused"
      Height          =   210
      Left            =   2070
      TabIndex        =   8
      Top             =   4455
      Width           =   1380
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.CheckBox Shadow 
      Caption         =   "Shadow"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   1305
      IntegralHeight  =   0   'False
      Left            =   3765
      TabIndex        =   4
      Top             =   4290
      Width           =   6690
   End
   Begin VB.CheckBox HFace 
      Caption         =   "HF"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.CheckBox wire 
      Caption         =   "Solid"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   735
   End
   Begin VB.CheckBox ckB 
      Caption         =   "Border"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   4440
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   701
      TabIndex        =   5
      Top             =   0
      Width           =   10575
      Begin VB.Timer ReMessage 
         Interval        =   5000
         Left            =   1080
         Top             =   120
      End
      Begin VB.ListBox Messages 
         Height          =   1230
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.PictureBox Radar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         DrawWidth       =   4
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   8640
         ScaleHeight     =   95
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   6
         ToolTipText     =   "urytu"
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Timer Timer2 
         Left            =   600
         Tag             =   "0"
         Top             =   120
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu m1 
      Caption         =   "Views"
      Visible         =   0   'False
      Begin VB.Menu o1 
         Caption         =   "Left side"
         Index           =   1
      End
      Begin VB.Menu o1 
         Caption         =   "Right side"
         Index           =   2
      End
      Begin VB.Menu o1 
         Caption         =   "Forward view"
         Index           =   3
      End
      Begin VB.Menu o1 
         Caption         =   "Rear View"
         Index           =   4
      End
      Begin VB.Menu o1 
         Caption         =   "Brakes"
         Index           =   5
      End
      Begin VB.Menu o1 
         Caption         =   "Hide your tank"
         Index           =   6
      End
      Begin VB.Menu o1 
         Caption         =   "Short Range Radar"
         Checked         =   -1  'True
         Index           =   7
      End
      Begin VB.Menu o1 
         Caption         =   "Long Range Radar"
         Index           =   8
      End
   End
   Begin VB.Menu m2 
      Caption         =   "Commands"
      Begin VB.Menu p1 
         Caption         =   "Plane 1 attack..."
         Begin VB.Menu p1o 
            Caption         =   "Tower"
            Index           =   1
         End
         Begin VB.Menu p1o 
            Caption         =   "The other tower"
            Index           =   2
         End
         Begin VB.Menu p1o 
            Caption         =   "Baddy tank"
            Index           =   3
         End
         Begin VB.Menu p1o 
            Caption         =   "Me (Camakarzi)"
            Index           =   4
         End
      End
      Begin VB.Menu p2 
         Caption         =   "Plane 2 attack..."
         Begin VB.Menu p2o 
            Caption         =   "Tower"
            Index           =   1
         End
         Begin VB.Menu p2o 
            Caption         =   "The other tank"
            Index           =   2
         End
         Begin VB.Menu p2o 
            Caption         =   "Baddy tank"
            Index           =   3
         End
         Begin VB.Menu p2o 
            Caption         =   "Me (Camakarzi)"
            Index           =   4
         End
      End
      Begin VB.Menu Ap 
         Caption         =   "All"
         Begin VB.Menu pAo 
            Caption         =   "Tower"
            Index           =   1
         End
         Begin VB.Menu pAo 
            Caption         =   "The other tower"
            Index           =   2
         End
         Begin VB.Menu pAo 
            Caption         =   "Baddy tank"
            Index           =   3
         End
         Begin VB.Menu pAo 
            Caption         =   "Me (Camakarzi)"
            Index           =   4
            Visible         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "Tank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using data Types makes things alot easier, Especially in VB, where
'you have that Auto-finish thing...
Private Type Colour
  Red As Byte
  Green As Byte
  Blue As Byte
End Type

Private Type Co
  x As Double
  y As Double
  Z As Double
End Type

Private Type WorldDis
  Position As Co
  Angle As Co
  Type As Integer
  Scale As Co
  Direction As Integer
  Speed As Integer
  GibSpeed As Co
  Delay As Integer
  Visible As Boolean
  Health As Integer
  Alert As Boolean
  Owner As Integer
  PlayerName As String
  Target As Integer
  Dead As Boolean
End Type

Private Type ThreeD
  Shadow(20) As Co
  Coord(20) As Co
  Scale As Co
  Colour As Colour
  Poly As Byte
  Index As Integer
End Type

Private Type PInfo
  PPos As Co
  OPPos As Co
  PAng As Co
  CPos As Co
  CAng As Co
End Type

Private Type Corner
  Points(20) As Co
  EdgeCount As Byte
End Type
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim player(1 To 1) As PInfo
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Corner As Corner
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim totframes, starttime As Double
Dim unitinfo(20) As Double
Dim pi As Double
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim World(2500) As WorldDis
Dim TemporaryWorld(2500) As WorldDis
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim WorldOrder(2500, 2), ORDER(99, 2) As Single
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Rand(20, 2) As Single
Dim x1, x2, y1, y2 As Integer
Dim TotalFaces, Cu, totalunits As Integer
Dim Zeye, Col As Integer
Dim FaceCount, Xofs, Yofs As Integer
Dim XMouse, YMouse As Integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim Pyramid(5) As ThreeD
Dim Plate(2) As ThreeD
Dim Scrap(5) As ThreeD
Dim Cube(6) As ThreeD
Dim Tank(16) As ThreeD
Dim Mamoth(21) As ThreeD
Dim Plane(15) As ThreeD
Dim Level(21) As ThreeD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Dim MouseButton, ShiftButton As Boolean

Private Sub Form_Load()

  'Tanxs for looking at this, and I hope you can figure it out, and
  'learn somthing from it. It should be easier to understand than some
  'other 3D demos, cos it dosn't use smart-ass little VB specific
  'things, that no-one could figure out in a million years without a
  'degree in talking bollocks. This code will basicly work in any langauge
  'without any major redesigning. I wrote this first in QBasic, but VB has
  'the advantage of smooth screens
  '
  'AAnnyway..
  'To tell you quickly what happens, each shape is represented in the
  'world array, by a 3D (X,Y & Z) co-ordanate. All the co-ordanates are
  'rotated around 0,0,0 and sorted into order. Then it starts with the
  'furthest away from you, and depending what it is (Say its a pyramid), it
  'copies the pyramid data from Pyramid() to Level()
  '
  'It then rotates all the points of each face in the Level() and puts then
  'into order.(adding the Z distances of each point on a face and deviding
  'by the number of points, and soring these averages) It then turns the
  '3D (X,Y & Z) co-ordanate into a 2D co-ordanate that can be ploted on the
  'screen, and draws lines from one point to the next, making a shape!
  '
  'Once this has been done for all the faces in that shape, it goes back
  'to the world order(), and starts again with the next shape.
  
  FilePathName = ""
  
  For n = 1 To 20
   Rand(n, 1) = Rnd * 1
   Rand(n, 2) = Rnd * 1
  Next n
   
  pi = (22 / 7) * 18.2
  Xofs = View.ScaleWidth / 2
  Yofs = View.ScaleHeight / 2
  totalunits = 100
  Zeye = 800
  For n = 0 To 90
    World(n).Scale.x = 1
    World(n).Scale.y = 1
    World(n).Scale.Z = 1
    World(n).Visible = True
  Next n
   
  'On Error GoTo wrongpath
'ThatsBetter:
  
  Open FilePathName + Form1.lstLevelName.FileName For Input As #1
    Input #1, temp: Input #1, totalunits
    Do
    Input #1, temp
    If Mid(temp, 1, 1) <> "~" Then Input #1, Number
    temp = LCase(temp)
    If temp = "world" Then cw = Number
    If temp = "position.x" Then World(cw).Position.x = Number
    If temp = "position.y" Then World(cw).Position.y = Number
    If temp = "position.z" Then World(cw).Position.Z = Number
    If temp = "angle.x" Then World(cw).Angle.x = Number
    If temp = "angle.y" Then World(cw).Angle.y = Number
    If temp = "angle.z" Then World(cw).Angle.Z = Number
    If temp = "scale.x" Then World(cw).Scale.x = Number
    If temp = "scale.y" Then World(cw).Scale.y = Number
    If temp = "scale.z" Then World(cw).Scale.Z = Number
    If temp = "type" Then World(cw).Type = Number
    If temp = "alert" Then World(cw).Alert = Number
    If temp = "health" Then World(cw).Health = Number
    If temp = "target" Then World(cw).Target = Number
    If temp = "direction" Then World(cw).Direction = Number
    If temp = "player" Then World(cw).PlayerName = Number
    Loop While EOF(1) = False
  Close
  
  player(1).PPos.y = World(1).Position.Z
  player(1).PPos.x = World(1).Position.x
  player(1).CPos.Z = -50
  
  
  XMouse = Xofs: YMouse = Yofs
  
  '------    Load Shape from file   -------
  Open FilePathName + "pyramid.3d" For Input As #1
    PyramidFaces = 5
    For M = 1 To PyramidFaces
      Input #1, Pyramid(M).Poly
      Input #1, Pyramid(M).Colour.Red
      Input #1, Pyramid(M).Colour.Green
      Input #1, Pyramid(M).Colour.Blue
      For n = 1 To Pyramid(M).Poly
        Input #1, Pyramid(M).Coord(n).x
        Input #1, Pyramid(M).Coord(n).y
        Input #1, Pyramid(M).Coord(n).Z
      Next n
    Next M
  Close
  '-------------------------------------------
  
  '------    Load Shape from file   -------
  Open FilePathName + "Cube.3d" For Input As #1
  CubeFaces = 6
  For M = 1 To CubeFaces
    Input #1, Cube(M).Poly
    Input #1, Cube(M).Colour.Red
    Input #1, Cube(M).Colour.Green
    Input #1, Cube(M).Colour.Blue
    For n = 1 To Cube(M).Poly
      Input #1, Cube(M).Coord(n).x
      Input #1, Cube(M).Coord(n).y
      Input #1, Cube(M).Coord(n).Z
    Next n
  Next M: Close
  '-------------------------------------------
  
  '------    Load Shape from file   -------
  Open FilePathName + "tank.3d" For Input As #1
    TankFaces = 16
    For M = 1 To TankFaces
      Input #1, Tank(M).Poly
      Input #1, Tank(M).Colour.Red
      Input #1, Tank(M).Colour.Green
      Input #1, Tank(M).Colour.Blue
      For n = 1 To Tank(M).Poly
        Input #1, Tank(M).Coord(n).Z
        Input #1, Tank(M).Coord(n).y
        Input #1, Tank(M).Coord(n).x
        Tank(M).Coord(n).x = Tank(M).Coord(n).x * 2
        Tank(M).Coord(n).y = Tank(M).Coord(n).y * 2
        Tank(M).Coord(n).Z = Tank(M).Coord(n).Z * -2
      Next n
    Next M
  Close
  '-------------------------------------------
  
  
  '------    Load Shape from file   -------
  Open FilePathName + "mamoth.3d" For Input As #1
  MamothFaces = 21
    For M = 1 To MamothFaces
      Input #1, Mamoth(M).Poly
      Input #1, Mamoth(M).Colour.Red
      Input #1, Mamoth(M).Colour.Green
      Input #1, Mamoth(M).Colour.Blue
      For n = 1 To Mamoth(M).Poly
        Input #1, Mamoth(M).Coord(n).Z
        Input #1, Mamoth(M).Coord(n).y
        Input #1, Mamoth(M).Coord(n).x
        Mamoth(M).Coord(n).x = Mamoth(M).Coord(n).x * 2
        Mamoth(M).Coord(n).y = Mamoth(M).Coord(n).y * 2
        Mamoth(M).Coord(n).Z = Mamoth(M).Coord(n).Z * -2
      Next n
    Next M
  Close
  '-------------------------------------------
  
  '------    Load Shape from file   -------
  Open FilePathName + "Scrap.3d" For Input As #1
    ScrapFace = 4
    For M = 1 To ScrapFace
      Input #1, Scrap(M).Poly
      Input #1, Scrap(M).Colour.Red
      Input #1, Scrap(M).Colour.Green
      Input #1, Scrap(M).Colour.Blue
      For n = 1 To Scrap(M).Poly
        Input #1, Scrap(M).Coord(n).Z
        Input #1, Scrap(M).Coord(n).y
        Input #1, Scrap(M).Coord(n).x
      Next n
    Next M
  Close
  '-------------------------------------------
  
  
  
  '------    Load Shape from file   -------
  Open FilePathName + "Plate.3d" For Input As #1
    PlateFace = 2
    For M = 1 To PlateFace
      Input #1, Plate(M).Poly
      Input #1, Plate(M).Colour.Red
      Input #1, Plate(M).Colour.Green
      Input #1, Plate(M).Colour.Blue
      For n = 1 To Plate(M).Poly
        Input #1, Plate(M).Coord(n).Z
        Input #1, Plate(M).Coord(n).y
        Input #1, Plate(M).Coord(n).x
        Next n
    Next M
  Close
  '-------------------------------------------
  
  '------    Load Shape from file   -------
  Open FilePathName + "Plane.3d" For Input As #1
    PlaneFaces = 15
    For M = 1 To PlaneFaces
      Input #1, Plane(M).Poly
      Input #1, Plane(M).Colour.Red
      Input #1, Plane(M).Colour.Green
      Input #1, Plane(M).Colour.Blue
      For n = 1 To Plane(M).Poly
        Input #1, Plane(M).Coord(n).Z
        Input #1, Plane(M).Coord(n).y
        Input #1, Plane(M).Coord(n).x
        Plane(M).Coord(n).x = Plane(M).Coord(n).x * 20
        Plane(M).Coord(n).y = Plane(M).Coord(n).y * 20
        Plane(M).Coord(n).Z = Plane(M).Coord(n).Z * -20
      Next n
    Next M
  Close
  
  For n = 0 To totalunits
    WorldOrder(n, 2) = 0
    WorldOrder(n, 1) = n
  Next n
  
  '-------------------------------------------
'  Exit Sub
  
'wrongpath:
'  x = MsgBox("The 3D files could not be found. Do you want to find them yourself?", vbYesNo, "Mising files")
'  If x = 7 Then End
'  FilePathName = InputBox("Enter the location of the *.3d files..", , FilePathName)
'  Resume ThatsBetter

End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub o1_Click(Index As Integer)
  ' o1 is a popup menu
  If Index = 1 Then player(1).CAng.y = 270   ' these bits change the own view angle
  If Index = 2 Then player(1).CAng.y = 90
  If Index = 3 Then player(1).CAng.y = 0
  If Index = 4 Then player(1).CAng.y = 180
  If Index = 5 Or Index = 6 Then
    If o1(Index).Checked = True Then
      o1(Index).Checked = False
    Else
      o1(Index).Checked = True
    End If
  End If
  World(1).Visible = False
  If o1(6).Checked = False Then World(1).Visible = True ' o1(6) is own tank menu option
  
  ' RADAR options
  If Index = 7 Then
    o1(7).Checked = True
    o1(8).Checked = False
    Radar.DrawWidth = 4
  End If
  
  If Index = 8 Then
    o1(8).Checked = True
    o1(7).Checked = False
    Radar.DrawWidth = 2
  End If
  
  
End Sub

Private Sub p1o_Click(Index As Integer)
  ' plane one attack ....
  If Index = 1 Then World(3).Target = 6
  If Index = 3 Then World(3).Target = 2
  If Index = 2 Then World(3).Target = 8
  If Index = 4 Then World(3).Target = 1
End Sub

Private Sub p2o_Click(Index As Integer)
  ' plane two attack ....
  If Index = 1 Then World(4).Target = 6
  If Index = 3 Then World(4).Target = 2
  If Index = 2 Then World(4).Target = 8
  If Index = 4 Then World(4).Target = 1
End Sub

Private Sub pAo_Click(Index As Integer)
  ' All attack ....
  If Index = 1 Then World(3).Target = 6
  If Index = 1 Then World(4).Target = 6
  If Index = 1 Then World(10).Target = 6
  If Index = 1 Then World(11).Target = 6
  If Index = 1 Then World(12).Target = 6
  
  If Index = 3 Then World(3).Target = 2
  If Index = 3 Then World(4).Target = 2
  If Index = 3 Then World(10).Target = 2
  If Index = 3 Then World(11).Target = 2
  If Index = 3 Then World(12).Target = 2
  
  If Index = 2 Then World(3).Target = 8
  If Index = 2 Then World(4).Target = 8
  If Index = 2 Then World(10).Target = 8
  If Index = 2 Then World(11).Target = 8
  If Index = 2 Then World(12).Target = 8
  
  If Index = 4 Then World(3).Target = 1
  If Index = 4 Then World(4).Target = 1
  If Index = 4 Then World(10).Target = 1
  If Index = 4 Then World(11).Target = 1
  If Index = 4 Then World(12).Target = 1
End Sub

Private Sub ReMessage_Timer()
  ReMessage.Interval = 5000
  If Messages.ListCount > 0 Then Messages.RemoveItem (0)
  If Messages.ListCount > 20 Then ReMessage.Interval = 100
End Sub


Private Sub Timer1_Timer()
  ' this fires every 10 milliseconds.. well tries to anyway
  Const sd = 197
  If Pause.Value = 1 Then Exit Sub
  player(1).OPPos.x = player(1).PPos.x
  player(1).OPPos.y = player(1).PPos.y
  player(1).OPPos.Z = player(1).PPos.Z
  
  '                          Views Brakes
  If ShiftButton = False And o1(5).Checked = False Then 'And World(1).Dead = False Then
    tangle = player(1).PAng.y
    If XMouse < 50 Then tangle = tangle + 4
    If XMouse > View.ScaleWidth - 50 Then tangle = tangle - 4
    If XMouse < 100 Then tangle = tangle + 1
    If XMouse > View.ScaleWidth - 100 Then tangle = tangle - 1
    If tangle < 0 Then tangle = tangle + 360
    If tangle > 359 Then tangle = tangle - 360
    player(1).PAng.y = tangle
    x = Sin(player(1).PAng.y / pi)
    y = Cos(player(1).PAng.y / pi)
    x = x * 50: y = y * 50
    If YMouse > 100 Then
      player(1).PPos.x = player(1).PPos.x - x
      player(1).PPos.y = player(1).PPos.y - y
    End If
    If YMouse > 50 Then
      player(1).PPos.x = player(1).PPos.x - x
      player(1).PPos.y = player(1).PPos.y - y
    End If
    If YMouse < View.ScaleHeight - 100 Then
      player(1).PPos.x = player(1).PPos.x + x
      player(1).PPos.y = player(1).PPos.y + y
    End If
    If YMouse < View.ScaleHeight - 50 Then
      player(1).PPos.x = player(1).PPos.x + x
      player(1).PPos.y = player(1).PPos.y + y
    End If
    World(1).Position.x = -player(1).PPos.x
    World(1).Position.Z = -player(1).PPos.y
    World(1).Angle.y = -player(1).PAng.y / pi
  Else
    tangle = player(1).CAng.y
    If XMouse < 50 Then tangle = tangle + 5
    If XMouse > View.ScaleWidth - 50 Then tangle = tangle - 5
    If XMouse < 100 Then tangle = tangle + 5
    If XMouse > View.ScaleWidth - 100 Then tangle = tangle - 5
    If tangle < 0 Then tangle = tangle + 360
    If tangle > 359 Then tangle = tangle - 360
    player(1).CAng.y = tangle
    If YMouse > 50 Then
      player(1).CPos.Z = player(1).CPos.Z + 10
      If player(1).CPos.Z > 100 Then player(1).CPos.Z = 100
    End If
    If YMouse < View.ScaleHeight - 50 Then
      player(1).CPos.Z = player(1).CPos.Z - 10
    End If
  End If
  
'  For Ai = 1 To totalunits
'    Call AIHandler(Ai)
'  Next Ai
  
  'This bit does not actually do any 3D, it just calls 'WorldOrderRender',
  'but it has loads of other rubish around it, that do times and
  'print locations and stuff like that. And Clears the screen. Important.
  
  
  '---------------  Draw the Sky and Ground  ------------------
  For n = 0 To Yofs
    View.Line (0, n)-(View.ScaleWidth, n), RGB(n + 50, n + 50, 255)
  Next n
  For n = Yofs To View.ScaleHeight
    View.Line (0, n)-(View.ScaleWidth, n), RGB(200 - (n - Yofs), 200, 200 - (n - Yofs))
  Next n
  '----------  Start and time the level renderer  ------------
  List1.Clear
  Stime = Timer
  FaceCount = 0
  Call WorldOrderRender
  Etime = Timer
  ttime = Etime - Stime
  'List1.AddItem "Time taken " + Str(ttime)
  If ttime <> 0 Then List1.AddItem "Frames/Second " + Str(1 / ttime)
  'If ttime = 0 Then List1.AddItem "-------"
  'List1.AddItem "Total Faces " + Str(FaceCount)
  'List1.AddItem "Your health : " + Str(World(1).Health)
  List1.AddItem "                Time left :" + Str(30 - Timer2.Tag)
  
  x1 = World(1).Position.x
  x2 = World(1).Position.y
  x3 = World(2).Position.x
  x4 = World(2).Position.y
  
  List1.AddItem "World(1) x " & Format(World(1).Position.x, "0000.000") & " y " & Format(World(1).Position.y, "0000.000") & " z " & Format(World(1).Position.Z, "0000.000")
  '---------------  Draw the radar view  -----------------
  
'  Dim DirName As String
'  Dim DirLoc As Byte
'  Dim range As Integer
'  Radar.Cls
'  range = 300
'  If o1(7).Checked = True Then range = 100
'  For n = 1 To totalunits
'    If World(n).Type <> 0 Then
'      xxx = (TemporaryWorld(n).Position.x / range) + (Radar.ScaleWidth / 2)
'      yyy = (TemporaryWorld(n).Position.Z / range) + (Radar.ScaleHeight / 2)
'      Select Case World(n).Type
'        Case 6, 7
'          Radar.PSet (xxx, yyy), RGB(255, 0, 255)
'        Case 3
'          Radar.PSet (xxx, yyy), RGB(255, 0, 0)
'        Case 4, 8
'          Radar.PSet (xxx, yyy), RGB(0, 200, 100)
'        Case 5
'        Case 11
'          Radar.PSet (xxx, yyy), RGB(155, 100, 0)
'        Case Else
'          Radar.PSet (xxx, yyy), 0
'      End Select
'    End If
'  Next n
  '-------------------  Draw the compass  ---------------------
  View.Circle (55, 55), 40, 0
  For n = 1 To 360 Step 45
    x = Sin((n - player(1).PAng.y) / pi) * 20
    y = Cos((n - player(1).PAng.y) / pi) * 20
    View.Line (55 + x, 55 + y)-(55 + x + x, 55 + y + y), 0
    DirName = ""
    If DirLoc = 0 Then DirName = "S"
    If DirLoc = 1 Then DirName = "SE"
    If DirLoc = 2 Then DirName = "E"
    If DirLoc = 3 Then DirName = "NE"
    If DirLoc = 4 Then DirName = "N"
    If DirLoc = 5 Then DirName = "NW"
    If DirLoc = 6 Then DirName = "W"
    If DirLoc = 7 Then DirName = "SW"
    View.PSet (55 + x * 3, 55 + y * 3)
    View.Print DirName
    DirLoc = DirLoc + 1
  Next n
  View.DrawStyle = 2:
  '-----  Draw the speed boxes and print the messages  -------
  View.Line (50, 50)-(View.ScaleWidth - 50, View.ScaleHeight - 50), RGB(sd, sd, sd), B
  View.Line (100, 100)-(View.ScaleWidth - 100, View.ScaleHeight - 100), RGB(sd, sd, sd), B
  View.DrawStyle = 0
  View.PSet (0, 0)
'  View.FontSize = 12
'  For n = 0 To Messages.ListCount
'    View.Print Messages.List(n)
'  Next n
  View.FontSize = 7
  
'  If Hlp.Value = 1 Then
'    View.FontSize = 10
'    View.PSet (350, 20): View.Print "Move the mouse here to go forward"
'    View.PSet (350, 220): View.Print "Move the mouse here to go backwards"
'    View.PSet (50, 120): View.Print "Move the mouse here to go left"
'    View.PSet (500, 120): View.Print "Move the mouse here to go right"
'  End If
End Sub

Private Sub cmdLeft_Click()
  World(1).Position.x = World(1).Position.x - 1
End Sub

Private Sub cmdRight_Click()
  World(1).Position.x = World(1).Position.x + 1
End Sub

'This is the start of the actual 3D engine. The subs are in the order
'that they get used. Ie. Starts with WorldOrderRender, to sort all
'the shapes into order. Then each shape is drawn one at a time, in
'DrawScreen. The shape is rotated in the small subs at the bottom,
'and DrawLine is just for filling solid faces, which is sloooow.
'
'Note that any of the maths stuff is not mine, but uses basic
'3D maths stuff, from any TextBook
'   3D to 2D translation
'   Hidden face removal
'   Rotations
'are used here, and I don't really understand how they work, but I
'know then do, and so I use them. Greater understanding will be needed
'for more complex things like Phong Shading and texture maps.

Sub WorldOrderRender()
  Dim startdis As Double
  For n = 0 To totalunits
    TemporaryWorld(n).Position.x = World(n).Position.x + player(1).PPos.x
    TemporaryWorld(n).Position.y = World(n).Position.y + 100
    TemporaryWorld(n).Position.Z = World(n).Position.Z + player(1).PPos.y
  Next n
  viewangle2 = (player(1).PAng.y + player(1).CAng.y) / pi
  GoSub WorldRotateY
'  For n = 0 To totalunits
'    WorldOrder(n, 2) = TemporaryWorld(n).Position.Z
'    WorldOrder(n, 1) = n
'  Next n

  For n = 0 To totalunits
    startdis = WorldOrder(n, 2)
    lok = n
    For M = n To totalunits
      If WorldOrder(M, 2) < startdis Then
        lok = M
        startdis = WorldOrder(M, 2)
      End If
    Next M
    temp = WorldOrder(n, 1)
    WorldOrder(n, 1) = WorldOrder(lok, 1)
    WorldOrder(lok, 1) = temp
    temp = WorldOrder(n, 2)
    WorldOrder(n, 2) = WorldOrder(lok, 2)
    WorldOrder(lok, 2) = temp
  Next n
  
  For M = 0 To totalunits
    Cu = WorldOrder(M, 1)
    unitinfo(0) = World(Cu).Type
    If unitinfo(0) <> 0 And World(Cu).Visible = True Then
      unitinfo(1) = TemporaryWorld(Cu).Position.x
      unitinfo(2) = TemporaryWorld(Cu).Position.y - player(1).CPos.Z
      unitinfo(3) = -TemporaryWorld(Cu).Position.Z
      unitinfo(5) = viewangle1 + World(Cu).Angle.x
      unitinfo(6) = viewangle2 + World(Cu).Angle.y
      unitinfo(7) = viewangle3 + World(Cu).Angle.Z
      unitinfo(8) = World(Cu).Scale.x
      unitinfo(9) = World(Cu).Scale.y
      unitinfo(10) = World(Cu).Scale.Z
      unitinfo(11) = Cu
      Call DrawScreen
    End If
  Next M
  Exit Sub

  '==========================================
WorldRotateY:
  For Cu = 0 To totalunits
    M = 1
    x = TemporaryWorld(Cu).Position.x
    y = TemporaryWorld(Cu).Position.y
    Z = TemporaryWorld(Cu).Position.Z
    XROTATED = Cos(viewangle2) * x - Sin(viewangle2) * Z
    YROTATED = y
    ZROTATED = Sin(viewangle2) * x + Cos(viewangle2) * Z
    TemporaryWorld(Cu).Position.x = XROTATED
    TemporaryWorld(Cu).Position.y = YROTATED
    TemporaryWorld(Cu).Position.Z = ZROTATED
  Next Cu
  Return
  '==========================================
End Sub

'First, depending on what the shape we want to draw is, we copy the data
'from the main array to a temporary array named level. Its then rotated
'around each axis. After that, each face is put into order, and then
'drawn, starting with the furtherest away. This doesnt work all the time,
'but with hidden face removal on, you'll hardly notice the mistakes.
'I'm not sure if its truly possible to get a perfect orderer without
'using a ZBuffer thing or not...

Sub DrawScreen()
  Dim n, M As Integer
  Dim lx, ly, lz, temp As Single
  
Select Case unitinfo(0)
    
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 1
  TotalFaces = 5
  For M = 1 To TotalFaces
   Level(M).Poly = Pyramid(M).Poly
   Level(M).Colour.Red = Pyramid(M).Colour.Red
   Level(M).Colour.Green = Pyramid(M).Colour.Green
   Level(M).Colour.Blue = Pyramid(M).Colour.Blue
   For n = 1 To Level(M).Poly
    Level(M).Coord(n).x = Pyramid(M).Coord(n).x * unitinfo(8)
    Level(M).Coord(n).y = Pyramid(M).Coord(n).y * unitinfo(9)
    Level(M).Coord(n).Z = Pyramid(M).Coord(n).Z * unitinfo(10)
   Next n
  Next M
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 2
  TotalFaces = 6
  For M = 1 To TotalFaces
   Level(M).Poly = Cube(M).Poly
   Level(M).Colour.Red = Cube(M).Colour.Red
   Level(M).Colour.Green = Cube(M).Colour.Green
   Level(M).Colour.Blue = Cube(M).Colour.Blue
   For n = 1 To Level(M).Poly
    Level(M).Coord(n).x = Cube(M).Coord(n).x * unitinfo(8)
    Level(M).Coord(n).y = Cube(M).Coord(n).y * unitinfo(9)
    Level(M).Coord(n).Z = Cube(M).Coord(n).Z * unitinfo(10)
   Next n
  Next M
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 Case 3
     x = unitinfo(1) - 5
     y = unitinfo(2) - 5
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx1 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy1 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     x = unitinfo(1) + 5
     y = unitinfo(2) + 5
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx2 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy2 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     For n = 1 To Abs(sx1 - sx2) Step 0.5
     View.Circle (sx1 + 2, sy1 + 2), n, RGB(255, 40, 40)
     Next n
     Exit Sub
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 Case 4
     TotalFaces = 16
     For M = 1 To TotalFaces
      Level(M).Poly = Tank(M).Poly
     Level(M).Colour.Red = Tank(M).Colour.Red
      Level(M).Colour.Green = Tank(M).Colour.Green
      Level(M).Colour.Blue = Tank(M).Colour.Blue
      For n = 1 To Level(M).Poly
       Level(M).Coord(n).x = Tank(M).Coord(n).x * unitinfo(8)
       Level(M).Coord(n).y = Tank(M).Coord(n).y * unitinfo(9)
       Level(M).Coord(n).Z = Tank(M).Coord(n).Z * unitinfo(10)
      Next n
     Next M
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Case 5, 10
    For n = 1 To 5
     xo = Rand(n, 1) * 40 - World(unitinfo(11)).Delay
     yo = Rand(n, 2) * 40 - World(unitinfo(11)).Delay
     x = unitinfo(1) - 3 + xo
     y = unitinfo(2) - 3 + yo
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx1 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy1 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     x = unitinfo(1) + 3 + xo
     y = unitinfo(2) + 3 + yo
     Z = -unitinfo(3)
     sx2 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy2 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     View.Line (sx1, sy1)-(sx2, sy2), RGB(0, 0, 0), BF
     Next n
     Exit Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 6
     TotalFaces = 4
     For M = 1 To TotalFaces
      Level(M).Poly = Scrap(M).Poly
      Level(M).Colour.Red = Scrap(M).Colour.Red
      Level(M).Colour.Green = Scrap(M).Colour.Green
      Level(M).Colour.Blue = Scrap(M).Colour.Blue
      For n = 1 To Level(M).Poly
       Level(M).Coord(n).x = Scrap(M).Coord(n).x * unitinfo(8)
       Level(M).Coord(n).y = Scrap(M).Coord(n).y * unitinfo(9)
       Level(M).Coord(n).Z = Scrap(M).Coord(n).Z * unitinfo(10)
      Next n
     Next M
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 7
     TotalFaces = 2
     For M = 1 To TotalFaces
      Level(M).Poly = Plate(M).Poly
      Level(M).Colour.Red = Plate(M).Colour.Red
      Level(M).Colour.Green = Plate(M).Colour.Green
      Level(M).Colour.Blue = Plate(M).Colour.Blue
      For n = 1 To Level(M).Poly
       Level(M).Coord(n).x = Plate(M).Coord(n).x * unitinfo(8)
       Level(M).Coord(n).y = Plate(M).Coord(n).y * unitinfo(9)
       Level(M).Coord(n).Z = Plate(M).Coord(n).Z * unitinfo(10)
      Next n
     Next M
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 8
     TotalFaces = 21
     For M = 1 To TotalFaces
      Level(M).Poly = Mamoth(M).Poly
      Level(M).Colour.Red = Mamoth(M).Colour.Red
      Level(M).Colour.Green = Mamoth(M).Colour.Green
      Level(M).Colour.Blue = Mamoth(M).Colour.Blue
      For n = 1 To Level(M).Poly
       Level(M).Coord(n).x = Mamoth(M).Coord(n).x * unitinfo(8)
       Level(M).Coord(n).y = Mamoth(M).Coord(n).y * unitinfo(9)
       Level(M).Coord(n).Z = Mamoth(M).Coord(n).Z * unitinfo(10)
      Next n
     Next M
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Case 9
     x = unitinfo(1)
     y = unitinfo(2)
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy = Yofs + Int(y * (Zeye / (Zeye - Z)))
     For n = 1 To 30 Step 0.5
     View.Circle (sx, sy), n, RGB(255, 255, 0)
     Next n
     For n = 1 To 360 Step 36
     x = Sin((n + (Timer * 10)) / pi) * 50
     y = Cos((n + (Timer * 10)) / pi) * 50
     View.Line (sx, sy)-(sx + x, sy + y), RGB(255, 255, 0)
     Next n
     Exit Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Case 11
     TotalFaces = 15
     For M = 1 To TotalFaces
      Level(M).Poly = Plane(M).Poly
      Level(M).Colour.Red = Plane(M).Colour.Red
      Level(M).Colour.Green = Plane(M).Colour.Green
      Level(M).Colour.Blue = Plane(M).Colour.Blue
      For n = 1 To Level(M).Poly
       Level(M).Coord(n).x = Plane(M).Coord(n).x * unitinfo(8)
       Level(M).Coord(n).y = Plane(M).Coord(n).y * unitinfo(9)
       Level(M).Coord(n).Z = Plane(M).Coord(n).Z * unitinfo(10)
      Next n
     Next M
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 Case 12
     x = unitinfo(1) - 5
     y = unitinfo(2) - 5
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx1 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy1 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     x = unitinfo(1) + 5
     y = unitinfo(2) + 5
     Z = -unitinfo(3)
     If Z > 700 Then Exit Sub
     sx2 = Xofs + Int(x * (Zeye / (Zeye - Z)))
     sy2 = Yofs + Int(y * (Zeye / (Zeye - Z)))
     For n = 1 To Abs(sx1 - sx2) Step 0.5
     View.Circle (sx1 + 2, sy1 + 2), n, RGB(255, 255, 255)
     Next n
     Exit Sub
 
 
 End Select
 
 Call RotateX
 Call RotateY
 Call RotateZ
 
 For n = 1 To TotalFaces
  ORDER(n, 1) = n: ZDistance = 0
  For M = 1 To Level(n).Poly
    ZDistance = ZDistance + Level(n).Coord(M).Z
  Next M
  ZDistance = ZDistance / Level(n).Poly
  ORDER(n, 2) = ZDistance
 Next n
 
 Dim startdis As Single
 
 
 For n = 1 To TotalFaces
  startdis = ORDER(n, 2)
  lok = n
  For M = n To TotalFaces
   If ORDER(M, 2) < startdis Then
    lok = M
    startdis = ORDER(M, 2)
   End If
  Next M
  temp = ORDER(n, 2)
  ORDER(n, 2) = ORDER(lok, 2)
  ORDER(lok, 2) = temp
  temp = ORDER(n, 1)
  ORDER(n, 1) = ORDER(lok, 1)
  ORDER(lok, 1) = temp
 Next n
 

'###################################
'####          Shadow          #####
'###################################
' If Shadow.Value = 1 Then
'  For shd = 1 To TotalFaces
'   For sc = 1 To Level(shd).Poly
'
'    temp = -Level(shd).Coord(sc).y + (100 - player(1).CPos.Z) - unitinfo(2)
'    lx = TemporaryWorld(0).Position.x
'    ly = TemporaryWorld(0).Position.y
'    lz = TemporaryWorld(0).Position.Z
'
'    x = (Level(shd).Coord(sc).x + unitinfo(1)) + (temp * lx) \ ly
'    y = 100 - player(1).CPos.Z
'    Z = (Level(shd).Coord(sc).Z - unitinfo(3)) + (temp * lz) \ ly
'
'     If Z > 700 Then Z = 700
'    Corner.Points(sc).x = Xofs + Int(x * (Zeye / (Zeye - Z)))
'    Corner.Points(sc).y = Yofs + Int(y * (Zeye / (Zeye - Z)))
'   Next sc
'   Col = -1
'   Corner.EdgeCount = Level(shd).Poly
'   FaceCount = FaceCount + 1
'   Call DrawPolygon
'NotBotherWithShadowz:
'  Next shd
' End If
'###################################
'####          Shadow          #####
'###################################

 
For tempnum = 1 To TotalFaces
  CURFACE = ORDER(tempnum, 1)
  Col = CURFACE

  For n = 1 To Level(CURFACE).Poly
     x = Level(CURFACE).Coord(n).x + unitinfo(1)
     y = Level(CURFACE).Coord(n).y + unitinfo(2)
     Z = Level(CURFACE).Coord(n).Z - unitinfo(3)
      If Z > 700 Then GoTo FaceBehindCamera
     Corner.Points(n).x = Xofs + Int(x * (Zeye / (Zeye - Z)))
     Corner.Points(n).y = Yofs + Int(y * (Zeye / (Zeye - Z)))
   
   Next n
   
   'Hidden Face removal.
   'Take three points on you polygon, and detemin if they go clockwise
   'or anti-clockwise on the screen. If you define your shapes with
   'all faces going clockwise, any anti-clockwise faces must but facing
   'away from you, and so are hidden, and don't need to be drawn.
   
   If HFace.Value = 1 Then
     xx1 = Corner.Points(1).x
     xx2 = Corner.Points(2).x
     xx3 = Corner.Points(3).x
     yy1 = Corner.Points(1).y
     yy2 = Corner.Points(2).y
     yy3 = Corner.Points(3).y
     Normal = ((yy1 - yy3) * (xx2 - xx1) - (xx1 - xx3) * (yy2 - yy1))
     If Normal > 0 Then GoTo FaceBehindCamera
   End If

   FaceCount = FaceCount + 1

  If wire.Value = 1 Then
   Corner.EdgeCount = Level(CURFACE).Poly
   Call DrawPolygon
  End If

  If ckB.Value = 1 Then
  For n = 1 To Level(CURFACE).Poly
  If n = Level(CURFACE).Poly Then
      x1 = Corner.Points(n).x
      y1 = Corner.Points(n).y
      x2 = Corner.Points(1).x
      y2 = Corner.Points(1).y
      View.Line (x1, y1)-(x2, y2), View.ForeColor
  Else
     x1 = Corner.Points(n).x
     y1 = Corner.Points(n).y
     x2 = Corner.Points(n + 1).x
     y2 = Corner.Points(n + 1).y
     View.Line (x1, y1)-(x2, y2), View.ForeColor
  End If
  Next n
  End If

FaceBehindCamera:
Next tempnum

Exit Sub
ProblemPerspective:
Resume FaceBehindCamera
 
End Sub

Sub DrawPolygon()
 'Lets say the screen is 200 pixels high. Therfore, if we draw 200 lines,
 'we can fill the whole screen up. Depending on where we start and stop
 'each line, we can produce any convex shape we want.
 'This bit takes the array StartStop() and by scanning along each line
 'of the face, fills in the start and stopping point of each horizontal
 'line.
 
Dim nn As Integer
Dim StartStop(300, 1 To 3) As Long
 
 For nn = 1 To Corner.EdgeCount
  If nn = Corner.EdgeCount Then
      x1 = Corner.Points(nn).x
      y1 = Corner.Points(nn).y
      x2 = Corner.Points(1).x
      y2 = Corner.Points(1).y
   GoSub DrawLine
  Else
      x1 = Corner.Points(nn).x
      y1 = Corner.Points(nn).y
      x2 = Corner.Points(nn + 1).x
      y2 = Corner.Points(nn + 1).y
   GoSub DrawLine
  End If
 Next nn
Exit Sub
 
DrawLine:

 mox = x2 - x1: moy = y2 - y1
 If moy <> 0 Then
  Disd = mox / moy
 Else
  Return
 End If

 If y1 < y2 Then Negpos = 1:  Disd = Disd
 If y1 > y2 Then Negpos = -1: Disd = -Disd
 Tag = y1: XX = x1
 For n = y1 To y2 Step Negpos
  If n > -1 And n < View.ScaleHeight Then
   If StartStop(n, 3) = 1 Then
    StartStop(n, 3) = 1
     If Col = -1 Then
      View.Line (StartStop(n, 1), n)-(XX, n), RGB(60, 140, 60)
     Else
      View.Line (StartStop(n, 1), n)-(XX, n), RGB(Level(Col).Colour.Red, Level(Col).Colour.Green, Level(Col).Colour.Blue)
     End If
    Else
    StartStop(n, 3) = 1
    StartStop(n, 1) = XX
   End If
  End If
  XX = XX + Disd
 Next n
 
 Return
End Sub



Sub RotateZ()
  Dim Cu, M As Integer
  For Cu = 1 To TotalFaces
    For M = 1 To 4
      x = Level(Cu).Coord(M).x
      y = Level(Cu).Coord(M).y
      Z = Level(Cu).Coord(M).Z
      XROTATED = Cos(unitinfo(5)) * x - Sin(unitinfo(5)) * y
      YROTATED = Sin(unitinfo(5)) * x + Cos(unitinfo(5)) * y
      ZROTATED = Z
      Level(Cu).Coord(M).x = XROTATED
      Level(Cu).Coord(M).y = YROTATED
      Level(Cu).Coord(M).Z = ZROTATED
    Next M
  Next Cu
End Sub

Sub RotateY()
  Dim Cu, M As Integer
  For Cu = 1 To TotalFaces
    For M = 1 To 4
      x = Level(Cu).Coord(M).x
      y = Level(Cu).Coord(M).y
      Z = Level(Cu).Coord(M).Z
      XROTATED = Cos(unitinfo(6)) * x - Sin(unitinfo(6)) * Z
      YROTATED = y
      ZROTATED = Sin(unitinfo(6)) * x + Cos(unitinfo(6)) * Z
      Level(Cu).Coord(M).x = XROTATED
      Level(Cu).Coord(M).y = YROTATED
      Level(Cu).Coord(M).Z = ZROTATED
    Next M
  Next Cu
End Sub

Sub RotateX()
  Dim Cu, M As Integer
  For Cu = 1 To TotalFaces
    For M = 1 To 4
      x = Level(Cu).Coord(M).x
      y = Level(Cu).Coord(M).y
      Z = Level(Cu).Coord(M).Z
      XROTATED = x
      YROTATED = Cos(unitinfo(7)) * y - Sin(unitinfo(7)) * Z
      ZROTATED = Sin(unitinfo(7)) * y + Cos(unitinfo(7)) * Z
      Level(Cu).Coord(M).x = XROTATED
      Level(Cu).Coord(M).y = YROTATED
      Level(Cu).Coord(M).Z = ZROTATED
    Next M
  Next Cu
End Sub

Private Sub Timer2_Timer()
'  Timer2.Tag = Timer2.Tag + 1
'  If Timer2.Tag = 30 Then
'    a = a + "Ha Ha!!" + Chr$(13)
'    a = a + "Your have been running the demo of 3D tanx" + Chr$(13)
'    a = a + "" + Chr$(13)
'    a = a + "There, wasn't that fun?" + Chr$(13)
'    a = a + "However, thats all your going to get, till the" + Chr$(13)
'    a = a + "full version is made, and put on sale, only £999.99!" + Chr$(13)
'    a = a + "If you want to play for more that 30 seconds, and have teams" + Chr$(13)
'    a = a + "and command your tanks and stuff like in UnReal, then start" + Chr$(13)
'    a = a + "saving those pennies now!!" + Chr$(13)
'
'    MsgBox a, 48, "Demo Version"
'    End
'  End If
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'  Dim Counta As Byte
'  If Button = 1 Then
'    For n = 1 To totalunits
'      If World(n).Type = 0 Then
'        Counta = Counta + 1
'
'        x = Sin(player(1).PAng.y)
'        y = Cos(player(1).PAng.y)
'        x = x * 30: y = y * 30
'
'        If Counta = 2 Then y = -y: x = -x
'
'        World(n).Position.x = -player(1).PPos.x + y
'        World(n).Position.Z = -player(1).PPos.y + x
'        World(n).Position.y = -70
'        World(n).Angle.y = -player(1).PAng.y / pi
'
'        World(n).Type = 3
'        World(n).Speed = 1
'        World(n).Delay = 40
'        World(n).Owner = 1
'        If Counta = 2 Then Exit For
'      End If
'     Next n
'  End If
  If Button = 2 Then PopupMenu m1
End Sub
Private Sub View_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  XMouse = x '   701 max   scalewidth
  YMouse = y '   276 max   scaleheight
  ShiftButton = Shift
End Sub

Sub AIHandler(Ai)

  Dim x, y As Single
  
  '=====================  Reduce Timer  =========================
  If World(Ai).Delay > 0 Then
    World(Ai).Delay = World(Ai).Delay - 1
    If World(Ai).Delay = 0 Then World(Ai).Type = 0
  End If
  '==============================================================
  Select Case World(Ai).Type
  '=====================     Tanks   ==========================
'  Case 4, 8
'    If World(Ai).Alert = True Then
'      TXX = World(World(Ai).Target).Position.x - World(Ai).Position.x
'      Tyy = World(World(Ai).Target).Position.Z - World(Ai).Position.Z
'
'      PAngle = (-World(Ai).Angle.y)
'      Angle = FindAngle(TXX, Tyy, PAngle)
'
'      If Angle > 0 Then World(Ai).Angle.y = World(Ai).Angle.y - 4 / pi
'      If Angle < 0 Then World(Ai).Angle.y = World(Ai).Angle.y + 4 / pi
'
'      x1 = World(World(Ai).Target).Position.x
'      x2 = World(World(Ai).Target).Position.Z
'      x3 = World(Ai).Position.x
'      x4 = World(Ai).Position.Z
'      If Distance(x1, x2, x3, x4) > 60 Then
'        x = Sin(-World(Ai).Angle.y)
'        y = Cos(-World(Ai).Angle.y)
'        x = x * 25: y = y * 25
'        World(Ai).Position.x = World(Ai).Position.x - x
'        World(Ai).Position.Z = World(Ai).Position.Z - y
'      End If
'
'      If Angle < 8 And Angle > -8 And Rnd > 0.9 Then
'        For n = 1 To totalunits
'          If World(n).Type = 0 Then
'            World(n).Position.x = World(Ai).Position.x
'            World(n).Position.Z = World(Ai).Position.Z
'            World(n).Position.y = -70
'            World(n).Angle.y = World(Ai).Angle.y
'            World(n).Direction = World(Ai).Angle.y
'            World(n).Type = 3
'            World(n).Speed = 1
'            World(n).Delay = 40
'            World(n).Owner = Ai
'            Exit For
'          End If
'        Next n
'      End If
'    End If
'
    '=====================     Planes   ==========================
    
'    Case 11
'      For n = 1 To totalunits
'        If World(n).Type = 0 Then
'          World(n).Type = 5
'          World(n).Position.x = World(Ai).Position.x
'          World(n).Position.y = World(Ai).Position.y
'          World(n).Position.Z = World(Ai).Position.Z
'          World(n).Delay = 7
'          Exit For
'        End If
'      Next n
'
'      TXX = World(World(Ai).Target).Position.x - World(Ai).Position.x
'      Tyy = World(World(Ai).Target).Position.Z - World(Ai).Position.Z
'
'      PAngle = (-World(Ai).Angle.y)
'      Angle = FindAngle(TXX, Tyy, PAngle)
'
'      If Angle > 0 Then World(Ai).Angle.y = World(Ai).Angle.y - 10 / pi
'      If Angle < 0 Then World(Ai).Angle.y = World(Ai).Angle.y + 10 / pi
'
'      x = Sin(-World(Ai).Angle.y)
'      y = Cos(-World(Ai).Angle.y)
'      x = x * 125: y = y * 125
'
'      x1 = World(World(Ai).Target).Position.x
'      x2 = World(World(Ai).Target).Position.Z
'      x3 = World(Ai).Position.x
'      x4 = World(Ai).Position.Z
'      If Distance(x1, x2, x3, x4) < 40 Then
'        For n = 1 To totalunits
'          If World(n).Type = 0 Then
'            World(n).Type = 12
'            World(n).Position.x = World(Ai).Position.x
'            World(n).Position.y = World(Ai).Position.y
'            World(n).Position.Z = World(Ai).Position.Z
'            World(n).Delay = 0
'            World(n).Owner = Ai
'            Exit For
'          End If
'        Next n
'        For n = 1 To totalunits
'          If World(n).Type = 0 Then
'            World(n).Type = 12
'            World(n).Position.x = World(Ai).Position.x - x
'            World(n).Position.y = World(Ai).Position.y
'            World(n).Position.Z = World(Ai).Position.Z - y
'            World(n).Delay = 0
'            World(n).Owner = Ai
'            Exit For
'          End If
'        Next n
'      End If
'
'      For n = 1 To totalunits
'        If World(n).Type = 0 Then
'          World(n).Type = 5
'          World(n).Position.x = World(Ai).Position.x - x
'          World(n).Position.y = World(Ai).Position.y
'          World(n).Position.Z = World(Ai).Position.Z - y
'          World(n).Delay = 10
'          Exit For
'        End If
'      Next n
'
'      World(Ai).Position.x = World(Ai).Position.x - x - x
'      World(Ai).Position.Z = World(Ai).Position.Z - y - y
'
      '=====================     Missiles   ==========================
'
'    Case 3, 12
'      If World(Ai).Type = 12 Then
'        World(Ai).Position.y = World(Ai).Position.y + 25
'        If World(Ai).Position.y > 0 Then
'          World(Ai).Type = 10
'          World(Ai).Delay = 5
'        End If
'      Else
'        x = Sin(-World(Ai).Angle.y)
'        y = Cos(-World(Ai).Angle.y)
'        x = x * 300: y = y * 300
'        World(Ai).Position.x = World(Ai).Position.x - x
'        World(Ai).Position.Z = World(Ai).Position.Z - y
'      End If
'
'
'      For ft = 1 To totalunits
'        If World(ft).Type = 1 Or World(ft).Type = 2 Or World(ft).Type = 4 Or World(ft).Type = 8 Then
'          If World(Ai).Position.x > World(ft).Position.x - (100 * World(ft).Scale.x) Then
'            If World(Ai).Position.x < World(ft).Position.x + (100 * World(ft).Scale.x) Then
'              If World(Ai).Position.Z > World(ft).Position.Z - (100 * World(ft).Scale.Z) Then
'                If World(Ai).Position.Z < World(ft).Position.Z + (100 * World(ft).Scale.Z) Then
'                  If World(Ai).Position.y > World(ft).Position.y - (100 * World(ft).Scale.Z) Then
'                    If World(Ai).Position.y < World(ft).Position.y + (100 * World(ft).Scale.Z) Then
'                      If World(Ai).Owner = ft Then Exit For
'                      World(Ai).Type = 0
'                      If World(Ai).Type = 12 Then World(ft).Health = World(ft).Health - 1
'                      World(ft).Health = World(ft).Health - 1
'                      If ft <> 1 Then World(ft).Alert = True
'                      Dim ScrapCount, Counta As Byte
'                      ScrapCount = 1
'
'                      'World(n).Type = 0
'
'                      'If World(ft).Health = 0 Then
'                      ' If ft = 1 Then
'                      '   World(1).Dead = True
'                      '   Timer2.Interval = 0
'                      '   Messages.AddItem "You just got blown to bits!!!"
'                      '   Messages.AddItem "You just got blown to bits!!!"
'                      '   Messages.AddItem "You just got blown to bits!!!"
'                      '   Messages.AddItem "You just got blown to bits!!!"
'                      '   Messages.AddItem "You just got blown to bits!!!"
'                      ' End If
'
'                      '-----------------------------
'
'                      'World(ft).Type = 0
'
'                      '-----------------------------
'
'                      'ScrapCount = 15
'                      'mes = World(World(Ai).Owner).PlayerName
'                      'mes = mes + " just smashed "
'                      'mes = mes + World(ft).PlayerName
'                      'Mid(mes, 1, 1) = UCase(Mid(mes, 1, 1))
'                      'Messages.AddItem "    " + mes
'                      'End If
'
'                      Counta = 0
'
'                      For nxn = 1 To totalunits
'                        If World(nxn).Type = 0 Then
'                          Counta = Counta + 1
'                          World(nxn).Type = 6 + Int(Rnd * 2)
'                          World(nxn).Position.x = World(ft).Position.x + Rnd * 300 - 150
'                          World(nxn).Position.y = World(ft).Position.y
'                          World(nxn).Position.Z = World(ft).Position.Z + Rnd * 300 - 150
'                          World(nxn).GibSpeed.x = (Rnd * 40) - 20
'                          World(nxn).GibSpeed.y = (-Rnd * 10) - 15
'                          World(nxn).GibSpeed.Z = (Rnd * 40) - 20
'                          World(nxn).Scale.x = 0.2 + Rnd * 0.2
'                          World(nxn).Scale.y = 0.2 + Rnd * 0.2
'                          World(nxn).Scale.Z = 0.2 + Rnd * 0.2
'                          World(nxn).Delay = 40 + Int(Rnd * 10)
'                          If World(ft).Health <> 0 Then
'                            World(nxn).Scale.x = 0.05 + Rnd * 0.05
'                            World(nxn).Scale.y = 0.05 + Rnd * 0.05
'                            World(nxn).Scale.Z = 0.05 + Rnd * 0.05
'                            World(nxn).Position.x = World(Ai).Position.x
'                            World(nxn).Position.y = World(Ai).Position.y
'                            World(nxn).Position.Z = World(Ai).Position.Z
'                          End If
'                          If Counta > ScrapCount Then Exit For
'                        End If
'                      Next nxn
'hag:
'                    End If
'                  End If
'                End If
'              End If
'            End If
'          End If
'        End If
'      Next ft
'
'
'    '=====================     Waste     ==========================
'    Case 6, 7
'      World(Ai).Position.x = World(Ai).Position.x + World(Ai).GibSpeed.x
'      World(Ai).Position.y = World(Ai).Position.y + World(Ai).GibSpeed.y
'      World(Ai).Position.Z = World(Ai).Position.Z + World(Ai).GibSpeed.Z
'      World(Ai).Angle.x = World(Ai).Angle.x + World(Ai).GibSpeed.x
'      World(Ai).Angle.y = World(Ai).Angle.y + World(Ai).GibSpeed.y
'      World(Ai).Angle.Z = World(Ai).Angle.Z + World(Ai).GibSpeed.Z
'      World(Ai).GibSpeed.y = World(Ai).GibSpeed.y + 2
'      If World(Ai).Position.y > 0 Then
'        World(Ai).GibSpeed.x = 0
'        World(Ai).GibSpeed.y = 0
'        World(Ai).GibSpeed.Z = 0
'      End If
  End Select
End Sub

