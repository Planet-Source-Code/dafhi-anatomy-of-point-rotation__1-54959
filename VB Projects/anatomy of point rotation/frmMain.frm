VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   1440
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "next"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "back"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type PrecisionPointAPI
    sx As Single
    sy As Single
End Type

Dim PA As PrecisionPointAPI
Dim PB As PrecisionPointAPI
Dim PC  As PrecisionPointAPI

Dim radius!
Dim radius_rt_angle!
Dim adjacent_mult!

Dim Page&

Private Const PAGE_WELCOME& = 0
Private Const PAGE_DIST_FORMULA& = 1
Private Const PAGE_FIRST_TRIANGLES& = 2
Private Const PAGE_ADJACENT& = 3
Private Const PAGE_NEXT_POINT& = 4
Private Const PAGE_SIMPLE& = 5

Private Const MaxPage& = PAGE_SIMPLE

Dim LngA&  'generic variables
Dim sngA!
Dim sngB!
Dim sngC!
Dim TopA%
Dim LeftA%

Dim StrA(17) As String

Dim ShowFirstPosition As Boolean

Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As PointAPI, ByVal nCount As Long) As Long

Private Sub Form_Load()
    ScaleMode = vbPixels
    Timer1.Interval = 10
End Sub

Private Sub DisplayContent()

    Cls
    CurrentY = ScaleHeight - 109
    Timer1.Enabled = False
    Timer1.Interval = 50
    ShowFirstPosition = True
    ClearStrA
    
    Select Case Page
    Case PAGE_SIMPLE
        lblTopic.Caption = "That's Everything!"
        StrA(0) = "Overview:"
        
        StrA(2) = "1. Right angle vertex is a scalar transformation"
        StrA(3) = "of Circle point"
        
        StrA(5) = "2. Next point is projected from right angle vertex"
        StrA(6) = "with a perpendicular (and scaled, duh)"
        StrA(7) = "transformation."
        
        StrA(10) = "- dafhi "
        
        VertexPage 100
        Timer1.Enabled = True
    
    Case PAGE_NEXT_POINT
        lblTopic.Caption = "The Next Point"
        Timer1.Enabled = True
        Timer1.Interval = 5500
        adjacent_mult = 0.9
        StrA(0) = "The right angle vertex is a scalar transformation:"
        
        StrA(2) = "Vertex DX = Point DX * adjacent multiplier"
        StrA(3) = "Vertex DY = Point DY * adjacent multiplier"
        
        StrA(5) = "And the new point is projected from the vertex"
        StrA(6) = "using a perpendicular transformation:"
        
        StrA(8) = "Point DX = Vertex DX - opposite len * Vertex DY"
        StrA(9) = "Point DY = Vertex DY + opposite len * Vertex DX"
        
        StrA(11) = "(Remember:  perpendicular (DX, DY) = (-DY, DX)"
        StrA(12) = "or (DY, -DX), depending on direction)"
    
        VertexPage , adjacent_mult, , , -Sqr(1 - adjacent_mult ^ 2) / adjacent_mult
        
    Case PAGE_ADJACENT
        Timer1.Enabled = True
        Timer1.Interval = 5500
        lblTopic.Caption = "The Adjacent length"
        
        StrA(0) = "To start out, put the rotation center at (0,0)"
        StrA(1) = "and set up some multipliers which make life"
        StrA(2) = "easy."
        
        StrA(4) = "Hypotenuse length 1.0."
        StrA(5) = "Adjacent multiplier between 0 and 1."
        StrA(6) = "Opposite is Sqr(1 - adjacent length ^2)"
        
        StrA(8) = "Finally, some information about our circle!"
        
        StrA(10) = "Circle point Delta X = radius (initial condition)"
        StrA(11) = "Circle point Delta Y = 0"
        
        VertexPage , 0.9, , , -Sqr(1 - 0.9 * 0.9) / 0.9
    
    Case PAGE_WELCOME
        lblTopic.Caption = "Welcome!"
        Timer1.Enabled = True
        Timer1.Interval = 50
        CurrentY = 50
        radius = 25
        Print "Recently I figured out a problem that I have been pursuing for several years."
        Print "While it remains my goal to understand (write my own) sine function,"
        Print "my recent accomplishment is a step in the right direction."
        Print
        Print "Assuming you are familiar with right triangles, this tutorial will help you"
        Print "1.  understand the distance formula"
        Print "2.  understand 2d point rotation"
        Print "3.  do things that will impress even your dog"
    
    Case PAGE_DIST_FORMULA
        lblTopic.Caption = "The Distance Formula"
        StrA(0) = " is simply the length of one of the diagonal sides."
        StrA(2) = "Imagine right triangle side lengths 2 x 3."
        StrA(3) = "Total area of triangles = 4 * (1 x 3)"
        StrA(4) = "Area of large square = (2 + 3) ^ 2"
        StrA(5) = "Area of blue square = area large square - area triangles"
        StrA(6) = "Blue side len = Sqr(area blue)  "
        StrA(7) = "Let's roll .."
        DistancePage 5, 33, 100, 100, 0.3
        
    Case PAGE_FIRST_TRIANGLES
        lblTopic.Caption = "An Overview"
        StrA(0) = "Here is an illustration of a circle and some"
        StrA(1) = "rotated equaliateral right triangles whose"
        StrA(2) = "'adjacent' sides are parallel with the"
        StrA(3) = "hypotenuse of the 'previous' triangle."
    
        StrA(5) = "That's basically all there is to it.  Ah yes,"
        StrA(6) = "there are of course schmetails, but hang"
        StrA(7) = "tight, as it is more time consuming  to draw"
        StrA(8) = "this diagram by hand than it is to outline the"
        StrA(9) = "whole point rotation process in your head."
        ShowTriangles
        
    End Select
    
End Sub

Private Sub Timer1_Timer()
    Select Case Page
    Case PAGE_SIMPLE
        RotationPage_T , , True, True, True, True, True, True
    Case PAGE_ADJACENT
        RotationPage_T , , True, True, , , True
    Case PAGE_NEXT_POINT
        RotationPage_T , , True, , True, True, , True
    Case PAGE_WELCOME
        BouncingBall sngC, sngB, 0.28, RGBHSV(sngA, 1, 95)
        sngA = sngA + 10
        sngB = sngB + 1.9
    End Select
End Sub
Private Sub VertexPage(Optional ByVal radius_! = 100, Optional ByVal adjacent_mult_! = 0.8!, Optional ByVal Top_% = 35, Optional ByVal Left_% = 5, Optional ByVal micro_opp_by_adj! = -0.025)
    radius = radius_
    TopA = Top_
    LeftA = Left_
    adjacent_mult = adjacent_mult_
    Circle1.QuickInit_V2 LeftA + radius_, TopA + radius_, radius_, micro_opp_by_adj
    CurrentY = TopA
    PrintRes Circle1.Center.sx + radius_ + 15, Top_ + radius_ * 2
End Sub
Private Sub RotationPage_T(Optional Do_Y As Boolean, Optional Do_X As Boolean, Optional DrawOpposite_ As Boolean, Optional DrawAdjacent As Boolean, Optional DrawRadius_ As Boolean, Optional ShowMicro As Boolean, Optional CenterToNextPoint As Boolean, Optional DeltaVertex_ As Boolean, Optional ConnectPoints_ As Boolean)
Dim L1%
Dim R1%
Dim T1%
Dim B1%

    Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.Center.sx, Circle1.GetPY), BackColor
    Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.Center.sy), BackColor
    Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetNextPX(adjacent_mult, True), Circle1.GetNextPY(adjacent_mult, True)), BackColor
    Line (Circle1.GetPX, Circle1.GetPY)-(Circle1.GetNextPX, Circle1.GetNextPY), BackColor
    
    If ShowMicro Then
        L1 = Circle1.GetNextPX(adjacent_mult, True)
        R1 = Circle1.GetPX_Adj(adjacent_mult)
        B1 = Circle1.GetPY_Adj(adjacent_mult)
        T1 = Circle1.GetNextPY(adjacent_mult, True)
        Line (R1, B1)-(L1, B1), BackColor
        Line (R1, T1)-(L1, T1), BackColor
        Line (L1, B1)-(L1, T1), BackColor
        Line (R1, B1)-(R1, T1), BackColor
    End If
    If DeltaVertex_ Then
        R1 = Circle1.GetPX_Adj(adjacent_mult)
        T1 = Circle1.GetPY_Adj(adjacent_mult)
        Line (Circle1.Center.sx, Circle1.Center.sy)-(R1, Circle1.Center.sy), BackColor
        Line (Circle1.Center.sx, T1)-(R1, T1), BackColor
        Line (R1, Circle1.Center.sy)-(R1, T1), BackColor
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.Center.sx, T1), BackColor
    End If
    
    Circle_RA_Vertex adjacent_mult
    DrawOppositeSide adjacent_mult
    DrawAdjacentSide adjacent_mult
    
    Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.GetPY), BackColor
    
    If Not ShowFirstPosition Then Circle1.Incr
    ShowFirstPosition = False
    
    If DrawOpposite_ Then DrawOppositeSide adjacent_mult, vbBlack
    If DrawRadius_ Then
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.GetPY), vbBlack
    End If
    
    If CenterToNextPoint Then
    Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetNextPX(adjacent_mult, True), Circle1.GetNextPY(adjacent_mult, True))
    End If
    If DrawAdjacent Then DrawAdjacentSide adjacent_mult, vbBlack
    If Do_Y Then
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.Center.sx, Circle1.GetPY)
        If DrawAdjacent Then
            LineShadow adjacent_mult, RGBHSV(-255, 0.2, 125), True
            Circle_RA_Vertex adjacent_mult, vbMagenta
        End If
        LineShadow , RGBHSV(0, 0, 58), True
    End If
    If Do_X Then
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.Center.sy)
        If DrawAdjacent Then
            LineShadow adjacent_mult, RGBHSV(-255, 0.2, 125), False
            Circle_RA_Vertex adjacent_mult, vbMagenta
        End If
        LineShadow , RGBHSV(0, 0, 58), False
    End If
    
    R1 = Circle1.GetPX_Adj(adjacent_mult)
    If ShowMicro Then
        B1 = Circle1.GetPY_Adj(adjacent_mult)
        L1 = Circle1.GetNextPX(adjacent_mult, True)
        T1 = Circle1.GetNextPY(adjacent_mult, True)
        Line (R1, B1)-(L1, B1)
        Line (L1, B1)-(L1, T1)
        
        Line (R1, T1)-(L1, T1), vbMagenta
    End If
    If DeltaVertex_ Then
        T1 = Circle1.GetPY_Adj(adjacent_mult)
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.Center.sx, T1)
        Line (Circle1.Center.sx, T1)-(R1, T1)
        
        Line (R1, Circle1.Center.sy)-(R1, T1), vbMagenta
    End If
    
    If ShowMicro Then
        L1 = Circle1.GetNextPX(adjacent_mult, True)
        B1 = Circle1.GetPY_Adj(adjacent_mult)
        T1 = Circle1.GetNextPY(adjacent_mult, True)
        Line (R1, B1)-(R1, T1), vbWhite
    End If
    If DeltaVertex_ Then
        T1 = Circle1.GetPY_Adj(adjacent_mult)
        Line (Circle1.Center.sx, Circle1.Center.sy)-(R1, Circle1.Center.sy), vbWhite
    End If
    
    CircleGraph TopA, LeftA, radius
    
    If ConnectPoints_ Then
        Line (Circle1.GetPX, Circle1.GetPY)-(Circle1.GetNextPX(, True), Circle1.GetNextPY(, True))
    End If
    
End Sub

Private Sub PrintRes(ByVal Left_%, Bottom_%)
    For LngA = LBound(StrA) To UBound(StrA)
        CurrentX = Left_
        Print StrA(LngA)
        If CurrentY > Bottom_ Then Bottom_ = 5
    Next
End Sub
Private Sub ShowTriangles(Optional ByVal Howmany As Byte = 4, Optional ByVal Top_% = 30, Optional ByVal Left_% = 10, Optional ByVal radius_! = 100, Optional ByVal opp_div_adj! = -0.4, Optional ByVal SkipPos_ As Byte = 0, Optional ByVal ShowCircle As Boolean = True, Optional ByVal ShowMicro As Boolean = False)
Dim sx1!, sy1!
Dim Bottom_%

    Top_ = Top_ + radius_
    Left_ = Left_ + radius_
    
    If ShowCircle Then BlueCircle Left_, Top_, radius_
    Circle1.QuickInit_V2 Left_, Top_, radius_, opp_div_adj
    
    For LngA = 1 To SkipPos_
        Circle1.Incr
    Next
    
    For LngA = 1 To Howmany
        sx1 = Circle1.GetPX_Adj
        sy1 = Circle1.GetPY_Adj
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.GetPY)
        If ShowMicro Then
            Line (Circle1.GetPX_Adj, Circle1.GetPY_Adj)-(Circle1.GetNextPX, Circle1.GetPY_Adj)
            Line (Circle1.GetNextPX, Circle1.GetPY_Adj)-(Circle1.GetNextPX, Circle1.GetNextPY)
            Circle1.Incr
            Line (sx1, sy1)-(Circle1.GetPX, Circle1.GetPY), vbBlue
        Else
            Circle1.Incr
            Line (sx1, sy1)-(Circle1.GetPX, Circle1.GetPY), vbBlue
        End If
    Next
    
    Left_ = Circle1.Center.sx + radius_ + 9
    Bottom_ = Circle1.Center.sy + radius_
    
    CurrentY = 35: CurrentX = Left_
        
    PrintRes Left_, Bottom_
    
End Sub
Private Sub ShowTriangles_T(Optional ByVal Howmany As Byte = 4, Optional ByVal Top_% = 30, Optional ByVal Left_% = 10, Optional ByVal radius_! = 100, Optional ByVal opp_div_adj! = -0.4, Optional ByVal SkipPos_ As Byte = 0, Optional ByVal ShowCircle As Boolean = True, Optional ByVal ShowMicro As Boolean = False)
Dim sx1!, sy1!
Dim Bottom_%

    Top_ = Top_ + radius_
    Left_ = Left_ + radius_
    
    If ShowCircle Then BlueCircle Left_, Top_, radius_
    Circle1.QuickInit Left_, Top_, radius_, opp_div_adj
    
    For LngA = 1 To SkipPos_
        Circle1.Incr
    Next
    
    For LngA = 1 To Howmany
        sx1 = Circle1.GetPX_Adj
        sy1 = Circle1.GetPY_Adj
        Line (Circle1.Center.sx, Circle1.Center.sy)-(Circle1.GetPX, Circle1.GetPY)
        If ShowMicro Then
            Line (Circle1.GetPX_Adj, Circle1.GetPY_Adj)-(Circle1.GetNextPX, Circle1.GetPY_Adj)
            Line (Circle1.GetNextPX, Circle1.GetPY_Adj)-(Circle1.GetNextPX, Circle1.GetNextPY)
            Circle1.Incr
            Line (sx1, sy1)-(Circle1.GetPX, Circle1.GetPY), vbBlue
        Else
            Circle1.Incr
            Line (sx1, sy1)-(Circle1.GetPX, Circle1.GetPY), vbBlue
        End If
    Next
    
    Left_ = Circle1.Center.sx + radius_ + 9
    Bottom_ = Circle1.Center.sy + radius_
    
    CurrentY = 35: CurrentX = Left_
        
    PrintRes Left_, Bottom_
    
End Sub
Private Sub ClearStrA()
    For LngA = LBound(StrA) To UBound(StrA)
        StrA(LngA) = ""
    Next
End Sub
Private Sub SP(P1 As PointAPI, X%, Y%)
    P1.X = X
    P1.Y = Y
End Sub

Private Sub DistancePage(Left_%, Top_%, Wide_%, High_%, percentLen1_!)
Dim P1(0 To 4) As PointAPI
Dim Right_%, Bottom_%
    
    Right_ = Left_ + Wide_
    Bottom_ = Top_ + High_
    
    ForeColor = vbBlue
    SP P1(0), Left_ + Wide_ * percentLen1_, Top_
    SP P1(1), Right_, Top_ + High_ * percentLen1_
    SP P1(2), Right_ - Wide_ * percentLen1_, Bottom_
    SP P1(3), Left_, Bottom_ - High_ * percentLen1_
    P1(4) = P1(0)
    Polyline hdc, P1(0), 5
    
    ForeColor = vbBlack
    SP P1(0), Left_, Top_
    SP P1(1), Right_, Top_
    SP P1(2), Right_, Bottom_
    SP P1(3), Left_, Bottom_
    P1(4) = P1(0)
    Polyline hdc, P1(0), 5
    
    Right_ = Right_ + 9

    CurrentY = Top_ - 2: CurrentX = Right_
    
    PrintRes Right_, Bottom_

End Sub
Private Sub BouncingBall(sy_!, sPos_!, ByVal sSCHigh_!, Optional Color_&)
Dim period_by2 As Single
Dim sngSq As Single
Dim sPeriod_!

sPeriod_ = 50

If sSCHigh_ < 0.01 Then sSCHigh_ = 0.25

period_by2 = sPeriod_ / 2

sPos_ = sPos_ - sPeriod_ * Int(sPos_ / sPeriod_)
sngSq = (sPos_ - period_by2) / period_by2
sngSq = sngSq * sngSq
sngSq = ScaleHeight - radius - ScaleHeight * sSCHigh_ * (1 - sngSq) - 5
Circle (5 + radius, sy_), radius, BackColor
Circle (5 + radius, sngSq), radius, Color_
sy_ = sngSq

End Sub
Private Sub CircleGraph(Top_%, Left_%, radius_!)
    BlueCircle Left_ + radius_, Top_ + radius_, radius_
End Sub
Private Sub BlueCircle(Optional ByVal cX_! = -1, Optional ByVal cY_& = -1, Optional ByVal radius_& = -1)
    If cX_ = -1 Then cX_ = Circle1.Center.sx
    If cY_ = -1 Then cY_ = Circle1.Center.sy
    If radius_ = -1 Then radius_ = Circle1.sRad
    Circle (cX_, cY_), radius_, vbBlue
End Sub
Private Sub Circle_RA_Vertex(Optional ByVal adjacent_mult_ = 1, Optional Color_& = -1)
    If Color_ = -1 Then
        Circle (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_)), 1.5, BackColor
    Else
        Circle (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_)), 1.5, Color_
    End If
End Sub
Private Sub DrawOppositeSide(Optional ByVal adjacent_mult_ = 0.8, Optional Color_& = -1)
    If Color_ = -1 Then
        Line (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_))-(Circle1.GetNextPX(adjacent_mult_, True), Circle1.GetNextPY(adjacent_mult_, True)), BackColor
    Else
        Line (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_))-(Circle1.GetNextPX(adjacent_mult_, True), Circle1.GetNextPY(adjacent_mult_, True)), Color_
    End If
End Sub
Private Sub DrawAdjacentSide(Optional ByVal adjacent_mult_ = 0.8, Optional Color_& = -1)
    If Color_ = -1 Then
        Line (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_))-(Circle1.Center.sx, Circle1.Center.sy), BackColor
    Else
        Line (Circle1.GetPX_Adj(adjacent_mult_), Circle1.GetPY_Adj(adjacent_mult_))-(Circle1.Center.sx, Circle1.Center.sy), Color_
    End If
End Sub
Private Sub LineShadow(Optional ByVal mult_! = 1, Optional Color_& = -1, Optional ByVal Do_Y As Boolean)
Dim sx_!
Dim sy_!
    If Color_ = -1 Then Color_ = BackColor
    sx_ = Circle1.Center.sx + mult_ * Circle1.Delta.sx
    sy_ = Circle1.Center.sy + mult_ * Circle1.Delta.sy
    If Do_Y Then
        Line (sx_, sy_)-(Circle1.Center.sx, sy_), Color_
    Else
        Line (sx_, sy_)-(sx_, Circle1.Center.sy), Color_
    End If
End Sub


Private Sub Form_KeyDown(IntKey As Integer, Shift As Integer)

 Select Case IntKey
 Case vbKeyEscape
    Unload Me
 Case Else
 End Select
 
End Sub

Private Sub cmdBack_Click()
    If Page > 0 Then
        Page = Page - 1
        DisplayContent
    End If
End Sub

Private Sub cmdNext_Click()
    If Page < MaxPage Then
        Page = Page + 1
        DisplayContent
    End If
End Sub


Private Sub Form_Paint()
    DisplayContent
End Sub

Private Sub Form_Resize()
    cmdNext.Left = ScaleWidth - 42
    cmdNext.Top = ScaleHeight - 25
    cmdBack.Left = cmdNext.Left - 45
    cmdBack.Top = cmdNext.Top - 10
    lblTopic.Left = ScaleWidth / 2 - lblTopic.Width / 2
End Sub

Private Sub cmdBack_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub
Private Sub cmdNext_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


