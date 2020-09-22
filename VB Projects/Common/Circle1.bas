Attribute VB_Name = "Circle1"
Option Explicit

'Circle1.bas 1.0  by dafhi
 
' = Sample Code Section =
 
'Dim I As Long
'Dim px As Single
'Dim py As Single
'Dim hue As Single
 
'Private Sub Form_Paint()
' BackColor = RGBHSV(0, 0.1, 175)
' ScaleMode = vbPixels
' Circle1.QuickInit 100, 100, , 5
' For I = 1 To 100
'    px = Center.sx + Delta.sx
'    py = Center.sy + Delta.sy
'    PSet (px, py), RGBHSV(hue, 0.3, 110)
'    Circle1.Incr
'    hue = hue + 1.2
' Next
'End Sub

' ==================================

Private Type PrecisionPointAPI
    sx As Double
    sy As Double
End Type

Public Center As PrecisionPointAPI
Public Delta As PrecisionPointAPI
Public sRad As Single

Dim dTmp#

Private adjacent_mult As Double
Private opposite_mult As Double
Private adjacent_by_radius As Double

Public Function OppositeMult() As Double
    OppositeMult = opposite_mult
End Function
Public Function AdjacentMult() As Double
    AdjacentMult = adjacent_mult
End Function
Public Function AdjacentByRadius() As Double
    AdjacentByRadius = adjacent_by_radius
End Function

Public Sub QuickInit_V2(Optional ByVal center_x As Single = 100, Optional ByVal center_y As Single = 100, Optional ByVal radius_! = 100, Optional ByVal opp_div_adjacent# = 0.01)
    adjacent_mult = Sqr(1 / (1 + opp_div_adjacent ^ 2))
    opposite_mult = opp_div_adjacent * adjacent_mult
    InitCommon center_x, center_y, radius_
End Sub
Public Sub QuickInit(Optional ByVal center_x As Single = 100, Optional ByVal center_y As Single = 100, Optional ByVal radius_! = 100, Optional ByVal point_spacing! = 1)
    adjacent_mult = (2 * radius_ * radius_ - point_spacing * point_spacing) / (2 * radius_ * radius_)
    opposite_mult = Sqr(1 - adjacent_mult ^ 2)
    InitCommon center_x, center_y, radius_
End Sub
Private Sub InitCommon(ByVal center_x As Single, ByVal center_y As Single, ByVal radius_!)
    Center.sx = center_x
    Center.sy = center_y
    SetDelta radius_, 0
    adjacent_by_radius = adjacent_mult / radius_
End Sub

Public Sub Incr() 'This is where the rotation happens
    dTmp = Delta.sx
    Delta.sx = Delta.sx * adjacent_mult - Delta.sy * opposite_mult
    Delta.sy = Delta.sy * adjacent_mult + dTmp * opposite_mult
End Sub
Public Sub IncrN(ByVal HowManyTimes_&, Optional OppositeDirection_ As Boolean)
    If OppositeDirection_ Then
        For HowManyTimes_ = 1 To HowManyTimes_
            dTmp = Delta.sx
            Delta.sx = Delta.sx * adjacent_mult + Delta.sy * opposite_mult
            Delta.sy = Delta.sy * adjacent_mult - dTmp * opposite_mult
        Next
    Else
        For HowManyTimes_ = 1 To HowManyTimes_
            dTmp = Delta.sx
            Delta.sx = Delta.sx * adjacent_mult - Delta.sy * opposite_mult
            Delta.sy = Delta.sy * adjacent_mult + dTmp * opposite_mult
        Next
    End If
End Sub

Public Sub SetRadius(radius_!)
    sRad = radius_
    SetDelta radius_, 0
End Sub
Public Function GetAdjacentMult(ByVal opp_div_adjacent As Double) As Double
    GetAdjacentMult = Sqr(1 / (1 + opp_div_adjacent ^ 2))
End Function
Public Function GetAdjacentLen(ByVal lenMultRadius#, ByVal lenmultA_#, ByVal lenmultB_#) As Double
    GetAdjacentLen = sRad * (lenMultRadius * lenMultRadius - lenmultA_ * lenmultA_ + lenmultB_ * lenmultB_) / (2 * lenMultRadius)
End Function
Public Sub SetDelta(ByVal px_!, ByVal py_!)
    Delta.sx = px_
    Delta.sy = py_
    sRad = Sqr(Delta.sx ^ 2 + Delta.sy ^ 2)
End Sub


Public Function GetPX() As Double
    GetPX = Center.sx + Delta.sx
End Function
Public Function GetPY() As Double
    GetPY = Center.sy + Delta.sy
End Function
Public Function GetPX_Adj(Optional ByVal adjacent_#) As Double
    If adjacent_ = 0 Then
        GetPX_Adj = Center.sx + Delta.sx * adjacent_mult
    Else
        GetPX_Adj = Center.sx + Delta.sx * adjacent_
    End If
End Function
Public Function GetPY_Adj(Optional ByVal adjacent_#) As Double
    If adjacent_ = 0 Then
        GetPY_Adj = Center.sy + Delta.sy * adjacent_mult
    Else
        GetPY_Adj = Center.sy + Delta.sy * adjacent_
    End If
End Function
Public Function GetNextPX(Optional ByVal adjacent_#, Optional OppositeDir As Boolean) As Double
    If adjacent_ = 0 Then
        If OppositeDir Then
            GetNextPX = Circle1.Center.sx + Circle1.Delta.sx * adjacent_mult + Circle1.Delta.sy * opposite_mult
        Else
            GetNextPX = Circle1.Center.sx + Circle1.Delta.sx * adjacent_mult - Circle1.Delta.sy * opposite_mult
        End If
    ElseIf OppositeDir Then
        GetNextPX = Circle1.Center.sx + Circle1.Delta.sx * adjacent_ + Circle1.Delta.sy * Sqr(1 - adjacent_ ^ 2)
    Else
        GetNextPX = Circle1.Center.sx + Circle1.Delta.sx * adjacent_ - Circle1.Delta.sy * Sqr(1 - adjacent_ ^ 2)
    End If
End Function
Public Function GetNextPY(Optional ByVal adjacent_#, Optional OppositeDir As Boolean) As Double
    If adjacent_ = 0 Then
        If OppositeDir Then
            GetNextPY = Circle1.Center.sy + Circle1.Delta.sy * adjacent_mult - Circle1.Delta.sx * opposite_mult
        Else
            GetNextPY = Circle1.Center.sy + Circle1.Delta.sy * adjacent_mult + Circle1.Delta.sx * opposite_mult
        End If
    ElseIf OppositeDir Then
        GetNextPY = Circle1.Center.sy + Circle1.Delta.sy * adjacent_ - Circle1.Delta.sx * Sqr(1 - adjacent_ ^ 2)
    Else
        GetNextPY = Circle1.Center.sy + Circle1.Delta.sy * adjacent_ + Circle1.Delta.sx * Sqr(1 - adjacent_ ^ 2)
    End If
End Function

Public Function RGBHSV(ByVal hue_0_To_1530!, ByVal saturation_0_To_1!, ByVal value_0_To_255!) As Long
Dim hue_and_sat!
Dim value1!
Dim diff1!
Dim subt!
Dim minim!
Dim maxim!
Dim BGRed&
Dim BGGrn&
Dim BGBlu&

 'This function doesn't have error checking, so keep
 'value_0_To_255 between 0 and 255, and keep
 'saturation_0_To_1 between 0 and 1.
 
 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    BGBlu = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     BGRed = Int(value1)
     BGGrn = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     BGGrn = Int(value1)
     BGRed = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    BGRed = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     BGGrn = Int(value1)
     BGBlu = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     BGBlu = Int(value1)
     BGGrn = Int(value1 - hue_and_sat)
    End If
   Else
    BGGrn = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     BGBlu = Int(value1)
     BGRed = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     BGRed = Int(value1)
     BGBlu = Int(value1 - hue_and_sat)
    End If
   End If
   RGBHSV = BGRed Or BGGrn * 256& Or BGBlu * 65536
  Else 'saturation_0_To_1 <= 0
   RGBHSV = Int(value1) * 65793 '1 + 256 + 65536
  End If
 Else 'value_0_To_255 <= 0
  RGBHSV = 0&
 End If
End Function


