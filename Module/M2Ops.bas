Attribute VB_Name = "M2Ops"
' Routines for manipulating 2-dimensional
' vectors and matrices.
Type Coord
     x As Single
     y As Single
End Type

Option Explicit

' Create a 2-dimensional identity matrix.
Public Sub m2Identity(M() As Single)
Dim I As Integer
Dim J As Integer

    For I = 1 To 3
        For J = 1 To 3
            If I = J Then
                M(I, J) = 1
            Else
                M(I, J) = 0
            End If
        Next J
    Next I
End Sub

' Create a translation matrix for translation by
' distances tx and ty.
Public Sub m2Translate(Result() As Single, ByVal tx As Single, ByVal ty As Single)
    m2Identity Result
    Result(3, 1) = tx
    Result(3, 2) = ty
End Sub

' Create a scaling matrix for scaling by factors
' of sx and sy.
Public Sub m2Scale(Result() As Single, ByVal sx As Single, ByVal sy As Single)
    m2Identity Result
    Result(1, 1) = sx
    Result(2, 2) = sy
End Sub

' Create a Skew matrix for Skew by factors
' of sx and sy.
Public Sub m2Skew(Result() As Single, ByVal sx As Single, ByVal sy As Single)
    m2Identity Result
    Result(1, 2) = 1 - sy
    Result(2, 1) = 1 - sx
End Sub

' Create a rotation matrix for rotating by the
' given angle (in radians).
Public Sub m2Rotate(Result() As Single, ByVal theta As Single)
    m2Identity Result
    Result(1, 1) = Cos(theta)
    Result(1, 2) = Sin(theta)
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub

' Create a rotation matrix that rotates the point
' (x, y) onto the X axis.
Public Sub m2RotateIntoX(Result() As Single, ByVal x As Single, ByVal y As Single)
Dim d As Single

    m2Identity Result
    d = Sqr(x * x + y * y)
    Result(1, 1) = x / d
    Result(1, 2) = -y / d
    Result(2, 1) = -Result(1, 2)
    Result(2, 2) = Result(1, 1)
End Sub

' Create a scaling matrix for scaling by factors
' of sx and sy at the point (x, y).
Public Sub m2ScaleAt(Result() As Single, _
                     ByVal sx As Single, ByVal sy As Single, _
                     ByVal x As Single, ByVal y As Single)
                    
Dim t(1 To 3, 1 To 3) As Single
Dim S(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -x, -y

    ' Compute the inverse translation.
    m2Translate T_Inv, x, y

    ' Scale.
    m2Scale S, sx, sy

    ' Combine the transformations.
    m2MatMultiply M, t, S           ' T * S
    m2MatMultiply Result, M, T_Inv  ' T * S * T_Inv
End Sub

' Create a matrix for reflecting across the line
' passing through (x, y) in direction <dx, dy>.
Public Sub m2ReflectAcross(Result() As Single, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single)
Dim t(1 To 3, 1 To 3) As Single
Dim r(1 To 3, 1 To 3) As Single
Dim S(1 To 3, 1 To 3) As Single
Dim R_Inv(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M1(1 To 3, 1 To 3) As Single
Dim M2(1 To 3, 1 To 3) As Single
Dim M3(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -x, -y

    ' Compute the inverse translation.
    m2Translate T_Inv, x, y

    ' Rotate so the direction vector lies in the Y axis.
    m2RotateIntoX r, dx, dy

    ' Compute the inverse translation.
    m2RotateIntoX R_Inv, dx, -dy

    ' Reflect across the X axis.
    m2Scale S, 1, -1

    ' Combine the transformations.
    m2MatMultiply M1, t, r     ' T * R
    m2MatMultiply M2, S, R_Inv ' S * R_Inv
    m2MatMultiply M3, M1, M2   ' T * R * S * R_Inv

    ' T * R * S * R_Inv * T_Inv
    m2MatMultiply Result, M3, T_Inv
End Sub

' Create a Skew matrix
' of sx and sy at the point (x, y).
Public Sub m2SkewAt(Result() As Single, ByVal sx As Single, ByVal sy As Single, ByVal x As Single, ByVal y As Single)
                    
Dim t(1 To 3, 1 To 3) As Single
Dim S(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -x, -y

    ' Compute the inverse translation.
    m2Translate T_Inv, x, y

    ' Skew.
    m2Skew S, sx, sy

    ' Combine the transformations.
    m2MatMultiply M, t, S           ' T * S
    m2MatMultiply Result, M, T_Inv  ' T * S * T_Inv
End Sub

' Create a rotation matrix for rotating through
' angle theta around the point (x, y).
Public Sub m2RotateAround(Result() As Single, ByVal theta As Single, ByVal x As Single, ByVal y As Single)
Dim t(1 To 3, 1 To 3) As Single
Dim r(1 To 3, 1 To 3) As Single
Dim T_Inv(1 To 3, 1 To 3) As Single
Dim M(1 To 3, 1 To 3) As Single

    ' Translate the point to the origin.
    m2Translate t, -x, -y

    ' Compute the inverse translation.
    m2Translate T_Inv, x, y

    ' Rotate.
    m2Rotate r, theta

    ' Combine the transformations.
    m2MatMultiply M, t, r
    m2MatMultiply Result, M, T_Inv
End Sub

' Multiply a point and a matrix.
Public Sub m2PointMultiply(ByRef x As Single, ByRef y As Single, a() As Single)
Dim newx As Single
Dim newy As Single

    newx = x * a(1, 1) + y * a(2, 1) + a(3, 1)
    newy = x * a(1, 2) + y * a(2, 2) + a(3, 2)
    x = newx
    y = newy
End Sub
' Set copy = orig.
Public Sub m2PointCopy(copy() As Single, orig() As Single)
Dim I As Integer

    For I = 1 To 3
        copy(I) = orig(I)
    Next I
End Sub

' Set copy = orig.
Public Sub m2MatCopy(copy() As Single, orig() As Single)
Dim I As Integer
Dim J As Integer

    For I = 1 To 3
        For J = 1 To 3
            copy(I, J) = orig(I, J)
        Next J
    Next I
End Sub

' Apply a transformation matrix to a point.
Public Sub m2Apply(Result() As Single, v() As Single, a() As Single)
    Result(1) = v(1) * a(1, 1) + v(2) * a(2, 1) + a(3, 1)
    Result(2) = v(1) * a(1, 2) + v(2) * a(2, 2) + a(3, 2)
    Result(3) = 1#
End Sub

' Multiply two transformation matrices.
Public Sub m2MatMultiply(Result() As Single, a() As Single, B() As Single)
    Result(1, 1) = a(1, 1) * B(1, 1) + a(1, 2) * B(2, 1)
    Result(1, 2) = a(1, 1) * B(1, 2) + a(1, 2) * B(2, 2)
    Result(1, 3) = 0#
    Result(2, 1) = a(2, 1) * B(1, 1) + a(2, 2) * B(2, 1)
    Result(2, 2) = a(2, 1) * B(1, 2) + a(2, 2) * B(2, 2)
    Result(2, 3) = 0#
    Result(3, 1) = a(3, 1) * B(1, 1) + a(3, 2) * B(2, 1) + B(3, 1)
    Result(3, 2) = a(3, 1) * B(1, 2) + a(3, 2) * B(2, 2) + B(3, 2)
    Result(3, 3) = 1#
End Sub

Public Function m2Atn2(ByVal y As Single, ByVal x As Single) As Single
   If x = 0 Then
      m2Atn2 = IIf(y = 0, PI / 4, Sgn(y) * PI / 2)
   Else
      m2Atn2 = Atn(y / x) + (1 - Sgn(x)) * PI / 2
   End If
End Function

'Public Function m2GetAngle3P(P1 As POINTAPI, P2 As POINTAPI, P3 As POINTAPI) As Single
Public Function m2GetAngle3P(X1 As Single, Y1 As Single, _
                             X2 As Single, Y2 As Single, _
                             X3 As Single, Y3 As Single) As Single

' Retreive the angle formed by this 3 points
' P1<---->P2 e P1<----->P3
' rather than GetAngle3P this function doesn't need of a
' parallel edge and it returns the real internal angle close to the P1 point

      ' / P2
     ' /
    ' /
   ' /
  ' /
'P1 \ <-a°
   ' \
    ' \
     ' \
      ' \ P3
'

Dim I As Integer, An As Single
Dim Alfa As Single
Dim a As Double, B As Double, C As Double, PS As Double
Dim Ds1 As Double, Ds2 As Double, Ds3 As Double
Dim Q1 As Coord, Q2 As Coord, Q3 As Coord
Dim P1 As Coord, P2 As Coord, P3 As Coord

   P1.x = X1
   P1.y = Y1
   P2.x = X2
   P2.y = Y2
   P3.x = X3
   P3.y = Y3
   
Const Rg# = 200 / PI
    Ds1 = Dist(P1.x, P1.y, P2.x, P2.y)
    Ds2 = Dist(P1.x, P1.y, P3.x, P3.y)
    Ds3 = Dist(P3.x, P3.y, P2.x, P2.y)

    a = Ds3
    B = Ds1
    C = Ds2

    If a = 0 Or B = 0 Or C = 0 Then Exit Function

    PS = (a + B + C) * 0.5
    If PS < C Then GoTo ErrorAngle
    If PS < a Or PS < B Then GoTo ErrorAngle

    On Error Resume Next
    Alfa = 2 * Atn(((PS - B) * (PS - C) / PS / (PS - a)) ^ 0.5) * Rg#

    Alfa = m2An360(Alfa)

    Q1 = P1
    Q2 = P2
    Q3.x = P2.x
    Q3.y = P1.y

    An = GetAngle3P(Q1, Q2, Q3)
    If An <> 0 Then
        Q3.x = P3.x - P1.x
        Q3.y = P3.y - P1.y
        Q3 = m2RotatePoint(Q3.x, Q3.y, -An)
        Q3.x = Q3.x + P1.x
        Q3.y = Q3.y + P1.y
    End If

    If Q3.y < Q1.y Then Alfa = 360 - Alfa
    m2GetAngle3P = Alfa

Exit Function

ErrorAngle:

    m2GetAngle3P = 0

End Function

Private Function GetAngle3P(P1 As Coord, P2 As Coord, P3 As Coord) As Single

' Calculate angle from edges
' P1<---->P2 e P1<----->P3

' Note:
' It returns the angle 0-360 referred by the edge P1-P3 always parallel to the X axe
' if that edge (P1-P3) is not parallel the function will wrong the result value
'
' Next checks in wich square P2 is contained
' to set the relative angle (0-90 , 91-180, 181-270 or 271,360)
'

Dim I As Integer, K As Integer, M As Integer
Dim X1 As Double, Y1 As Double
Dim X2 As Double, Y2 As Double
Dim Alfa As Single
Dim a As Double, B As Double, C As Double, PS As Double
Dim Fd As Boolean
Dim Q1 As Coord, Q2 As Coord
Dim Ds1 As Single, Ds2 As Single, Ds3 As Single

Const Rg# = 200 / PI

X1 = P1.x
Y1 = P1.y

X2 = P2.x
Y2 = P2.y

Ds1 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

X2 = P3.x
Y2 = P3.y

Ds2 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

X1 = P2.x
Y1 = P2.y

Ds3 = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)

a = Ds3
B = Ds1
C = Ds2

If a = 0 Or B = 0 Or C = 0 Then GoTo Parallel
PS = (a + B + C) * 0.5
If PS < C Then GoTo ErrorAngle
If PS < a Or PS < B Then GoTo ErrorAngle

On Error Resume Next
Alfa = 2 * Atn(((PS - B) * (PS - C) / PS / (PS - a)) ^ 0.5) * Rg#

' Alfa now is in centesimal units (0-400) need to convert it with An360
Alfa = m2An360(Alfa)

' Check Sqares

Parallel:

X1 = P1.x
Y1 = P1.y

X2 = P2.x
Y2 = P2.y


If X1 = X2 Then
    If Y1 > Y2 Then
        Alfa = 270
    ElseIf Y1 < Y2 Then
        Alfa = 90
    End If
ElseIf Y1 = Y2 Then
    If X1 > X2 Then
        Alfa = 180
    ElseIf X1 < X2 Then
        Alfa = 0
    End If
ElseIf X1 > X2 And Y1 < Y2 Then ' II°
    Alfa = 90 - Alfa + 90
ElseIf X1 > X2 And Y1 > Y2 Then ' III°
    Alfa = Alfa + 180
ElseIf X1 < X2 And Y1 > Y2 Then ' IV°
    Alfa = 90 - Alfa + 270
End If

GetAngle3P = Alfa

Exit Function

ErrorAngle:
    GetAngle3P = 0
End Function

Public Function m2An360(An As Single) As Single
' Transform an Angle from Centesimal 0,400 to
' 0, 360
If An <> 0 Then
   m2An360 = An / 1.11111111111111
Else
   m2An360 = 0
End If

End Function

Function Dist(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
        Dist = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

' Rotate a single Point using Rad function to converts Degree to Radians
Private Function m2RotatePoint(x As Single, y As Single, Angle As Single) As Coord
Dim xA As Single, yA As Single
Dim mSin As Single, mCos As Single
Dim P As Coord
   P.x = x
   P.y = y
If Angle <> 0 Then
   mSin = Sin(Angle * PI / 180): mCos = Cos(Angle * PI / 180)
   xA = mCos * P.x - mSin * P.y
   yA = mSin * P.x + mCos * P.y
   m2RotatePoint.x = xA
   m2RotatePoint.y = yA
Else
   m2RotatePoint = P
End If

End Function
