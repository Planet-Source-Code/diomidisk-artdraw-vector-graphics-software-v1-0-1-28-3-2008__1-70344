Attribute VB_Name = "vbdStuff"
Option Explicit

Private m_OldPen As Long
Private m_OldBrush As Long
Private m_NewBrush As Long
Private m_NewPen As Long

' Bound the objects in the collection.
Public Sub BoundObjects(ByVal the_objects As Collection, ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim Obj As vbdObject
Dim R As RECT
    
    Set Obj = the_objects(1)
   ' Obj.Bound xMin, yMin, xMax, yMax
   
    GetRgnBox Obj.hRegion, R
    xmin = R.Left
    ymin = R.Top
    xmax = R.Right
    ymax = R.Bottom
    
    'For Each Obj In the_objects
    '    Obj.Bound X1, Y1, X2, Y2
    '    If xMin > X1 Then xMin = X1
    '    If xMax < X2 Then xMax = X2
    '    If yMin > Y1 Then yMin = Y1
    '    If yMax < Y2 Then yMax = Y2
    'Next Obj
    
End Sub

Public Sub NewTransformation()
    Dim the_scene As vbdScene
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.NewTransformation
    Set the_scene = Nothing
End Sub

' Return this object's bounds.
Public Sub BoundText(ByRef Points() As POINTAPI, ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim I As Integer, m_NumPoints As Long
    m_NumPoints = UBound(Points)
    If m_NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With Points(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For I = 2 To m_NumPoints
            With Points(I)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next I
    End If
End Sub



' StartBound the objects in the collection.
Public Sub StartBoundObjects(ByVal the_objects As Collection, ByRef xmin As Single, ByRef ymin As Single)
    Dim Obj As vbdObject
    Set Obj = the_objects(1)
    Obj.StartBound xmin, ymin
End Sub

' Initialize default drawing properties.
Public Sub InitializeDrawingProperties(ByVal Obj As vbdObject)
    Obj.DrawWidth = 1
    Obj.DrawStyle = vbSolid
    Obj.ForeColor = vbBlack
    Obj.FillColor = vbBlack
    Obj.FillStyle = vbFSTransparent
    Obj.FillMode = fALTERNATE
    Obj.TextDraw = ""
    Obj.TypeDraw = 0
    Obj.TypeFill = 0
    Obj.Bold = False
    Obj.Charset = 0
    Obj.Italic = False
    Obj.Name = "Arial"
    Obj.Size = 20
    Obj.Strikethrough = False
    Obj.Underline = False
    Obj.Weight = 400
    Obj.Angle = 0
    
End Sub
' Return the drawing property serialization
' for this object.
Public Function DrawingPropertySerialization(ByVal Obj As vbdObject) As String
Dim txt As String

    txt = txt & " DrawWidth(" & Format$(Obj.DrawWidth) & ")"
    txt = txt & " DrawStyle(" & Format$(Obj.DrawStyle) & ")"
    txt = txt & " ForeColor(" & Format$(Obj.ForeColor) & ")"
    txt = txt & " FillColor(" & Format$(Obj.FillColor) & ")"
    txt = txt & " FillColor2(" & Format$(Obj.FillColor2) & ")"
    txt = txt & " FillMode(" & Format$(Obj.FillMode) & ")"
    txt = txt & " Pattern(" & Trim(Obj.Pattern) & ")"
    txt = txt & " Gradient(" & Format$(Obj.Gradient) & ")"
    txt = txt & " FillStyle(" & Format$(Obj.FillStyle) & ")"
    txt = txt & " TextDraw(" & Format$(Obj.TextDraw) & ")"
    txt = txt & " TypeDraw(" & Format$(Obj.TypeDraw) & ")"
    txt = txt & " CurrentX(" & Format$(Obj.CurrentX) & ")"
    txt = txt & " CurrentY(" & Format$(Obj.CurrentY) & ")"
    txt = txt & " TypeFill(" & Format$(Obj.TypeFill) & ")"
    txt = txt & " Shade(" & Format$(Obj.Shade) & ")"
    txt = txt & " ObjLock(" & Format$(Obj.ObjLock) & ")"
    txt = txt & " Blend(" & Format$(Obj.Blend) & ")"
    
    txt = txt & " Bold(" & Format$(Obj.Bold) & ")"
    txt = txt & " Charset(" & Format$(Obj.Charset) & ")"
    txt = txt & " Italic(" & Format$(Obj.Italic) & ")"
    txt = txt & " Name(" & Format$(Obj.Name) & ")"
    txt = txt & " Size(" & Format$(Obj.Size) & ")"
    txt = txt & " Strikethrough(" & Format$(Obj.Strikethrough) & ")"
    txt = txt & " Underline(" & Format$(Obj.Underline) & ")"
    txt = txt & " Weight(" & Format$(Obj.Weight) & ")"
    txt = txt & " Angle(" & Format$(Obj.Angle) & ")"
    
    DrawingPropertySerialization = txt & vbCrLf & "    "
End Function

' Read the token name and value and to see
' if it is drawing property information.
Public Sub ReadDrawingPropertySerialization(ByVal Obj As vbdObject, ByVal token_name As String, ByVal token_value As String)
    
    Select Case token_name
        Case "DrawWidth"
            Obj.DrawWidth = CInt(token_value)
        Case "DrawStyle"
            Obj.DrawStyle = CInt(token_value)
        Case "ForeColor"
            Obj.ForeColor = CLng(token_value)
        Case "FillColor"
            Obj.FillColor = CLng(token_value)
        Case "FillColor2"
           Obj.FillColor2 = CLng(token_value)
         Case "FillMode"
            Obj.FillMode = CSng(token_value)
        Case "Pattern"
           Obj.Pattern = Trim(token_value)
        Case "Gradient"
           Obj.Gradient = CInt(token_value)
        Case "FillStyle"
            Obj.FillStyle = CInt(token_value)
        Case "TextDraw"
            Obj.TextDraw = Trim(token_value)
        Case "TypeDraw"
            Obj.TypeDraw = CInt(token_value)
        Case "TypeFill"
            Obj.TypeFill = CInt(token_value)
        Case "Shade"
            Obj.Shade = Val(token_value)
        Case "ObjLock"
            Obj.ObjLock = CBool(token_value)
        Case "Blend", "Opacity"
            Obj.Blend = Val(token_value)
        Case "Angle"
            Obj.Angle = CSng(token_value)
        Case "Charset"
            Obj.Charset = CInt(token_value)
        Case "Italic"
            Obj.Italic = CBool(token_value)
        Case "Name"
            Obj.Name = Trim(token_value)
        Case "Size"
            Obj.Size = CInt(token_value)
        Case "Bold"
            Obj.Bold = CBool(token_value)
        Case "Strikethrough"
            Obj.Strikethrough = CBool(token_value)
        Case "Underline"
            Obj.Underline = CBool(token_value)
        Case "Weight"
            Obj.Weight = CLng(token_value)
        Case "CurrentX"
            Obj.CurrentX = CSng(token_value)
        Case "CurrentY"
            Obj.CurrentY = CSng(token_value)
        Case "AlingText"
            Obj.AlingText = CSng(token_value)
        Case Else
         ' Stop
    End Select
End Sub


' Set the drawing properties for the canvas.
Public Sub SetCanvasDrawingParameters(ByVal Obj As vbdObject, ByVal canvas As PictureBox)
    Dim Newfonts As New StdFont
    Dim hBrush As Long, OldBrush As Long
    On Error Resume Next
     '  canvas.DrawWidth = Obj.DrawWidth * gZoomFactor
     '  canvas.DrawStyle = Obj.DrawStyle
     '  canvas.ForeColor = Obj.ForeColor
     '  Canvas.FillColor = Obj.FillColor
      
     '  canvas.FillStyle = Obj.FillStyle

'       Newfonts.Bold = Obj.Bold
'       Newfonts.Charset = Obj.Charset
'       Newfonts.Italic = Obj.Italic
'       Newfonts.Name = Obj.Name
'       Newfonts.Size = Obj.Size * gZoomFactor
'       Newfonts.Strikethrough = Obj.Strikethrough
'       Newfonts.Underline = Obj.Underline
'       Newfonts.Weight = Obj.Weight
       
      ' canvas.Font = Newfonts
       
       'Debug.Print Obj.TypeDraw
       'Debug.Print Obj.TypeFill
       'Debug.Print Obj.TextDraw
    On Error GoTo 0
End Sub

' Set the drawing properties for the metafile.
Public Sub SetMetafileDrawingParameters(ByVal Obj As vbdObject, ByVal mf_dc As Long)
Dim log_brush As LogBrush
Dim new_brush As Long
Dim new_pen As Long

    With log_brush
        If Obj.FillStyle = vbFSTransparent Then
            .lbStyle = BS_HOLLOW
        ElseIf Obj.FillStyle = vbFSSolid Then
            .lbStyle = BS_SOLID
        Else
            .lbStyle = BS_HATCHED
            Select Case Obj.FillStyle
                Case vbCross
                    .lbHatch = HS_CROSS
                Case vbDiagonalCross
                    .lbHatch = HS_DIAGCROSS
                Case vbDownwardDiagonal
                    .lbHatch = HS_BDIAGONAL
                Case vbHorizontalLine
                    .lbHatch = HS_HORIZONTAL
                Case vbUpwardDiagonal
                    .lbHatch = HS_FDIAGONAL
                Case vbVerticalLine
                    .lbHatch = HS_VERTICAL
            End Select
        End If
        .lbColor = Obj.FillColor
    End With

    m_NewPen = CreatePen(Obj.DrawStyle, Obj.DrawWidth, Obj.ForeColor)
    m_NewBrush = CreateBrushIndirect(log_brush)
    m_OldPen = SelectObject(mf_dc, m_NewPen)
    m_OldBrush = SelectObject(mf_dc, m_NewBrush)
End Sub

' Restore the drawing properties for the metafile.
Public Sub RestoreMetafileDrawingParameters(ByVal mf_dc As Long)
    SelectObject mf_dc, m_OldBrush
    SelectObject mf_dc, m_OldPen
    DeleteObject m_NewBrush
    DeleteObject m_NewPen
End Sub

' Return the serialization for this transformation matrix.
Public Function TransformationSerialization(M() As Single) As String
Dim I As Integer
Dim J As Integer
Dim txt As String

    For I = 1 To 3
        For J = 1 To 3
            txt = txt & Format$(M(I, J)) & " "
        Next J
    Next I

    TransformationSerialization = "Transformation(" & txt & ")"
End Function

' initialize the transformation matrix using this serialization.
Public Sub SetTransformationSerialization(ByVal txt As String, M() As Single)
Dim I As Integer
Dim J As Integer
Dim token As String

    For I = 1 To 3
        For J = 1 To 3
            token = GetDelimitedToken(txt, " ")
            token = Replace(token, ",", ".")
            M(I, J) = CSng(Val(token))
        Next J
    Next I
End Sub

Public Function LoadPatternPic(FileImage As String, _
                              Optional dWidth As Long = 8, _
                              Optional dHeight As Long = 8) As Long
     Const LR_LOADFROMFILE = &H10
     Const IMAGE_BITMAP = 0
     
     If FileExists(FileImage) Then
        LoadPatternPic = LoadImage(App.hInstance, FileImage, IMAGE_BITMAP, dWidth, dHeight, LR_LOADFROMFILE)
     End If
End Function

'Create pen Style
Public Function PenCreate(mDrawStyle As Integer, mWidthLine As Integer, mColorLine As Long) As Long
Dim BrushInf As LogBrush
Dim StyleArr() As Long
Dim wLine As Long
Dim PenStyle As Long
    
    wLine = mWidthLine
    
    
    Select Case mDrawStyle
    Case 0 'vbSolid
'        ReDim StyleArr(3)
'        StyleArr(0) = 1
'        StyleArr(1) = 0
'        StyleArr(2) = 0
'        StyleArr(3) = 0
        PenCreate = CreatePen(PS_SOLID, wLine, mColorLine)
        Exit Function
    Case 1 'vbDash
       ReDim StyleArr(1)
        StyleArr(0) = 18 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)
        'StyleArr(2) = 6 * wLine
        'StyleArr(3) = 4 * wLine

    Case 2 'vbDot
       ReDim StyleArr(3)
        StyleArr(0) = 3 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
        
    Case 3 'vbDashDot
       ReDim StyleArr(3)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 6 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 6 * (wLine / 2)
    
    Case 4 'vbDashDotDot
        ReDim StyleArr(5)
        StyleArr(0) = 9 * (wLine / 2)
        StyleArr(1) = 3 * (wLine / 2)
        StyleArr(2) = 3 * (wLine / 2)
        StyleArr(3) = 3 * (wLine / 2)
        StyleArr(4) = 3 * (wLine / 2)
        StyleArr(5) = 3 * (wLine / 2)
        
    Case 5 'vbInvisible
'        ReDim StyleArr(3)
'        StyleArr(0) = 0
'        StyleArr(1) = 0
'        StyleArr(2) = 0
'        StyleArr(3) = 0
        PenCreate = CreatePen(PS_NULL, wLine, mColorLine)
        Exit Function
    End Select
    
    BrushInf.lbColor = mColorLine
    PenCreate = ExtCreatePen(PS_GEOMETRIC Or PS_USERSTYLE, wLine, BrushInf, UBound(StyleArr()) + 1, StyleArr(0))
    
    Erase StyleArr
    
End Function

Public Function BitmapFromDC(ByVal lhDC As Long, _
                             ByVal lLeft As Long, _
                             ByVal lTop As Long, _
                             ByVal lWidth As Long, _
                             ByVal lHeight As Long) As Long ', _
                             ByVal lAngle As Single) , _
                             ByVal picWidth As Long, _
                             ByVal picheight As Long) As Long

   ' Copy the bitmap in lHDC:
   Dim lhDCCopy As Long
   Dim lhBmpCopy As Long
   Dim lhBmpCopyOld As Long
   Dim lhDCC As Long
   Dim tBM As Bitmap
   'Dim PlgPts(1 To 4) As POINTAPI
'   Dim PlgPts(0 To 4) As POINTAPI
'   Dim PicWidth As Long, PicHeight As Long
'   Dim HalfWidth As Single, HalfHeight As Single
'   Dim AngleRad As Single
'   Const HalfPi As Single = PI * 0.5
   
   lhDCC = CreateDCA("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   lhDCCopy = CreateCompatibleDC(lhDCC)
   
   lhBmpCopy = CreateCompatibleBitmap(lhDCC, lWidth, lHeight)
   lhBmpCopyOld = SelectObject(lhDCCopy, lhBmpCopy)
   Call BitBlt(lhDCCopy, lLeft, lTop, lWidth, lHeight, lhDC, 0, 0, vbSrcCopy)
   
'  HalfWidth = lWidth / 2
'  HalfHeight = lHeight / 2
'  AngleRad = (lAngle / 180) * PI

  ' Get picture size
  '  picWidth = hWidth
  'picheight = hheight

  ' Get half picture size and angle in radians
'  HalfWidth = picWidth / 2
'  HalfHeight = picheight / 2
'  AngleRad = (lAngle / 180) * PI
'
'    PlgPts(0).X = Cos(AngleRad) * HalfWidth
'    PlgPts(0).Y = Sin(AngleRad) * HalfWidth
'    PlgPts(1).X = Cos(AngleRad + HalfPi) * HalfHeight
'    PlgPts(1).Y = Sin(AngleRad + HalfPi) * HalfHeight
'
'    ' Project parallelogram points for rotated area
'    PlgPts(2).X = HalfWidth + lLeft - PlgPts(0).X - PlgPts(1).X
'    PlgPts(2).Y = HalfHeight + lTop - PlgPts(0).Y - PlgPts(1).Y
'    PlgPts(3).X = HalfWidth + lLeft - PlgPts(1).X + PlgPts(0).X
'    PlgPts(3).Y = HalfHeight + lTop - PlgPts(1).Y + PlgPts(0).Y
'    PlgPts(4).X = HalfWidth + lLeft - PlgPts(0).X + PlgPts(1).X
'    PlgPts(4).Y = HalfHeight + lTop - PlgPts(0).Y + PlgPts(1).Y
'    PlgBlt lhDCCopy, PlgPts(2), lhDC, 0, 0, picWidth, picheight, 0, 0, 0
    
   
   If Not (lhDCC = 0) Then
      DeleteDC lhDCC
   End If
   If Not (lhBmpCopyOld = 0) Then
      SelectObject lhDCCopy, lhBmpCopyOld
   End If
   If Not (lhDCCopy = 0) Then
      DeleteDC lhDCCopy
   End If

   BitmapFromDC = lhBmpCopy

End Function

' Return the next delimited token from txt.
' Trim blanks.
Public Function GetDelimitedToken(ByRef txt As String, ByVal delimiter As String) As String
Dim pos As Integer

    pos = InStr(txt, delimiter)
    If pos < 1 Then
        ' The delimiter was not found. Return
        ' the rest of txt.
        GetDelimitedToken = Trim$(txt)
        txt = ""
    Else
        ' We found the delimiter. Return the token.
        GetDelimitedToken = Trim$(Left$(txt, pos - 1))
        txt = Trim$(Mid$(txt, pos + Len(delimiter)))
    End If
End Function
' Replace non-printable characters in txt with spaces.
Public Function NonPrintingToSpace(ByVal txt As String) As String
Dim I As Integer
Dim cH As String

    For I = 1 To Len(txt)
        cH = Mid$(txt, I, 1)
        If (cH < " ") Or (cH > "~") Then Mid$(txt, I, 1) = " "
    Next I
    NonPrintingToSpace = txt
End Function


' Remove comments starting with  from the end of lines.
Public Function RemoveComments(ByVal txt As String) As String
Dim pos As Integer
Dim new_txt As String

    Do While Len(txt) > 0
        ' Find the next '.
        pos = InStr(txt, "'")
        If pos = 0 Then
            new_txt = new_txt & txt
            Exit Do
        End If

        ' Add this part to the result.
        new_txt = new_txt & Left$(txt, pos - 1)

        ' Find the end of the line.
        pos = InStr(pos + 1, txt, vbCrLf)
        If pos = 0 Then
            ' There was no vbCrLf.
            ' Remove the rest of the text.
            txt = ""
        Else
            txt = Mid$(txt, pos + Len(vbCrLf))
        End If
    Loop

    RemoveComments = new_txt
End Function


Public Function PolygonPoints(cPtsQty As Integer, cLeft As Single, cTop As Single, cWidth As Single, cHeight As Single) As POINTAPI()

Dim POINT() As POINTAPI
Dim n As Integer
Dim RadiusW As Single
Dim RadiusH As Single
Dim iCounter As Integer
Dim R As Single
Dim Alfa As Single

RadiusW = (cWidth - cLeft) / 2
RadiusH = (cHeight - cTop) / 2

ReDim POINT(cPtsQty)
iCounter = 0
For n = 0 To 360 Step 360 / cPtsQty
    POINT(iCounter).X = RadiusW + Sin(n * PI / 180) * RadiusW
    POINT(iCounter).Y = RadiusH + Cos(n * PI / 180) * RadiusH
    R = Sqr(POINT(iCounter).X ^ 2 + POINT(iCounter).Y ^ 2)
    Alfa = m2Atn2(POINT(iCounter).Y, POINT(iCounter).X)
    POINT(iCounter).X = cLeft + R * Cos(Alfa)
    POINT(iCounter).Y = cTop + R * Sin(Alfa)
    iCounter = iCounter + 1
Next

PolygonPoints = POINT

End Function


' Find the distance from the point (x1, y1) to the
' line passing through (x1, y1) and (x2, y2).
Public Function DistPointToLine(ByVal A As Single, ByVal b As Single, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
Dim vx As Single
Dim vy As Single
Dim t As Single
Dim dx As Single
Dim dy As Single
Dim close_x As Single
Dim close_y As Single
'On Error GoTo errnum
    ' Get the vector component for the segment.
    ' The segment is given by:
    '       x(t) = x1 + t * vx
    '       y(t) = y1 + t * vy
    ' where 0.0 <= t <= 1.0
    vx = x2 - x1
    vy = y2 - y1

    ' Find the best t value.
    If (vx = 0) And (vy = 0) Then
        ' The points are the same. There is no segment.
        t = 0
    Else
        ' Calculate the minimal value for t.
        If (vx * vx + vy * vy) <> 0 Then
        t = -((x1 - A) * vx + (y1 - b) * vy) / (vx * vx + vy * vy)
        End If
    End If

    ' Keep the point on the segment.
    If t < 0# Then
        t = 0#
    ElseIf t > 1# Then
        t = 1#
    End If

    ' Set the return values.
    close_x = x1 + t * vx
    close_y = y1 + t * vy
    dx = A - close_x
    dy = b - close_y
    DistPointToLine = Sqr(dx * dx + dy * dy)
errnum:
   On Error GoTo 0
End Function

' Return True if the polygon is at this location.
Public Function PolygonIsAt(ByVal is_closed As Boolean, ByVal X As Single, ByVal Y As Single, Points() As POINTAPI) As Boolean
Const HIT_DIST = 10
Dim start_i As Integer
Dim I As Integer
Dim num_points As Integer
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim Dist As Single

    PolygonIsAt = False
     
    num_points = UBound(Points)
    If is_closed Then
        x2 = Points(num_points).X
        y2 = Points(num_points).Y
        start_i = 1
    Else
        x2 = Points(1).X
        y2 = Points(1).Y
        start_i = 2
    End If

    ' Check each segment in the Polyline.
    For I = start_i To num_points
        With Points(I)
            x1 = .X
            y1 = .Y
        End With
        Dist = DistPointToLine(X, Y, x1, y1, x2, y2)
        If Dist <= HIT_DIST Then
            PolygonIsAt = True
            Exit For
        End If
        x2 = x1
        y2 = y1
    Next I
End Function

' Return True if the point is inside the object.
Public Function PointIsInPolygon(ByVal X As Single, ByVal Y As Single, Points() As POINTAPI) As Boolean
Dim polygon_region As Long

    polygon_region = CreatePolygonRgn(Points(1), UBound(Points), fALTERNATE)
    PointIsInPolygon = PtInRegion(polygon_region, X, Y)
    DeleteObject polygon_region
End Function

 'Return a named token from the string txt.
' Tokens have the form TokenName(TokenValue).
Public Sub GetNamedToken(ByRef txt As String, ByRef token_name As String, ByRef token_value As String)
Dim pos1 As Long
Dim pos2 As Long
Dim open_parens As Long
Dim cH As String

    ' Find the "(".
    pos1 = InStr(txt, "(")
    If pos1 = 0 Then
        ' No "(" found. Return the rest as the token name.
        token_name = Trim$(txt)
        token_value = ""
        txt = ""
        Exit Sub
    End If

    ' Find the corresponding ")". Note that
    ' parentheses may be nested.
    open_parens = 1
    pos2 = pos1 + 1
    Do While pos2 <= Len(txt)
        cH = Mid$(txt, pos2, 1)
        If cH = "(" Then
            open_parens = open_parens + 1
        ElseIf cH = ")" Then
            open_parens = open_parens - 1
            If open_parens = 0 Then
                ' This is the corresponding ")".
                Exit Do
            End If
        End If
        pos2 = pos2 + 1
    Loop

    ' At this point, pos1 points to the ( and
    ' pos2 points to the ).
    token_name = Trim$(Left$(txt, pos1 - 1))
    token_value = Trim$(Mid$(txt, pos1 + 1, pos2 - pos1 - 1))
    txt = Trim$(Mid$(txt, pos2 + 1))
End Sub

' Replace non-printable characters with spaces.
Public Function RemoveNonPrintables(ByVal txt As String) As String
Dim pos As Integer
Dim cH As String
  On Error Resume Next
    'For pos = 1 To Len(txt)
    '    cH = Mid$(txt, pos, 1)
    '    If (cH < " ") Or (cH > "~") Then
    '       Mid$(txt, pos, 1) = " "
    '    End If
    'Next pos
    For pos = 1 To 32
        txt = Replace(txt, Chr(pos), " ")
    Next
    For pos = 126 To 255
        txt = Replace(txt, Chr(pos), " ")
    Next
    RemoveNonPrintables = txt
End Function



