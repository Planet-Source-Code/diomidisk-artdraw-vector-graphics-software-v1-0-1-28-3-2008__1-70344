VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbdDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VbDraw Polygon/Polyline object.

Implements vbdObject

Private Enum vbdDrawErrors
    errInvalidIndex = 381   ' Invalid property array index.
End Enum

' Indicates a closed polygon rather than a polyline.
Public IsClosed As Boolean

' The surface on which the user is clicking
' to define the object. This is set only during
' creation of this object.
'Public WithEvents m_Canvas As PictureBox
Private m_DrawingStarted As Boolean

' Drawing properties.
Public m_DrawWidth As Integer
Public m_DrawStyle As Integer
Public m_ForeColor As OLE_COLOR
Public m_FillColor As OLE_COLOR
Public m_FillColor2 As OLE_COLOR
Public m_FillMode As FillMode
Public m_Pattern As String
Public m_Gradient As Integer
Public m_FillStyle As Integer
Public m_TypeDraw As DrawType
Public m_TextDraw As String
Public m_AlingText As Integer
Public m_TypeFill As DrawTypeFill
Public m_Shade As Boolean
Public m_ObjLock As Boolean
Public m_Blend As Integer
'Public m_TextProperty As New StdFont
Public m_Bold As Boolean
Public m_Charset As Integer
Public m_Italic As Boolean
Public m_Name As String
Public m_Size As Integer
Public m_Strikethrough As Boolean
Public m_Underline As Boolean
Public m_Weight As Integer

Public m_angle As Single
Public m_Selected As Boolean
Public m_hRegion As Long
Public m_Picture As StdPicture

' Data variables.
Private m_NumPoints As Long
Private m_TypePoint() As Byte
Private m_OriginalPoints() As POINTAPI
Private m_TransformedPoints() As POINTAPI
Private m_M(1 To 3, 1 To 3) As Single

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single

Public Property Let TypePoint(ByVal Index As Integer, ByVal new_value As Integer)
       m_TypePoint(Index) = new_value
End Property

' Return the number of points.
Public Property Get NumPoints() As Integer
    NumPoints = m_NumPoints
End Property

' Set the number of points.
Public Property Let NumPoints(ByVal new_value As Integer)
    m_NumPoints = new_value
    If m_NumPoints < 1 Then
        Erase m_OriginalPoints
        Erase m_TypePoint
    Else
        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
        ReDim Preserve m_TypePoint(1 To m_NumPoints)
    End If
End Property
' Return an X coordinate.
Property Get X(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.X"
    End If

    X = m_OriginalPoints(Index).X
End Property
' Set an X coordinate.
Property Let X(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.X"
    End If

    m_OriginalPoints(Index).X = new_value
End Property

' Return a Y coordinate.
Property Get Y(ByVal Index As Integer) As Single
    If (Index < 1) Or (Index > m_NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.Y"
    End If

    Y = m_OriginalPoints(Index).Y
End Property
' Set a Y coordinate.
Property Let Y(ByVal Index As Integer, ByVal new_value As Single)
    If (Index < 1) Or (Index > NumPoints) Then
        Err.Raise errInvalidIndex, "Draw.Y"
    End If
    m_OriginalPoints(Index).Y = new_value
End Property

' Start with an identity transformation.
Private Sub Class_Initialize()
    ' Initialize the drawing parameters.
    InitializeDrawingProperties Me
    m2Identity m_M
End Sub

' The user has selected a point.
Private Sub m_Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    ' If we are drawing a rubberband line,
'    ' erase it.
'    If m_NumPoints > 0 Then m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
'
'    ' If this is the first point, start using dotted vbInvert mode.
'    If m_NumPoints = 0 Then
'        m_Canvas.DrawMode = vbInvert
'        m_Canvas.DrawStyle = vbDot
'        m_DrawingStarted = True
'    End If
'
'    ' See if this is the left or right button.
'    If Button And vbLeftButton Then
'        ' It's the left button.
'        ' Add this point to the polygon.
'
'        ' Add the new point.
'        m_NumPoints = m_NumPoints + 1
'        ReDim Preserve m_OriginalPoints(1 To m_NumPoints)
'        ReDim Preserve m_TypePoint(1 To m_NumPoints)
'        m_OriginalPoints(m_NumPoints).X = X / gZoomFactor
'        m_OriginalPoints(m_NumPoints).Y = Y / gZoomFactor
'        m_TypePoint(m_NumPoints) = 2
'        ' Draw the line permanently.
'        m_Canvas.DrawMode = vbCopyPen
'        m_Canvas.DrawStyle = vbSolid
'        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
'        m_Canvas.DrawMode = vbInvert
'        m_Canvas.DrawStyle = vbDot
'
'        ' Start the next rubberband line.
'        m_StartX = X
'        m_StartY = Y
'        m_LastX = X
'        m_LastY = Y
'        m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
'    Else
'        ' It's the right button.
'        ' Stop building the polygon.
'
'        ' Go back to vbCopyPen drawing mode.
'        m_Canvas.DrawMode = vbCopyPen
'
'        ' Stop receiving events from the canvas.
'        Set m_Canvas = Nothing
'
'        ' If we have at least 2 points, tell the form to save us.
'        If m_NumPoints >= 2 Then
'            m_TypePoint(1) = 6
'            m_TypePoint(m_NumPoints) = 3
'            Me.vbdObject_DrawStyle = m_DrawStyle
'            Me.vbdObject_DrawWidth = m_DrawWidth
'            Me.vbdObject_FillColor = RGB(255, 255, 255) 'm_FillColor
'            Me.vbdObject_FillStyle = m_FillStyle
'            Me.vbdObject_ForeColor = RGB(0, 0, 0) 'm_ForeColor
'            Me.vbdObject_TextDraw = m_TextDraw
'            Me.vbdObject_TypeDraw = m_TypeDraw
'            Me.vbdObject_TypeDraw = m_TypeDraw
'            Me.vbdObject_TypeFill = m_Blend
'            Me.vbdObject_Shade = m_Shade
'
'            Me.vbdObject_MakeTransformation
'            ' We have at least 2 points. Tell the form to save us.
'            AddObject Me
'            'm_TransformedPoints = m_OriginalPoints
'        Else
'            ' We do not have 3 points. Tell the
'            ' form to cancel us.
'            frmVbDraw.CancelObject
'        End If
'    End If
End Sub

' Continue drawing the rubberband line.
Private Sub m_Canvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If Not m_DrawingStarted Then Exit Sub
'
'    ' Start using dotted vbInvert mode.
'    m_Canvas.DrawMode = vbInvert
'    m_Canvas.DrawStyle = vbDot
'
'    ' Erase the old line.
'    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
'
'    ' Update the point.
'    m_LastX = X
'    m_LastY = Y
'
'    ' Draw the new line.
'    m_Canvas.Line (m_StartX, m_StartY)-(m_LastX, m_LastY)
End Sub


Private Property Get vbdObject_AlingText() As Integer
        vbdObject_AlingText = m_AlingText
End Property

Private Property Let vbdObject_AlingText(ByVal RHS As Integer)
       m_AlingText = RHS
End Property

Private Property Let vbdObject_Angle(ByVal RHS As Single)
        m_angle = RHS
End Property

Private Property Get vbdObject_Angle() As Single
         vbdObject_Angle = m_angle
End Property

Public Property Let vbdObject_Bold(ByVal RHS As Boolean)
       m_Bold = RHS
End Property

Public Property Get vbdObject_Bold() As Boolean
       vbdObject_Bold = m_Bold
End Property

Private Property Set vbdObject_Canvas(ByVal RHS As PictureBox)
    'Set m_Canvas = RHS
End Property

Private Property Get vbdObject_Canvas() As PictureBox
    'Set vbdObject_Canvas = m_Canvas
End Property

Private Property Let vbdObject_Charset(ByVal RHS As Integer)
         m_Charset = RHS
End Property

Private Property Get vbdObject_Charset() As Integer
          vbdObject_Charset = m_Charset
End Property

' Clear the object's transformation.
Private Sub vbdObject_ClearTransformation()
    m2Identity m_M
    m_TransformedPoints = m_OriginalPoints
End Sub

' Add this transformation to the current one.
Private Sub vbdObject_AddTransformation(M() As Single)
    Dim t(1 To 3, 1 To 3) As Single
    m2MatMultiply t, m_M, M
    m2MatCopy m_M, t
    
End Sub

Private Property Let vbdObject_CurrentX(ByVal RHS As Single)

End Property

Private Property Get vbdObject_CurrentX() As Single

End Property

Private Property Let vbdObject_CurrentY(ByVal RHS As Single)

End Property

Private Property Get vbdObject_CurrentY() As Single

End Property

' Draw the object in a metafile.
Private Sub vbdObject_DrawInMetafile(ByVal mf_dc As Long)
Dim xmin As Single
Dim xmax As Single
Dim ymin As Single
Dim ymax As Single
Dim Br As LogBrush
Dim hBrush As Long, lPen As Long, nL As Long, nT As Long
Dim NRECT As RECT, oldfont As Long, NewFont As Long
Dim OldBrush As Long, hOldPen As Long, mDC As Long, mBitmap As Long
Dim imgData As Long, nW As Long, nH As Long

    On Error GoTo CheckPoints
    If UBound(m_TransformedPoints) = 0 Then
       m_TransformedPoints = m_OriginalPoints
    End If
    
    On Error GoTo 0
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
    
    TransformPoints gZoomFactor
    
    If m_ViewSimple Then
        Br.lbColor = Me.m_FillColor
        Br.lbHatch = -1
        Br.lbStyle = 2
        hBrush = CreateBrushIndirect(Br)
        lPen = PenCreate(Me.m_DrawStyle, 1, Me.m_ForeColor)
        StrokePath mf_dc
    Else
        Select Case Me.m_FillStyle
        Case 0  'Solid
            hBrush = CreateSolidBrush(Me.m_FillColor)
        Case 8  'Bitmap pattern
            imgData = LoadPatternPic(App.Path + "\Pattern\" + Me.m_Pattern)
            hBrush = CreatePatternBrush(imgData)
        Case 10  'Bitmap
            'GetRgnBox Me.m_hRegion, R
           
            'nL = R.Left
            'nT = R.Top
            'nW = R.Right - R.Left
            'nH = R.Bottom - R.Top
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax - xmin
            nH = ymax - ymin
            nL = xmin
            nT = ymin
            mDC = CreateCompatibleDC(0)
            imgData = LoadImage(mDC, Me.m_Pattern, 0, nW, nH, &H10)
            SelectObject mDC, imgData
            If nL < 0 Then nL = 0
            If nT < 0 Then nT = 0
            imgData = BitmapFromDC(mDC, nL, nT, xmax, ymax) ', Me.m_angle, R.Right, R.Bottom)
            
            SelectObject mDC, imgData
            hBrush = CreatePatternBrush(imgData)
           
        Case 9  'Gradient
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax
            nH = ymax
            'Create a device context, compatible with the screen
            mDC = CreateCompatibleDC(GetDC(0))
            'Create a bitmap, compatible with the screen
            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
            SelectObject mDC, mBitmap
            GradientFillRectDC mDC, xmin - 5, ymin - 5, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
            hBrush = CreatePatternBrush(mBitmap)
            
        Case Else   'Other Patterns
            Br.lbColor = Me.m_FillColor
            Br.lbHatch = Me.m_FillStyle - 2
            Br.lbStyle = 2
            hBrush = CreateBrushIndirect(Br)
     End Select
         lPen = PenCreate(Me.m_DrawStyle, Me.m_DrawWidth, Me.m_ForeColor)
    End If
    
    StrokePath mf_dc
    OldBrush = SelectObject(mf_dc, hBrush)
    hOldPen = SelectObject(mf_dc, lPen)
       
    'Make Path for Region
     BeginPath mf_dc
     DrawObjectWMF mf_dc  ', True
     EndPath mf_dc
     
     'Delete  old Region
     DeleteObject Me.m_hRegion
     'New Region
     Me.m_hRegion = PathToRegion(mf_dc)

     'Draw
     If m_TypeDraw = dText Or m_TypeDraw = dPolydraw Or m_TypeDraw = dCurve Or m_TypeDraw = dEllipse Then   'Or m_TypeDraw = dText Then  'Or m_TypeDraw = dTextFrame Then
        DrawObjectWMF mf_dc, True      'for fill object with PolyDraw
        DrawObjectWMF mf_dc             'for fill object with PolyDraw
     ElseIf m_TypeDraw = dTextFrame Then
        DrawObjectWMF mf_dc, True      'for fill object with PolyDraw
    Else
        'DrawObjectWMF mf_dc, True      'for fill object with PolyDraw
        DrawObjectWMF mf_dc, True
     End If
     
'    If m_TypeDraw = dTextFrame Then
'       Call ModifyWorldTransform(mf_dc, XForm, MWT_IDENTITY)
'    End If
       
       If m_ViewSimple = False Then
         ' If m_TypeDraw = dText Or m_TypeDraw = dTextFrame Then
         '     FillPath mf_dc
         ' Else
              StrokeAndFillPath mf_dc
         ' End If
       End If
       
'    If m_TypeDraw = dTextFrame Then
'       'Return default settings
'       SetWorldTransform mf_dc, OldXForm
'       SetMapMode mf_dc, OldMapMode
'       SetGraphicsMode mf_dc, OldMode
'    End If
    'De-select Brush-pen from DC
    Call SelectObject(mf_dc, OldBrush)
    Call SelectObject(mf_dc, hOldPen)
    
    DeleteObject hBrush
    DeleteObject lPen
    DeleteObject imgData
    DeleteObject mBitmap
    DeleteDC mDC
    Exit Sub
    
CheckPoints:
   If Err = 9 Then
      Resume Next
   End If
End Sub

Private Sub vbdObject_DrawPoint()

End Sub

' Draw box on the object.
Private Sub vbdObject_DrawSelect(ByVal pic As PictureBox)

    Dim xmin As Single
    Dim xmax As Single
    Dim ymin As Single
    Dim ymax As Single
    Dim OldDrawWirdth As Integer
    Dim OldDrawStyle As Integer
    Dim OldFillStyle As Integer
    Dim OldFillColor As Long
    Dim olddrawmode As Integer
    Dim OldForeColor As Long
    Dim R As RECT
    Dim oldmode As Long, OldMapMode As Long
    Dim XForm As XForm
    Dim OldXForm As XForm
    
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
 
    ' Highlight if necessary.
    If m_Selected Then
       'OldMode = SetGraphicsMode(Pic.hdc, GM_ADVANCED)
       'OldMapMode = SetMapMode(Pic.hdc, MM_TEXT)
       'GetWorldTransform Pic.hdc, OldXForm
       ' XForm.eM11 = m_M(1, 1)
       ' XForm.eM12 = m_M(1, 2)
       ' XForm.eM21 = m_M(2, 1)
       ' XForm.eM22 = m_M(2, 2)
       ' XForm.eDx = m_M(3, 1)
       ' XForm.eDy = m_M(3, 2)
       ' 'm_TransformedPoints = m_OriginalPoints
       ' SetWorldTransform Pic.hdc, XForm
       ' Call ModifyWorldTransform(Pic.hdc, XForm, MWT_IDENTITY)
        If Me.m_TypeDraw = dTextFrame Then
           vbdObject_Bound xmin, ymin, xmax, ymax
        Else
           GetRgnBox Me.m_hRegion, R
           xmin = R.Left
           ymin = R.Top
           xmax = R.Right
           ymax = R.Bottom
           If R.Left = 0 And R.Top = 0 And R.Right = 0 And R.Bottom = 0 Then
              vbdObject_Bound xmin, ymin, xmax, ymax
           End If
        End If
        'Return default settings
'        SetWorldTransform Pic.hDC, OldXForm
'        SetMapMode Pic.hDC, OldMapMode
'        SetGraphicsMode Pic.hDC, OldMode
        
        'Backup parameter
        OldDrawWirdth = pic.DrawWidth
        OldDrawStyle = pic.DrawStyle
        OldFillStyle = pic.FillStyle
        OldFillColor = pic.FillColor
        olddrawmode = pic.DrawMode
        OldForeColor = pic.ForeColor
        
        'New parameter
        pic.DrawWidth = 1
        pic.DrawStyle = vbSolid
        pic.FillStyle = vbFSSolid 'vbFSSolid
        pic.FillColor = vbRed
      '  Pic.ForeColor = vbRed
        pic.DrawMode = vbMaskPenNot
         
        If m_ObjLock = False Then
           'Top left
           DrawBoxGap pic, xmin - GAP, ymin - GAP, xmin, ymin
           'Botton left
           DrawBoxGap pic, xmin - GAP, ymax, xmin, ymax + GAP
           'Right top
           DrawBoxGap pic, xmax, ymin - GAP, xmax + GAP, ymin
           'Right botton
           DrawBoxGap pic, xmax, ymax, xmax + GAP, ymax + GAP
        Else
           'Top left
           DrawLock pic, xmin - GAP, ymin - GAP
           'Botton left
           DrawLock pic, xmin - GAP, ymax
           'Right top
           DrawLock pic, xmax, ymin - GAP
           'Right botton
           DrawLock pic, xmax, ymax
        End If
        
        If m_SelectedObjects.Count = 1 Then
          If m_ObjLock = False Then
            'center
            DrawX pic, xmin + (xmax - xmin) / 2, ymin + (ymax - ymin) / 2
            If m_TypeDraw <> dTextFrame Then
            'middle top
            DrawArrowH pic, xmin + (xmax - xmin) / 2, ymin - GAP
            'middle right
            DrawArrowV pic, xmax + GAP, ymin + (ymax - ymin) / 2
            'middle bottom
            DrawArrowH pic, xmin + (xmax - xmin) / 2, ymax + GAP
            'middle right
            DrawArrowV pic, xmin - GAP, ymin + (ymax - ymin) / 2
            End If
            'left for rotate
            DrawCircleGap pic, xmax, ymin + (ymax - ymin) / 2, 20
          End If
        End If
        'Default parameter
        pic.DrawWidth = OldDrawWirdth
        pic.DrawStyle = OldDrawStyle
        pic.FillStyle = OldFillStyle
        pic.FillColor = OldFillColor
        pic.DrawMode = olddrawmode
        pic.ForeColor = OldForeColor
'       'Return default settings
'        SetWorldTransform Pic.hdc, OldXForm
'        SetMapMode Pic.hdc, OldMapMode
'        SetGraphicsMode Pic.hdc, OldMode

    End If
End Sub

' Return the object's DrawWidth.
Public Property Get vbdObject_DrawWidth() As Integer
    vbdObject_DrawWidth = m_DrawWidth
End Property

' Set the object's DrawWidth.
Public Property Let vbdObject_DrawWidth(ByVal new_value As Integer)
    m_DrawWidth = new_value
End Property

' Return the object's DrawStyle.
Public Property Get vbdObject_DrawStyle() As Integer
    vbdObject_DrawStyle = m_DrawStyle
End Property
' Set the object's DrawStyle.
Public Property Let vbdObject_DrawStyle(ByVal new_value As Integer)
    m_DrawStyle = new_value
End Property

Private Property Let vbdObject_FillColor2(ByVal RHS As stdole.OLE_COLOR)
     m_FillColor2 = RHS
End Property

Private Property Get vbdObject_FillColor2() As stdole.OLE_COLOR
    vbdObject_FillColor2 = m_FillColor2
End Property

Private Property Let vbdObject_FillMode(ByVal RHS As FillMode)
     m_FillMode = RHS
End Property

Private Property Get vbdObject_FillMode() As FillMode
        vbdObject_FillMode = m_FillMode
End Property

' Return the object's ForeColor.
Public Property Get vbdObject_ForeColor() As OLE_COLOR
    vbdObject_ForeColor = m_ForeColor
End Property
' Set the object's ForeColor.
Public Property Let vbdObject_ForeColor(ByVal new_value As OLE_COLOR)
    m_ForeColor = new_value
End Property

' Return the object's FillColor.
Public Property Get vbdObject_FillColor() As OLE_COLOR
    vbdObject_FillColor = m_FillColor
End Property
' Set the object's FillColor.
Public Property Let vbdObject_FillColor(ByVal new_value As OLE_COLOR)
    m_FillColor = new_value
End Property

' Return the object's FillStyle.
Public Property Get vbdObject_FillStyle() As Integer
    vbdObject_FillStyle = m_FillStyle
End Property
' Set the object's FillStyle.
Public Property Let vbdObject_FillStyle(ByVal new_value As Integer)
    m_FillStyle = new_value
End Property

' Return this object's bounds.
Public Sub vbdObject_Bound(ByRef xmin As Single, ByRef ymin As Single, ByRef xmax As Single, ByRef ymax As Single)
Dim I As Integer

    If m_NumPoints < 1 Then
        xmin = 0
        xmax = 0
        ymin = 0
        ymax = 0
    Else
        With m_TransformedPoints(1)
            xmin = .X
            xmax = xmin
            ymin = .Y
            ymax = ymin
        End With

        For I = 2 To m_NumPoints
            With m_TransformedPoints(I)
                If xmin > .X Then xmin = .X
                If xmax < .X Then xmax = .X
                If ymin > .Y Then ymin = .Y
                If ymax < .Y Then ymax = .Y
            End With
        Next I
    End If
End Sub

' Draw the object on the canvas.
Public Sub vbdObject_Draw(ByVal pic As PictureBox)

Dim xmin As Single, x1 As Single
Dim xmax As Single, y1 As Single
Dim ymin As Single
Dim ymax As Single
Dim hRgn As Long
Dim oldmode As Long, OldMapMode As Long
Dim XForm As XForm
Dim OldXForm As XForm
Dim Br As LogBrush
Dim hBrush As Long, lPen As Long
Dim OldBrush As Long, hOldPen As Long, mDC As Long, tmDC As Long, mBitmap As Long
Dim imgData As Long, nW As Long, nH As Long, nL As Long, nT As Long, imgData1 As Long
Dim Flags As Long, R As RECT, OldWhd As POINTAPI, OldExt As POINTAPI

    On Error GoTo CheckPoints
    If m_NumPoints < 2 Then Exit Sub
    
    If UBound(m_TransformedPoints) = 0 Then
       m_TransformedPoints = m_OriginalPoints
    End If
    
    On Error GoTo 0
    ' Make sure we have at least 2 points.
    If m_NumPoints < 2 Then Exit Sub
    
    'SetCanvasDrawingParameters Me, Pic
    TransformPoints gZoomFactor
  
'       XForm.eM11 = m_M(1, 1)
'       XForm.eM12 = m_M(1, 2)
'       XForm.eM21 = m_M(2, 1)
'       XForm.eM22 = m_M(2, 2)
'       XForm.eDx = m_M(3, 1)
'       XForm.eDy = m_M(3, 2)
'       OldMode = SetGraphicsMode(Pic.hdc, GM_ADVANCED)
'       OldMapMode = SetMapMode(Pic.hdc, MM_TEXT)
'       GetWorldTransform Pic.hdc, OldXForm
    
    If m_ViewSimple Then
        Br.lbColor = Me.m_FillColor
        Br.lbHatch = -1
        Br.lbStyle = 2
        hBrush = CreateBrushIndirect(Br)
        lPen = PenCreate(Me.m_DrawStyle, 1, Me.m_ForeColor)
        StrokePath pic.hDC
    Else
        Select Case Me.m_FillStyle
        Case 0  'Solid
            hBrush = CreateSolidBrush(Me.m_FillColor)
        Case 8  'Bitmap pattern
            imgData = LoadPatternPic(App.Path + "\Pattern\" + Me.m_Pattern)
            hBrush = CreatePatternBrush(imgData)
        Case 10  'Bitmap
            'GetRgnBox Me.m_hRegion, R
           
            'nL = R.Left
            'nT = R.Top
            'nW = R.Right - R.Left
            'nH = R.Bottom - R.Top
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax - xmin
            nH = ymax - ymin
            nL = xmin
            nT = ymin
            mDC = CreateCompatibleDC(0)
            imgData = LoadImage(mDC, Me.m_Pattern, 0, nW, nH, &H10)
            SelectObject mDC, imgData
            If nL < 0 Then nL = 0
            If nT < 0 Then nT = 0
            imgData = BitmapFromDC(mDC, nL, nT, xmax, ymax) ', Me.m_angle, R.Right, R.Bottom)
            
            SelectObject mDC, imgData
            hBrush = CreatePatternBrush(imgData)
           
        Case 9  'Gradient
            vbdObject_Bound xmin, ymin, xmax, ymax
            nW = xmax
            nH = ymax
            'Create a device context, compatible with the screen
            mDC = CreateCompatibleDC(GetDC(0))
            'Create a bitmap, compatible with the screen
            mBitmap = CreateCompatibleBitmap(GetDC(0), nW, nH)
            SelectObject mDC, mBitmap
            GradientFillRectDC mDC, xmin - 5, ymin - 5, xmax, ymax, Me.m_FillColor, Me.m_FillColor2, Me.m_Gradient
            hBrush = CreatePatternBrush(mBitmap)
            
        Case Else   'Other Patterns
            Br.lbColor = Me.m_FillColor
            Br.lbHatch = Me.m_FillStyle - 2
            Br.lbStyle = 2
            hBrush = CreateBrushIndirect(Br)
     End Select
         lPen = PenCreate(Me.m_DrawStyle, Me.m_DrawWidth, Me.m_ForeColor)
    End If
    
    StrokePath pic.hDC
    OldBrush = SelectObject(pic.hDC, hBrush)
    hOldPen = SelectObject(pic.hDC, lPen)
       
'    If m_TypeDraw = dTextFrame Then
'        SetWorldTransform Pic.hdc, XForm
'    End If
       
    'Make Path for Region
     BeginPath pic.hDC
     DrawObject pic ', True
     EndPath pic.hDC
     
     'Delete  old Region
     DeleteObject Me.m_hRegion
     'New Region
     Me.m_hRegion = PathToRegion(pic.hDC)

     'Draw
     If m_TypeDraw = dText Or m_TypeDraw = dPolydraw Or m_TypeDraw = dCurve Or m_TypeDraw = dEllipse Then   'Or m_TypeDraw = dText Then  'Or m_TypeDraw = dTextFrame Then
        DrawObject pic, True      'for fill object with PolyDraw
        DrawObject pic             'for fill object with PolyDraw
     ElseIf m_TypeDraw = dTextFrame Then
        DrawObject pic, True      'for fill object with PolyDraw
    Else
        'DrawObject Pic, True      'for fill object with PolyDraw
        DrawObject pic, True
     End If
     
'    If m_TypeDraw = dTextFrame Then
'       Call ModifyWorldTransform(Pic.hdc, XForm, MWT_IDENTITY)
'    End If
       
       If m_ViewSimple = False Then
         ' If m_TypeDraw = dText Or m_TypeDraw = dTextFrame Then
         '     FillPath Pic.hdc
         ' Else
              StrokeAndFillPath pic.hDC
         ' End If
       End If
       
'    If m_TypeDraw = dTextFrame Then
'       'Return default settings
'       SetWorldTransform Pic.hdc, OldXForm
'       SetMapMode Pic.hdc, OldMapMode
'       SetGraphicsMode Pic.hdc, OldMode
'    End If
    'De-select Brush-pen from DC
    Call SelectObject(pic.hDC, OldBrush)
    Call SelectObject(pic.hDC, hOldPen)
    
    DeleteObject hBrush
    DeleteObject lPen
    DeleteObject imgData
    DeleteObject mBitmap
    DeleteDC mDC
    Exit Sub
    
CheckPoints:
   If Err = 9 Then
      Resume Next
   End If
End Sub

'Draw
Private Sub DrawObject(ByVal pic As PictureBox, Optional WithPath As Boolean = False)
   Dim xmin As Single, x1 As Single
   Dim xmax As Single, y1 As Single
   Dim ymin As Single, I As Long
   Dim ymax As Single, lpXform As XForm
   Dim mx1 As Long, mx2 As Long, my1 As Long, my2 As Long
   Dim BackBuffer As clsBackBuffer, mBuffer As Boolean
   Dim mBlend As Byte
   Dim PointCoords() As POINTAPI
   Dim PointTypes() As Byte, NumPoints As Long
   
   
    SetPolyFillMode pic.hDC, m_FillMode
    
    vbdObject_Bound xmin, ymin, xmax, ymax
    mx1 = xmin
    mx2 = xmax
    my1 = ymin
    my2 = ymax
    
    If WithPath = True And Me.m_Blend <= 255 And m_TypeDraw <> dTextFrame Then
        ' Create back-buffer
        Set BackBuffer = New clsBackBuffer
        ' Create a snapshot of the current surface big enough for the drawing
        mBuffer = BackBuffer.CreateSnapshop(pic.hDC, mx1, my1, mx2, my2)
    End If
    
    'Text
    If m_TypeDraw = dText Or m_TypeDraw = dTextFrame Then
       pic.Font.Bold = m_Bold
        pic.Font.Charset = m_Charset
         pic.Font.Italic = m_Italic
          pic.Font.Name = m_Name
           pic.Font.Size = m_Size * gZoomFactor
            pic.Font.Strikethrough = m_Strikethrough
             pic.Font.Underline = m_Underline
             pic.Font.Weight = m_Weight
          
          If m_TypeDraw = dText Then
             vbdObject_Bound xmin, ymin, xmax, ymax
              pic.ForeColor = Me.m_ForeColor 'Me.m_FillColor
              pic.FillColor = Me.m_FillColor 'Me.m_ForeColor
              If WithPath Then BeginPath pic.hDC

                  CenterText pic, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), _
                             m_TextDraw, m_Size * gZoomFactor, , -m_angle * 10, m_Weight, _
                             m_Italic, m_Underline, m_Strikethrough, m_Charset, , , , , m_Name
              If WithPath Then Call EndPath(pic.hDC)
''              Call PolyDraw(Pic.hDC, PointCoords(0), PointTypes(0), NumPoints)
'               If WithPath Then BeginPath Pic.hDC
'               PolyDraw Pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'               If WithPath Then EndPath Pic.hDC
'               If WithPath Then PolyDraw Pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'               If WithPath Then FillPath Pic.hDC
              
          ElseIf m_TypeDraw = dTextFrame Then
             'If WithPath Then BeginPath Pic.hDC
                
               lpXform.eM11 = m_M(1, 1)
               lpXform.eM12 = m_M(1, 2)
               lpXform.eM21 = m_M(2, 1)
               lpXform.eM22 = m_M(2, 2)
               lpXform.eDx = m_M(3, 1)
               lpXform.eDy = m_M(3, 2)
               m_TransformedPoints = m_OriginalPoints
               xmin = m_TransformedPoints(1).X
               ymin = m_TransformedPoints(1).Y
               xmax = IIf(m_OriginalPoints(3).X > m_TransformedPoints(3).X, m_OriginalPoints(3).X, m_TransformedPoints(3).X)
               ymax = IIf(m_OriginalPoints(3).Y > m_TransformedPoints(3).Y, m_OriginalPoints(3).Y, m_TransformedPoints(3).Y)
               DrawFontFrame pic.hDC, m_TextDraw, m_angle, Me.m_ForeColor, m_AlingText, xmin, ymin, xmax, ymax, lpXform
             'If WithPath Then EndPath Pic.hDC
'               NumPoints = GetPathAPI(Pic.hDC, ByVal 0&, ByVal 0&, 0)
'            If NumPoints > 0 Then
'                ReDim PointCoords(NumPoints - 1) As POINTAPI
'                ReDim PointTypes(NumPoints - 1) As Byte
'                ' Get the path data from the DC
'                Call GetPathAPI(Pic.hDC, PointCoords(0), PointTypes(0), NumPoints) '<> 0
'
'            Call BeginPath(Pic.hDC)
'                Call PolyDraw(Pic.hDC, PointCoords(0), PointTypes(0), NumPoints)
'            Call EndPath(Pic.hDC)
'            Call PolyDraw(Pic.hDC, PointCoords(0), PointTypes(0), NumPoints)
'            End If
'             FillPath Pic.hDC
          End If
          
          'FillPath Pic.hDC
    Else
       'Polydraw (Curve object)
      ' If m_TypeDraw = dEllipse Then
      '     If WithPath Then BeginPath Pic.hDC
      '     PolyBezier Pic.hDC, m_TransformedPoints(1), m_NumPoints
      '     If WithPath Then EndPath Pic.hDC
      ' End If
      If m_TypeDraw = dPolydraw Or m_TypeDraw = dEllipse Or m_TypeDraw = dCurve Then
        If WithPath Then BeginPath pic.hDC
             'Draw fill
             PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
'             For i = 1 To m_NumPoints - 3 Step 3
'                Pic.Line (m_TransformedPoints(i).X - 4, m_TransformedPoints(i).Y - 4)-(m_TransformedPoints(i).X + 4, m_TransformedPoints(i).Y + 4), vbRed, BF
'                If i + 1 > m_NumPoints Then Exit For
'                Pic.Line (m_TransformedPoints(i + 1).X - 4, m_TransformedPoints(i + 1).Y - 4)-(m_TransformedPoints(i + 1).X + 4, m_TransformedPoints(i + 1).Y + 4), vbGreen, BF
'                Pic.Line (m_TransformedPoints(i).X, m_TransformedPoints(i).Y)-(m_TransformedPoints(i + 1).X, m_TransformedPoints(i + 1).Y), vbMagenta
'                If i + 2 > m_NumPoints Then Exit For
'                Pic.Line (m_TransformedPoints(i + 2).X - 4, m_TransformedPoints(i + 2).Y - 4)-(m_TransformedPoints(i + 2).X + 4, m_TransformedPoints(i + 2).Y + 4), vbBlue, BF
'                If i + 3 > m_NumPoints Then
'                    Pic.Line (m_TransformedPoints(i + 2).X, m_TransformedPoints(i + 2).Y)-(m_TransformedPoints(1).X, m_TransformedPoints(1).Y), vbMagenta
'                Else
'                    Pic.Line (m_TransformedPoints(i + 2).X, m_TransformedPoints(i + 2).Y)-(m_TransformedPoints(i + 3).X, m_TransformedPoints(i + 3).Y), vbMagenta
'                End If
'                Debug.Print i, m_TransformedPoints(i).X, m_TransformedPoints(i).Y
'             Next
        If WithPath Then EndPath pic.hDC
        If WithPath Then PolyDraw pic.hDC, m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then FillPath pic.hDC
      Else
        'Polyline - Polygon
         If IsClosed Then
            Polygon pic.hDC, m_TransformedPoints(1), m_NumPoints
         Else
            Polyline pic.hDC, m_TransformedPoints(1), m_NumPoints
         End If
      End If
      
    End If
    
    If WithPath = True And mBuffer And Me.m_Blend < 255 And m_TypeDraw <> dTextFrame Then
       ' Blend back-buffer to foreground buffer
        mBlend = Me.m_Blend
        Call BackBuffer.Draw(pic.hDC, mx1, my1, mBlend)
        Debug.Print "mBlend:" + Str(mBlend)
       ' Destroy back-buffer
        Set BackBuffer = Nothing
    End If
    
End Sub

Sub DrawObjectWMF(ByVal picwm As Long, Optional WithPath As Boolean = False)
  Dim xmin As Single, x1 As Single
   Dim xmax As Single, y1 As Single
   Dim ymin As Single, I As Long
   Dim ymax As Single, lpXform As XForm
   Dim mx1 As Long, mx2 As Long, my1 As Long, my2 As Long
   Dim BackBuffer As clsBackBuffer, mBuffer As Boolean
   Dim mBlend As Byte
   Dim PointCoords() As POINTAPI
   Dim PointTypes() As Byte, NumPoints As Long
   
   
    SetPolyFillMode picwm, m_FillMode
    
    vbdObject_Bound xmin, ymin, xmax, ymax
    mx1 = xmin
    mx2 = xmax
    my1 = ymin
    my2 = ymax
    
    If WithPath = True And Me.m_Blend <= 255 And m_TypeDraw <> dTextFrame Then
        ' Create back-buffer
        Set BackBuffer = New clsBackBuffer
        ' Create a snapshot of the current surface big enough for the drawing
        mBuffer = BackBuffer.CreateSnapshop(picwm, mx1, my1, mx2, my2)
    End If
    
    'Text
    If m_TypeDraw = dText Or m_TypeDraw = dTextFrame Then
'       pic.Font.Bold = m_Bold
'        pic.Font.Charset = m_Charset
'         pic.Font.Italic = m_Italic
'          pic.Font.Name = m_Name
'           pic.Font.Size = m_Size * gZoomFactor
'            pic.Font.Strikethrough = m_Strikethrough
'             pic.Font.Underline = m_Underline
'             pic.Font.Weight = m_Weight
          
          If m_TypeDraw = dText Then
'             vbdObject_Bound xmin, ymin, xmax, ymax
'              pic.ForeColor = Me.m_ForeColor 'Me.m_FillColor
'              pic.FillColor = Me.m_FillColor 'Me.m_ForeColor
'              If WithPath Then BeginPath picwm
'
'                  CenterText pic, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), _
'                             m_TextDraw, m_Size * gZoomFactor, , -m_angle * 10, m_Weight, _
'                             m_Italic, m_Underline, m_Strikethrough, m_Charset, , , , , m_Name
'              If WithPath Then Call EndPath(picwm)
              
'          ElseIf m_TypeDraw = dTextFrame Then
'               lpXform.eM11 = m_M(1, 1)
'               lpXform.eM12 = m_M(1, 2)
'               lpXform.eM21 = m_M(2, 1)
'               lpXform.eM22 = m_M(2, 2)
'               lpXform.eDx = m_M(3, 1)
'               lpXform.eDy = m_M(3, 2)
'               m_TransformedPoints = m_OriginalPoints
'               xmin = m_TransformedPoints(1).X
'               ymin = m_TransformedPoints(1).Y
'               xmax = IIf(m_OriginalPoints(3).X > m_TransformedPoints(3).X, m_OriginalPoints(3).X, m_TransformedPoints(3).X)
'               ymax = IIf(m_OriginalPoints(3).Y > m_TransformedPoints(3).Y, m_OriginalPoints(3).Y, m_TransformedPoints(3).Y)
'               DrawFontFrame picwm, m_TextDraw, m_angle, Me.m_ForeColor, m_AlingText, xmin, ymin, xmax, ymax, lpXform
          End If
    
    Else
    
      If m_TypeDraw = dPolydraw Or m_TypeDraw = dEllipse Or m_TypeDraw = dCurve Then
        If WithPath Then BeginPath picwm
             'Draw fill
             PolyDrawVB picwm ', m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then EndPath picwm
        If WithPath Then PolyDrawVB picwm ', m_TransformedPoints(1), m_TypePoint(1), m_NumPoints
        If WithPath Then FillPath picwm
      Else
        'Polyline - Polygon
         If IsClosed Then
            Polygon picwm, m_TransformedPoints(1), m_NumPoints
         Else
            Polyline picwm, m_TransformedPoints(1), m_NumPoints
         End If
      End If
      
    End If
    
    If WithPath = True And mBuffer And Me.m_Blend < 255 And m_TypeDraw <> dTextFrame Then
       ' Blend back-buffer to foreground buffer
        mBlend = Me.m_Blend
        Call BackBuffer.Draw(picwm, mx1, my1, mBlend)
        Debug.Print "mBlend:" + Str(mBlend)
       ' Destroy back-buffer
        Set BackBuffer = Nothing
    End If
    
End Sub
' Apply the transformation matrix.
Public Sub TransformPoints(pZoomFactor As Single)
    Dim I As Integer
    If m_NumPoints = 0 Then Exit Sub
    ReDim m_TransformedPoints(1 To m_NumPoints)
    For I = 1 To m_NumPoints
        With m_OriginalPoints(I)
            m_TransformedPoints(I).X = (.X * m_M(1, 1) + .Y * m_M(2, 1) + m_M(3, 1)) * pZoomFactor
            m_TransformedPoints(I).Y = (.X * m_M(1, 2) + .Y * m_M(2, 2) + m_M(3, 2)) * pZoomFactor
        End With
    Next I
End Sub

Private Property Let vbdObject_Gradient(ByVal RHS As Integer)
      m_Gradient = RHS
End Property

Private Property Get vbdObject_Gradient() As Integer
         vbdObject_Gradient = m_Gradient
End Property

Private Property Let vbdObject_hRegion(ByVal RHS As Long)
          m_hRegion = RHS
End Property

Private Property Get vbdObject_hRegion() As Long
         vbdObject_hRegion = m_hRegion
End Property

Private Function vbdObject_Info() As String
 
     vbdObject_Info = "Max Points:" + Str(Me.NumPoints)
         
End Function

Private Property Get vbdObject_Italic() As Boolean
         vbdObject_Italic = m_Italic
End Property

Private Property Let vbdObject_Italic(ByVal RHS As Boolean)
           m_Italic = RHS
End Property

Public Sub vbdObject_MakeTransformation()
       TransformPoints gZoomFactor
End Sub

Private Property Get vbdObject_Name() As String
         vbdObject_Name = m_Name
End Property

Private Property Let vbdObject_Name(ByVal RHS As String)
        m_Name = RHS
End Property

Private Sub vbdObject_NewPoint(NumPoints As Long, X() As Single, Y() As Single, TypePoint() As Byte)

End Sub

Private Sub vbdObject_NewTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)

End Sub

Private Property Let vbdObject_ObjLock(ByVal RHS As Boolean)
         m_ObjLock = RHS
End Property

Private Property Get vbdObject_ObjLock() As Boolean
         vbdObject_ObjLock = m_ObjLock
End Property

Private Property Let vbdObject_Blend(ByVal RHS As Integer)
          m_Blend = RHS
End Property

Private Property Get vbdObject_Blend() As Integer
         vbdObject_Blend = m_Blend
End Property

Private Property Let vbdObject_Pattern(ByVal RHS As String)
       m_Pattern = RHS
End Property

Private Property Get vbdObject_Pattern() As String
       vbdObject_Pattern = m_Pattern
End Property

Private Property Get vbdObject_Picture() As stdole.Picture
       Set vbdObject_Picture = m_Picture
End Property

Private Property Set vbdObject_Picture(ByVal RHS As stdole.Picture)
       Set m_Picture = RHS
End Property

Private Sub vbdObject_ReadPoint(ByRef NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       Dim I As Long
       NumPoints = m_NumPoints
       ReDim X(1 To NumPoints)
       ReDim Y(1 To NumPoints)
       ReDim TPoint(1 To NumPoints)
       TPoint = m_TypePoint
       For I = 1 To m_NumPoints
           X(I) = m_OriginalPoints(I).X
           Y(I) = m_OriginalPoints(I).Y
       Next
       'm_Object.ReadPoint NumPoints, X(), Y(), TPoint()
End Sub

Private Sub vbdObject_ReadTrPoint(NumPoints As Long, X() As Single, Y() As Single, TPoint() As Byte)
       Dim I As Long
       NumPoints = m_NumPoints
       ReDim X(1 To NumPoints)
       ReDim Y(1 To NumPoints)
       ReDim TPoint(1 To NumPoints)
       TPoint = m_TypePoint
       TransformPoints gZoomFactor
       For I = 1 To m_NumPoints
           X(I) = m_TransformedPoints(I).X
           Y(I) = m_TransformedPoints(I).Y
       Next
End Sub

' Set the object's Selected status.
Private Property Let vbdObject_Selected(ByVal RHS As Boolean)
    m_Selected = RHS
End Property
' Return the object's Selected status.
Private Property Get vbdObject_Selected() As Boolean
    vbdObject_Selected = m_Selected
End Property

' Return True if the object is at this location.
Private Function vbdObject_IsAt(ByVal X As Single, ByVal Y As Single) As Boolean
Dim is_at As Boolean
Dim m_Point() As POINTAPI
    
    'm_TransformedPoints = m_OriginalPoints
    On Error GoTo CheckPoints
    If UBound(m_TransformedPoints) = 0 Then m_TransformedPoints = m_OriginalPoints
    
    On Error GoTo vbErr
    m_Point = m_TransformedPoints

    is_at = PolygonIsAt(IsClosed, X, Y, m_Point)
    
    If (Not is_at) And IsClosed And (m_FillStyle <> vbFSTransparent) Then
        is_at = PointIsInPolygon(X, Y, m_Point)
    End If

    vbdObject_IsAt = is_at
    
    On Error GoTo 0
    Exit Function
    
CheckPoints:
    If Err = 9 Then
       'm_TransformedPoints = m_OriginalPoints
       Resume Next
    End If
vbErr:
    vbdObject_IsAt = False
    On Error GoTo 0
End Function

' Initialize the object using a serialization string.
' The serialization does not include the
' ObjectType(...) part.
Private Property Let vbdObject_Serialization(ByVal RHS As String)
Dim token_name As String
Dim token_value As String
Dim next_x As Integer
Dim next_y As Integer
Dim nextpoint As Integer

    InitializeDrawingProperties Me
    m2Identity m_M

    ' Read tokens until there are no more.
    Do While Len(RHS) > 0
        ' Read a token.
        GetNamedToken RHS, token_name, token_value
        Select Case token_name
            Case "IsClosed"
                IsClosed = CBool(token_value)
            Case "NumPoints"
                ' This allocates the m_X and m_Y arrays.
                NumPoints = CLng(token_value)
                next_x = 1
                next_y = 1
                nextpoint = 1
            Case "X"
                X(next_x) = CSng(token_value)
                next_x = next_x + 1
            Case "Y"
                Y(next_y) = CSng(token_value)
                next_y = next_y + 1
            Case "P"
                TypePoint(nextpoint) = CSng(token_value)
                nextpoint = nextpoint + 1
            Case "Transformation"
                SetTransformationSerialization token_value, m_M
            Case Else
                ReadDrawingPropertySerialization Me, token_name, token_value
        End Select
    Loop
End Property

' Return a serialization string for the object.
Public Property Get vbdObject_Serialization() As String
Dim txt As String
Dim I As Integer

    txt = DrawingPropertySerialization(Me)
    txt = txt & TransformationSerialization(m_M)
    txt = txt & " IsClosed(" & Format$(IsClosed) & ")"
    txt = txt & " NumPoints(" & Format$(NumPoints) & ")"
    For I = 1 To NumPoints
        With m_OriginalPoints(I)
            txt = txt & vbCrLf & "    X(" & Format$(.X) & ")"
            txt = txt & " Y(" & Format$(.Y) & ")"
            txt = txt & " P(" & Format$(m_TypePoint(I)) & ")"
        End With
    Next I
    'Debug.Print vbdObject_TypeDraw
    Select Case Me.m_TypeDraw
    Case dPolygon
        vbdObject_Serialization = "Polygon(" & txt & ")"
    Case dPolyline
        vbdObject_Serialization = "Polyline(" & txt & ")"
    Case dScribble
        vbdObject_Serialization = "Scribble(" & txt & ")"
    Case dFreePolygon
        vbdObject_Serialization = "FreePolygon(" & txt & ")"
    Case dRectAngle
        vbdObject_Serialization = "RectAngle(" & txt & ")"
    Case dEllipse
        vbdObject_Serialization = "Ellipse(" & txt & ")"
    Case dText
        vbdObject_Serialization = "Text(" & txt & ")"
    Case dTextFrame
        vbdObject_Serialization = "TextFrame(" & txt & ")"
    Case dPolydraw
        vbdObject_Serialization = "PolyDraw(PolyDraw(" & txt & "))"
    Case dCurve
        vbdObject_Serialization = "Curve(" & txt & ")"
    End Select
    'Stop
    'vbdObject_Serialization = "Polygon(" & Txt & ")"
    Debug.Print vbdObject_Serialization
End Property

Public Property Let vbdObject_Shade(ByVal RHS As Boolean)
       m_Shade = RHS
End Property

Public Property Get vbdObject_Shade() As Boolean
      vbdObject_Shade = m_Shade
End Property

Private Property Let vbdObject_Size(ByVal RHS As Integer)
        m_Size = RHS
End Property

Private Property Get vbdObject_Size() As Integer
          vbdObject_Size = m_Size
End Property

Private Sub vbdObject_StartBound(xmin As Single, ymin As Single)
Dim I As Integer

    If m_NumPoints < 1 Then
        xmin = 0
        ymin = 0
    Else
        With m_OriginalPoints(1)
            xmin = .X
            ymin = .Y
        End With
    End If
End Sub

Private Property Let vbdObject_Strikethrough(ByVal RHS As Boolean)
        m_Strikethrough = RHS
End Property

Private Property Get vbdObject_Strikethrough() As Boolean
         vbdObject_Strikethrough = m_Strikethrough
End Property

Private Property Get vbdObject_TextDraw() As String
        vbdObject_TextDraw = m_TextDraw
End Property

Public Property Let vbdObject_TextDraw(ByVal RHS As String)
     m_TextDraw = RHS
End Property

' Set the object's TypeDraw.
Public Property Let vbdObject_TypeDraw(ByVal RHS As DrawType)
     m_TypeDraw = RHS
End Property

' Return the object's TypeDraw.
Public Property Get vbdObject_TypeDraw() As DrawType
       vbdObject_TypeDraw = m_TypeDraw
End Property

Public Property Let vbdObject_TypeFill(ByVal RHS As DrawTypeFill)
       m_TypeFill = RHS
End Property

Public Property Get vbdObject_TypeFill() As DrawTypeFill
       vbdObject_TypeFill = m_TypeFill
End Property

Private Sub DrawBoxGap(pic As Object, cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single)
     Dim Ls As Single, StepGap As Integer
     StepGap = (GAP / 2) - 1
     If StepGap > 3 Then StepGap = 3
     pic.Line (cX1 + Ls, cY1 + Ls)-(cX2 - Ls, cY2 - Ls), vbRed, BF
End Sub

Private Sub DrawArrowH(pic As Object, cX As Single, cY As Single)
     Dim N_Point(1 To 3) As POINTAPI
     Dim cX1 As Long, cY1 As Long
     cX1 = cX
     cY1 = cY
     
     pic.Line (cX1 - 4, cY1)-(cX1 + 4, cY1), vbRed
     
     N_Point(1).X = cX1 - 8
     N_Point(1).Y = cY1
     N_Point(2).X = cX1 - 4
     N_Point(2).Y = cY1 - 4
     N_Point(3).X = cX1 - 4
     N_Point(3).Y = cY1 + 4
     Polygon pic.hDC, N_Point(1), 3

     N_Point(1).X = cX1 + 8
     N_Point(1).Y = cY1
     N_Point(2).X = cX1 + 4
     N_Point(2).Y = cY1 - 4
     N_Point(3).X = cX1 + 4
     N_Point(3).Y = cY1 + 4
     Polygon pic.hDC, N_Point(1), 3
     
End Sub

Private Sub DrawLock(pic As Object, cX As Single, cY As Single)
     Dim N_Point() As POINTAPI
     Dim NumPoints(0 To 1) As Long
     Dim cX1 As Long, cY1 As Long
     cX1 = cX
     cY1 = cY
     
     ReDim N_Point(1 To 6)
     N_Point(1).X = cX1 + 1
     N_Point(1).Y = cY1 + 3
     N_Point(2).X = cX1 + 1
     N_Point(2).Y = cY1 + 2
     N_Point(3).X = cX1 + 2
     N_Point(3).Y = cY1
     N_Point(4).X = cX1 + 4
     N_Point(4).Y = cY1
     N_Point(5).X = cX1 + 6
     N_Point(5).Y = cY1 + 3
     N_Point(6).X = cX1 + 6
     N_Point(6).Y = cY1 + 4
     Polyline pic.hDC, N_Point(1), 6
     
     ReDim N_Point(1 To 8)
     N_Point(1).X = cX1
     N_Point(1).Y = cY1 + 3
     N_Point(2).X = cX1 + 6
     N_Point(2).Y = cY1 + 3
     N_Point(3).X = cX1 + 6
     N_Point(3).Y = cY1 + 6
     N_Point(4).X = cX1
     N_Point(4).Y = cY1 + 6
     NumPoints(0) = 4
     
     N_Point(5).X = cX1 + 2
     N_Point(5).Y = cY1 + 3
     N_Point(6).X = cX1 + 4
     N_Point(6).Y = cY1 + 3
     N_Point(7).X = cX1 + 4
     N_Point(7).Y = cY1 + 6
     N_Point(8).X = cX1 + 2
     N_Point(8).Y = cY1 + 6
     NumPoints(1) = 4
     
     PolyPolygon pic.hDC, N_Point(1), NumPoints(0), 2
 
End Sub
Private Sub DrawArrowV(pic As Object, cX As Single, cY As Single)
     Dim N_Point(0 To 4) As POINTAPI
     Dim cX1 As Long, cY1 As Long
     cX1 = cX
     cY1 = cY
     pic.Line (cX1, cY1 - 4)-(cX1, cY1 + 4), vbRed
      
     N_Point(0).Y = cY1 - 8
     N_Point(0).X = cX1
     N_Point(1).Y = cY1 - 4
     N_Point(1).X = cX1 - 4
     N_Point(2).Y = cY1 - 4
     N_Point(2).X = cX1 + 4
     N_Point(3).Y = cY1 - 8
     N_Point(3).X = cX1
     Polygon pic.hDC, N_Point(0), 4

     N_Point(0).Y = cY1 + 8
     N_Point(0).X = cX1
     N_Point(1).Y = cY1 + 4
     N_Point(1).X = cX1 - 4
     N_Point(2).Y = cY1 + 4
     N_Point(2).X = cX1 + 4
     N_Point(3).Y = cY1 + 8
     N_Point(3).X = cX1
     Polygon pic.hDC, N_Point(0), 4

End Sub

'Private Sub DrawArrowC(Pic As Object, cX1 As Single, cY1 As Single, cX2 As Single, cY2 As Single)
'     Dim Ls As Single, StepGap As Integer
'     StepGap = (GAP / 2) - 1
'     If StepGap > 3 Then StepGap = 3
'      Pic.Line (cX1 + Ls, cY1 + Ls)-(cX2 - Ls, cY2 - Ls), RGB(Ls * 60 + 43, Ls * 60 + 43, Ls * 60 + 43), BF
'End Sub

Private Sub DrawX(pic As Object, cX As Single, cY As Single)
      Dim OldWidth As Integer

      OldWidth = pic.DrawWidth
      pic.DrawWidth = 2
  
      pic.Line (cX - 3, cY - 3)-(cX + 3, cY + 3), vbRed
      pic.Line (cX - 3, cY + 3)-(cX + 3, cY - 3), vbRed
      
      pic.DrawWidth = OldWidth
End Sub

Private Sub DrawCenter(pic As Object, cX As Single, cY As Single)
      Dim OldWidth As Integer
      OldWidth = pic.DrawWidth
      pic.PSet (cX, cY)
      pic.DrawWidth = 1
      pic.FillColor = vbRed
      pic.Circle (cX, cY), 6
      pic.DrawWidth = OldWidth
End Sub

Private Sub DrawCircleGap(pic As Object, cX As Single, cY As Single, Optional cLenght As Integer = 20)
     Dim Ls As Single, StepGap As Integer
     
     StepGap = (GAP / 2)
     If StepGap > 3 Then StepGap = 3
     pic.FillColor = vbRed
     pic.Line (cX + GAP, cY)-(cX + cLenght - StepGap, cY), vbRed
     pic.FillColor = vbRed
     pic.Circle (cX + cLenght, cY), StepGap, 0

End Sub

Private Property Let vbdObject_Underline(ByVal RHS As Boolean)
     m_Underline = RHS
End Property

Private Property Get vbdObject_Underline() As Boolean
          vbdObject_Underline = m_Underline
End Property

Private Property Let vbdObject_Weight(ByVal RHS As Integer)
        m_Weight = RHS
End Property

Private Property Get vbdObject_Weight() As Integer
        vbdObject_Weight = m_Weight
End Property

Private Sub PolyDrawVB(ByVal wmhdc As Long) ', ByRef lpPt() As POINTAPI, _
                       ByRef lpbTypes() As Byte, ByVal cCount As Long)
                       
    Dim LoopPts As Long
    Dim BezIdx As Long
    Dim lpPt() As POINTAPI, lpbTypes() As Byte, cCount As Long
    lpPt = m_TransformedPoints '(1)
    lpbTypes = m_TypePoint '(1)
    cCount = m_NumPoints
    Const PT_CLOSEFIGURE As Long = &H1
    Const PT_LINETO As Long = &H2
    Const PT_BEZIERTO As Long = &H4
    Const PT_MOVETO As Long = &H6
    
    For LoopPts = 1 To cCount '- 1
        ' Clear bezier point index for non-bezier point
        If ((lpbTypes(LoopPts) And PT_BEZIERTO) = 0) Then BezIdx = 0

        Select Case lpbTypes(LoopPts) And Not PT_CLOSEFIGURE
            Case PT_LINETO    ' Straight line segment
                Call LineTo(wmhdc, lpPt(LoopPts).X, lpPt(LoopPts).Y)
        Case PT_BEZIERTO    ' Curve segment
                '//
                Select Case BezIdx
                Case 0, 1   ' Bezier control handles
                   '  Call DrawBox(lpPt(LoopPts), 2, vbBlue)
                   ' Me.Caption = LoopPts + 1 & "/" & cCount & " Bezier control handles :" & BezIdx
                Case 2    ' Bezier end point
                   '//Connecting lines betweenn (start to 1st control pt) and (2nd control pt to end point)
                   ' Move to first point where we will start curve
                   Call MoveToEx(wmhdc, lpPt(LoopPts - 3).X, lpPt(LoopPts - 3).Y, ByVal 0&)
                   Call PolyBezierTo(wmhdc, lpPt(LoopPts - 2), 3)
                End Select
                   BezIdx = (BezIdx + 1) Mod 3 '//Reset counter after 3 Bezier points
            Case PT_MOVETO    ' Move current drawing point
                 Call MoveToEx(wmhdc, lpPt(LoopPts).X, lpPt(LoopPts).Y, ByVal 0&)
        End Select

        If (lpbTypes(LoopPts) And PT_CLOSEFIGURE) Then
            Call CloseFigure(wmhdc)
        End If
    Next LoopPts
    
End Sub



