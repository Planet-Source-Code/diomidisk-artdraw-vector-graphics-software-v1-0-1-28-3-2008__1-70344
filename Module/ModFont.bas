Attribute VB_Name = "ModFonts"
Option Explicit


'Font enumeration types
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte 'OR STRING *33

        lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte

        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1

Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
'Private Const TRUETYPE_FONTTYPE = &H4

Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" _
                            (ByVal hDC As Long, ByVal lpszFamily As String, _
                            ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFont Lib "GDI32.dll" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetPathAPI Lib "GDI32.dll" Alias "GetPath" (ByVal hDC As Long, ByRef lpPoints As Any, ByRef lpTypes As Any, ByVal nSize As Long) As Long
Private Declare Function GetPath Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32" (ByVal hDC As Long, lpPt As POINTAPI, ByVal cCount As Long) As Long
Private Declare Function PolyDraw Lib "gdi32" (ByVal hDC As Long, lpPt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long

Private Declare Function FillPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Dim m_PointCoords() As POINTAPI
Dim m_PointTypes() As Byte
Dim m_NumPoints As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const LOGPIXELSY = 90                    'For GetDeviceCaps - returns the height of a logical pixel
'Private Const ANSI_CHARSET = 0                   'Use the default Character set
'Private Const CLIP_LH_ANGLES = 16                ' Needed for tilted fonts.
'Private Const OUT_TT_PRECIS = 9                  'Tell it to use True Types when Possible
'Private Const PROOF_QUALITY = 9                  'Make it as clean as possible.
'Private Const DEFAULT_PITCH = 0                  'We want the font to take whatever pitch it defaults to
'Private Const FF_DONTCARE = 0                    'Use whatever fontface it is.

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

'drawtext
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const DC_GRADIENT = &H20

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, ByVal lpRect As Any, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Public Const ETO_OPAQUE = 2
' Font weight constants.
Private Const FW_DONTCARE = 0
Enum FontWeight
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
    FW_ULTRALIGHT = FW_EXTRALIGHT
    FW_REGULAR = FW_NORMAL
    FW_DEMIBOLD = FW_SEMIBOLD
    FW_ULTRABOLD = FW_EXTRABOLD
End Enum
Private Const FW_BLACK = FW_HEAVY

' Character set constants.
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const OEM_CHARSET = 255

' Output precision constants.
Private Const OUT_CHARACTER_PRECIS = 2
Private Const OUT_DEFAULT_PRECIS = 0
Private Const OUT_DEVICE_PRECIS = 5
Private Const OUT_RASTER_PRECIS = 6
Private Const OUT_STRING_PRECIS = 1
Private Const OUT_STROKE_PRECIS = 3
Private Const OUT_TT_ONLY_PRECIS = 7
Private Const OUT_TT_PRECIS = 4

' Clipping precision constants.
Private Const CLIP_CHARACTER_PRECIS = 1
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_EMBEDDED = &H80
Private Const CLIP_LH_ANGLES = &H10
Private Const CLIP_STROKE_PRECIS = 2
Private Const CLIP_TO_PATH = 4097
Private Const CLIP_TT_ALWAYS = &H20

' Character quality constants.
Private Const DEFAULT_QUALITY = 0
Private Const DRAFT_QUALITY = 1
Private Const PROOF_QUALITY = 2

' Pitch and family constants.
Private Const DEFAULT_PITCH = 0
Private Const FIXED_PITCH = 1
Private Const VARIABLE_PITCH = 2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const FF_DECORATIVE = 80  '  Old English, etc.
Private Const FF_DONTCARE = 0     '  Don't care or don't know.
Private Const FF_MODERN = 48      '  Constant stroke width, serifed or sans-serifed.
Private Const FF_ROMAN = 16       '  Variable stroke width, serifed.
Private Const FF_SCRIPT = 64      '  Cursive, etc.
Private Const FF_SWISS = 32



'Draw a rotated string centered at the indicated
'position using the indicated font parameters.
Public Sub CenterText(ByVal Pic As PictureBox, _
                       ByVal xmid As Single, ByVal ymid As Single, _
                       ByVal txt As String, _
                       ByVal nHeight As Long, _
                       Optional ByVal nWidth As Long = 0, _
                       Optional ByVal nEscapement As Long = 0, _
                       Optional ByVal fnWeight As Long = FW_NORMAL, _
                       Optional ByVal fbItalic As Long = False, _
                       Optional ByVal fbUnderline As Long = False, _
                       Optional ByVal fbStrikeOut As Long = False, _
                       Optional ByVal fbCharSet As Long = DEFAULT_CHARSET, _
                       Optional ByVal fbOutputPrecision As Long = OUT_TT_ONLY_PRECIS, _
                       Optional ByVal fbClipPrecision As Long = CLIP_EMBEDDED, _
                       Optional ByVal fbQuality As Long = DEFAULT_QUALITY, _
                       Optional ByVal fbPitchAndFamily As Long = TRUETYPE_FONTTYPE, _
                       Optional ByVal lpszFace As String = "Arial")

Dim NewFont As Long
Dim oldfont As Long
Dim text_metrics As TEXTMETRIC
Dim internal_leading As Single
Dim total_hgt As Single
Dim text_wid As Long
Dim text_hgt As Single
Dim text_bound_wid As Single
Dim text_bound_hgt As Single
Dim total_bound_wid As Single
Dim total_bound_hgt As Single
Dim theta As Single
Dim phi As Single
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim x4 As Single
Dim y4 As Single
Dim NRECT As RECT
Dim Flags As Long
    
    ' Create the font.
    NewFont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, _
                         fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, _
                         fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
                         
    oldfont = SelectObject(Pic.hDC, NewFont)

    ' Get the font metrics.
    GetTextMetrics Pic.hDC, text_metrics
    internal_leading = Pic.ScaleY(text_metrics.tmInternalLeading, vbPixels, Pic.ScaleMode)
    total_hgt = Pic.ScaleY(text_metrics.tmHeight, vbPixels, Pic.ScaleMode)
    text_hgt = total_hgt - internal_leading
    text_wid = 0
    text_wid = CLng(Pic.TextWidth(txt))

    ' Get the bounding box geometry.
    theta = nEscapement / 10 / 180 * PI
    phi = PI / 2 - theta
    text_bound_wid = text_hgt * Cos(phi) + text_wid * Cos(theta)
    text_bound_hgt = text_hgt * Sin(phi) + text_wid * Sin(theta)
    total_bound_wid = total_hgt * Cos(phi) + text_wid * Cos(theta)
    total_bound_hgt = total_hgt * Sin(phi) + text_wid * Sin(theta)

    ' Find the desired center point.
    X1 = xmid
    Y1 = ymid

    ' Subtract half the height and width of the text
    ' bounding box. This puts (x1, y2) in the upper
    ' left corner of the text bounding box.
    X1 = X1 - text_bound_wid / 2
    Y1 = Y1 - text_bound_hgt / 2

    ' The start position's X coordinate belongs at
    ' the left edge of the text bounding box, so
    ' x1 is correct. Move the Y coordinate down to
    ' its start position.
    Y1 = Y1 + text_wid * Sin(theta)

'    Find the other points on the text bounding box.
'    X2 = X1 + text_wid * Cos(theta) '
'    Y2 = Y1 - text_wid * Sin(theta) '
'    X3 = X2 + text_hgt * Cos(phi) '
'    Y3 = Y2 + text_hgt * Sin(phi) '
'    x4 = X3 + -text_wid * Cos(theta) '
'    y4 = Y3 + text_wid * Sin(theta) '
'
'    NRECT.Left = X1
'    NRECT.Top = Y1
'    NRECT.Right = X2
'    NRECT.Bottom = Y2

    ' See if we should draw the bounding box.
'    If chkShowBoundingBox.Value = vbChecked Then '
'        ' Draw the text bounding box.'
'        Pic.Line (X1, Y1)-(X2, Y2) '
'        Pic.Line -(X3, Y3) '
'        Pic.Line -(x4, y4) '
'        Pic.Line -(X1, Y1) '
'    End If

    ' See if we should mark the text and PictureBox
    ' center positions.
''    If chkMarkCenters.Value = vbChecked Then '
''        ' Draw lines to mark the center of the PictureBox.
''        Pic.Line (0, 0)-(Pic.ScaleWidth, Pic.ScaleHeight) '
''        Pic.Line (0, Pic.ScaleHeight)-(Pic.ScaleWidth, 0) '
''
''        ' Draw lines to mark the center of the text rectangle.
''        Pic.Line (x1, y1)-(x3, y3) '
''        Pic.Line (x2, y2)-(x4, y4) '
''    End If

    ' Move (x1, y1) to the start corner of the
    ' outer bounding box.
    X1 = X1 - (total_bound_wid - text_bound_wid)
    Y1 = Y1 - (total_bound_hgt - text_bound_hgt)
   
    ' Display the text.
    'Pic.CurrentX = X1
    'Pic.CurrentY = Y1
    'Pic.Print txt
    'ExtTextOut Pic.hdc, X1, Y1, ETO_OPAQUE, ByVal 0&, txt, Len(txt), ByVal 0&
    'Flags = DT_LEFT Or DT_WORD_ELLIPSIS Or DT_EXPANDTABS Or DT_WORDBREAK Or DT_TOP
    'DrawText Pic.hdc, txt, Len(txt), NRECT, Flags
    
    TextOut Pic.hDC, X1, Y1, txt, Len(txt)
    
'    Dim OLDForeColor As Long, tX As Integer, tY As Integer, m_Text3DType As Integer
'    OLDForeColor = Pic.ForeColor
'    Pic.ForeColor = Pic.FillColor
'    TextOut Pic.hdc, X1, Y1, txt, Len(txt)
'    m_Text3DType = 0
'    Select Case m_Text3DType
'    Case 0
'    Case 1: tX = -1: tY = -1
'    Case 2: tX = 0: tY = -1
'    Case 3: tX = 1: tY = -1
'    Case 4: tX = 1: tY = 0
'    Case 5: tX = 1: tY = 1
'    Case 6: tX = 0: tY = 1
'    Case 7: tX = -1: tY = 1
'    Case 8: tX = -1: tY = 0
'    End Select
'    Pic.ForeColor = OLDForeColor
'    TextOut Pic.hdc, X1 + tX, Y1 + tY, txt, Len(txt)
'
    
    ' Reselect the old font and delete the new one.
    NewFont = SelectObject(Pic.hDC, oldfont)
    Call DeleteObject(NewFont)
        
End Sub

Public Function CreateFontWmf(ByVal hDC As Long, _
                       ByVal nHeight As Long, _
                       Optional ByVal nWidth As Long = 0, _
                       Optional ByVal nEscapement As Long = 0, _
                       Optional ByVal fnWeight As Long = FW_NORMAL, _
                       Optional ByVal fbItalic As Long = False, _
                       Optional ByVal fbUnderline As Long = False, _
                       Optional ByVal fbStrikeOut As Long = False, _
                       Optional ByVal fbCharSet As Long = DEFAULT_CHARSET, _
                       Optional ByVal fbOutputPrecision As Long = OUT_TT_ONLY_PRECIS, _
                       Optional ByVal fbClipPrecision As Long = CLIP_EMBEDDED, _
                       Optional ByVal fbQuality As Long = PROOF_QUALITY, _
                       Optional ByVal fbPitchAndFamily As Long = TRUETYPE_FONTTYPE, _
                       Optional ByVal lpszFace As String = "Arial") As Long

    ' Create the font.
    CreateFontWmf = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, lpszFace)
            
End Function


'Read Path text and make PointCoolds and Type for draw
Public Sub ReadPathText(ByVal Obj As PictureBox, _
                        ByVal txt As String, _
                        ByRef Point_Coords() As POINTAPI, _
                        ByRef Point_Types() As Byte, _
                        ByVal NumPoints As Long)
    Dim ret As Long
    ret = BeginPath(Obj.hDC)
    Obj.Print txt
    ret = EndPath(Obj.hDC)
    NumPoints = 0
    NumPoints = GetPathAPI(Obj.hDC, ByVal 0&, ByVal 0&, 0)

    If (NumPoints) Then
        ReDim Point_Coords(NumPoints - 1)
        ReDim Point_Types(NumPoints - 1)
        'Get the path data from the DC
        Call GetPathAPI(Obj.hDC, Point_Coords(0), Point_Types(0), NumPoints)
    End If

End Sub

Public Sub LoadFonts(ByVal ComboBox As Object)
    Dim hDC As Long
    ComboBox.Clear
    hDC = GetDC(ComboBox.hwnd)
    Call EnumFontFamilies(hDC, vbNullString, AddressOf EnumFontFamProc, ComboBox)
    Call ReleaseDC(ComboBox.hwnd, hDC)
End Sub

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
    ByVal FontType As Long, lParam As ComboBox) As Long
    On Local Error Resume Next
    Dim FaceName As String
    Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    Call lParam.AddItem(Left$(FaceName, InStr(FaceName, vbNullChar) - 1))
    EnumFontFamProc = 1
End Function

'
'Function RotateText(inObj As Object, inText As String, inFontName As String, _
'                    inBold As Boolean, inItalic As Boolean, inFontSize As Integer, _
'                    inAngle As Long, inStyle As Integer, _
'                    firstClr As Long, secondClr As Long, mainClr As Long, _
'                    X As Single, Y As Single, _
'                    Optional inDepth As Integer = 1) As Boolean
'
'    On Error GoTo errHandler
'    RotateText = False
'
'    Dim L As LOGFONT
'    Dim mFont As Long
'    Dim mPrevFont As Long
'    Dim i As Integer
'    Dim origMode As Integer
'    Dim tmpX As Single, tmpY As Single
'    Dim mresult
'
'     ' For Windows NT to work
'    mresult = SetGraphicsMode(inObj.hDC, GM_ADVANCED)
'
'    origMode = inObj.ScaleMode
'    inObj.ScaleMode = vbPixels
'
'    'If inBold = True And inItalic = True Then
'    '    L.lfFaceName = inFontName & Space(1) & "Bold" & Space(1) & "Italic" & Chr(0)    ' Must be null terminated
'    'ElseIf inBold = True And inItalic = False Then
'    '    L.lfFaceName = inFontName & Space(1) & "Bold" + Chr$(0)
'    'ElseIf inBold = False And inItalic = True Then
'    '    L.lfFaceName = inFontName & Space(1) & "Italic" + Chr$(0)
'    'Else
'        L.lfFaceName = inFontName & Chr$(0)
'    'End If
'    L.lfCharSet = 1
'    If inBold Then L.lfWeight = 800 Else L.lfWeight = 400
'    L.lfItalic = True
'    L.lfEscapement = inAngle * 10
'    L.lfHeight = inFontSize * -20 / Screen.TwipsPerPixelY
'
'    mFont = CreateFontIndirect(L)
'    mPrevFont = SelectObject(inObj.hDC, mFont)
'
'    inObj.CurrentX = X
'    inObj.CurrentY = Y
'    inObj.Font.Charset = 161
'    tmpX = X
'    tmpY = Y
'    Select Case inStyle
'         Case 0          ' Ordinary shade
'            If firstClr <> -1 Then         ' Minus 1 indicate N/A
'                inObj.ForeColor = firstClr
'                For i = 1 To inDepth
'                    tmpX = tmpX + 1: tmpY = tmpY + 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If secondClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                tmpX = X
'                tmpY = Y
'                inObj.ForeColor = secondClr
'                For i = 1 To inDepth
'                    tmpX = tmpX - 1: tmpY = tmpY - 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If mainClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                inObj.ForeColor = mainClr
'                inObj.Print inText
'            End If
'
'        Case 1             'Embossed effect - text horizontal
'            If firstClr <> -1 Then
'                inObj.ForeColor = firstClr
'                For i = 1 To inDepth
'                    tmpX = tmpX - 1: tmpY = tmpY - 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If secondClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                tmpX = X
'                tmpY = Y
'                inObj.ForeColor = secondClr
'                For i = 1 To inDepth
'                    tmpX = tmpX + 1: tmpY = tmpY + 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If mainClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                inObj.ForeColor = mainClr
'                inObj.Print inText
'            End If
'
'         Case 2          ' Engroved effect - text horizontal
'            If firstClr <> -1 Then
'                inObj.ForeColor = firstClr
'                For i = 1 To inDepth
'                    tmpX = tmpX + 1: tmpY = tmpY + 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'
'            If secondClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                tmpX = X
'                tmpY = Y
'                inObj.ForeColor = secondClr
'                For i = 1 To inDepth
'                    tmpX = tmpX - 1: tmpY = tmpY - 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If mainClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                inObj.ForeColor = mainClr
'                inObj.Print inText
'            End If
'
'         Case 3          ' Embossed effect - text vertical
'            If firstClr <> -1 Then
'                inObj.ForeColor = secondClr
'                For i = 1 To inDepth
'                    tmpX = tmpX + 1: tmpY = tmpY + 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If secondClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                tmpX = X
'                tmpY = Y
'                inObj.ForeColor = firstClr
'                For i = 1 To inDepth
'                    tmpX = tmpX + 1: tmpY = tmpY + 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If mainClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                inObj.ForeColor = mainClr
'                inObj.Print inText
'            End If
'
'        Case 4             'Engraved effect - text vertical
'            If firstClr <> -1 Then
'                inObj.ForeColor = secondClr
'                For i = 1 To inDepth
'                    tmpX = tmpX - 1: tmpY = tmpY - 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If secondClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                tmpX = X
'                tmpY = Y
'                inObj.ForeColor = firstClr
'                For i = 1 To inDepth
'                    tmpX = tmpX - 1: tmpY = tmpY - 1
'                    inObj.CurrentX = tmpX
'                    inObj.CurrentY = tmpY
'                    inObj.Print inText
'                Next i
'            End If
'
'            If mainClr <> -1 Then
'                inObj.CurrentX = X
'                inObj.CurrentY = Y
'                inObj.ForeColor = mainClr
'                inObj.Print inText
'            End If
'    End Select
'
'    mresult = SelectObject(inObj.hDC, mPrevFont)
'    mresult = DeleteObject(mFont)
'    inObj.ScaleMode = origMode
'    RotateText = True
'    Exit Function
'errHandler:
'    inObj.ScaleMode = origMode
'    MsgBox "RotateText"
'End Function

Public Sub Draw_Example(Pic As PictureBox, txt As String, _
                         m3Deffect As Boolean, mRaised As Boolean, _
                         mColor1 As Long, mColor2 As Long, _
                         Optional m3DAngle As Single = 0, Optional mAngle As Single = 0)
    Dim XX As Single, yy As Single, Ang As Single, A As Integer, X As Single, Y As Single
    'Pic.Cls
    XX = Pic.CurrentX '20
    yy = Pic.CurrentY '20
    If mRaised = True Then GoTo raisedd
    If m3Deffect = False Then GoTo ddd1
'    chrs = 0
'    For a = 1 To Len(Text2.Text)
'        If Asc(Mid$(Text2.Text, a, 1)) < 48 Or Asc(Mid$(Text2.Text, a, 1)) > 57 Then chrs = 1
'    Next a
'    If chrs = 1 Then MsgBox "The angle text box should only contain numbers", vbCritical, "Error": Exit Sub
hh1:
    If mAngle > 360 Then mAngle = mAngle - 360: GoTo hh1
    'VALU = Slider1.Value
    If mAngle <= 0 Then mAngle = 0: GoTo hhjh
    Ang = 360 / mAngle
    Ang = ((PI * 2) / Ang) + (PI / 2)
hhjh:
    For A = 1 To m3DAngle
        Pic.ForeColor = mColor1
        Pic.DrawWidth = 1
        Pic.Line (XX + A * Cos(Ang), yy + A * Sin(Ang))-(XX + A * Cos(Ang), yy + A * Sin(Ang)), mColor1
        Pic.Print txt
    Next A
    GoTo ddff1
raisedd:
    For X = XX - 1 To XX Step 1
        For Y = yy - 1 To yy Step 1
            Pic.ForeColor = RGB(255, 255, 255)
            Pic.DrawWidth = 1: Pic.Line (X, Y)-(X, Y), RGB(255, 255, 255): Pic.Print txt
    Next Y, X
   For X = XX To XX + 1 Step 1
       For Y = yy To yy + 1 Step 1
           Pic.ForeColor = RGB(0, 0, 0)
           Pic.DrawWidth = 1: Pic.Line (X, Y)-(X, Y), RGB(0, 0, 0): Pic.Print txt
    Next Y, X
GoTo ddff1
ddd1:
'If Check1.Value = False Then GoTo ddff1
   For X = XX - 1 To XX + 1
       For Y = yy - 1 To yy + 1
           Pic.ForeColor = mColor1
           Pic.DrawWidth = 1: Pic.Line (X, Y)-(X, Y), mColor1: Pic.Print txt
   Next Y, X
ddff1:
    Pic.ForeColor = mColor2
    Pic.DrawWidth = 1: Pic.Line (XX, yy)-(XX, yy), mColor2: Pic.Print txt
    'Pic.DrawWidth = Form1.Slider1.Value
 End Sub

'dIST 1-100, Pxx=-1 - 1 ,Pyy=-1 - 1
Public Sub Print3D(Ob As Object, txt As String, Dist As Integer, _
                   mColor1 As Long, mColor2 As Long, _
                   Pxx As Single, PYY As Single)
                   
On Error Resume Next
Dim Sr As Single, Sg As Single, Sb As Single
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim TMidX As Long, TMidY As Long, XX As Long, yy As Long
   ' Ob.Cls
    'do 3D
    Pxx = Pxx / 10
    PYY = PYY / 10
    
    'make sure text is always centerred
    TMidX = (Ob.Width / 2) - (Ob.TextWidth(txt$) / 2)
    TMidY = (Ob.Height / 2) - (Ob.TextHeight(txt$) / 2)
    TMidX = TMidX - ((Pxx * Dist) / 2)
    TMidY = TMidY - ((PYY * Dist) / 2)
    SplitRGB mColor1, R1, G1, B1
    SplitRGB mColor2, R2, G2, B2
    
    Sr = (R2 - R1) / Dist
    Sg = (G2 - G1) / Dist
    Sb = (B2 - B1) / Dist
    
    'print a lot of text
    For XX = 0 To Dist - 1
        Ob.CurrentX = TMidX + (XX * Pxx)
        Ob.CurrentY = TMidY + (XX * PYY)
        R1 = R1 + Sr
        G1 = G1 + Sg
        B1 = B1 + Sb
        'the values cannot be < 0
        If Int(R1) < 0 Then R1 = 0
        If Int(G1) < 0 Then G1 = 0
        If Int(B1) < 0 Then B1 = 0
        Ob.ForeColor = RGB(R1, G1, B1)
        Ob.Print txt
    Next XX
    

End Sub

