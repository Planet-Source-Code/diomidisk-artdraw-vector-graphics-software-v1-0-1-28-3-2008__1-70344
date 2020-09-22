Attribute VB_Name = "ModFontFrame"
Option Explicit

Public Sub DrawFontFrame(hdc As Long, DrawString As String, cAngle As Single, cColor As Long, cAling As Integer, _
                         xMin As Single, yMin As Single, xMax As Single, yMax As Single, lpXform As XForm)
    Dim OldGM As Long
    Dim OldXForm As XForm
    Dim RotXForm As XForm
    Dim OldOrg As PointApi
    Dim OldCol As Long
    Dim LoopAngs As Long
    Dim OldMapMode As Long
    Dim Flags As Long
    Dim R As RECT
    Const PI = 3.14159265358979
    
    ' Buffer existing DC properties
    OldCol = SetTextColor(hdc, cColor)
    OldGM = SetGraphicsMode(hdc, GM_ADVANCED)
    OldMapMode = SetMapMode(hdc, MM_TEXT)
    Call GetWorldTransform(hdc, OldXForm)
    Call SetViewportOrgEx(hdc, xMin, yMin, OldOrg)
    SetWorldTransform hdc, lpXform
    ' Create rotation transformation
    'RotXForm = NewRotationXForm(cAngle)
'
    'Call SetWorldTransform(hdc, CombineXForm(NewReflectionXForm(CBool(LoopAngs And &H2), _
                                              CBool(LoopAngs And &H1)), RotXForm))
     Select Case cAling
     Case 0
        Flags = DT_LEFT
     Case 1
        Flags = DT_CENTER
     Case 2
        Flags = DT_RIGHT
     Case Else
        Flags = DT_CENTER
     End Select
     Flags = Flags Or DT_WORD_ELLIPSIS Or DT_EXPANDTABS Or DT_WORDBREAK Or DT_TOP
    
    R.Right = xMax '- (Sin((cAngle - 90 Mod 360) * PI / 180)) '+ xMin
    R.Bottom = yMax '- (Cos((cAngle - 90 Mod 360) * PI / 180)) '+ yMin
    
    DrawText hdc, DrawString, Len(DrawString), R, Flags
'    ' Re-set DC properties
    Call SetViewportOrgEx(hdc, OldOrg.x, OldOrg.y, ByVal 0&)
    Call SetWorldTransform(hdc, OldXForm)
    Call SetGraphicsMode(hdc, OldGM)
    Call SetTextColor(hdc, OldCol)
End Sub

Private Function NEWXForm(ByVal inM11 As Single, ByVal inM12 As Single, _
                          ByVal inM21 As Single, ByVal inM22 As Single, _
                          ByVal inDx As Single, ByVal inDy As Single) As XForm
    
    With NEWXForm ' Set all the members of this structure
        .eM11 = inM11
        .eM12 = inM12
        .eM21 = inM21
        .eM22 = inM22
        .eDx = inDx
        .eDy = inDy
    End With
    
End Function

Private Function NewIdentityXForm() As XForm
    NewIdentityXForm = NEWXForm(1, 0, 0, 1, 0, 0)
End Function

Private Function NewRotationXForm(ByVal inAngle As Single) As XForm
    Dim AngRad As Single

    AngRad = (inAngle / 180) * 3.14159
    NewRotationXForm = NEWXForm(Cos(AngRad), Sin(AngRad), -Sin(AngRad), Cos(AngRad), 0, 0)
End Function

Private Function NewReflectionXForm(ByVal inHoriz As Boolean, ByVal inVert As Boolean) As XForm
    NewReflectionXForm = NEWXForm( _
        IIf(inHoriz, -1, 1), 0, 0, IIf(inVert, -1, 1), 0, 0)
End Function

Private Function CombineXForm(ByRef inA As XForm, ByRef inB As XForm) As XForm
    Call CombineTransform(CombineXForm, inA, inB)
End Function

