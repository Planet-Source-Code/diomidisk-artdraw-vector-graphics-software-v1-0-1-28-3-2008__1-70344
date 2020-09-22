VERSION 5.00
Begin VB.UserControl ColorPalette 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ColorPalette.ctx":0000
   Begin VB.CommandButton CommandOpen 
      Height          =   270
      Left            =   3015
      Picture         =   "ColorPalette.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open Palette"
      Top             =   30
      Width           =   270
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   0
      Max             =   255
      SmallChange     =   10
      TabIndex        =   0
      Top             =   270
      Width           =   3285
   End
   Begin VB.PictureBox PicturePalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   15
      MouseIcon       =   "ColorPalette.ctx":05C4
      MousePointer    =   99  'Custom
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   199
      TabIndex        =   1
      Top             =   0
      Width           =   2985
   End
End
Attribute VB_Name = "ColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_FileNamePalette = ""

'Property Variables:
Dim m_FileNamePalette As String

Dim ColorList() As Long
Dim MaxCol As Integer
Dim cH As Long
Dim cView As Integer

'Event Declarations:
Event Click() 'MappingInfo=PicturePalette,PicturePalette,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=PicturePalette,PicturePalette,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicturePalette,PicturePalette,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicturePalette,PicturePalette,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicturePalette,PicturePalette,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event ColorSelected(Button As Integer, cColor As Long)
Event ColorOver(cColor As Long)

Private Sub CommandOpen_Click()
Dim user_canceled As Boolean
Dim FileNamePalette As String

    user_canceled = FrmPalette.ShowForm(FileNamePalette)
    Unload FrmPalette

    ' If the user canceled, do no more.
    If user_canceled Then Exit Sub
 
    LoadPalette FileNamePalette
    PicturePalette.SetFocus
    
End Sub

Private Sub HScroll1_Change()
       cView = HScroll1.Value
       LoadPalette
End Sub

Private Sub UserControl_Resize()
      'UserControl.Height = HScroll1.Height * 2
      cH = 18 'HScroll1.Height
      PicturePalette.Move 0, 0, 256 * cH + 18, cH - 1
      HScroll1.Move 0, 18, UserControl.ScaleWidth  ', cH
      CommandOpen.Move UserControl.ScaleWidth - 18, 0, 18, 18
      LoadPalette
End Sub
Public Sub LoadPalette(Optional PalFile As String)
   
On Error Resume Next ' GoTo ErrLoad
Dim FF As Integer
Dim tStr As String
Dim n As Integer
Dim cQty As Integer
Dim Row As Integer
Dim Col As Integer

   If PalFile <> "" Then FileNamePalette = PalFile

FF = FreeFile

If PalFile = "" Or Dir(PalFile) = "" Then
    If FileNamePalette <> "" Then
      If FileExists(FileNamePalette) Then
         PalFile = FileNamePalette
      Else
        PalFile = App.Path & "\Palette\Default.pal"
      End If
   Else
      PalFile = App.Path & "\Palette\Default.pal"
   End If
End If
'UserControl.Line (0, 0)-(UserControl.Width, UserControl.Height), QBColor(15), BF
If Dir(PalFile) <> "" Then
Open PalFile For Input As #FF
    Input #FF, tStr$ 'JASC-PAL
    If UCase(tStr) <> "JASC-PAL" Then
       Close #FF
        Exit Sub
    End If
    Input #FF, tStr$ '0010
    Input #FF, tStr$ '256 (color qty)
    cQty = Int(tStr)
    
    ReDim ColorList(Int(cQty))
    n = 0
    While Not EOF(FF)
       Input #FF, tStr$
ragain:
       tStr$ = Replace(tStr$, "  ", " ")
       If InStr(1, tStr$, "  ") Then GoTo ragain
       ColorList(n) = RGB(Split(tStr, " ")(0), Split(tStr, " ")(1), Split(tStr, " ")(2))
       n = n + 1
    Wend
Close #FF

   HScroll1.Max = cQty - (UserControl.ScaleWidth \ cH) + 1
   PicturePalette.Line (0, 0)-(PicturePalette.Width, PicturePalette.Height), QBColor(15), BF
   PicturePalette.Line (0, 0)-(18, 18), QBColor(0)
   PicturePalette.Line (0, 18)-(18, 0), QBColor(0)
   Col = 1
   MaxCol = cQty
    For n = cView To cQty - 1
       PicturePalette.Line (Col * cH, 0)-Step(cH, cH), ColorList(n), BF
       PicturePalette.Line (Col * cH, 0)-Step(cH, cH), , B
       Col = Col + 1
    Next n
    
End If

Exit Sub

ErrLoad:
   Close #FF
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileNamePalette = m_def_FileNamePalette
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_FileNamePalette = PropBag.ReadProperty("FileNamePalette", m_def_FileNamePalette)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("FileNamePalette", m_FileNamePalette, m_def_FileNamePalette)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
End Sub

Private Sub PicturePalette_Click()
    RaiseEvent Click
End Sub

Private Sub PicturePalette_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub PicturePalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim tColor As Long
Dim tInd As Integer
    If (X \ cH) - 1 = -1 Then
       tInd = -1
    Else
       tInd = (cView + X \ cH) - 1
    End If
    'Debug.Print tInd
    If tInd > UBound(ColorList) Then Exit Sub
    If tInd >= 0 Then
        tColor = ColorList(tInd)
    End If
    If tInd = -1 Then tColor = -1
    RaiseEvent ColorSelected(Button, tColor)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub PicturePalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Dim tColor As Long
Dim tInd As Integer

    tInd = (cView + X \ cH) - 1
    If tInd > UBound(ColorList) Or tInd = -1 Then Exit Sub
    tColor = ColorList(tInd)
    If tInd = -1 Then tColor = -1
    RaiseEvent ColorOver(tColor)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PicturePalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileNamePalette() As String
    FileNamePalette = m_FileNamePalette
End Property

Public Property Let FileNamePalette(ByVal New_FileNamePalette As String)
    m_FileNamePalette = New_FileNamePalette
    PropertyChanged "FileNamePalette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

