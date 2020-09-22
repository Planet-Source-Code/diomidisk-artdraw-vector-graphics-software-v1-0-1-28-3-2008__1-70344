VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl DrawControl 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   606
   Begin ArtDraw.MeForm MeForm3 
      Height          =   2475
      Left            =   4485
      TabIndex        =   16
      Top             =   735
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   4366
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin ArtDraw.CtrTranform CtrTranform1 
         Height          =   2130
         Left            =   105
         TabIndex        =   17
         Top             =   285
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   3757
      End
   End
   Begin ArtDraw.MeForm MeForm2 
      Height          =   3150
      Left            =   6765
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton ComDropperFill 
         Height          =   315
         Left            =   270
         Picture         =   "DrawControl.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2295
         Width           =   345
      End
      Begin ArtDraw.CtlFill CtlFill1 
         Height          =   2790
         Left            =   15
         TabIndex        =   19
         Top             =   255
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   4921
         Color2          =   16777215
      End
   End
   Begin ArtDraw.MeForm MeForm1 
      Height          =   2850
      Left            =   6750
      TabIndex        =   5
      Top             =   705
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ImageList imlDrawWidths 
         Left            =   1635
         Top             =   1275
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":059C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":07AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":09C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0BD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0DE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":0FF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1208
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":141A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":162C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlDrawStyles 
         Left            =   1665
         Top             =   1695
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   40
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":183E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1BE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1DB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":1F86
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "DrawControl.ctx":2158
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo icbDrawStyle 
         Height          =   330
         Left            =   90
         TabIndex        =   11
         Top             =   1935
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin MSComctlLib.ImageCombo icbDrawWidth 
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   1290
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.CommandButton ComDropperPen 
         Height          =   315
         Left            =   120
         Picture         =   "DrawControl.ctx":232A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2340
         Width           =   360
      End
      Begin VB.CommandButton cmdSysColorsPen 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   1650
         Picture         =   "DrawControl.ctx":26B4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " System colors "
         Top             =   555
         Width           =   375
      End
      Begin VB.CommandButton CommandPen 
         Caption         =   "Apply"
         Height          =   330
         Left            =   495
         TabIndex        =   7
         Top             =   2325
         Width           =   1485
      End
      Begin VB.PictureBox PicPenColor 
         BackColor       =   &H00000000&
         Height          =   465
         Left            =   180
         ScaleHeight     =   405
         ScaleWidth      =   1380
         TabIndex        =   6
         Top             =   495
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pen Color "
         Height          =   180
         Left            =   195
         TabIndex        =   14
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label LbDrawWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Width"
         Height          =   180
         Left            =   165
         TabIndex        =   13
         Top             =   1035
         Width           =   1935
      End
      Begin VB.Label LblDrawStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Style"
         Height          =   195
         Left            =   165
         TabIndex        =   12
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.PictureBox picHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   105
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   4
      Top             =   1245
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1335
      LargeChange     =   50
      Left            =   6375
      Max             =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton ComCorner 
      Height          =   240
      Left            =   6405
      TabIndex        =   0
      Top             =   5535
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   50
      Left            =   4860
      Max             =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5535
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox PicCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   855
      MousePointer    =   99  'Custom
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   19
      Left            =   2835
      Picture         =   "DrawControl.ctx":287E
      Top             =   570
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   18
      Left            =   1605
      Picture         =   "DrawControl.ctx":29D0
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   17
      Left            =   1110
      Picture         =   "DrawControl.ctx":2B22
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   16
      Left            =   660
      Picture         =   "DrawControl.ctx":2C74
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   15
      Left            =   420
      Picture         =   "DrawControl.ctx":2DC6
      Top             =   615
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   14
      Left            =   2460
      Picture         =   "DrawControl.ctx":2F18
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   13
      Left            =   2040
      Picture         =   "DrawControl.ctx":306A
      Top             =   690
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   12
      Left            =   7380
      Picture         =   "DrawControl.ctx":31BC
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   11
      Left            =   6705
      Picture         =   "DrawControl.ctx":330E
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   10
      Left            =   6015
      Picture         =   "DrawControl.ctx":3460
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   9
      Left            =   5460
      Picture         =   "DrawControl.ctx":376A
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   8
      Left            =   4875
      Picture         =   "DrawControl.ctx":3A74
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   7
      Left            =   4335
      Picture         =   "DrawControl.ctx":433E
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   6
      Left            =   3675
      Picture         =   "DrawControl.ctx":4648
      Top             =   75
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   5
      Left            =   3060
      Picture         =   "DrawControl.ctx":4952
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   4
      Left            =   2385
      Picture         =   "DrawControl.ctx":4C5C
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   3
      Left            =   1725
      Picture         =   "DrawControl.ctx":4F66
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   2
      Left            =   1035
      Picture         =   "DrawControl.ctx":5270
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   1
      Left            =   450
      Picture         =   "DrawControl.ctx":557A
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageMouse 
      Height          =   480
      Index           =   0
      Left            =   45
      Picture         =   "DrawControl.ctx":5E44
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuTransform 
      Caption         =   "Transform"
      Begin VB.Menu mnuClearTransform 
         Caption         =   "Clear Transform"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCurve 
         Caption         =   "Make Curve"
      End
      Begin VB.Menu mnuFillMode 
         Caption         =   "FillMode"
         Begin VB.Menu mnuAlternate 
            Caption         =   "Alternate"
         End
         Begin VB.Menu mnuWinding 
            Caption         =   "Winding"
         End
      End
      Begin VB.Menu seplock 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLock 
         Caption         =   "Lock"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUnlock 
         Caption         =   "UnLock"
      End
      Begin VB.Menu sepEditPoints 
         Caption         =   "-"
      End
      Begin VB.Menu mnueditPoints 
         Caption         =   "Edit points"
      End
      Begin VB.Menu sepProperty 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperty 
         Caption         =   "Property"
      End
   End
End
Attribute VB_Name = "DrawControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim idredraw As Long

'Default Property Values:
Const m_def_LockObject = 0
Const m_def_BackImage = ""
Const m_def_Blend = 0
Const m_def_FillColor2 = 0
Const m_def_Pattern = ""
Const m_def_TypeGradient = 0
Const m_def_ShowPenProperty = 0
Const m_def_ShowFillProperty = 0
Const m_def_m_ShowTranformProperty = False
Const m_def_FileTitle = ""
Const m_def_FileName = ""
Const m_def_ForeColor = 0
Const m_def_DrawWidth = 1
Const m_def_DrawStyle = 0
Const m_def_FillStyle = 0
Const m_def_FillColor = 0
Const m_def_ShowCanvasSize = False
Const m_def_CanvasWidth = 640 '2480
Const m_def_CanvasHeight = 480 '3508
Const m_def_ZoomFactor = 1

'Property Variables:
Dim m_LockObject As Boolean
Dim m_BackImage As String
Dim m_Blend As Integer
Dim m_FillColor2 As OLE_COLOR
Dim m_Pattern As String
Dim m_TypeGradient As Integer
Dim m_ShowPenProperty As Boolean
Dim m_ShowFillProperty As Boolean
Dim m_ShowTranformProperty As Boolean
Dim m_FileTitle As String
Dim m_FileName As String
Dim m_ForeColor As OLE_COLOR
Dim m_DrawWidth As Integer
Dim m_DrawStyle As Integer
Dim m_FillStyle As Integer
Dim m_FillColor As OLE_COLOR
Dim m_ShowCanvasSize As Boolean
Dim m_CanvasLeft As Long
Dim m_CanvasTop As Long
Dim m_CanvasWidth As Long
Dim m_CanvasHeight As Long
Dim m_Image As Picture
Dim m_ZoomFactor As Single

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event SetDirty()
Event EnableMenusForSelection()
Event ColorSelected(tColor As Integer, cColor As Long)
Event MsgControl(txt As String)

Private Obj As vbdObject
Private TmpObj As vbdObject
Private m_DrawingObject As Boolean
Private mEditObject As Boolean

' Rubberband variables.
Private m_StartX As Single
Private m_StartY As Single
Private m_LastX As Single
Private m_LastY As Single
Private x1 As Single
Private x2 As Single
Private y1 As Single
Private y2 As Single

Public XminBox As Single
Public YminBox As Single
Public XmaxBox As Single
Public YmaxBox As Single

Public mPicture As StdPicture

Private Xmin_Box As Single
Private Ymin_Box As Single
Private Xmax_Box As Single
Private Ymax_Box As Single

Private m_ReadFillProperty As Boolean
Private m_ReadPenProperty As Boolean

Private m_Rotate As Boolean
Private m_Move As Boolean
Private m_Scale As Boolean
Private m_Skew As Boolean

Private m_ScaleType As Integer
Private Ortho As RectAngle

Dim M(1 To 3, 1 To 3) As Single
Dim MeFormView1 As Boolean
Dim MeFormView2 As Boolean
Dim MeFormView3 As Boolean

Public Enum m_Order
    BringToFront = 0
    SendToBack = 1
    BringFoward = 2
    SendBackward = 3
End Enum


Private Sub cmdSysColorsPen_Click()
     OpenColorDialog PicPenColor
End Sub

Private Sub ComCorner_Click()
    HScroll1.Value = (HScroll1.Max - HScroll1.Min) \ 2
    VScroll1.Value = (VScroll1.Max - VScroll1.Min) \ 2
End Sub

Private Sub ComDropperPen_Click()
      m_ReadPenProperty = True
      SelectTool 20 '"DropperFill"
End Sub

Private Sub CommandPen_Click()
     DrawStyle = icbDrawStyle.SelectedItem.Index - 1
     DrawWidth = icbDrawWidth.SelectedItem.Index
     ForeColor = PicPenColor.BackColor
     Redraw
End Sub

Private Sub CtlFill1_Apply(nTypeFill As Integer, nFillStyle As Integer, nColor1 As Long, nColor2 As Long, nPattern As String, nTypeGradient As Integer, mBlend As Integer)
      Select Case nTypeFill
      Case 1
            FillStyle = nFillStyle  'icbFillStyle.SelectedItem.Index - 1
            FillColor = nColor1
      Case 2
            FillStyle = nFillStyle
            FillColor = nColor1
            FillColor2 = nColor2
            Pattern = nPattern
      Case 3
            FillStyle = nFillStyle
            FillColor = nColor1
            FillColor2 = nColor2
            Gradient = nTypeGradient
      Case 4
            FillStyle = nFillStyle
            Pattern = nPattern
      End Select
      
      Blend = mBlend
      Redraw
End Sub


Private Sub CtlFill1_ApplyImage(nTypeFill As Integer, nFillStyle As Integer, nPattern As String, nPicture As stdole.StdPicture, mBlend As Integer)
        
       If Obj Is Nothing Then Exit Sub
        FillStyle = nFillStyle
        Pattern = nPattern
        Blend = mBlend
      '  mPicture = nPicture
        
         ' Add the transformation to the selected objects.
        For Each Obj In m_SelectedObjects
           If Obj.Selected = True Then
              Obj.Pattern = nPattern
              Set Obj.Picture = nPicture
           End If
        Next Obj
        Redraw
End Sub

Private Sub CtrTranform1_TransformMove(X_Move As Single, Y_Move As Single)
       TransformPoint X_Move, Y_Move
End Sub

Private Sub CtrTranform1_TransformMirror(X_Skew As Integer, Y_Skew As Integer)
       TransformMirror X_Skew, Y_Skew
End Sub

Private Sub CtrTranform1_TransformRotate(t_Angle As Single, xmin As Single, ymin As Single, xmax As Single, ymax As Single)
       If Obj Is Nothing Then Exit Sub
       Obj.Bound XminBox, YminBox, XmaxBox, YmaxBox
       TransformRotate t_Angle, XminBox, YminBox, XmaxBox, YmaxBox
End Sub

Private Sub CtrTranform1_TransformScale(X_scale As Single, Y_Scale As Single)
       m_ScaleType = 9
       TransformScale X_scale, Y_Scale
       m_ScaleType = 0
End Sub

Private Sub CtrTranform1_TransformSkew(X_Skew As Single, Y_Skew As Single)
       m_ScaleType = 9
       TransformSkew X_Skew, Y_Skew
       m_ScaleType = 0
End Sub

Private Sub HScroll1_Change()
 On Error Resume Next
    'picCanvas.Left = HScroll1.Value
    m_CanvasLeft = HScroll1.Value
    PicCanvas.Visible = False
    PicCanvas.Left = m_CanvasLeft
    ReDrawPage
    PicCanvas.Visible = True
End Sub

Private Sub ReDrawPage()
'     picCanvas.FontBold = True
'     'picCanvas.BackColor = &H8000000F
'     picCanvas.DrawWidth = 1
'
'     UserControl.Line (picCanvas.Left + 4, picCanvas.Top + 4)-Step(picCanvas.Width + 2, picCanvas.Height + 2), &H80000015, BF
'     'UserControl.Line (picCanvas.Left + 2, picCanvas.Top + 2)-Step(picCanvas.Width + 2, picCanvas.Height + 2), &H80000015, BF
'
'     If m_ShowCanvasSize = True Then
'        UserControl.CurrentX = picCanvas.Left + picCanvas.Width - UserControl.TextWidth(m_CanvasWidth & " X " & m_CanvasHeight) + 7
'        UserControl.CurrentY = picCanvas.Top + picCanvas.Height + 7
'        UserControl.Print m_CanvasWidth & " X " & m_CanvasHeight
'     End If
     
        UserControl.Cls
        UserControl.FontBold = True
        UserControl.DrawWidth = 1
        PicCanvas.Visible = True
        
        PicCanvas.Move m_CanvasLeft, m_CanvasTop '
        
        UserControl.Line (PicCanvas.Left + 4, PicCanvas.Top + 4)-Step(PicCanvas.Width + 2, PicCanvas.Height + 2), &H80000015, BF
        UserControl.Line (PicCanvas.Left - 1, PicCanvas.Top - 1)-Step(PicCanvas.Width + 1, PicCanvas.Height + 1), vbWhite, BF
        UserControl.Line (PicCanvas.Left - 1, PicCanvas.Top - 1)-Step(PicCanvas.Width + 1, PicCanvas.Height + 1), , B

        If m_ShowCanvasSize = True Then
            UserControl.CurrentX = PicCanvas.Left + PicCanvas.Width - PicCanvas.TextWidth(m_CanvasWidth & " X " & m_CanvasHeight) + 7
            UserControl.CurrentY = PicCanvas.Top + PicCanvas.Height + 7
            UserControl.Print m_CanvasWidth & " X " & m_CanvasHeight
        End If

End Sub

Private Sub icbDrawStyle_Click()
     icbDrawStyle.ToolTipText = icbDrawStyle.SelectedItem.Key
End Sub

Private Sub icbDrawWidth_Click()
    icbDrawWidth.ToolTipText = icbDrawWidth.SelectedItem.Key
End Sub

Private Sub MeForm1_Hide()
       MeForm1.Visible = False
       m_ShowPenProperty = False
End Sub

Private Sub MeForm2_Hide()
     MeForm2.Visible = False
     m_ShowFillProperty = False
End Sub


Private Sub MeForm3_Hide()
     MeForm3.Visible = False
     m_ShowTranformProperty = False
End Sub

Private Sub mnuAlternate_Click()
        If Not Obj Is Nothing Then
           Obj.FillMode = fALTERNATE
           Redraw
        End If
End Sub

Private Sub mnuClearTransform_Click()
   ClearTransform
End Sub

Private Sub mnuCurve_Click()
     Dim PointCoords() As POINTAPI
     Dim PointType() As Byte, tmpType() As Byte
     Dim iCounter As Long, StartCounter As Long, EndCounter As Long
     Dim tx() As Single, ty() As Single, TPoint() As Byte
     Dim OldObj As vbdObject
     Dim txt As String, OldTxt As String, I As Long
     Dim xmin As Single, ymin As Single, xmax As Single, ymax As Single
          
     If Not Obj Is Nothing Then
     
       iCounter = 0
       Select Case Obj.TypeDraw
       Case dText
          Debug.Print Obj.Serialization
          Obj.Bound xmin, ymin, xmax, ymax
          PicCanvas.ForeColor = Obj.ForeColor
          PicCanvas.FillColor = Obj.FillColor
          BeginPath PicCanvas.hDC
          CenterText PicCanvas, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), _
                     Obj.TextDraw, Obj.Size * gZoomFactor, , -Obj.Angle * 10, Obj.Weight, _
                     Obj.Italic, Obj.Underline, Obj.Strikethrough, Obj.Charset, , , , , Obj.Name
          EndPath PicCanvas.hDC
          
          iCounter = GetPathAPI(PicCanvas.hDC, ByVal 0&, ByVal 0&, 0)
           If (iCounter) Then
             ReDim PointCoords(iCounter - 1)
             ReDim PointType(iCounter - 1)
             'Get the path data from the DC
             Call GetPathAPI(PicCanvas.hDC, PointCoords(0), PointType(0), iCounter)
               StartCounter = 0
               EndCounter = iCounter - 1
          End If
          
       Case dEllipse, dFreePolygon, dPolygon, dPolyline, dScribble, dRectAngle
            Obj.ReadPoint iCounter, tx, ty, tmpType
            ReDim PointCoords(1 To iCounter)
            ReDim PointType(1 To iCounter)
            StartCounter = 1
            EndCounter = iCounter
            PointType = tmpType
            I = 0
            For I = 1 To iCounter
                PointCoords(I).X = tx(I)
                PointCoords(I).Y = ty(I)
            Next
       End Select
       
       If (iCounter) Then
             txt = txt & " DrawWidth(" + Trim(Str(Obj.DrawWidth)) + ")"
             txt = txt & " DrawStyle(" + Trim(Str(Obj.DrawStyle)) + ")"
             txt = txt & " ForeColor(" + Trim(Str(Obj.ForeColor)) + ")"
             txt = txt & " FillColor(" + Trim(Str(Obj.FillColor)) + ")"
             txt = txt & " FillColor2(" + Trim(Str(Obj.FillColor2)) + ")"
             txt = txt & " FillMode(" + Trim(Str(Obj.FillMode)) + ")"
             txt = txt & " Pattern(" + Obj.Pattern + ")"
             txt = txt & " Gradient(" + Trim(Str(Obj.Gradient)) + ")"
             txt = txt & " FillStyle(" + Trim(Str(Obj.FillStyle)) + ")"
             txt = txt & " TypeDraw(" + Format$(dPolydraw) + ")"
             txt = txt & " TextDraw()"
             txt = txt & " CurrentX(" + Trim(Str(Obj.CurrentX)) + ")"
             txt = txt & " CurrentY(" + Trim(Str(Obj.CurrentY)) + ")"
             txt = txt & " TypeFill(" + Trim(Str(Obj.TypeFill)) + ")"
             txt = txt & " ObjLock(" + Trim(Str(Obj.ObjLock)) + ")"
             txt = txt & " Blend(" + Trim(Str(Obj.Blend)) + ")"
             txt = txt & " Shade(False)"
             txt = txt & " AlingText(0)"
             txt = txt & " Bold(0)"
             txt = txt & " Charset(0)"
             txt = txt & " Italic(0)"
             txt = txt & " Name()"
             txt = txt & " Size(0)"
             txt = txt & " Strikethrough(0)"
             txt = txt & " Underline(0)"
             txt = txt & " Weight(400)"
             txt = txt & " Angle(" + Trim(Str(Obj.Angle)) + ")"
             txt = txt & vbCr & "Transformation(1 0 0 0 1 0 0 0 1 )"
             txt = txt & " IsClosed(True)"
             txt = txt & " NumPoints(" & Format$(iCounter) & ")"
    
             For I = StartCounter To EndCounter
                 txt = txt & vbCrLf & "    X(" & Format$(PointCoords(I).X) & ")"
                 txt = txt & " Y(" & Format$(PointCoords(I).Y) & ")"
                 txt = txt & " P(" & Format$(PointType(I)) & ")"
             Next I
             txt = "PolyDraw(PolyDraw(" & txt & "))"
             ObjectDelete
             OldTxt = Clipboard.GetText
             Clipboard.SetText txt
             PasteObject
             Clipboard.SetText OldTxt
       End If
       End If
End Sub

'only Polygon
Private Sub mnueditPoints_Click()

Dim PointType() As Byte, tx() As Single, ty() As Single, PointCoods() As POINTAPI
Dim iCounter As Long, xmin As Single, ymin As Single, xmax As Single, ymax As Single
Dim NumPoints As Integer, F As String, n As Single, Alfa As Single, np As Long

       Obj.Bound xmin, ymin, xmax, ymax
       
       F = InputBox("Number of points (3-20)", "Polygon")
       If F <> "" Then
        If Val(F) >= 3 And Val(F) <= 20 Then
          NumPoints = Val(F)
            ReDim tx(1 To NumPoints)
            ReDim ty(1 To NumPoints)
            ReDim PointType(1 To NumPoints)
            PointCoods = PolygonPoints(NumPoints, xmin, ymin, xmax, ymax)
            For iCounter = 1 To NumPoints
               tx(iCounter) = PointCoods(iCounter - 1).X
               ty(iCounter) = PointCoods(iCounter - 1).Y
               PointType(iCounter) = 2
            Next
            PointType(1) = 6
            PointType(NumPoints) = 3
            np = NumPoints
            Obj.NewPoint np, tx, ty, PointType
         End If
       End If
End Sub

Private Sub mnuFillMode_Click()
     mnutransform_Click
End Sub

Private Sub mnuLock_Click()
       ObjectLock True
End Sub

Private Sub mnuProperty_Click()
   Dim msg As String
   Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single
      If Not Obj Is Nothing Then
         Obj.Bound x1, y1, x2, y2
         msg = "StartX:" + Str(x1) + " - StartY:" + Str(y1) + vbCr
         msg = msg + "EndX:" + Str(x1) + " - EndY:" + Str(y1) + vbCr
         msg = msg + Obj.Info + vbCr
         msg = msg + "FillColor1:" + Str(Obj.FillColor) + vbCr
         msg = msg + "FillColor2:" + Str(Obj.FillColor2) + vbCr
         msg = msg + "ForeColor:" + Str(Obj.ForeColor) + vbCr
         msg = msg + "FillMode:" + Str(Obj.FillMode) + vbCr
         msg = msg + "FillStyle:" + Str(Obj.FillStyle) + vbCr
         msg = msg + "Gradient:" + Str(Obj.Gradient) + vbCr
         msg = msg + "Blend:" + Str(Obj.Blend) + vbCr
         msg = msg + "Pattern:" + Obj.Pattern
         MsgBox msg, vbInformation
      End If
End Sub

Private Sub mnutransform_Click()
      If Not Obj Is Nothing Then
         Select Case Obj.FillMode
         Case 1
            mnuAlternate.Checked = True
            mnuWinding.Checked = False
         Case 2
            mnuAlternate.Checked = False
            mnuWinding.Checked = True
         Case Else
            mnuAlternate.Checked = False
            mnuWinding.Checked = False
         End Select
         
         If Obj.ObjLock = True Then
            mnuLock.Enabled = False
            mnuUnlock.Enabled = True
         Else
            mnuLock.Enabled = True
            mnuUnlock.Enabled = False
         End If
         If Obj.TypeDraw = dText Or Obj.TypeDraw = dEllipse _
            Or Obj.TypeDraw = dFreePolygon Or Obj.TypeDraw = dPolygon _
            Or Obj.TypeDraw = dPolyline Or Obj.TypeDraw = dRectAngle _
            Or Obj.TypeDraw = dTextFrame Then
            mnuCurve.Enabled = True
         Else
            mnuCurve.Enabled = False
         End If
         
         mnuProperty.Enabled = True
      Else
         mnuCurve.Enabled = False
         mnuAlternate.Checked = False
         mnuWinding.Checked = False
         mnuLock.Enabled = False
         mnuUnlock.Enabled = False
         mnuProperty.Enabled = False
      End If
End Sub

Private Sub mnuUnlock_Click()
      ObjectLock False
       
End Sub

Private Sub mnuWinding_Click()
      If Not Obj Is Nothing Then
           Obj.FillMode = fWINDING
           Redraw
      End If
End Sub

Private Sub picCanvas_DblClick()
Dim nfonts As New StdFont
Dim PointCoords() As POINTAPI
Dim PointType() As Byte, tx() As Single, ty() As Single, TypePoint() As Byte
Dim iCounter As Long, NewText As String, xmin As Single, ymin As Single, xmax As Single, ymax As Single, mAlingText As Integer

    If Obj Is Nothing Then Exit Sub
    
      If Obj.TypeDraw = dText Or Obj.TypeDraw = dTextFrame Then
          NewText = Obj.TextDraw
          nfonts.Charset = Obj.Charset
          nfonts.Italic = Obj.Italic
          nfonts.Name = Obj.Name
          nfonts.Size = Obj.Size
          nfonts.Strikethrough = Obj.Strikethrough
          nfonts.Underline = Obj.Underline
          nfonts.Weight = Obj.Weight
          If FrmFonts.ShowForm(nfonts, NewText, mAlingText) = False Then
              Obj.Bold = nfonts.Bold
              Obj.Charset = nfonts.Charset
              Obj.Italic = nfonts.Italic
              Obj.Name = nfonts.Name
              Obj.Size = nfonts.Size
              Obj.Strikethrough = nfonts.Strikethrough
              Obj.Underline = nfonts.Underline
              Obj.Weight = nfonts.Weight
              Obj.TextDraw = NewText
              With PicCanvas
                 .Font.Bold = nfonts.Bold
                 .Font.Charset = nfonts.Charset
                 .Font.Italic = nfonts.Italic
                 .Font.Name = nfonts.Name
                 .Font.Size = nfonts.Size
                 .Font.Strikethrough = nfonts.Strikethrough
                 .Font.Underline = nfonts.Underline
                 .Font.Weight = nfonts.Weight
              End With
              If Obj.TypeDraw = dText Then
                 PicCanvas.CurrentX = Obj.CurrentX
                 PicCanvas.CurrentY = Obj.CurrentY
                 ReadPathText PicCanvas, NewText, PointCoords(), PointType(), iCounter
                 ReDim tx(1 To iCounter), ty(1 To iCounter), TypePoint(1 To iCounter)
                 With m_Polygon
                    For I = 1 To iCounter
                       tx(I) = PointCoords(I - 1).X
                       ty(I) = PointCoords(I - 1).Y
                       TypePoint(I) = PointType(I - 1)
                    Next
                 End With
                 Obj.NewPoint iCounter, tx, ty, PointType
                
              End If
              UnSelectAllObject
              Redraw
          End If
      ElseIf Obj.TypeDraw = dPolygon Then
          mnueditPoints_Click
      End If
End Sub

Private Sub picCanvas_KeyUp(KeyCode As Integer, Shift As Integer)
      Dim msg As String
      Dim X_min As Single
      Dim Y_min As Single
      If Obj Is Nothing Then Exit Sub
      
      Select Case KeyCode
      Case vbKeyLeft '37 LEFT ARROW key
          X_min = -1
          Y_min = 0
      Case vbKeyUp '38 UP ARROW key
          Y_min = -1
          X_min = 0
      Case vbKeyRight '39 RIGHT ARROW key
          X_min = 1
          Y_min = 0
      Case vbKeyDown '40
          Y_min = 1
          X_min = 0
      End Select

      TransformPoint X_min / m_ZoomFactor, Y_min / m_ZoomFactor
      
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IsSelect As Boolean
Dim R As RECT
    If Not (m_NewObject Is Nothing) Or m_DrawingObject Then Exit Sub
    ' See where we clicked.
    
    If Not Obj Is Nothing Then IsSelect = True
    
    If m_Rotate = False And m_Scale = False And m_Skew = False And m_Move = False Then
       Set Obj = FindObjectAt(X, Y)
    End If
    
    If (Obj Is Nothing) Then 'And m_SelectedObjects.Count <= 1 Then
        'Deselect all objects.
        DeselectAll
        m_Rotate = False
        m_Scale = False
        m_Move = False
        m_Skew = False
        LockObject = False
        If IsSelect = False Then Exit Sub

    Else
       If Button = 2 Then
           ViewMenu
          Exit Sub
       ElseIf Button = 1 Then
        'See if the Shift key is pressed.
        If (Shift And vbShiftMask) Then
            ' Shift is pressed. Toggle this object's selection.
            If Obj.Selected Then
                DeselectVbdObject Obj
                m_Rotate = False
                m_Scale = False
                m_Move = False
                ClearBox
            Else
                SelectVbdObject Obj
                'GoTo MD1
            End If
        Else
            If m_SelectedObjects.Count > 1 Then
               m_Move = True: mEditObject = True
                m_StartX = X
               m_StartY = Y
               m_LastX = X
               m_LastY = Y
               Exit Sub
            End If
            ' Shift is not pressed. Select only this object.
            DeselectAllVbdObjects
MD1:
            SelectVbdObject Obj
            
            LockObject = Obj.ObjLock
            
            If m_ReadPenProperty Then
               PicPenColor.BackColor = Obj.ForeColor
               ' Select the 1 pixel DrawWidth.
               icbDrawWidth.SelectedItem = icbDrawWidth.ComboItems(Obj.DrawWidth)
               icbDrawWidth.ToolTipText = icbDrawWidth.ComboItems(Obj.DrawWidth).Key
               ' Select the solid DrawStyle.
               icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(Obj.DrawStyle + 1)
               icbDrawStyle.ToolTipText = icbDrawStyle.ComboItems(Obj.DrawStyle + 1).Key
               m_ReadPenProperty = False
               SelectTool 1 '"Arrow"
            End If
            
            If m_ReadFillProperty Then
               CtlFill1.Color1 = Obj.FillColor
               CtlFill1.Color2 = Obj.FillColor2
               CtlFill1.FillStyle = Obj.FillStyle '+ 1
               CtlFill1.NamePattern = Obj.Pattern
               CtlFill1.TypeGradient = Obj.Gradient
               CtlFill1.Blend = Obj.Blend
               m_ReadFillProperty = False
               SelectTool 1 '"Arrow"
            End If
            
            DrawWidth = Obj.DrawWidth
            DrawStyle = Obj.DrawStyle
            FillStyle = Obj.FillStyle
            'Blend = Obj.Blend
            m_StartX = X
            m_StartY = Y
            m_LastX = X
            m_LastY = Y
              
            GetRgnBox Obj.hRegion, R
            XminBox = R.Left
            YminBox = R.Top
            XmaxBox = R.Right
            YmaxBox = R.Bottom
            If R.Left = 0 And R.Top = 0 And R.Right = 0 And R.Bottom = 0 Then
              Obj.Bound XminBox, YminBox, XmaxBox, YmaxBox
            End If
            
            Xmin_Box = XminBox - m_StartX
            Xmax_Box = XmaxBox - m_StartX
            Ymin_Box = YminBox - m_StartY
            Ymax_Box = YmaxBox - m_StartY
            
            If m_Scale = False And m_Rotate = False And m_Skew = False Then
               m_Move = True:
               mEditObject = True
            End If
            If m_Move Then
               PicCanvas.Line (m_StartX + Xmin_Box, m_StartY + Ymin_Box)-(m_StartX + Xmax_Box, m_StartY + Ymax_Box), , B
            End If
        End If
       End If
    End If
    
    If Not Obj Is Nothing Then
         RaiseEvent ColorSelected(1, Obj.FillColor)
         RaiseEvent ColorSelected(2, Obj.ForeColor)
    End If
    
    Redraw
        
    ' See if any objects are selected.
    RaiseEvent EnableMenusForSelection
    
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ang As Single
Dim msg As String
'Const GAP = 6
       
    RaiseEvent MouseMove(Button, Shift, X / m_ZoomFactor, Y / m_ZoomFactor)

    If Not (m_NewObject Is Nothing) Or m_DrawingObject Then Exit Sub
      
    If Not (Obj Is Nothing) Then
       If Obj.ObjLock = True Then GoTo NoSelect:
    End If
    
    'move point
    If Not (Obj Is Nothing) And Button = 1 And m_Move = True And mEditObject Then
mm1:
          PicCanvas.DrawMode = vbInvert
          PicCanvas.DrawStyle = vbDot
          
          If m_Move Then
             PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
             m_LastX = X
             m_LastY = Y
             PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
             msg = "Move X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
             RaiseEvent MsgControl(msg)
             Exit Sub
          End If
    
    'Scale point
    ElseIf Not (Obj Is Nothing) And Button = 1 And (m_Scale = True Or m_Skew = True) And mEditObject Then
          PicCanvas.DrawMode = vbInvert
          PicCanvas.DrawStyle = vbDot
         
          Select Case m_ScaleType
          Case 1 'Left top Corner
                  PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YmaxBox), , B
          Case 2 'Middle top
               If m_Scale Then
                  PicCanvas.Line (XminBox, m_LastY)-(XmaxBox, YmaxBox), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew (100 + ((m_LastX - m_StartX) * 100) / (XmaxBox - XminBox)), 100
               End If
          Case 3 'Right top Corner
                  PicCanvas.Line (XminBox, YmaxBox)-(m_LastX, m_LastY), , B
          Case 4 'Middle Right
               If m_Scale Then
                  PicCanvas.Line (XminBox, YminBox)-(m_LastX, YmaxBox), , B
               Else
                  If YmaxBox - YminBox <= 0 Then Exit Sub
                  mDrawSkew 100, (100 + ((m_StartY - m_LastY) * 100) / (YmaxBox - YminBox))
               End If
          Case 5 'Bottom Right corner
                  PicCanvas.Line (XminBox, YminBox)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
          Case 6 'Middle Bottom
               If m_Scale Then
                  PicCanvas.Line (XminBox, YminBox)-(XmaxBox, m_LastY), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew (100 + ((m_StartX - m_LastX) * 100) / (XmaxBox - XminBox)), 100
               End If
          Case 7 'Left bottom corner
                  PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YminBox), , B
          Case 8 'Middle left
               If m_Scale Then
                  PicCanvas.Line (m_LastX, YminBox)-(XmaxBox, YmaxBox), , B
               Else
                  If XmaxBox - XminBox <= 0 Then Exit Sub
                  mDrawSkew 100, (100 + ((m_LastY - m_StartY) * 100) / (YmaxBox - YminBox))
               End If
          End Select
            m_LastX = X
            m_LastY = Y
          Select Case m_ScaleType
            Case 1 'Left top Corner
                   PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YmaxBox), , B
            Case 2 'Middle top
                If m_Scale Then
                   PicCanvas.Line (XminBox, m_LastY)-(XmaxBox, YmaxBox), , B
                Else
                   If XmaxBox - XminBox <= 0 Then Exit Sub
                   mDrawSkew (100 + ((m_LastX - m_StartX) * 100) / (XmaxBox - XminBox)), 100
                End If
            Case 3 'Right top Corner
                   PicCanvas.Line (XminBox, YmaxBox)-(m_LastX, m_LastY), , B
            Case 4 'Middle Right
                If m_Scale Then
                   PicCanvas.Line (XminBox, YminBox)-(m_LastX, YmaxBox), , B
                Else
                   If YmaxBox - YminBox <= 0 Then Exit Sub
                   mDrawSkew 100, (100 + ((m_StartY - m_LastY) * 100) / (YmaxBox - YminBox))
                End If
            Case 5 'Bottom Right corner
                   PicCanvas.Line (XminBox, YminBox)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
            Case 6 'Middle Bottom
                If m_Scale Then
                   PicCanvas.Line (XminBox, YminBox)-(XmaxBox, m_LastY), , B
                Else
                   If XmaxBox - XminBox <= 0 Then Exit Sub
                   mDrawSkew (100 + ((m_StartX - m_LastX) * 100) / (XmaxBox - XminBox)), 100
                End If
            Case 7 'Left bottom corner
                   PicCanvas.Line (m_LastX, m_LastY)-(XmaxBox, YminBox), , B
            Case 8 'Middle left
                If m_Scale Then
                   PicCanvas.Line (m_LastX, YminBox)-(XmaxBox, YmaxBox), , B
                Else
                   If YmaxBox - YminBox <= 0 Then Exit Sub
                   mDrawSkew 100, (100 + ((m_LastY - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
            End Select
            If m_Scale Then
               msg = "Scale X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
            ElseIf m_Skew Then
               msg = "Skew X:" + Format(m_StartX, "0.0") + " Y:" + Format(m_StartY, "0.0") + " DX:" + Format(m_LastX, "0.0") + " DY:" + Format(m_LastY, "0.0")
            End If
            RaiseEvent MsgControl(msg)
            Exit Sub

     'Rotate point
     ElseIf Not (Obj Is Nothing) And Button = 1 And m_Rotate = True And mEditObject Then
            PicCanvas.DrawMode = vbInvert
            PicCanvas.DrawStyle = vbDot
            ' Create the rotation transformation.
            mDrawRotate m_LastX, m_LastY
            m_LastX = X
            m_LastY = Y
            mDrawRotate m_LastX, m_LastY
               
     'Change state and mousepointer
     ElseIf Not (Obj Is Nothing) And mEditObject = True Then
        
         'Move object
        If XminBox + ((XmaxBox - XminBox) / 2) - GAP <= X And XminBox + ((XmaxBox - XminBox) / 2) + GAP >= X And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            PicCanvas.MousePointer = 99
            PicCanvas.MouseIcon = ImageMouse(19).Picture
            GoSub State_Move
            Exit Sub
        End If
        
        'Point rotate
        If (XmaxBox + 18 <= X And XmaxBox + 22 >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            PicCanvas.MousePointer = 99
            PicCanvas.MouseIcon = ImageMouse(12).Picture
            GoTo State_Rotate
            Exit Sub
        End If
        
        'If Obj.TypeDraw = dText Then GoTo NoSelect: Exit Sub
        
        'Left top Corner
        If (XminBox + GAP >= X And XminBox - GAPn <= X) And _
           (YminBox + GAP >= Y And YminBox - GAP <= Y) Then
           'picCanvas.MousePointer = 8
           m_ScaleType = 1
           m_Skew = False
           GoSub State_Scale
           If m_Scale Then
              PicCanvas.MousePointer = 8
            End If
            Exit Sub
        End If
        
        'Middle top
        If ((XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) >= X And _
            ((XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) - GAP) <= X) And _
           (YminBox >= Y And YminBox - GAP <= Y) Then
            m_ScaleType = 2
            m_Skew = True
             GoSub State_Scale
              PicCanvas.MousePointer = 99
              PicCanvas.MouseIcon = ImageMouse(13).Picture
            Exit Sub
        End If
                
       'Right top corner
        If (XmaxBox - GAP <= X And XmaxBox + GAP >= X) And _
           (YminBox - GAP <= Y And YminBox + GAP >= Y) Then
             m_ScaleType = 3
             m_Skew = False
            GoSub State_Scale
            If m_Scale Then
               PicCanvas.MousePointer = 6
            End If
            Exit Sub
        End If
                
        'Middle right
        If (XmaxBox - GAP <= X And XmaxBox + GAP >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
             m_ScaleType = 4
             m_Skew = True
            GoSub State_Scale
              PicCanvas.MousePointer = 99
              PicCanvas.MouseIcon = ImageMouse(14).Picture
            Exit Sub
        End If
        
        'Right botton corner
        If (XmaxBox <= X And X <= XmaxBox + GAP) And _
           (YmaxBox <= Y And Y <= YmaxBox + GAP) Then
             m_ScaleType = 5
             m_Skew = False
            GoSub State_Scale
            If m_Scale Then
               PicCanvas.MousePointer = 8
            End If
            Exit Sub
        End If
        
        'Middle botton
        If (XminBox + (XmaxBox - XminBox) / 2 + GAP / 2 >= X And _
            X >= (XminBox + (XmaxBox - XminBox) / 2 + GAP / 2) - GAP) And _
           (YmaxBox <= Y And YmaxBox + GAP >= Y) Then
            m_ScaleType = 6
            m_Skew = True
           GoSub State_Scale
             PicCanvas.MousePointer = 99
             PicCanvas.MouseIcon = ImageMouse(13).Picture
            Exit Sub
        End If
        
        'Botton left corner
        If (XminBox - GAP <= X And XminBox + GAP >= X) And _
           (YmaxBox - GAP <= Y And YmaxBox + GAP >= Y) Then
            m_ScaleType = 7
            m_Skew = False
            GoSub State_Scale
               PicCanvas.MousePointer = 6
            Exit Sub
        End If
                
        'Middle left
        If (XminBox - GAP <= X And XminBox + GAP >= X) And _
           ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) >= Y And _
           Y >= ((YminBox + (YmaxBox - YminBox) / 2 + GAP / 2) - GAP)) Then
            m_Skew = True
            m_ScaleType = 8
            GoSub State_Scale
              PicCanvas.MousePointer = 99
              PicCanvas.MouseIcon = ImageMouse(14).Picture
            Exit Sub
        End If
                
        
        If m_ReadPenProperty = False And m_ReadFillProperty = False And Button = 0 Then
NoSelect:
             m_Scale = False
             m_Rotate = False
             m_Skew = False
             m_Move = False
             PicCanvas.MousePointer = 99
             PicCanvas.MouseIcon = ImageMouse(0).Picture
        End If
        
     'If not select state is move
     ElseIf Not (Obj Is Nothing) And mEditObject Then

         m_Scale = False
         m_Rotate = False
         m_Skew = False
         m_Move = True
     End If
    
    Exit Sub
    
State_Scale:
   If m_Skew = False And m_Scale = False Then m_Scale = True
   If m_Skew = True Then
      m_Scale = False
      m_Skew = True
    Else
      m_Scale = True
      m_Skew = False
    End If
    m_Move = False
    m_Rotate = False
    mEditObject = True
    Return
Exit Sub

State_Rotate:
    m_Scale = False
    m_Skew = False
    m_Move = False
    m_Rotate = True
    mEditObject = True
Exit Sub

State_Move:
    m_Scale = False
    m_Skew = False
    m_Move = True
    m_Rotate = False
    mEditObject = True
    Return
Exit Sub

End Sub


Private Sub PicCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim Ang As Single
Dim msg As String
Dim R As RECT
          
     If Not (m_NewObject Is Nothing) Or m_DrawingObject Then Exit Sub
     
     If m_SelectedObjects.Count > 1 And m_Move Then
          m_LastX = X
          m_LastY = Y
        GoTo MU1:
     End If
     
     If Not (Obj Is Nothing) Then
     
MU1:
      If Obj.ObjLock = True Then GoTo NoSelect
       PicCanvas.DrawMode = vbCopyPen
       
       If m_Move And Button = 1 Then
          If (X - m_StartX) / m_ZoomFactor = 0 And (Y - m_StartY) / m_ZoomFactor = 0 Then Exit Sub
             PicCanvas.Line (m_LastX + Xmin_Box, m_LastY + Ymin_Box)-(m_LastX + Xmax_Box, m_LastY + Ymax_Box), , B
             TransformPoint (X - m_StartX) / m_ZoomFactor, (Y - m_StartY) / m_ZoomFactor
             msg = "Move X:" + Format((X - m_StartX) / m_ZoomFactor, "0.0") + _
                       " Y:" + Format((Y - m_StartY) / m_ZoomFactor, "0.0")
             RaiseEvent MsgControl(msg)
 
       ElseIf m_Rotate And Button = 1 Then
            m_LastX = X / m_ZoomFactor
            m_LastY = Y / m_ZoomFactor
            Ang = m2GetAngle3P(XminBox + (XmaxBox - XminBox) / 2, YminBox + (YmaxBox - YminBox) / 2, _
                               XminBox + (XmaxBox - XminBox), m_StartY, _
                               m_LastX, m_LastY)
            TransformRotate Ang, XminBox / m_ZoomFactor, YminBox / m_ZoomFactor, XmaxBox / m_ZoomFactor, YmaxBox / m_ZoomFactor
            msg = "Rotate angle:" + Format(Ang, "0.0")
            RaiseEvent MsgControl(msg)
             
       ElseIf m_Scale Or m_Skew And Button = 1 Then
           Select Case m_ScaleType
           Case 1 'Left top Corner
              If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                     TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                    (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 End If
              End If
           Case 2 'Middle top
              If (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                    TransformScale 100, _
                                   (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 Else
                    TransformSkew (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  100
                 End If
              End If
           Case 3 'Right top Corner
              If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                 If m_Scale Then
                    TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                   (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                 End If
              End If
           Case 4 'Middle Right
             If (XmaxBox - XminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  100
                Else
                   TransformSkew 100, _
                                 (100 + ((m_StartY - Y) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 5 'Bottom Right corner
             If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((X - m_StartX) * 100) / (XmaxBox - XminBox)), _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 6 'Middle Bottom
             If (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale 100, _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                Else
                   TransformSkew (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                   100
                End If
             End If
           Case 7 'Left bottom corner
             If (XmaxBox - XminBox) <> 0 And (YmaxBox - YminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                  (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           Case 8 'Middle left
             If (XmaxBox - XminBox) <> 0 Then
                If m_Scale Then
                   TransformScale (100 + ((m_StartX - X) * 100) / (XmaxBox - XminBox)), _
                                   100
                Else
                   TransformSkew 100, _
                                 (100 + ((Y - m_StartY) * 100) / (YmaxBox - YminBox))
                End If
             End If
           End Select
       End If
       Redraw
        
     End If
    
     SelectTool 1 '"Arrow"
NoSelect:
     m_Move = False
     m_Rotate = False
     m_Scale = False
     m_Skew = False
     If Not Obj Is Nothing Then
         GetRgnBox Obj.hRegion, R
         XminBox = R.Left
         YminBox = R.Top
         XmaxBox = R.Right
         YmaxBox = R.Bottom
     End If
         
     RaiseEvent MouseUp(Button, Shift, X, Y)
     
End Sub

Private Sub ClearBox()
    XminBox = 0
    XmaxBox = 0
    YminBox = 0
    YmaxBox = 0
End Sub

Private Sub picCanvas_Paint()
   Dim mDC As Long, tmpBmp As Long
     If m_TheScene Is Nothing Then Exit Sub
     PicCanvas.Cls
    PicCanvas.DrawMode = 13
    LockWindowUpdate UserControl.hwnd
     m_TheScene.Draw PicCanvas
    LockWindowUpdate False
     idredraw = idredraw + 1
     Debug.Print idredraw
End Sub

Sub ViewMenu()
    If Obj.TypeDraw = dText Or Obj.TypeDraw = dEllipse _
       Or Obj.TypeDraw = dFreePolygon Or Obj.TypeDraw = dPolygon _
       Or Obj.TypeDraw = dPolyline Or Obj.TypeDraw = dRectAngle _
       Or Obj.TypeDraw = dTextFrame Then
       sep1.Visible = True
       mnuCurve.Visible = True
       mnuFillMode.Enabled = False
    Else
       mnuCurve.Visible = False
       mnuFillMode.Enabled = True
    End If
    
    If Obj.TypeDraw = dPolygon Then
       sepEditPoints.Visible = True
       mnueditPoints.Visible = True
    Else
       sepEditPoints.Visible = False
       mnueditPoints.Visible = False
    End If
    
    PopupMenu mnutransform
End Sub

Private Sub PicPenColor_DblClick()
    cmdSysColorsPen_Click
End Sub



Private Sub UserControl_GotFocus()
     ' PicCanvas.SetFocus
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDelete
             ObjectDelete
    End Select
End Sub

Private Sub UserControl_Paint()
     Debug.Print "UserControl_Paint"
End Sub

Private Sub UserControl_Resize()
    Dim tW As Long
    Dim tH As Long
    Dim nW As Long
    Dim pcW As Long, pcH As Long, pcL As Long, pcT As Long
    
    PicCanvas.Visible = False
    
    HScroll1.Move 0, UserControl.ScaleHeight - HScroll1.Height, UserControl.ScaleWidth - ComCorner.Width
    VScroll1.Move UserControl.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, UserControl.ScaleHeight - ComCorner.Height
    ComCorner.Move VScroll1.Left, UserControl.ScaleHeight - ComCorner.Height, VScroll1.Width
    
    PicCanvas.Width = m_CanvasWidth * m_ZoomFactor
    PicCanvas.Height = m_CanvasHeight * m_ZoomFactor
    
    tW = UserControl.ScaleX(UserControl.Width, vbTwips, vbPixels) - 4 '- nW
    tH = UserControl.ScaleY(UserControl.Height, vbTwips, vbPixels) - 4
    
    gZoomFactor = m_ZoomFactor
     
    PicCanvas.Left = (tW / 2 - (CanvasWidth * m_ZoomFactor / 2))
    PicCanvas.Top = (tH / 2 - (CanvasHeight * m_ZoomFactor / 2))
    
    m_CanvasLeft = (tW / 2 - (CanvasWidth * m_ZoomFactor / 2)) '* m_ZoomFactor
    m_CanvasTop = (tH / 2 - (CanvasHeight * m_ZoomFactor / 2)) '* m_ZoomFactor
    
    PicCanvas.Move m_CanvasLeft, m_CanvasTop ', m_CanvasWidth, m_CanvasHeight
    
    pcL = m_CanvasLeft * m_ZoomFactor
    pcT = m_CanvasTop * m_ZoomFactor
    pcW = m_CanvasWidth * m_ZoomFactor
    pcH = m_CanvasHeight * m_ZoomFactor
    
    HScroll1.Visible = False
    VScroll1.Visible = False
    ComCorner.Visible = False

    If pcW + 18 > tW Then
        HScroll1.Left = 0
        HScroll1.Top = tH - 16
        HScroll1.Width = tW - 16
        HScroll1.Visible = True
        HScroll1.Max = -(pcW - tW) - 40
        HScroll1.Min = 20
        If toZoom = False Then
            HScroll1_Change
            HScroll1.Value = 20
        End If
    End If

    If pcH + 18 > tH Then
        VScroll1.Left = tW - 16
        VScroll1.Top = 0
        VScroll1.Height = tH
        VScroll1.Visible = True
        VScroll1.Max = -(pcH - tH) - 40
        VScroll1.Min = 20
        If toZoom = False Then
            VScroll1_Change
            VScroll1.Value = 20
        End If
    End If

    toZoom = False

    If pcW + 18 > tW And pcH + 18 > tH Then
        HScroll1.Width = tW - 16
        VScroll1.Height = tH - 16
        ComCorner.Left = tW - 16
        ComCorner.Top = tH - 16
        ComCorner.Visible = True
    Else
        ReDrawPage
    End If
     
    If m_ShowPenProperty = True Then MeForm1.Visible = True
    If m_ShowFillProperty = True Then MeForm2.Visible = True
    If m_ShowTranformProperty = True Then MeForm3.Visible = True
    PicCanvas.Visible = True
    Redraw
    
End Sub

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = PicCanvas.Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
End Property

Public Property Get ZoomFactor() As Single
    ZoomFactor = m_ZoomFactor
End Property

Public Property Let ZoomFactor(ByVal New_ZoomFactor As Single)
    If New_ZoomFactor < 0.05 Then Exit Property
    If New_ZoomFactor > 4 Then Exit Property
    m_ZoomFactor = New_ZoomFactor
    gZoomFactor = m_ZoomFactor
    
    PropertyChanged "ZoomFactor"
    
    NewTransformation
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Image = LoadPicture("")
    m_ZoomFactor = m_def_ZoomFactor
    m_CanvasWidth = m_def_CanvasWidth
    m_CanvasHeight = m_def_CanvasHeight
    m_ShowCanvasSize = m_def_ShowCanvasSize
    m_ForeColor = m_def_ForeColor
    m_DrawWidth = m_def_DrawWidth
    m_DrawStyle = m_def_DrawStyle
    m_FillStyle = m_def_FillStyle
    m_FillColor = m_def_FillColor
    m_FileName = m_def_FileName
    m_FileTitle = m_def_FileTitle
    m_ShowPenProperty = m_def_ShowPenProperty
    m_ShowFillProperty = m_def_ShowFillProperty
    m_ShowTranformProperty = m_def_ShowFillProperty
    m_FillColor2 = m_def_FillColor2
    m_Pattern = m_def_Pattern
    m_TypeGradient = m_def_TypeGradient
    m_Blend = m_def_Blend
    m_BackImage = m_def_BackImage
    m_LockObject = m_def_LockObject
    NewDraw
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    PicCanvas.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    m_ZoomFactor = PropBag.ReadProperty("ZoomFactor", m_def_ZoomFactor)
    m_CanvasWidth = PropBag.ReadProperty("CanvasWidth", m_def_CanvasWidth)
    m_CanvasHeight = PropBag.ReadProperty("CanvasHeight", m_def_CanvasHeight)
    m_ShowCanvasSize = PropBag.ReadProperty("ShowCanvasSize", m_def_ShowCanvasSize)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_DrawWidth = PropBag.ReadProperty("DrawWidth", m_def_DrawWidth)
    m_DrawStyle = PropBag.ReadProperty("DrawStyle", m_def_DrawStyle)
    m_FillStyle = PropBag.ReadProperty("FillStyle", m_def_FillStyle)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
       
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    m_FileTitle = PropBag.ReadProperty("FileTitle", m_def_FileTitle)
    m_ShowPenProperty = PropBag.ReadProperty("ShowPenProperty", m_def_ShowPenProperty)
    m_ShowFillProperty = PropBag.ReadProperty("ShowFillProperty", m_def_ShowFillProperty)
    m_ShowTranformProperty = PropBag.ReadProperty("ShowTranformProperty", m_def_ShowTranformProperty)
    m_FillColor2 = PropBag.ReadProperty("FillColor2", m_def_FillColor2)
    m_Pattern = PropBag.ReadProperty("Pattern", m_def_Pattern)
    m_TypeGradient = PropBag.ReadProperty("Gradient", m_def_TypeGradient)
    m_Blend = PropBag.ReadProperty("Blend", m_def_Blend)
    m_BackImage = PropBag.ReadProperty("BackImage", m_def_BackImage)
    m_LockObject = PropBag.ReadProperty("LockObject", m_def_LockObject)
    
    DrawPen
    NewDraw
    MeForm1.Caption = "Pen"
    MeForm1.Alignment = 0
    MeForm2.Caption = "Fill"
    MeForm2.Alignment = 0
    MeForm3.Caption = "Transform"
    MeForm3.Alignment = 0
    MeForm1.BackColor = &HF1E2DC
    MeForm2.BackColor = &HF1E2DC
    MeForm3.BackColor = &HF1E2DC
    CtlFill1.BackColor = &HF1E2DC
    CtrTranform1.BackColor = &HF1E2DC
    MeForm1.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, 0
    MeForm1.Height = 190 'UserControl.ScaleHeight / 3
    MeForm2.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height
    MeForm2.Height = 210 'UserControl.ScaleHeight / 3
    MeForm3.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height + MeForm2.Height
    MeForm3.Height = 165 'UserControl.ScaleHeight / 3
    
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", PicCanvas.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("ZoomFactor", m_ZoomFactor, m_def_ZoomFactor)
    Call PropBag.WriteProperty("CanvasWidth", m_CanvasWidth, m_def_CanvasWidth)
    Call PropBag.WriteProperty("CanvasHeight", m_CanvasHeight, m_def_CanvasHeight)
    Call PropBag.WriteProperty("ShowCanvasSize", m_ShowCanvasSize, m_def_ShowCanvasSize)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("DrawWidth", m_DrawWidth, m_def_DrawWidth)
    Call PropBag.WriteProperty("DrawStyle", m_DrawStyle, m_def_DrawStyle)
    Call PropBag.WriteProperty("FillStyle", m_FillStyle, m_def_FillStyle)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Call PropBag.WriteProperty("FileTitle", m_FileTitle, m_def_FileTitle)
    Call PropBag.WriteProperty("ShowPenProperty", m_ShowPenProperty, m_def_ShowPenProperty)
    Call PropBag.WriteProperty("ShowFillProperty", m_ShowFillProperty, m_def_ShowFillProperty)
    Call PropBag.WriteProperty("ShowTranformProperty", m_ShowTranformProperty, m_def_ShowTranformProperty)
    Call PropBag.WriteProperty("FillColor2", m_FillColor2, m_def_FillColor2)
    Call PropBag.WriteProperty("Pattern", m_Pattern, m_def_Pattern)
    Call PropBag.WriteProperty("Gradient", m_TypeGradient, m_def_TypeGradient)
    Call PropBag.WriteProperty("Blend", m_Blend, m_def_Blend)
    Call PropBag.WriteProperty("BackImage", m_BackImage, m_def_BackImage)
    Call PropBag.WriteProperty("LockObject", m_LockObject, m_def_LockObject)
End Sub

Public Property Get CanvasWidth() As Long
    CanvasWidth = m_CanvasWidth
End Property

Public Property Let CanvasWidth(ByVal New_CanvasWidth As Long)
    m_CanvasWidth = New_CanvasWidth
    PropertyChanged "CanvasWidth"
End Property

Public Property Get CanvasHeight() As Long
    CanvasHeight = m_CanvasHeight
End Property

Public Property Let CanvasHeight(ByVal New_CanvasHeight As Long)
    m_CanvasHeight = New_CanvasHeight
    PropertyChanged "CanvasHeight"
End Property

Private Sub VScroll1_Change()
    On Error Resume Next
    m_CanvasTop = VScroll1.Value
    PicCanvas.Visible = False
    PicCanvas.Top = m_CanvasTop
    ReDrawPage
    PicCanvas.Visible = True
End Sub

Public Property Get ShowCanvasSize() As Boolean
    ShowCanvasSize = m_ShowCanvasSize
End Property

Public Property Let ShowCanvasSize(ByVal New_ShowCanvasSize As Boolean)
    m_ShowCanvasSize = New_ShowCanvasSize
    PropertyChanged "ShowCanvasSize"
End Property

Public Sub Redraw()
    picCanvas_Paint
End Sub

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    ChangeForeColor m_ForeColor
    RaiseEvent ColorSelected(2, m_ForeColor)
End Property

Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
    DrawWidth = m_DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    m_DrawWidth = New_DrawWidth
    PropertyChanged "DrawWidth"
    ChangeDrawWidth m_DrawWidth
End Property

Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = m_DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    m_DrawStyle = New_DrawStyle
    PropertyChanged "DrawStyle"
    ChangeDrawstyle m_DrawStyle
End Property

Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = m_FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    m_FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
    ChangeFillstyle m_FillStyle
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal new_FillColor As OLE_COLOR)
    m_FillColor = new_FillColor
    PropertyChanged "FillColor"
    ChangeFillColor 1, m_FillColor
    RaiseEvent ColorSelected(1, m_FillColor)
End Property

Public Sub SelectTool(ByVal Key As Integer)
    Dim msg As String
    Dim new_pgon As vbdPolygon
    Dim new_plgon As VbPolygon
    Dim new_line As vbdLine
    Dim new_text As VbText
    Dim new_Scribble As vbdScribble
    Dim new_Ellipse As vbdEllipse
    Dim new_curve As vbdCurve
    
    Dim nPoint As Long, X() As Single, Y() As Single, TPoint() As Byte
    ' Free any previously started object.
    Set m_NewObject = Nothing
        
    m_DrawingObject = True
    If Key <> 1 Then
        DeselectAll
    End If
    ' Create the new object.
   ' m_ToolKey = Key
    Select Case Key
        Case 1 '"Arrow"
            PicCanvas.MouseIcon = ImageMouse(0).Picture
            PicCanvas.MousePointer = 99
            msg = "Select object"
            m_DrawingObject = False
            
        Case 2 '"Point"
            ' Let the new object receive picCanvas events.
            If Obj Is Nothing Then Exit Sub
            If Obj.ObjLock = True Then Exit Sub
            If Obj.TypeDraw = dPolygon Or _
               Obj.TypeDraw = dRectAngle Or _
               Obj.TypeDraw = dEllipse Then
               Exit Sub
            End If
            If Obj Is Nothing Then Exit Sub
            gZoomLock = True
            Set m_EditObject = Obj
            SelectVbdObject Obj
            Set m_EditObject.canvas = PicCanvas
            m_EditObject.DrawStyle = Obj.DrawStyle
            m_EditObject.DrawWidth = Obj.DrawWidth
            m_EditObject.FillColor = Obj.FillColor
            m_EditObject.FillColor2 = Obj.FillColor2
            m_EditObject.FillMode = Obj.FillMode
            m_EditObject.FillStyle = Obj.FillStyle
            m_EditObject.ForeColor = Obj.ForeColor
            m_EditObject.Gradient = Obj.Gradient
            m_EditObject.Pattern = Obj.Pattern
            m_EditObject.Blend = Obj.Blend
            Set m_EditObject.Picture = Obj.Picture
            m_EditObject.Shade = Obj.Shade
            m_EditObject.TypeFill = Obj.TypeFill

            Obj.ReadTrPoint nPoint, X(), Y(), TPoint
            DeletevbdObject
            m_EditObject.NewTrPoint nPoint, X, Y, TPoint
            m_EditObject.DrawPoint
            m_DrawingObject = True
            Do Until m_EditObject.canvas Is Nothing
               DoEvents
            Loop
            gZoomLock = False
            Redraw

        Case 3 '"Polyline"
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = False
            new_pgon.m_DrawStyle = 0
            new_pgon.m_DrawWidth = 1
            new_pgon.m_FillColor = RGB(255, 255, 255)
            new_pgon.m_FillStyle = 1
            new_pgon.m_ForeColor = RGB(0, 0, 0)
            new_pgon.m_TypeDraw = dPolyline
            new_pgon.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(1).Picture
            msg = "Draw Line or Polyline"
            
        Case 4 '"FreePolygon"
            Set m_NewObject = New vbdPolygon
            Set new_pgon = m_NewObject
            new_pgon.IsClosed = True
            new_pgon.m_DrawStyle = 0
            new_pgon.m_DrawWidth = 1
            new_pgon.m_FillColor = RGB(255, 255, 255)
            new_pgon.m_FillStyle = 1
            new_pgon.m_ForeColor = RGB(0, 0, 0)
            new_pgon.m_TypeDraw = dFreePolygon
            new_pgon.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(3).Picture
            msg = "Draw free Polygon"
        
        Case 5 'Free line "Scribble"
            Set m_NewObject = New vbdScribble
            Set new_Scribble = m_NewObject
            new_Scribble.m_DrawStyle = 0
            new_Scribble.m_DrawWidth = 1
            new_Scribble.m_FillColor = RGB(255, 255, 255)
            new_Scribble.m_FillStyle = 1
            new_Scribble.m_ForeColor = RGB(0, 0, 0)
            new_Scribble.m_TypeDraw = dScribble
            new_Scribble.IsClosed = False
            new_Scribble.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(2).Picture
            msg = "Draw free line"
        Case 6 ' Free line close "Scribble"
            Set m_NewObject = New vbdScribble
            Set new_Scribble = m_NewObject
            new_Scribble.m_DrawStyle = 0
            new_Scribble.m_DrawWidth = 1
            new_Scribble.m_FillColor = RGB(255, 255, 255)
            new_Scribble.m_FillStyle = 1
            new_Scribble.m_ForeColor = RGB(0, 0, 0)
            new_Scribble.m_TypeDraw = dScribble
            new_Scribble.IsClosed = True
            new_Scribble.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(2).Picture
            msg = "Draw free line close"
              
        Case 7 '"Curve"
            Set m_NewObject = New vbdCurve
            Set new_curve = m_NewObject
            new_curve.IsClosed = False
            new_curve.m_DrawStyle = 0
            new_curve.m_DrawWidth = 1
            new_curve.m_FillColor = RGB(255, 255, 255)
            new_curve.m_FillStyle = 1
            new_curve.m_ForeColor = RGB(0, 0, 0)
            new_curve.m_TypeDraw = dCurve
            new_curve.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(6).Picture
            msg = "Select 4 point to draw curve"
            
        Case 8 '"RectAngle"
            Set m_NewObject = New vbdLine
            Set new_line = m_NewObject
            new_line.IsBox = True
            new_line.m_DrawStyle = 0
            new_line.m_DrawWidth = 1
            new_line.m_FillColor = RGB(255, 255, 255)
            new_line.m_FillStyle = 1
            new_line.m_ForeColor = RGB(0, 0, 0)
            new_line.m_TypeDraw = dRectAngle
            new_line.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(5).Picture
            msg = "Press and Hold ''Ctrl'' Button to make a Cube"
        Case 9 '"Polygon"
            Set m_NewObject = New VbPolygon
            Set new_plgon = m_NewObject
            new_plgon.IsBox = True
            new_plgon.m_DrawStyle = 0
            new_plgon.m_DrawWidth = 1
            new_plgon.m_FillColor = RGB(255, 255, 255)
            new_plgon.m_FillStyle = 1
            new_plgon.m_ForeColor = RGB(0, 0, 0)
            new_plgon.m_TypeDraw = dPolygon
            new_plgon.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(4).Picture
            msg = "Press and Hold ''Ctrl'' Button to make a Polygon"
        
        Case 10 '"Ellipse"
            Set m_NewObject = New vbdEllipse
            Set new_Ellipse = m_NewObject
            new_Ellipse.m_DrawStyle = 0
            new_Ellipse.m_DrawWidth = 1
            new_Ellipse.m_FillColor = RGB(255, 255, 255)
            new_Ellipse.m_FillStyle = 1
            new_Ellipse.m_ForeColor = RGB(0, 0, 0)
            new_Ellipse.m_TypeDraw = dEllipse
            new_Ellipse.m_Shade = False
            
            PicCanvas.MouseIcon = ImageMouse(6).Picture
            msg = "Press and Hold ''Ctrl'' Button to make a Circle"
            
        Case 11 '"Text"
            Set m_NewObject = New VbText
            Set new_text = m_NewObject
            new_text.IsBox = True
            new_text.m_DrawStyle = 0
            new_text.m_DrawWidth = 1
            new_text.m_FillColor = RGB(255, 255, 255)
            new_text.m_FillStyle = 1
            new_text.m_ForeColor = RGB(0, 0, 0)
            new_text.m_TypeDraw = dText
            new_text.m_Shade = False
            PicCanvas.MouseIcon = ImageMouse(7).Picture
            msg = "Select position for text"
'        Case 11 '"TextArt"
'            Set m_NewObject = New VbText
'            Set new_text = m_NewObject
'            new_text.IsBox = True
'            new_text.m_DrawStyle = 0
'            new_text.m_DrawWidth = 1
'            new_text.m_FillColor = RGB(255, 255, 255)
'            new_text.m_FillStyle = 1
'            new_text.m_ForeColor = RGB(0, 0, 0)
'            new_text.m_TypeDraw = dTextFrame
'            new_text.m_Shade = False
'            PicCanvas.MouseIcon = ImageMouse(7).Picture
'            msg = "Select position for text"
            
        Case 12 '"Pen"
             ShowPenProperty = True
             
        Case 13 '"Fill"
             ShowFillProperty = True
        
         Case 20 '"DropperPen", "DropperFill"
            PicCanvas.MouseIcon = ImageMouse(11).Picture
            msg = "Select object"
            m_DrawingObject = False
    End Select
        
    ' Let the new object receive picCanvas events.
    If Not (m_NewObject Is Nothing) Then
        Set m_NewObject.canvas = PicCanvas
    End If
    
    RaiseEvent MsgControl(msg)
    
End Sub

Sub ChoiceColorForControl()
    RaiseEvent ColorSelected(1, m_FillColor)
    RaiseEvent ColorSelected(2, m_ForeColor)
End Sub

' Move this object to the front,back,forward,Backward of the scene's
' object list.
Public Function SetObjectOrder(mOrder As m_Order)
   Dim the_scene As vbdScene
        
   Set the_scene = m_TheScene
   
   Select Case mOrder
   Case BringToFront
       the_scene.MoveToFront m_SelectedObjects
   Case SendToBack
        the_scene.MoveToBack m_SelectedObjects
   Case BringFoward
       the_scene.MoveToFoward m_SelectedObjects
   Case SendBackward
       the_scene.MoveToBackward m_SelectedObjects
   End Select
   
   Set_Dirty
   
   Redraw
   
End Function

Public Sub SelectAllObject()
    Dim the_scene As vbdScene
     
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.SelectAllObject
    
    Redraw
    Set the_scene = Nothing
End Sub

Public Sub UnSelectAllObject()
    Dim the_scene As vbdScene
     
    ' Save the new object.
    Set the_scene = m_TheScene
    the_scene.DeselectAllObject
    
    Redraw
    Set the_scene = Nothing
End Sub

Public Sub NewDraw(Optional SelectPage As Boolean = False)
    Dim cW As Single, cH As Single, cImage As String, cColor As OLE_COLOR
    
    ' Create a new, empty scene object.
    Set m_TheScene = New vbdScene

    ' No objects are selected.
    Set m_SelectedObjects = New Collection
    PicCanvas.Cls
    PicCanvas.Line (0, 0)-(PicCanvas.ScaleWidth, PicCanvas.ScaleHeight), vbWhite, BF
    PrepareToEdit
    If SelectPage Then
        cW = CanvasWidth: cH = CanvasHeight
        cColor = BackColor
        cImage = BackImage
        FrmCanvas.ShowForm cW, cH, cImage, cColor
        CanvasHeight = cH
        CanvasWidth = cW
        LockObject = False
        BackImage = cImage
        PicCanvas.BackColor = cColor
        Unload FrmCanvas
        UserControl_Resize
    End If
    Filename = ""
End Sub

' Select default values and prepare to edit.
Public Sub PrepareToEdit()
    
    ' Start at normal (pixel) scale.
    PicCanvas.ScaleMode = vbPixels
    
    ' Save the initial snapshot.
    Set m_Snapshots = New Collection
    m_CurrentSnapshot = 0
    SaveSnapshot
    
    ' Enable/disable the undo and redo menus.
    RaiseEvent EnableMenusForSelection
    
    ' Select the solid DrawStyle.
    icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(1)
    icbDrawStyle.ToolTipText = icbDrawStyle.ComboItems(1).Key
    
    ' Select the 1 pixel DrawWidth.
    icbDrawWidth.SelectedItem = icbDrawStyle.ComboItems(1)
    icbDrawWidth.ToolTipText = icbDrawStyle.ComboItems(1).Key
    'Redraw
End Sub

Public Function SaveDraw(ByVal File_name As String, ByVal file_title As String) As Boolean
Dim fnum As Integer

    On Error GoTo SaveError

    ' Open the file.
    fnum = FreeFile
    Open File_name For Output As fnum

    ' Write the scene serialization into the file.
    Print #fnum, "Page (W(" + Trim(Str(CanvasWidth)) + ") " + _
                      " H(" + Trim(Str(CanvasHeight)) + ")" + _
                      " C(" + Trim(Str(BackColor)) + "))" + vbCrLf + _
                      m_TheScene.Serialization

    ' Close the file.
    Close fnum
    
    m_FileName = File_name
    m_FileTitle = file_title

    m_DataModified = False
    SaveDraw = True
    Exit Function

SaveError:
    MsgBox "Error " & Format$(Err.Number) & " saving file " & File_name & "." & vbCrLf & Err.Description, vbCritical
    SaveDraw = False
    
End Function

'Open draw file
Public Function OpenDraw(ByVal File_name As String, ByVal file_title As String) As Boolean
    
Dim fnum As Integer
Dim txt As String
Dim token_name As String
Dim token_value As String
    Dim tmptxt As String
    
    On Error GoTo LoadError

    ' Open the file.
    fnum = FreeFile
    Open File_name For Input As fnum

    ' Read the scene serialization from the file.
    txt = Input$(LOF(fnum), fnum)

    ' Close the file.
    Close fnum
    If InStr(1, txt, "Page") > 0 Then
       'Do
        GetNamedToken txt, token_name, token_value
        If token_name = "Page" Then
          tmptxt = token_value
        Do
        GetNamedToken tmptxt, token_name, token_value
        If token_name = "W" Then
            CanvasWidth = CLng(token_value) 'Mid(tmptxt, InStr(1, tmptxt, "W:") + 2, InStr(1, tmptxt, " H:") - 9))
        ElseIf token_name = "H" Then
            CanvasHeight = CLng(token_value) 'Mid(tmptxt, InStr(1, tmptxt, "H:") + 2, InStr(1, tmptxt, " C:") - 15))
        ElseIf token_name = "C" Then
            PicCanvas.BackColor = CLng(token_value) 'Mid(tmptxt, InStr(1, tmptxt, "C:") + 2, Len(tmptxt) - InStr(1, tmptxt, "C:") + 2))
        Else
           Exit Do
         End If
       Loop
       End If
        txt = Replace(txt, vbCrLf, "")
    End If
    ' Initialize the scene.
    GetNamedToken txt, token_name, token_value
        
    If token_name <> "Scene" Then
        MsgBox "Error loading file " & File_name & "." & vbCrLf & "This is not a VbDraw file."
    Else
        m_TheScene.Serialization = token_value
        m_DataModified = False
    End If
   
    m_FileName = File_name
    m_FileTitle = file_title
    
   ' Save the initial snapshot.
    Set m_Snapshots = New Collection
    PrepareToEdit
    OpenDraw = True
    
    Exit Function
    
LoadError:
    MsgBox "Error " & Format$(Err.Number) & " loading file " & File_name & "." & vbCrLf & Err.Description
    OpenDraw = False
    Exit Function
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Filename() As String
    Filename = m_FileName
End Property

Public Property Let Filename(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileTitle() As String
    FileTitle = m_FileTitle
End Property

Public Property Let FileTitle(ByVal New_FileTitle As String)
    m_FileTitle = New_FileTitle
    PropertyChanged "FileTitle"
End Property

Public Sub DeselectAll()
       ' Deselect all objects.
       DeselectAllVbdObjects
       ChoiceColorForControl
       ClearBox
End Sub

'Fill object for draw
Private Sub DrawPen()
    Dim txt As String
    icbDrawStyle.ComboItems.Clear
    Set icbDrawStyle.ImageList = imlDrawStyles
    For I = 1 To 6
        Select Case I
        Case 1: txt = "Solid"
        Case 2: txt = "Dash"
        Case 3: txt = "Dot"
        Case 4: txt = "Dash-Dot"
        Case 5: txt = "Dash-Dot-Dot"
        Case 6: txt = "Transparent"
        End Select
        icbDrawStyle.ComboItems.Add I, txt, txt
        icbDrawStyle.ComboItems(I).Image = I
    Next I
    
    icbDrawWidth.ComboItems.Clear
    Set icbDrawWidth.ImageList = imlDrawWidths
    For I = 1 To 10
        icbDrawWidth.ComboItems.Add I, Str(I) + " point", Str(I) + " point"
        icbDrawWidth.ComboItems(I).Image = I
    Next I
End Sub


Public Property Get ShowPenProperty() As Boolean
    ShowPenProperty = m_ShowPenProperty
End Property

Public Property Let ShowPenProperty(ByVal New_ShowPenProperty As Boolean)
    m_ShowPenProperty = New_ShowPenProperty
    PropertyChanged "ShowPenProperty"
    
    If MeFormView1 = False Then
       MeForm1.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, 0
      MeFormView1 = True
    End If
       
    MeForm1.Visible = m_ShowPenProperty
    
End Property

Public Property Get ShowFillProperty() As Boolean
    ShowFillProperty = m_ShowFillProperty
End Property

Public Property Let ShowFillProperty(ByVal New_ShowFillProperty As Boolean)
    m_ShowFillProperty = New_ShowFillProperty
    PropertyChanged "ShowFillProperty"
    If MeFormView2 = False Then
       MeForm2.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height
       MeFormView2 = True
    End If
       
    MeForm2.Visible = m_ShowFillProperty
    
End Property

Public Property Get ShowTranformProperty() As Boolean
    ShowTranformProperty = m_ShowTranformProperty
End Property

Public Property Let ShowTranformProperty(ByVal New_ShowTranformProperty As Boolean)
    m_ShowTranformProperty = New_ShowTranformProperty
    PropertyChanged "ShowTranformProperty"
    If MeFormView3 = False Then
       MeForm3.Move UserControl.ScaleWidth - MeForm1.Width + 1 - VScroll1.Width, MeForm1.Height + MeForm2.Height
       MeFormView3 = True
    End If

    MeForm3.Visible = m_ShowTranformProperty
    
End Property

' Let the user scale the selected objects.
Private Sub TransformScale(X_scale As Single, Y_Scale As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single
Dim msg As String
    
     'Debug.Print "scale " + Str(x_scale) + " " + Str(y_scale)
     msg = "Scale X:" + Format(X_scale, "0.0") + " Y:" + Format(Y_Scale, "0.0")
     RaiseEvent MsgControl(msg)
 
     X_scale = X_scale / 100
     Y_Scale = Y_Scale / 100
    
    ' Bound the selected objects.
    'BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax
     xmin = XminBox
     ymin = YminBox
     xmax = XmaxBox
     ymax = YmaxBox
     
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 1 'Left top Corner
        xmid = xmax
        ymid = ymax
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 3 'Right top Corner
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 5 'Bottom Right corner
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 7 'Left bottom corner
       xmid = xmax
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case 9
       xmid = (xmin + xmax) / 2
       ymid = (ymin + ymax) / 2
    End Select
    
    m2ScaleAt M, X_scale, Y_Scale, xmid / m_ZoomFactor, ymid / m_ZoomFactor

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation M
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
      Set_Dirty
      Redraw
    End If
End Sub

' Let the user scale the selected objects.
Private Sub TransformSkew(X_scale As Single, Y_Scale As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single
Dim msg As String

     'Debug.Print "scale " + Str(x_scale) + " " + Str(y_scale)
     msg = "Skew X:" + Format(X_scale, "0.0") + " Y:" + Format(Y_Scale, "0.0")
     RaiseEvent MsgControl(msg)
 
     X_scale = X_scale / 100
     Y_Scale = Y_Scale / 100
    
    'Bound the selected objects.
     xmin = XminBox
     ymin = YminBox
     xmax = XmaxBox
     ymax = YmaxBox
     
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case 9 'Center
       xmid = (xmin + xmax) / 2
       ymid = (ymin + ymax) / 2
    End Select
    
    m2SkewAt M, X_scale, Y_Scale, xmid / m_ZoomFactor, ymid / m_ZoomFactor

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation M
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj
   
    ' The data has changed.
    If fSelect Then
       Set_Dirty
       Redraw
   End If
End Sub

' Let the user transform the selected objects.
Private Sub TransformPoint(X_Move As Single, Y_Move As Single)
Dim fSelect As Boolean
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single
    
    
    ' Make the transformation matrix.
    m2Translate M, X_Move, Y_Move

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
       If Obj.Selected = True And Obj.ObjLock = False Then
          Obj.AddTransformation M
          Obj.MakeTransformation
          fSelect = True
      End If
    Next Obj

    ' The data has changed.
    If fSelect Then
        Set_Dirty
        Redraw
    End If
End Sub

' Rotate the selected objects.
Private Sub TransformRotate(m_angle As Single, XminB As Single, YminB As Single, XmaxB As Single, YmaxB As Single)
Dim fSelect As Boolean
Dim Angle As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim Obj As vbdObject
Dim M(1 To 3, 1 To 3) As Single
    
    ' Get the angle of rotation.
    Angle = m_angle * PI / 180

    ' Bound the selected objects.
    xmin = XminB
    ymin = YminB
    xmax = XmaxB
    ymax = YmaxB
    
    ' Make the transformation matrix.
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
    m2RotateAround M, Angle, xmid, ymid

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.Angle = Obj.Angle + m_angle
           Obj.AddTransformation M
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
     Set_Dirty
    Redraw
    End If
End Sub

' Draw Reflect the transformed data.
Private Sub TransformMirror(rHor As Integer, rVer As Integer)
Dim fSelect As Boolean
Dim dx As Single
Dim dy As Single
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim M(1 To 3, 1 To 3) As Single
Dim I As Integer

    If Obj Is Nothing Then Exit Sub
    
    Obj.Bound xmin, ymin, xmax, ymax
    
    ' Transform the data.
     If rHor > 0 Then dx = 90 Else dx = 0
     If rVer > 0 Then dy = 90 Else dy = 0
     m2ReflectAcross M, xmin + ((xmax - xmin) / 2), ymin + ((ymax - ymin) / 2), dx, dy
    
    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        If Obj.Selected = True And Obj.ObjLock = False Then
           Obj.AddTransformation M
           Obj.MakeTransformation
           fSelect = True
        End If
    Next Obj

    ' The data has changed.
    If fSelect Then
       Set_Dirty
       Redraw
    End If
End Sub

Private Sub mDrawRotate(LastX As Single, LastY As Single)
Dim msg As String
Dim Ang As Single
Dim Points() As POINTAPI
ReDim Points(1 To 4)

    Set Ortho = New RectAngle
    Ortho.NumPoints = 4
    Ortho.X(1) = XminBox
    Ortho.X(2) = XmaxBox
    Ortho.X(3) = XmaxBox
    Ortho.X(4) = XminBox
    Ortho.Y(1) = YminBox
    Ortho.Y(2) = YminBox
    Ortho.Y(3) = YmaxBox
    Ortho.Y(4) = YmaxBox
    Ang = m2GetAngle3P(XminBox + (XmaxBox - XminBox) / 2, YminBox + (YmaxBox - YminBox) / 2, _
                       XminBox + (XmaxBox - XminBox), m_StartY, _
                       LastX, LastY)
    msg = "Rotate Angle:" + Format(360 - Ang, "0.0")
    RaiseEvent MsgControl(msg)
               
    If Ang = 0 Then Exit Sub
    m2Identity M
    m2RotateAround M, (Ang * PI / 180), (XmaxBox + XminBox) / 2, (YmaxBox + YminBox) / 2
    Ortho.Transform M
    
    Points(1).X = Ortho.X(1)
    Points(1).Y = Ortho.Y(1)
    PicCanvas.PSet (Points(1).X, Points(1).Y)
    For I = 2 To 4
       Points(I).X = Ortho.X(I)
       Points(I).Y = Ortho.Y(I)
       PicCanvas.Line -(Points(I).X, Points(I).Y)
    Next
    PicCanvas.Line -(Points(1).X, Points(1).Y)
    'Ortho.Draw PicCanvas
    'Ortho.ClearTransform M
    Set Ortho = Nothing
End Sub

Public Sub ViewTransform(TypeTrans As Integer)
      If MeForm3.Visible = False Then
         MeForm3.Visible = True
         ShowTranformProperty = True
      End If
      CtrTranform1.TypeTransform = TypeTrans
End Sub


Private Sub mDrawSkew(LastX As Single, LastY As Single)
Dim xmin As Single
Dim ymin As Single
Dim xmax As Single
Dim ymax As Single
Dim xmid As Single
Dim ymid As Single
Dim msg As String
Dim Ang As Single
Dim Points() As POINTAPI
ReDim Points(1 To 4)
    Set Ortho = New RectAngle
       
    Ortho.NumPoints = 4
    Ortho.X(1) = XminBox
    Ortho.X(2) = XmaxBox
    Ortho.X(3) = XmaxBox
    Ortho.X(4) = XminBox
    Ortho.Y(1) = YminBox
    Ortho.Y(2) = YminBox
    Ortho.Y(3) = YmaxBox
    Ortho.Y(4) = YmaxBox
    
     msg = "Skew X:" + Format(LastX, "0.0") + " Y:" + Format(LastY, "0.0")
     RaiseEvent MsgControl(msg)
 
     LastX = LastX / 100
     LastY = LastY / 100
    
    ' Bound the selected objects.
    BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax
    
    ' Make the transformation matrix.
    Select Case m_ScaleType
    Case 2 'Middle top
        xmid = xmin
        ymid = ymax
    Case 4 'Middle Right
       xmid = xmin
       ymid = ymin
    Case 6 'Middle Bottom
       xmid = xmin
       ymid = ymin
    Case 8 'Middle left
       xmid = xmax
       ymid = ymin
    Case Else
       Exit Sub
    End Select
    
    m2Identity M
    m2SkewAt M, LastX, LastY, xmid / m_ZoomFactor, ymid / m_ZoomFactor
    
    Ortho.Transform M
    Points(1).X = Ortho.X(1)
    Points(1).Y = Ortho.Y(1)
    PicCanvas.PSet (Points(1).X, Points(1).Y)
    For I = 2 To 4
       Points(I).X = Ortho.X(I)
       Points(I).Y = Ortho.Y(I)
       PicCanvas.Line -(Points(I).X, Points(I).Y)
    Next
    PicCanvas.Line -(Points(1).X, Points(1).Y)
    
    Set Ortho = Nothing
End Sub

Function DelObject()
       ObjectDelete
End Function

Public Sub Set_Dirty()
      RaiseEvent SetDirty
      RaiseEvent EnableMenusForSelection
End Sub

Public Function CutObject()
    If Obj Is Nothing Then Exit Function
    
    Clipboard.Clear
    Clipboard.SetText Obj.Serialization

    'Delete object
    ObjectDelete
    
End Function

Public Function CopyObject()
     If Obj Is Nothing Then Exit Function
     Clipboard.Clear
     Clipboard.SetText Obj.Serialization
     Set_Dirty
End Function

Public Function PasteObject()
Dim NewTxt As String, OldTxt As String, token_name As String, token_value As String
    
    NewTxt = Clipboard.GetText
    OldTxt = m_TheScene.Serialization
    GetNamedToken OldTxt, token_name, token_value
    If token_name = "Scene" Then
        m_TheScene.Serialization = token_value + vbCr + NewTxt
        m_DataModified = False
    End If
    
    Set_Dirty
   
    Redraw
End Function
 
Public Sub ClearTransform()
Dim Obj As vbdObject

    For Each Obj In m_SelectedObjects
       If Obj.Selected = True Then
          Obj.Angle = 0
          Obj.ClearTransformation
       End If
    Next Obj

    ' The data has changed.
    Set_Dirty
    Redraw
End Sub

Public Function ClearObject()
       Clipboard.Clear
End Function

Public Function IsSelectObject() As Boolean
       If Obj Is Nothing Then
          IsSelectObject = False
       Else
          IsSelectObject = True
       End If
End Function

Public Sub FileExport(Filename As String)
Dim mf_dc As Long
Dim hmf As Long
Dim old_size As POINTAPI

    ' Create the metafile.
    mf_dc = CreateMetaFile(ByVal Filename)
    If mf_dc = 0 Then
        MsgBox "Error creating the metafile.", vbExclamation
        Exit Sub
    End If

    ' Set the metafile's size to something reasonable.
    SetWindowExtEx mf_dc, PicCanvas.ScaleWidth, PicCanvas.ScaleHeight, old_size

    ' Draw in the metafile.
    m_TheScene.DrawInMetafile mf_dc

    ' Close the metafile.
    hmf = CloseMetaFile(mf_dc)
    If hmf = 0 Then
        MsgBox "Error closing the metafile.", vbExclamation
    End If

    ' Delete the metafile to free resources.
    If DeleteMetaFile(hmf) = 0 Then
        MsgBox "Error deleting the metafile.", vbExclamation
    End If
    If FileExists(Filename) Then
       MsgBox Filename & " Saved OK!"
    Else
       MsgBox "ERROR! " & Filename & " Not Saved!"
    End If
End Sub

Public Sub FileExportBitmap(Filename As String)
   
   ' Make picHidden big enough to hold everything.
    picHidden.Width = PicCanvas.Width
    picHidden.Height = PicCanvas.Height
    picHidden.ScaleWidth = PicCanvas.ScaleWidth
    picHidden.ScaleHeight = PicCanvas.ScaleHeight
    ' Erase the picture.
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF

    ' Deselect all the objects.
    DeselectAllVbdObjects
    Redraw
    picHidden.AutoRedraw = True
    ' Draw the bitmap on picHidden.
    m_TheScene.Draw picHidden
    picHidden.Picture = picHidden.Image
    ' Save the picture.
    If StartUpGDIPlus(GdiPlusVersion) Then
       If SavePictureFromHDC(picHidden, Filename) Then
          MsgBox Filename & " Saved OK!"
       Else
          MsgBox "ERROR! " & Filename & " Not Saved!"
       End If
       ShutdownGDIPlus
    End If
        
End Sub

Public Sub PrintDraw()
    picHidden.Width = PicCanvas.Width
    picHidden.Height = PicCanvas.Height
    picHidden.ScaleWidth = PicCanvas.ScaleWidth
    picHidden.ScaleHeight = PicCanvas.ScaleHeight
    ' Erase the picture.
    picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF

    ' Deselect all the objects.
    DeselectAllVbdObjects
    Redraw
    picHidden.AutoRedraw = True
    ' Draw the bitmap on picHidden.
    m_TheScene.Draw picHidden
    picHidden.Picture = picHidden.Image
    
     If Printers.Count < 1 Then
       MsgBox "No printer", vbInformation, "Printing"
     Else
       frmPrint.ShowForm picHidden.Picture
     End If
     
     picHidden.Picture = LoadPicture()
     picHidden.Line (picHidden.ScaleLeft, picHidden.ScaleTop)-Step(picHidden.ScaleWidth, picHidden.ScaleHeight), vbWhite, BF
End Sub

Private Sub ComDropperFill_Click()
      m_ReadFillProperty = True
      SelectTool 20 '"DropperFill"
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor2() As OLE_COLOR
    FillColor2 = m_FillColor2
End Property

Public Property Let FillColor2(ByVal New_FillColor2 As OLE_COLOR)
    m_FillColor2 = New_FillColor2
    ChangeFillColor 2, m_FillColor2
    PropertyChanged "FillColor2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Pattern() As String
    Pattern = m_Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    m_Pattern = New_Pattern
    ChangePattern m_Pattern
    PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Gradient() As Integer
    Gradient = m_TypeGradient
End Property

Public Property Let Gradient(ByVal New_TypeGradient As Integer)
    m_TypeGradient = New_TypeGradient
    ChangeGradient m_TypeGradient
    PropertyChanged "Gradient"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,255
Public Property Get Blend() As Integer
    Blend = m_Blend
End Property

Public Property Let Blend(ByVal New_Blend As Integer)
    m_Blend = New_Blend
    PropertyChanged "Blend"
    ChangeBlend m_Blend
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicCanvas,PicCanvas,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicCanvas.BackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get BackImage() As String
    BackImage = m_BackImage
End Property

Public Property Let BackImage(ByVal New_BackImage As String)
    m_BackImage = New_BackImage
    PropertyChanged "BackImage"
    If FileExists(m_BackImage) Then
       LoadPicBox m_BackImage, PicCanvas, False
       PicCanvas.Picture = PicCanvas.Image
       ScalePBox PicCanvas
       'PicCanvas.Refresh
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get LockObject() As Boolean
    LockObject = m_LockObject
End Property

Public Property Let LockObject(ByVal New_LockObject As Boolean)
    m_LockObject = New_LockObject
    PropertyChanged "LockObject"
    ObjectLock m_LockObject
End Property

Private Sub ObjectLock(isLock As Boolean)
      If Not Obj Is Nothing Then
         Obj.ObjLock = isLock
         m_LockObject = Obj.ObjLock
         Set_Dirty
         Redraw
      End If
End Sub

Private Sub ObjectDelete()
      DeletevbdObject
      ' Save the current snapshot.
      Set_Dirty
      Redraw
End Sub


'Read Path text and make PointCoolds and Type for draw
Private Sub ReadPathText(Obj As PictureBox, _
                         txt As String, _
                         Point_Coords() As POINTAPI, _
                         Point_Types() As Byte, _
                         NumPoints As Long)
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

