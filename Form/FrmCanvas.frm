VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmCanvas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Page Size"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   150
      TabIndex        =   9
      Top             =   375
      Width           =   1755
   End
   Begin VB.CommandButton ComOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1980
      TabIndex        =   0
      Top             =   2595
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom"
      Height          =   1305
      Left            =   2040
      TabIndex        =   1
      Top             =   150
      Width           =   3330
      Begin VB.TextBox TxtHeight 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1080
         TabIndex        =   6
         Text            =   "480"
         Top             =   675
         Width           =   690
      End
      Begin VB.TextBox TxtWidth 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1065
         TabIndex        =   5
         Text            =   "640"
         Top             =   285
         Width           =   690
      End
      Begin VB.OptionButton OptionType 
         Caption         =   "Pixels"
         Height          =   240
         Index           =   0
         Left            =   1860
         TabIndex        =   4
         Top             =   255
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton OptionType 
         Caption         =   "Inches"
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Top             =   525
         Width           =   1125
      End
      Begin VB.OptionButton OptionType 
         Caption         =   "Millimeters"
         Height          =   240
         Index           =   2
         Left            =   1860
         TabIndex        =   2
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Height:"
         Height          =   240
         Left            =   255
         TabIndex        =   8
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Page back color"
      Height          =   870
      Left            =   2055
      TabIndex        =   11
      Top             =   1500
      Width           =   1905
      Begin VB.TextBox TextImage 
         Height          =   255
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   675
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton CmdImage 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   345
         Left            =   1350
         Picture         =   "FrmCanvas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Insert image"
         Top             =   630
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   1140
         Picture         =   "FrmCanvas.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Back page color"
         Top             =   270
         Width           =   375
      End
      Begin VB.OptionButton OptionColor 
         Caption         =   "Color"
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   570
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.OptionButton OptionImage 
         Caption         =   "Image"
         Height          =   330
         Left            =   315
         TabIndex        =   12
         Top             =   630
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.PictureBox PicColor 
         Height          =   315
         Left            =   330
         ScaleHeight     =   255
         ScaleWidth      =   600
         TabIndex        =   16
         Top             =   300
         Width           =   660
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4290
      Top             =   2610
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label LabelSize 
      Caption         =   "Size :"
      Height          =   270
      Left            =   180
      TabIndex        =   10
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "FrmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myH As Single
Dim myW As Single
Dim mColor As OLE_COLOR
Dim mImage As String
Dim Canceled As Boolean
' Display the form. Return True if the user cancels.
Public Function ShowForm(cWidth As Single, _
                         cHeight As Single, _
                         Optional cImage As String = "", _
                         Optional cColor As OLE_COLOR) As Boolean
    
    myW = cWidth
    myH = cHeight
    mImage = cImage
    mColor = cColor
    PicColor.BackColor = mColor
    TxtHeight.Text = myH
    TxtWidth.Text = myW
    
    ' Display the form.
    Show vbModal
    ShowForm = Canceled
    cWidth = myW
    cHeight = myH
    If OptionImage.Value Then
       cImage = mImage
       cColor = vbWhite
    End If
    If OptionColor.Value Then
       cColor = PicColor.BackColor 'mColor
       cImage = ""
    End If
    If OptionColor.Value = False And OptionImage.Value = False Then
       cColor = vbWhite ' -1
       cImage = ""
    End If
    Unload Me
End Function

Private Sub cmdColor_Click()
        OpenColorDialog PicColor
End Sub

Private Sub CmdImage_Click()
                 
    dlgFile.Flags = cdlOFNExplorer Or _
                    cdlOFNHideReadOnly Or _
                    cdlOFNLongNames Or _
                    cdlOFNOverwritePrompt
    dlgFile.Filter = "Bitmap Files (bmp,gif,tif,png,jpg)|*.bmp;*.gif;*.tif;*.png;*.jpg"
    
    If PathExists(App.Path + "\Object") = False Then MkDir App.Path + "\Object"
    dlgFile.InitDir = App.Path + "\Object"
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If

    mImage = dlgFile.FileName
    dlgFile.InitDir = Left$(mImage, Len(mImage) - Len(dlgFile.FileTitle) - 1)
    TextImage.Text = mImage
End Sub

Private Sub ComOk_Click()

    Canceled = False
    Hide
    'Unload Me
    'Exit Sub

End Sub


Private Sub Form_Load()
   ' myH = Form1.ObjDraw1.CanvasHeight
   ' myW = Form1.ObjDraw1.CanvasWidth
    TxtHeight.Text = myH
    TxtWidth.Text = myW
    List1.AddItem "320 x 200"
    List1.AddItem "640 x 480"
    List1.AddItem "800 x 600"
    List1.AddItem "1024 x 768"
    List1.AddItem "1280 x 1024"
    
End Sub

Private Sub List1_Click()
    Select Case List1.ListIndex
    Case 0
       myW = 320
       myH = 200
    Case 1
       myW = 640
       myH = 480
    Case 2
       myW = 800
       myH = 600
    Case 3
       myW = 1024
       myH = 768
    Case 4
       myW = 1280
       myH = 1024
    End Select
    
     TxtWidth.Text = myW
     TxtHeight.Text = myH
End Sub

Private Sub OptionColor_Click()
       If OptionColor.Value = True Then
           cmdColor.Enabled = True
           CmdImage.Enabled = False
           cmdColor_Click
       End If
End Sub

Private Sub OptionImage_Click()
       If OptionImage.Value = True Then
           cmdColor.Enabled = False
           CmdImage.Enabled = True
           CmdImage_Click
       End If
End Sub

Private Sub OptionType_Click(Index As Integer)
    Select Case Index
    Case 0
        TxtWidth.Text = Round(myW)
        TxtHeight.Text = Round(myH)
    Case 1
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbInches), 2)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbInches), 2)
    Case 2
        TxtWidth.Text = Round(ScaleX(myW, vbPixels, vbMillimeters), 2)
        TxtHeight.Text = Round(ScaleY(myH, vbPixels, vbMillimeters), 2)
        
    End Select
End Sub



Private Sub TxtHeight_Change()
On Error Resume Next

    If OptionType(1).Value = True Then
        myH = Round(ScaleY(Format(TxtHeight.Text), vbInches, vbPixels), 2)
    ElseIf OptionType(2).Value = True Then
        myH = Round(ScaleY(Format(TxtHeight.Text), vbMillimeters, vbPixels), 2)
    Else
        myH = TxtHeight.Text
    End If
End Sub

Private Sub TxtWidth_Change()
On Error Resume Next

    If OptionType(1).Value = True Then
        myW = Round(ScaleX(Format(TxtWidth.Text), vbInches, vbPixels), 2)
    ElseIf OptionType(2).Value = True Then
        myW = Round(ScaleX(Format(TxtWidth.Text), vbMillimeters, vbPixels), 2)
    Else
        myW = TxtWidth.Text
    End If
End Sub


