VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   11040
      Left            =   7080
      ScaleHeight     =   732
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   10
      Top             =   2865
      Visible         =   0   'False
      Width           =   15420
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   4275
      TabIndex        =   8
      Top             =   2400
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   420
      Left            =   4275
      TabIndex        =   7
      Top             =   1905
      Width           =   1305
   End
   Begin VB.CheckBox chkVel 
      Appearance      =   0  'Flat
      Caption         =   "Custom Size"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4140
      TabIndex        =   6
      Top             =   420
      Width           =   1485
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4890
      TabIndex        =   3
      Top             =   780
      Width           =   705
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4890
      TabIndex        =   2
      Top             =   1080
      Width           =   705
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   405
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   690
      Width           =   3150
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   405
         MousePointer    =   15  'Size All
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   144
         TabIndex        =   11
         Top             =   615
         Width           =   2190
      End
      Begin VB.Image imgPreview 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1110
         Left            =   975
         MousePointer    =   15  'Size All
         Stretch         =   -1  'True
         Top             =   3045
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Shape shpMargin 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         Height          =   4230
         Left            =   60
         Top             =   60
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Picture"
      Height          =   420
      Left            =   4275
      TabIndex        =   9
      Top             =   1905
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4170
      TabIndex        =   5
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   4170
      TabIndex        =   4
      Top             =   780
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "PRINT PREVIEW - A4 PAPER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   435
      TabIndex        =   1
      Top             =   375
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   240
      Top             =   255
      Width           =   3495
   End
   Begin VB.Image imgBuffer 
      Height          =   945
      Left            =   1365
      Top             =   3240
      Visible         =   0   'False
      Width           =   1680
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aMouseDown As Boolean
Private xp1 As Long, yp1 As Long
Private ratio As Single
Dim Canceled As Boolean

Private Sub chkVel_Click()
      If chkVel.Value = 1 Then
         txtWidth.Enabled = True
         txtHeight.Enabled = True
      Else
        txtWidth.Enabled = False
        txtHeight.Enabled = False
      End If
End Sub

Private Sub cmdExit_Click()
 
   Canceled = False
    Hide
End Sub

Private Sub cmdOpen_Click()
   
                            
        On Error Resume Next
        'Set imgBuffer.Picture to the picture from the file
        imgBuffer.Picture = Picture1.Picture 'LoadPicture(FileName)
        ratio = imgBuffer.Width / imgBuffer.Height
        
        'Put the image to scale according to paper size
        pic.Width = imgBuffer.Width / 2.8
        pic.Height = imgBuffer.Height / 2.8
        pic.Picture = imgBuffer.Picture
        
        'If the image is too wide resize it but constrain proportions
        'You should add similar code for height
        If pic.Left + pic.Width > shpMargin.Left + shpMargin.Width Then
            If pic.Width > 560 / 2.8 Then
                pic.Width = 560 / 2.8
                pic.Height = pic.Width / ratio
            End If
            pic.Move shpMargin.Left
        End If
        'Set resize labels
        txtHeight.Text = Int(pic.Height * 2.8)
        txtWidth.Text = Int(pic.Width * 2.8)
        pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
        'StretchBlt pic, 0, 0, pic.Width, pic.Height, imgBuffer, 0, 0, imgBuffer.Width, imgBuffer.Height, vbSrcCopy
        'imgResize.Picture = imgBuffer.Picture
        
        cmdPrint.Enabled = True
        
        If Err Then
         '   MsgBox Err.Description, vbInformation, App.Title
        End If
    
  '  End If

End Sub

Private Sub cmdPrint_Click()
    
    'Print the image
    'SetLargePrinterScale pic
    If txtWidth.Text > 0 Then
        Printer.PaintPicture pic.Picture, ((pic.Left * 2.8) / 28) * 546.44, ((pic.Top * 2.8) / 28) * 546.44, ((pic.Width * 2.8) / 28) * 546.44, ((pic.Height * 2.8) / 28) * 546.44
        Printer.EndDoc
    End If
    cmdExit_Click
   
End Sub

Private Sub Form_Load()
'   SetLargePrinterScale pic
       
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload Me
    
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
aMouseDown = True
   xp1 = X
   yp1 = Y
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewX As Long, NewY As Long

   If aMouseDown Then
      NewX = pic.Left + (X - xp1)
      NewY = pic.Top + (Y - yp1)
       If NewX <= 5 Then NewX = 5
       If NewY <= 5 Then NewY = 5
       If NewX + pic.Width < picPaper.Width - 5 Then
       If NewY + pic.Height < picPaper.Height - 5 Then
            pic.Left = NewX
            pic.Top = NewY
       End If
       End If
   End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = False

End Sub

Private Sub txtHeight_Change()
    
    On Error Resume Next
    
    If Int(txtHeight.Text) <= 792 Then
        pic.Height = Int(txtHeight.Text) / 2.8
    Else
        txtHeight.Text = "792"
    End If
    pic.Width = pic.Height * ratio
    
    pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
    'txtWidth.Text = Int(pic.Width * 2.8)
End Sub

Private Sub txtWidth_Change()

    On Error Resume Next

    If Int(txtWidth.Text) <= 560 Then
        If Int(txtWidth.Text) > 0 Then
            pic.Width = Int(txtWidth.Text) / 2.8
        End If
    Else
        txtWidth.Text = "560"
    End If
    pic.Height = pic.Width / ratio
    
    pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
    'txtHeight.Text = Int(pic.Height * 2.8)
End Sub

Sub ChangeSize()

        ratio = imgBuffer.Width / imgBuffer.Height
        pic.Height = pic.Width / ratio
        
        pic.PaintPicture imgBuffer.Picture, 0, 0, pic.Width, pic.Height
        txtHeight.Text = Int(pic.Height * 2.8)
        txtWidth.Text = Int(pic.Width * 2.8)
        
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(cPic As Picture)
       Set Picture1.Picture = cPic
       cmdOpen_Click
       Show vbModal
       Unload Me
End Function

