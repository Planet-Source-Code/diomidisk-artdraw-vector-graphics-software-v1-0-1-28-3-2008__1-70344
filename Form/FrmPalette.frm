VERSION 5.00
Begin VB.Form FrmPalette 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palette"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   3345
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   5
      Top             =   390
      Width           =   1650
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2865
      TabIndex        =   4
      Top             =   3570
      Width           =   1065
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   135
      Pattern         =   "*.pal"
      TabIndex        =   1
      Top             =   375
      Width           =   2940
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   1500
      TabIndex        =   0
      Top             =   3570
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Color palette "
      Height          =   210
      Left            =   3345
      TabIndex        =   3
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select palette :"
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   105
      Width           =   2850
   End
End
Attribute VB_Name = "FrmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Canceled As Boolean

Dim ColorList() As Long
Const MaxCol = 12
Const TSize = 9

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Canceled = False
    Me.Hide
End Sub

Private Sub File1_Click()
   If FileExists(File1.Path + "\" + File1.Filename) Then
      'ColorPal1.LoadPalette File1.Path + "\" + File1.Filename
      LoadPalette File1.Path + "\" + File1.Filename
   End If
End Sub

Private Sub File1_KeyUp(KeyCode As Integer, Shift As Integer)
     
     Dim id As Long
     id = File1.ListIndex
     If KeyCode = 46 Then
        Kill File1.Path + "\" + File1.Filename
        File1.Refresh
        If File1.ListIndex = -1 Then Exit Sub
        File1.ListIndex = id
     End If
End Sub

Private Sub Form_Load()
    File1.Path = App.Path + "\Palette"
    File1.Refresh
End Sub

' Display the form. Return True if the user cancels.
Public Function ShowForm(FileNamePalette As String) As Boolean
    ' Assume we will cancel.
    Canceled = True

    ' Display the form.
    Show vbModal

    ShowForm = Canceled
    If Not Canceled Then
        On Error Resume Next
        FileNamePalette = File1.Path + "\" + File1.Filename
        On Error GoTo 0
    End If
End Function

Public Sub LoadPalette(Optional PalFile As String)
On Error Resume Next
Dim FF As Integer
Dim tStr As String
Dim n As Integer
Dim cQty As Integer
Dim Row As Integer
Dim Col As Integer

    FF = FreeFile

    If PalFile = "" Or Dir(PalFile) = "" Then PalFile = App.Path & "\Palette\Default.pal"

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
        ColorList(n) = RGB(Val(Split(tStr, " ")(0)), Val(Split(tStr, " ")(1)), Val(Split(tStr, " ")(2)))
        n = n + 1
    Wend
Close #FF
Col = 0
Row = 0
    For n = 0 To cQty - 1
    Picture1.Line (Col * TSize, Row * TSize)-(Col * TSize + TSize, Row * TSize + TSize), ColorList(n), BF
    Col = Col + 1
    If Col = MaxCol Then
    Col = 0
    Row = Row + 1
    End If
    Next n
 Picture1.ScaleMode = 3
'Picture1.Width = Picture1.ScaleX((MaxCol * TSize), vbPixels, vbContainerSize)
'Picture1.Height = Picture1.ScaleY((cQty / MaxCol * TSize) + TSize + 2, vbPixels, vbContainerSize)
End If
Exit Sub
ErrLoad:
Close #FF
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmColorPicker.ShowForm Picture1.POINT(X, Y)
End Sub
