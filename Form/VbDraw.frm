VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVbDraw 
   Caption         =   "Draw []"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10875
   Icon            =   "VbDraw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ArtDraw.DrawControl DrawControl1 
      Height          =   6285
      Left            =   2580
      TabIndex        =   2
      Top             =   1740
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   11086
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   3  'Align Left
      Height          =   6630
      Left            =   0
      TabIndex        =   14
      Top             =   750
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   11695
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   405
      _CBHeight       =   6630
      _Version        =   "6.7.9782"
      Child1          =   "drawToolbar"
      MinHeight1      =   345
      Width1          =   2940
      NewRow1         =   0   'False
      Begin ArtDraw.ucToolbar drawToolbar 
         Height          =   6570
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   11589
      End
   End
   Begin VB.PictureBox PicturePrint 
      Height          =   585
      Left            =   5490
      ScaleHeight     =   525
      ScaleWidth      =   1290
      TabIndex        =   13
      Top             =   1095
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox PicTollBar2 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   900
      Picture         =   "VbDraw.frx":5F32
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   11
      Top             =   1425
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTollBar3 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   945
      Picture         =   "VbDraw.frx":7D34
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTollBar1 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   870
      Picture         =   "VbDraw.frx":8AF6
      ScaleHeight     =   240
      ScaleWidth      =   3930
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.PictureBox PicTools 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   915
      Picture         =   "VbDraw.frx":BFB8
      ScaleHeight     =   240
      ScaleWidth      =   3690
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   3690
   End
   Begin ArtDraw.ColorPalette ColorPalette1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   7380
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   979
   End
   Begin MSComctlLib.ImageList ImToolBar 
      Left            =   750
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":F83A
            Key             =   "Arrow"
            Object.Tag             =   "Arrow"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":FBD4
            Key             =   "Line"
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":FF6E
            Key             =   "Point"
            Object.Tag             =   "Point"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":10308
            Key             =   "Free"
            Object.Tag             =   "Free"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":106A2
            Key             =   "FreePolygon"
            Object.Tag             =   "FreePolygon"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":10A3C
            Key             =   "Polygon"
            Object.Tag             =   "Polygon"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":10DD6
            Key             =   "RectAngle"
            Object.Tag             =   "RectAngle"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":11170
            Key             =   "Plegma"
            Object.Tag             =   "Plegma"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1150A
            Key             =   "Ellipse"
            Object.Tag             =   "Ellipse"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":118A4
            Key             =   "Spiral"
            Object.Tag             =   "Spiral"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":11C3E
            Key             =   "Fill"
            Object.Tag             =   "Fill"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":14948
            Key             =   "Pen"
            Object.Tag             =   "Pen"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":17652
            Key             =   "Polyline"
            Object.Tag             =   "Polyline"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":179EC
            Key             =   "Text"
            Object.Tag             =   "Text"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":17D86
            Key             =   "TextFrame"
         EndProperty
      EndProperty
   End
   Begin ArtDraw.CtrColor CtrColor1 
      Height          =   525
      Left            =   8655
      TabIndex        =   8
      Top             =   6855
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   926
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   765
      Top             =   3225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlToolBar 
      Left            =   735
      Top             =   2565
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   46
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":18120
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":18472
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":187C4
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":18B16
            Key             =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":18E68
            Key             =   "Export"
            Object.Tag             =   "Export"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":191BA
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1950C
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1985E
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":19BB0
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":19F02
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1A254
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1A5A6
            Key             =   "TextLeft"
            Object.Tag             =   "TextLeft"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1A8F8
            Key             =   "TextCenter"
            Object.Tag             =   "TextCenter"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1AC4A
            Key             =   "TextRight"
            Object.Tag             =   "TextRight"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1AF9C
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1B2EE
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1B640
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1B992
            Key             =   "Strikethru"
            Object.Tag             =   "Strikethru"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1BCE4
            Key             =   "SelectAll"
            Object.Tag             =   "SelectAll"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1C036
            Key             =   "UnselectAll"
            Object.Tag             =   "UnselectAll"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1C388
            Key             =   "AlignLeft"
            Object.Tag             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1C6DA
            Key             =   "AlignCenterVertical"
            Object.Tag             =   "AlignCenterVertical"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1CA2C
            Key             =   "AlignRight"
            Object.Tag             =   "AlignRight"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1CD7E
            Key             =   "AlignTop"
            Object.Tag             =   "AlignTop"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1D0D0
            Key             =   "AlignCenterHorizontal"
            Object.Tag             =   "AlignCenterHorizontal"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1D422
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1D774
            Key             =   "AlignCenterVerticalHorizontal"
            Object.Tag             =   "AlignCenterVerticalHorizontal"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1DAC6
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1DE18
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1E16A
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1E4BC
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1E80E
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1EB60
            Key             =   "Ungroup"
            Object.Tag             =   "Ungroup"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1EEB2
            Key             =   "Zoom100"
            Object.Tag             =   "Zoom100"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1F204
            Key             =   "Zoom-"
            Object.Tag             =   "Zoom-"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1F556
            Key             =   "Zoom+"
            Object.Tag             =   "Zoom+"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1F8A8
            Key             =   "ZoomAll"
            Object.Tag             =   "ZoomAll"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1FBFA
            Key             =   "IMG38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":1FF4C
            Key             =   "IMG39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":2029E
            Key             =   "IMG40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":205F0
            Key             =   "IMG41"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":20942
            Key             =   "IMG42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":20E94
            Key             =   "Trans"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":213E6
            Key             =   "Fill"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":21938
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VbDraw.frx":21E8A
            Key             =   "Symbol"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   7935
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10054
            MinWidth        =   10054
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Mouse Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "VbDraw.frx":223DC
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "2:10 ìì"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   1323
      _CBWidth        =   10875
      _CBHeight       =   750
      _Version        =   "6.7.9782"
      Child1          =   "drawToolbar1"
      MinWidth1       =   6405
      MinHeight1      =   330
      Width1          =   4305
      NewRow1         =   0   'False
      Child2          =   "drawToolbar3"
      MinWidth2       =   2205
      MinHeight2      =   330
      Width2          =   2205
      NewRow2         =   -1  'True
      Child3          =   "drawToolbar2"
      MinWidth3       =   5595
      MinHeight3      =   330
      Width3          =   1200
      NewRow3         =   0   'False
      Begin ArtDraw.ucToolbar drawToolbar2 
         Height          =   330
         Left            =   2595
         TabIndex        =   12
         Top             =   390
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   582
      End
      Begin ArtDraw.ucToolbar drawToolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   30
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   582
      End
      Begin ArtDraw.ucToolbar drawToolbar3 
         Height          =   330
         Left            =   165
         TabIndex        =   9
         Top             =   390
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   582
         Begin VB.ComboBox ComboZoom 
            Height          =   315
            Left            =   1260
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "ComboZoom"
            Top             =   0
            Width           =   960
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSaveBitmapSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import ..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileSaveBitmap 
         Caption         =   "Export ..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSaveMetafile 
         Caption         =   "Export &Metafile..."
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprintersetup 
         Caption         =   "Printer setup"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExitSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnusepcut 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusepclear 
         Caption         =   "-"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear ClipBoard"
      End
      Begin VB.Menu mnusepDel 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu Mnunormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSimpleWireframe 
         Caption         =   "Simple Wireframe"
      End
      Begin VB.Menu mnufullscreenpreview 
         Caption         =   "-"
      End
      Begin VB.Menu mnufullscreen 
         Caption         =   "Full screen preview"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusepSymbol 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSymbol 
         Caption         =   "Symbol"
      End
      Begin VB.Menu mnusepform 
         Caption         =   "-"
      End
      Begin VB.Menu mnupenform 
         Caption         =   "Pen form"
      End
      Begin VB.Menu mnufillform 
         Caption         =   "Fill form"
      End
      Begin VB.Menu mnutransformform 
         Caption         =   "Transform form"
      End
   End
   Begin VB.Menu mnutext 
      Caption         =   "&Text"
      Begin VB.Menu mnuedittext 
         Caption         =   "Edit text"
      End
      Begin VB.Menu mnuexptudetext 
         Caption         =   "Extrude text"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepparag 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvertparagraph 
         Caption         =   "Convert paragraph"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuConvertarttext 
         Caption         =   "Convert Art Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu sepfittext 
         Caption         =   "-"
      End
      Begin VB.Menu mnufitPath 
         Caption         =   "Fit Text To Path"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnufitframe 
         Caption         =   "Fit Text to Frame"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuArrange 
      Caption         =   "&Arrange"
      Begin VB.Menu mnutransform 
         Caption         =   "&Transform"
         Begin VB.Menu mnuMove 
            Caption         =   "&Position"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuTransformRotate 
            Caption         =   "&Rotate"
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnuTransformScale 
            Caption         =   "&Scale"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mnuskew 
            Caption         =   "S&kew"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu mnuReflect 
            Caption         =   "&Mirror"
            Shortcut        =   ^{F9}
         End
      End
      Begin VB.Menu mnuTransformClear 
         Caption         =   "&Clear Transformations"
      End
      Begin VB.Menu mnusepSelectAll 
         Caption         =   "-"
      End
      Begin VB.Menu mnuselectall 
         Caption         =   "Select All"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuunselectall 
         Caption         =   "UnSelect all"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuSepBringFront 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeSendToFront 
         Caption         =   "&Bring To Front"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuArrangeSendToForward 
         Caption         =   "&Bring To Forward"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuArrangeSendToBackward 
         Caption         =   "&Send To Backward"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuArrangeSendToBack 
         Caption         =   "&Send To Back"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnusepLock 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock Object"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuunloackobject 
         Caption         =   "Unlock Object"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUnlockAllObject 
         Caption         =   "Unlock All Object"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmVbDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'' The new object we are building.
'Private m_NewObject As vbdObject
'Private m_ToolKey As String
'
'' The selected object.
'Private m_SelectedObjects As Collection
'
'' Undo variables.
'Private Const MAX_UNDO = 50
'Private m_Snapshots As Collection
'Private m_CurrentSnapshot As Integer
'
'' The scene that holds all objects.
'Private m_TheScene As vbdObject

' The currently selected colors.
Private m_ForeColor As Integer
Private m_BackColor As Integer

'' The name and title of the current file.
'Private m_FileName As String
'Private m_FileTitle As String

' MRU list file names.
Private m_MruList As Collection


' Return True if it is safe to discard the
' current picture.
Private Function DataSafe() As Boolean
    If Not m_DataModified Then
        DataSafe = True
    Else
        Select Case MsgBox("The data has been modified. Do you want to save the changes?", vbYesNoCancel + vbInformation)
            Case vbYes
                mnuFileSave_Click
                DataSafe = Not m_DataModified
            Case vbNo
                DataSafe = True
            Case vbCancel
                DataSafe = False
        End Select
    End If
End Function

'' Save the picture.
'Private Sub DataSave(ByVal file_name As String, ByVal file_title As String)
'Dim fnum As Integer
'
'    On Error GoTo SaveError
'
'    ' Open the file.
'    fnum = FreeFile
'    Open file_name For Output As fnum
'
'    ' Write the scene serialization into the file.
'    Print #fnum, m_TheScene.Serialization
'
'    ' Close the file.
'    Close fnum
'
'    ' Update the caption.
'    SetFileName file_name, file_title
'
'    m_DataModified = False
'    Exit Sub
'
'SaveError:
'    MsgBox "Error " & Format$(Err.Number) & _
'        " saving file " & file_name & "." & _
'        vbCrLf & Err.Description
'    Exit Sub
'End Sub

' Load the picture.
Private Sub DataLoad(ByVal File_name As String, ByVal file_title As String)
Dim fnum As Integer
Dim txt As String
Dim token_name As String
Dim token_value As String

    On Error GoTo LoadError

    ' Open the file.
    fnum = FreeFile
    Open File_name For Input As fnum

    ' Read the scene serialization from the file.
    txt = Input$(LOF(fnum), fnum)

    ' Close the file.
    Close fnum

    ' Initialize the scene.
    GetNamedToken txt, token_name, token_value
    If token_name <> "Scene" Then
        MsgBox "Error loading file " & File_name & "." & vbCrLf & "This is not a VbDraw file."
    Else
        m_TheScene.Serialization = token_value

        ' Update the caption.
        SetFileName File_name, file_title
        m_DataModified = False

        ' Prepare to edit.
        DrawControl1.PrepareToEdit
    End If
    Exit Sub

LoadError:
    MsgBox "Error " & Format$(Err.Number) & " loading file " & File_name & "." & vbCrLf & Err.Description
    Exit Sub
End Sub

'' Deselect this object.
'Private Sub DeselectVbdObject(ByVal target As vbdObject)
'Dim obj As vbdObject
'Dim i As Integer
'
'    ' Remove the object from the
'    ' m_SelectedObjects collection.
'    i = 1
'    For Each obj In m_SelectedObjects
'        If obj Is target Then
'            m_SelectedObjects.Remove i
'            Exit For
'        End If
'        i = i + 1
'    Next obj
'
'    ' Mark the object as not selected.
'    target.Selected = False
'End Sub
'' Deselect all objects.
'Private Sub DeselectAllVbdObjects()
'Dim obj As vbdObject
'
'    ' Deselect all selected objects.
'    For Each obj In m_SelectedObjects
'        obj.Selected = False
'    Next obj
'
'    ' Empty the m_SelectedObjects collection.
'    Set m_SelectedObjects = New Collection
'End Sub


'' Select the arrow tool.
'Private Sub SelectArrowTool()
'    ' Make sure the arrow button is pressed.
'    tbrTools.Buttons("Arrow").Value = tbrPressed
'
'    ' Prepare to deal with this tool.
'    SelectTool "Arrow"
'End Sub

' Create an appropriate object for this tool.
'Private Sub SelectTool(ByVal Key As String)
'Dim new_pgon As vbdPolygon
'Dim new_line As vbdLine
'
'    ' Free any previously started object.
'    Set m_NewObject = Nothing
'
'    ' Create the new object.
'    m_ToolKey = Key
'    Select Case m_ToolKey
'        Case "Polyline"
'            Set m_NewObject = New vbdDraw
'            Set new_pgon = m_NewObject
'            new_pgon.IsClosed = False
'        Case "Polygon"
'            Set m_NewObject = New vbdDraw
'            Set new_pgon = m_NewObject
'            new_pgon.IsClosed = True
'        Case "Line"
'            Set m_NewObject = New vbdLine
'            Set new_line = m_NewObject
'            new_line.IsBox = False
'        Case "Rectangle"
'            Set m_NewObject = New vbdLine
'            Set new_line = m_NewObject
'            new_line.IsBox = True
'        Case "Scribble"
'            Set m_NewObject = New vbdScribble
''        Case "Ellipse"
''            Set m_NewObject = New vbdEllipse
'    End Select
'
'    ' Let the new object receive picCanvas events.
'    If Not (m_NewObject Is Nothing) Then
'        Set m_NewObject.Canvas = picCanvas
'    End If
'End Sub
'' Select this object.
'Private Sub SelectVbdObject(ByVal target As vbdObject)
'    ' See if it is aleady selected.
'    If target.Selected Then Exit Sub
'
'    ' Add the object to the
'    ' m_SelectedObjects collection.
'    m_SelectedObjects.Add target
'    Debug.Print target.Serialization
'    ' Mark the object as selected.
'    target.Selected = True
'End Sub


'' Find the object at this position.
'Private Function FindObjectAt(ByVal X As Single, ByVal Y As Single) As vbdObject
'    Dim the_scene As vbdScene
'
'    Set the_scene = m_TheScene
'    Set FindObjectAt = the_scene.FindObjectAt(X, Y)
'End Function

' Add this file name to the MRU list.
Private Sub MruAddName(ByVal File_name As String)
Dim I As Integer

    ' Remove any duplicates.
    For I = m_MruList.Count To 1 Step -1
        If m_MruList(I) = File_name Then
            m_MruList.Remove I
        End If
    Next I

    ' Add the new name at the front.
    If m_MruList.Count = 0 Then
        m_MruList.Add File_name
    Else
        m_MruList.Add File_name, , 1
    End If

    ' Only keep 4.
    Do While m_MruList.Count > 4
        m_MruList.Remove 5
    Loop

    ' Save the MRU list in the registry.
    For I = 1 To m_MruList.Count
        SaveSetting App.Title, "MRU", Format$(I), m_MruList(I)
    Next I
    For I = m_MruList.Count + 1 To 4
        SaveSetting App.Title, "MRU", Format$(I), ""
    Next I

    ' Display the MRU list.
    MruDisplay
End Sub
' Display the MRU list.
Private Sub MruDisplay()
Dim I As Integer

    mnuFileMRU(0).Visible = (m_MruList.Count > 0)
    For I = 1 To m_MruList.Count
        If I > mnuFileMRU.UBound Then
            Load mnuFileMRU(I)
        End If
        mnuFileMRU(I).Caption = "&" & _
            Format$(I) & " " & m_MruList(I)
        mnuFileMRU(I).Visible = True
    Next I
End Sub
' Load the MRU list.
Private Sub MruLoad()
Dim I As Integer
Dim File_name As String

    Set m_MruList = New Collection
    For I = 1 To 4
        File_name = GetSetting(App.Title, "MRU", _
            Format$(I), "")
        If Len(File_name) > 0 Then
            m_MruList.Add File_name
        End If
    Next I

    ' Display the list.
    MruDisplay
End Sub

'' Select default values and prepare to edit.
'Private Sub PrepareToEdit()
'    ' Select default colors.
'    picForeColorSample_Click 0  ' Black
'    picbackColorSample_Click 7  ' Gray
'
' '   m_CurrentSnapshot = 0
'    SaveSnapshot
'
'    ' Start at normal (pixel) scale.
'    picCanvas.ScaleMode = vbPixels
'
'    ' Select the arrow tool.
'    tbrTools.Buttons("Arrow").Value = tbrPressed
'
'    ' Select the solid DrawStyle.
'    icbDrawStyle.SelectedItem = icbDrawStyle.ComboItems(1)
'
'    ' Select the solid FillStyle.
'    icbFillStyle.SelectedItem = icbDrawStyle.ComboItems(1)
'
'    ' Select the 1 pixel DrawWidth.
'    icbDrawWidth.SelectedItem = icbDrawStyle.ComboItems(1)
'
'    ' Redraw.
'    'picCanvas.Refresh
'    DrawControl1.Redraw
'End Sub

' Flag the data as modified.
Private Sub SetDirty()
    If Not m_DataModified Then
        Caption = App.Title & "*[" & DrawControl1.FileTitle & "]"
    End If

    ' Save the current snapshot.
    SaveSnapshot
     
    m_DataModified = True
End Sub

' Set the file's name.
Private Sub SetFileName(ByVal File_name As String, ByVal file_title As String)
    ' Save the file's name and title.
    DrawControl1.Filename = File_name
    DrawControl1.FileTitle = file_title
    mnuFileSave.Enabled = Len(DrawControl1.FileTitle) > 0

    ' Update the caption.
    Caption = App.Title & " [" & DrawControl1.FileTitle & "]"

    ' Add the name to the MRU list.
    If Len(DrawControl1.Filename) > 0 Then MruAddName DrawControl1.Filename
End Sub

'' Enable or disable the undo and redo menus.
'Public Sub SetUndoMenus()
'    mnuEditUndo.Enabled = (m_CurrentSnapshot > 1)
'    mnuEditRedo.Enabled = (m_CurrentSnapshot < m_Snapshots.Count)
'End Sub
'
'' Save a snapshot for undo.
'Private Sub SaveSnapshot()
''    ' Remove any previously undone snapshots.
''    Do While m_Snapshots.Count > m_CurrentSnapshot
''        m_Snapshots.Remove m_Snapshots.Count
''    Loop
''
''    ' Save the current snapshot.
''    m_Snapshots.Add m_TheScene.Serialization
''    If m_Snapshots.Count > MAX_UNDO + 1 Then
''        m_Snapshots.Remove 1
''    End If
''    m_CurrentSnapshot = m_Snapshots.Count
'
'    ' Enable/disable the undo and redo menus.
'    SetUndoMenus
'End Sub
''
'' Add this object to the collection.
'Public Sub AddObject(ByVal obj As vbdObject)
'Dim the_scene As vbdScene
'
'    ' Give the object its drawing properties.
'    obj.ForeColor = QBColor(m_ForeColor)
'    obj.FillColor = QBColor(m_BackColor)
'    obj.DrawStyle = icbDrawStyle.SelectedItem.Index - 1
'    obj.FillStyle = icbFillStyle.SelectedItem.Index - 1
'    obj.DrawWidth = icbDrawWidth.SelectedItem.Index
'
'    ' Save the new object.
'    Set the_scene = m_TheScene
'    the_scene.SceneObjects.Add obj
'    Set m_NewObject = Nothing
'
'    ' Select the new object only.
'    DeselectAllVbdObjects
'    SelectVbdObject obj
'
'    ' See if any objects are selected.
'    EnableMenusForSelection
'
'    ' Select the arrow tool.
'    SelectArrowTool
'
'    ' The data has changed.
'    SetDirty
'
'    ' Redraw.
'    picCanvas.Refresh
'End Sub
' Cancel adding an object to the collection.
Public Sub CancelObject()
    Set m_NewObject = Nothing

    ' Select the arrow tool.
    SelectArrowTool
End Sub

' Restore the previous snapshot.
Private Sub Undo()
Dim token_name As String
Dim token_value As String

    If m_CurrentSnapshot <= 1 Then Exit Sub

    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot - 1
    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value

    
    ' Enable/disable the undo and redo menus.
    DrawControl1_EnableMenusForSelection 'SetUndoMenus
    DrawControl1.Redraw
End Sub

' Reapply a previously undone snapshot.
Private Sub Redo()
Dim token_name As String
Dim token_value As String
    
    If m_Snapshots Is Nothing Then Exit Sub
    
    If m_CurrentSnapshot >= m_Snapshots.Count Then Exit Sub

    ' Restore the previous snapshot.
    m_CurrentSnapshot = m_CurrentSnapshot + 1

    GetNamedToken m_Snapshots(m_CurrentSnapshot), token_name, token_value
    m_TheScene.Serialization = token_value
    
    ' Enable/disable the undo and redo menus.
    DrawControl1_EnableMenusForSelection 'SetUndoMenus
    DrawControl1.Redraw
End Sub




Private Sub ColorPalette1_ColorOver(cColor As Long)
    Dim sTmp As String
    sTmp = Right("000000" & Hex(cColor), 6)
    sTmp = "R:" + Str(Int("&H" & Right$(sTmp, 2))) + " - G:" + Str(Int("&H" & Mid$(sTmp, 3, 2))) + " - B:" + Str(Int("&H" & Left$(sTmp, 2)))
    StatusBar1.Panels(3).Text = sTmp
End Sub

Private Sub ColorPalette1_ColorSelected(Button As Integer, cColor As Long)
    If cColor <> -1 Then
       Select Case Button
       Case 1
           CtrColor1.ColorFill = cColor
           DrawControl1.FillStyle = 0
       Case 2
           CtrColor1.ColorBorder = cColor
           DrawControl1.DrawStyle = 0
       End Select
      
       CtrColor1.Redraw
       DrawControl1.ForeColor = CtrColor1.ColorBorder
       DrawControl1.FillColor = CtrColor1.ColorFill
       
       StatusBar1.Panels(4).Picture = LoadPicture()
       StatusBar1.Panels(4).Picture = CtrColor1.Image
    Else
       Select Case Button
       Case 1
           DrawControl1.FillStyle = 1
       Case 2
           DrawControl1.DrawStyle = 5
       End Select
    End If
    DrawControl1.Redraw
    
End Sub


Private Sub ComboZoom_Click()
    Dim mZoom As Single
    If gZoomLock = True Then Exit Sub
    If ComboZoom.ListIndex = -1 Then Exit Sub
    mZoom = Val(ComboZoom.Text) / 100
    If mZoom > 4 Then mZoom = 4
    If mZoom < 0.1 Then mZoom = 0.1
     DrawControl1.ZoomFactor = mZoom
     StatusBar1.Panels(1).Text = "Zoom (" & Round(mZoom * 100) & "%)"
End Sub


Private Sub DrawControl1_ColorSelected(tColor As Integer, cColor As Long)
    Select Case tColor
    Case 1
       CtrColor1.ColorFill = cColor
    Case 2
       CtrColor1.ColorBorder = cColor
    End Select
    CtrColor1.Redraw
    StatusBar1.Panels(4).Picture = LoadPicture()
    StatusBar1.Panels(4).Picture = CtrColor1.Image
End Sub

' Enable the appropriate transformation menus.
Public Sub DrawControl1_EnableMenusForSelection()

Dim objects_selected As Boolean

    objects_selected = (m_SelectedObjects.Count > 0)
    
    mnuArrangeSendToFront.Enabled = objects_selected
    mnuArrangeSendToBack.Enabled = objects_selected
    mnuArrangeSendToForward.Enabled = objects_selected
    mnuArrangeSendToBackward.Enabled = objects_selected
    
    mnuTransformClear.Enabled = objects_selected
    mnuTransformRotate.Enabled = objects_selected
    mnuTransformScale.Enabled = objects_selected
    mnuskew.Enabled = objects_selected
    mnuReflect.Enabled = objects_selected
    mnuMove.Enabled = objects_selected
    
    drawToolbar2.EnableButton 3, objects_selected
    drawToolbar2.EnableButton 4, objects_selected
    drawToolbar2.EnableButton 5, objects_selected
    drawToolbar2.EnableButton 6, objects_selected
    drawToolbar2.CheckButton 7, DrawControl1.LockObject
    
    mnuEditUndo.Enabled = (m_CurrentSnapshot > 1)
    mnuEditRedo.Enabled = (m_CurrentSnapshot < m_Snapshots.Count)
     
    drawToolbar1.EnableButton 9, mnuEditUndo.Enabled
    drawToolbar1.EnableButton 10, mnuEditRedo.Enabled
    
    mnuEdit_Click
    
End Sub

Private Sub DrawControl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      StatusBar1.Panels(2).Text = "X:" + Format(X, "0.00") + " Y:" + Format(Y, "0.00")
      If DrawControl1.XmaxBox - DrawControl1.XminBox <> 0 Or DrawControl1.YmaxBox - DrawControl1.YminBox <> 0 Then
         StatusBar1.Panels(3).Text = "X:" + Format(DrawControl1.XminBox, "0.00") + " Y:" + Format(DrawControl1.YminBox, "0.00") + "  W:" + Format(DrawControl1.XmaxBox - DrawControl1.XminBox, "0.00") + _
                                      "   H:" + Format(DrawControl1.YmaxBox - DrawControl1.YminBox, "0.00")
      Else
         StatusBar1.Panels(3).Text = ""
      End If
End Sub

Private Sub DrawControl1_MsgControl(txt As String)
     StatusBar1.Panels(1).Text = txt
End Sub

Public Sub DrawControl1_SetDirty()
      SetDirty
      m_DataModified = True
End Sub

Private Sub drawToolbar_ButtonCheck(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
     Debug.Print "ButtonCheck", Index, xLeft, yTop
End Sub

Private Sub drawToolbar_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    
    Debug.Print "ButtonClick", Index, MouseButton, xLeft, yTop
    Debug.Print drawToolbar.GetToolTips(Index)
    DrawControl1.SelectTool Index
    
    If drawToolbar.GetToolTips(Index) = "Pen" Or drawToolbar.GetToolTips(Index) = "Fill" Or Index = 2 Then
         ' Select the arrow tool.
          SelectArrowTool
    End If
End Sub

Private Sub drawToolbar1_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Select Case Index
    Case 1
       mnuFileNew_Click
    Case 2
       mnuFileOpen_Click
    Case 3
       mnuFileSave_Click
    Case 4
       mnuFileSaveBitmap_Click
    Case 5
       mnuPrint_Click
    Case 6
       DrawControl1.CutObject
    Case 7
       DrawControl1.CopyObject
    Case 8 '
        DrawControl1.PasteObject
    Case 9
       Undo
    Case 10
       Redo
    Case 11
       mnuDelete_Click
    Case 12
       mnuSymbol_Click
    Case 13
      mnupenform_Click
    Case 14
      mnufillform_Click
    Case 15
      mnutransformform_Click
    End Select
End Sub

Private Sub drawToolbar2_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    Select Case Index
    Case 1
        DrawControl1.SelectAllObject
    Case 2
       DrawControl1.UnSelectAllObject
    Case 3
        DrawControl1.SetObjectOrder BringToFront
    Case 4
        DrawControl1.SetObjectOrder SendToBack
    Case 5
        DrawControl1.SetObjectOrder BringFoward
    Case 6
        DrawControl1.SetObjectOrder SendBackward
    Case 7
        DrawControl1.LockObject = Not DrawControl1.LockObject
    Case 8
        'DrawControl1.GroupObjects
    Case 9
        'DrawControl1.UnGroupObjects
    Case 10
        'DrawControl1.AlignSelectedObjects mLeft
    Case 11
        'DrawControl1.AlignSelectedObjects mCenterV
    Case 12
        'DrawControl1.AlignSelectedObjects mRight
    Case 13
        'DrawControl1.AlignSelectedObjects mTop
    Case 14
        'DrawControl1.AlignSelectedObjects mCenterH
    Case 15
        'DrawControl1.AlignSelectedObjects mBottom
    Case 16
        'DrawControl1.AlignSelectedObjects mCenterVH
    End Select
   
End Sub

Private Sub drawToolbar3_ButtonClick(ByVal Index As Long, ByVal MouseButton As Integer, ByVal xLeft As Long, ByVal yTop As Long)
    If gZoomLock = True Then Exit Sub
    Select Case Index
    Case 1
        DrawControl1.ZoomFactor = 1
    Case 2
        DrawControl1.ZoomFactor = DrawControl1.ZoomFactor - 0.1
    Case 3
        DrawControl1.ZoomFactor = DrawControl1.ZoomFactor + 0.1
    End Select
    ComboZoom.Text = Str(Round(DrawControl1.ZoomFactor * 100)) & " %"
    StatusBar1.Panels(1).Text = "Zoom (" & Round(DrawControl1.ZoomFactor * 100) & "%)"
End Sub

' Process key presses.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim the_scene As vbdScene
'
'    Select Case KeyCode
'        Case vbKeyDelete
'            If m_SelectedObjects.Count > 0 Then
'                ' Delete the selected objects.
'                Set the_scene = m_TheScene
'                the_scene.RemoveObjects m_SelectedObjects
'
'                ' The data has changed.
'                SetDirty
'                PicCanvas.Refresh
'            End If
'    End Select
End Sub

Private Sub Form_Load()
    
    InitScreen
'
'    frmAbout.ShowForm
'    Unload frmAbout
    ' Load the MRU list.
    MruLoad

    ' Prepare the dialog.
    dlgFile.CancelError = True
    dlgFile.InitDir = App.Path
    
    ' Start a new picture.
    'mnuFileNew_Click
    Me.Width = Screen.Width - 1000
    Me.Height = Screen.Height - 1000
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'FrmMagnify.Show
    'FormOnTop FrmMagnify, True
    'Me.WindowState = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = (Not DataSafe())
    If m_FormSymbolView = True Then Unload FrmSymbols
    'If m_FormMagnify = True Then Unload FrmMagnify
    If Cancel = 0 Then Clipboard.Clear

End Sub


Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single
  
    wid = ScaleWidth - drawToolbar.Width
    If wid < 3000 Then wid = 3000
    hgt = ScaleHeight - StatusBar1.Height - CoolBar1.Height
    If hgt < 3000 Then hgt = 3000
    
    
    DrawControl1.Move drawToolbar.Width, CoolBar1.Height, wid, hgt - ColorPalette1.Height
    If Me.WindowState = 1 Then
       If m_FormSymbolView = True Then
          FrmSymbols.Hide
       End If
    Else
       If m_FormSymbolView = True Then
          FrmSymbols.Show
       End If
    End If
    
End Sub




Private Sub mnuabout_Click()
     frmAbout.ShowForm True
     Unload frmAbout
End Sub

' Move this object to the front of the scene's
' object list.
Private Sub mnuArrangeSendToBack_Click()
    
   DrawControl1.SetObjectOrder SendToBack
    
End Sub

' Move this object to the Backward of the scene's
' object list.
Private Sub mnuArrangeSendToBackward_Click()
    
    DrawControl1.SetObjectOrder SendBackward
    
End Sub
' Move this object Send To Forward of the scene's
' object list.
Private Sub mnuArrangeSendToForward_Click()

   DrawControl1.SetObjectOrder BringFoward
    
End Sub

' Move this object Bring To Front of the scene's
' object list.
Private Sub mnuArrangeSendToFront_Click()
       
    DrawControl1.SetObjectOrder BringToFront
    
End Sub

Private Sub mnuclear_Click()
    DrawControl1.ClearObject
End Sub

Private Sub MnuCopy_Click()
     DrawControl1.CopyObject
End Sub

Private Sub mnuCut_Click()
    DrawControl1.CutObject
End Sub

Private Sub mnuDelete_Click()
    DrawControl1.DelObject
End Sub

Private Sub mnuEdit_Click()
    Dim mnuenabled1 As Boolean
    Dim mnuenabled2 As Boolean
    
    mnuenabled1 = DrawControl1.IsSelectObject
    mnuenabled2 = FindObject(Clipboard.GetText)
    
'      If DrawControl1.IsSelectObject = True Then
         mnuCut.Enabled = mnuenabled1
         MnuCopy.Enabled = mnuenabled1
         mnuDelete.Enabled = mnuenabled1
'      Else
'         mnuCut.Enabled = False
'         MnuCopy.Enabled = False
'         mnuDelete.Enabled = False
'      End If
              
'      If FindObject(Clipboard.GetText) Then
         mnupaste.Enabled = mnuenabled2 'True
         mnuclear.Enabled = mnuenabled2 'True
'      Else
'         mnupaste.Enabled = False
'         mnuclear.Enabled = False
'      End If
      
      drawToolbar1.EnableButton 6, mnuenabled1 ' mnuCut.Enabled
      drawToolbar1.EnableButton 7, mnuenabled1 'MnuCopy.Enabled
      drawToolbar1.EnableButton 8, mnuenabled2 'mnupaste.Enabled
      drawToolbar1.EnableButton 11, mnuenabled1 ' mnuDelete.Enabled
End Sub

Private Sub mnuEditRedo_Click()
    Redo
End Sub

Private Sub mnuEditUndo_Click()
    Undo
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' Load the selected file.
Private Sub mnuFileMRU_Click(Index As Integer)
Dim pos As Integer
Dim file_title As String

    If Not DataSafe() Then Exit Sub

    pos = InStrRev(m_MruList(Index), "\")
    file_title = Mid$(m_MruList(Index), pos + 1)
    DataLoad m_MruList(Index), file_title
End Sub

' Start a new picture.
Private Sub mnuFileNew_Click()
    If Not DataSafe() Then Exit Sub

    'New draw
    DrawControl1.NewDraw True
     
    'Blank the file name.
    SetFileName "", ""

    'The data has not been modified.
    m_DataModified = False

    ' Prepare to edit.
    DrawControl1.PrepareToEdit
End Sub

' Load a file.
Private Sub mnuFileOpen_Click()
Dim File_name As String

    dlgFile.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNFileMustExist
    dlgFile.Filter = "ArtDraw Files (*.adrw)|*.adrw|" & "All Files (*.*)|*.*"
    If PathExists(App.Path + "\Samples") = False Then MkDir App.Path + "\Samples"
    dlgFile.InitDir = App.Path + "\Samples"
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If

    File_name = dlgFile.Filename
    dlgFile.InitDir = Left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
    If DrawControl1.OpenDraw(File_name, dlgFile.FileTitle) Then
       ' Update the caption.
       SetFileName File_name, dlgFile.FileTitle
        ' Prepare to edit.
       DrawControl1.PrepareToEdit

       DrawControl1.Redraw
    End If
End Sub

' Save the data using the current file name.
Private Sub mnuFileSave_Click()
    If Len(DrawControl1.Filename) = 0 Then
        ' There is no file name. Use Save As.
        mnuFileSaveAs_Click
    Else
        ' Save the data.
        If DrawControl1.SaveDraw(DrawControl1.Filename, DrawControl1.FileTitle) Then
           ' Update the caption.
           SetFileName DrawControl1.Filename, DrawControl1.FileTitle
           'SetFileName file_name, file_title

        End If
    End If
End Sub
' Save the picture with a new file name.
Private Sub mnuFileSaveAs_Click()
Dim File_name As String

    dlgFile.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
    
    If PathExists(App.Path + "\Samples") = False Then MkDir App.Path + "\Samples"
    dlgFile.InitDir = App.Path + "\Samples"
    dlgFile.Filename = DrawControl1.Filename
    dlgFile.Filter = "ArtDraw Files (*.adrw)|*.adrw|" & "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & " selecting file." & vbCrLf & Err.Description
        Exit Sub
    End If

    File_name = dlgFile.Filename
    dlgFile.InitDir = Left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
    'DataSave file_name, dlgFile.FileTitle
   If DrawControl1.SaveDraw(File_name, dlgFile.FileTitle) Then
     ' Update the caption.
      SetFileName DrawControl1.Filename, DrawControl1.FileTitle
   End If
End Sub

' Save a bitmap image.
Private Sub mnuFileSaveBitmap_Click()
Dim old_file_name As String
Dim pos As Integer
Dim File_name As String
Dim fDrive As String, fPath As String, fFileName As String, fFile As String, fExtension As String
                     
    old_file_name = dlgFile.Filename
    pos = InStrRev(old_file_name, ".")
    If pos > 0 Then dlgFile.Filename = Left$(old_file_name, pos) & "bmp"

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt
    dlgFile.Filter = "Bitmap Files (*.bmp)|*.bmp|" + _
                     "Graphics Interchange Format(*.gif)|*.gif|" + _
                     "Tagged Image Format(*.tif)|*.tif|" + _
                     "Portable Network Graphics(*.png)|*.png|" + _
                     "Joint Photographic Experts Group(*.jpg)|*.jpg|Metafiles (*.wmf)|*.wmf|" & _
                     "All Files (*.*)|*.*"
    If PathExists(App.Path + "\Export") = False Then MkDir App.Path + "\Export"
    dlgFile.InitDir = App.Path + "\Export"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    File_name = dlgFile.Filename
    dlgFile.InitDir = Left$(File_name, Len(File_name) - Len(dlgFile.FileTitle) - 1)
    SplitPath File_name, fDrive, fPath, fFileName, fFile, fExtension
    
    Select Case dlgFile.FilterIndex
    Case 1 'bmp
         fExtension = ".bmp"
    Case 2 'gif
         fExtension = ".gif"
    Case 3 'tif
         fExtension = ".tif"
    Case 4 'png
         fExtension = ".png"
    Case 5 'jpg
         fExtension = ".jpg"
    Case 6 'wmf
         fExtension = ".wmf"
    Case Else
         MsgBox "Error Extension " & File_name & " Not Saved!"
         Exit Sub
    End Select
    If Left(fExtension, 1) <> "." Then fExtension = "." + fExtension
    File_name = fPath + "\" + fFile + fExtension
    
    If dlgFile.FilterIndex = 6 Then
       DrawControl1.FileExport File_name
    Else
       DrawControl1.FileExportBitmap File_name
    End If
    
    dlgFile.Filename = old_file_name
End Sub

' Save the objects in a metafile.
Private Sub mnuFileSaveMetafile_Click()
Dim old_file_name As String
Dim pos As Integer
Dim File_name As String
'Dim mf_dc As Long
'Dim hmf As Long
'Dim old_size As size

    old_file_name = dlgFile.Filename
    pos = InStrRev(old_file_name, ".")
    If pos > 0 Then dlgFile.Filename = Left$(old_file_name, pos) & "wmf"

    dlgFile.Flags = cdlOFNExplorer Or _
        cdlOFNHideReadOnly Or _
        cdlOFNLongNames Or _
        cdlOFNOverwritePrompt
    dlgFile.Filter = "Metafiles (*.wmf)|*.wmf|" & _
        "All Files (*.*)|*.*"
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    File_name = dlgFile.Filename
    dlgFile.InitDir = Left$(File_name, Len(File_name) _
        - Len(dlgFile.FileTitle) - 1)

'    ' Create the metafile.
'    mf_dc = CreateMetaFile(ByVal file_name)
'    If mf_dc = 0 Then
'        MsgBox "Error creating the metafile.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Set the metafile's size to something reasonable.
'    SetWindowExtEx mf_dc, picCanvas.ScaleWidth, _
'        picCanvas.ScaleHeight, old_size
'
'    ' Draw in the metafile.
'    m_TheScene.DrawInMetafile mf_dc
'
'    ' Close the metafile.
'    hmf = CloseMetaFile(mf_dc)
'    If hmf = 0 Then
'        MsgBox "Error closing the metafile.", vbExclamation
'    End If
'
'    ' Delete the metafile to free resources.
'    If DeleteMetaFile(hmf) = 0 Then
'        MsgBox "Error deleting the metafile.", vbExclamation
'    End If

    dlgFile.Filename = old_file_name
End Sub


Private Sub mnufillform_Click()
       Open_Form 13 '"Fill"
End Sub

Private Sub mnuImport_Click()
    Dim old_file_name As String
    Dim pos As Integer
    Dim File_name As String ', file_Drive As String, file_path As String, file_filename As String, file_ext As String
    Dim fDrive As String, fPath As String, fFileName As String, fFile As String, fExtension As String
                     
    'old_file_name = dlgFile.FileName
    'pos = InStrRev(old_file_name, ".")
    'If pos > 0 Then dlgFile.FileName = Left$(old_file_name, pos) & "bmp"

    dlgFile.Flags = cdlOFNExplorer Or _
                    cdlOFNHideReadOnly Or _
                    cdlOFNLongNames Or _
                    cdlOFNOverwritePrompt
    dlgFile.Filter = "Bitmap Files (bmp,gif,tif,png,jpg)|*.bmp;*.gif;*.tif;*.png;*.jpg"
    
                     '"Graphics Interchange Format(*.gif)|*.gif|" + _
                     '"Tagged Image Format(*.tif)|*.tif|" + _
                     '"Portable Network Graphics(*.png)|*.png|" + _
                     '"Joint Photographic Experts Group(*.jpg)|*.jpg|" & _
                     '"All Files (*.*)|*.*"
    If PathExists(App.Path + "\Object") = False Then MkDir App.Path + "\Object"
    dlgFile.InitDir = App.Path + "\Object"
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number = cdlCancel Then
        Exit Sub
    ElseIf Err.Number <> 0 Then
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If

    File_name = dlgFile.Filename
    dlgFile.InitDir = Left$(File_name, Len(File_name) _
        - Len(dlgFile.FileTitle) - 1)

    'LoadPicBox file_name, PicCanvas
              
End Sub

Private Sub mnuMove_Click()
      DrawControl1.ViewTransform 0
End Sub

Private Sub Mnunormal_Click()
     m_ViewSimple = False
     Mnunormal.Checked = True
     mnuSimpleWireframe.Checked = False
     DrawControl1.Redraw
End Sub

Private Sub mnupaste_Click()
     DrawControl1.PasteObject
End Sub

Private Sub mnupenform_Click()
     Open_Form 12 '"Pen"
End Sub

Private Sub mnuPrint_Click()
     DrawControl1.PrintDraw
    
End Sub

Private Sub mnuprintersetup_Click()
      dlgFile.Flags = cdlPDNoSelection Or cdlPDHidePrintToFile Or _
                      cdlPDNoWarning 'Or cdlPDReturnDefault 'Or cdlPDUseDevModeCopies
      dlgFile.ShowPrinter
     ' Printer.Orientation = dlgFile.Orientation
End Sub

Private Sub mnuReflect_Click()
     DrawControl1.ViewTransform 4
End Sub

Private Sub mnuSimpleWireframe_Click()
     m_ViewSimple = True
     mnuSimpleWireframe.Checked = True
     Mnunormal.Checked = False
     DrawControl1.Redraw
End Sub

Private Sub mnuskew_Click()
     DrawControl1.ViewTransform 3
End Sub

Private Sub mnuSymbol_Click()
     m_FormSymbolView = True
     FrmSymbols.Show
End Sub

' Clear the selected objects' transformations.
Private Sub mnuTransformClear_Click()
    DrawControl1.ClearTransform
   ' DrawControl1.Set_Dirty
End Sub

Private Sub mnutransformform_Click()
      DrawControl1.ViewTransform 0
End Sub

' Rotate the selected objects.
Private Sub mnuTransformRotate_Click()
'Const PI = 3.14159265
    DrawControl1.ViewTransform 1
    Exit Sub
Dim txt As String
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
    txt = InputBox("Angle (degrees)", "Angle", "")
    txt = Trim$(txt)
    If Len(txt) = 0 Then Exit Sub
    If Not IsNumeric(txt) Then Exit Sub
    Angle = CSng(txt) * PI / 180

    ' Bound the selected objects.
    BoundObjects m_SelectedObjects, xmin, ymin, xmax, ymax

    ' Make the transformation matrix.
    xmid = (xmin + xmax) / 2
    ymid = (ymin + ymax) / 2
    m2RotateAround M, Angle, xmid, ymid

    ' Add the transformation to the selected objects.
    For Each Obj In m_SelectedObjects
        Obj.AddTransformation M
    Next Obj

    ' The data has changed.
    SetDirty
    
End Sub

' Let the user scale the selected objects.
Private Sub mnuTransformScale_Click()
 
     DrawControl1.ViewTransform 2
     
End Sub


Private Sub InitScreen()
Dim I As Integer

    ComboZoom.AddItem "10 %"
    ComboZoom.AddItem "25 %"
    ComboZoom.AddItem "50 %"
    ComboZoom.AddItem "100 %"
    ComboZoom.AddItem "150 %"
    ComboZoom.AddItem "200 %"
    ComboZoom.AddItem "400 %"
    ComboZoom.ListIndex = 3
    
    ColorPalette1_ColorSelected 1, 16777215
    ColorPalette1_ColorSelected 2, 0
    
    drawToolbar.BarOrientation = tbVertical
    drawToolbar.BuildToolbar PicTools.Picture, vbButtonFace, 16, "OOOOOOOOOOOOO"
    drawToolbar.SetTooltips "Arrow|Point|Polyline|FreePolygon|Free Line|Free Line Closed|Curve|RectAngle|Polygon|Ellipse|Text Art|Pen|Fill"
    drawToolbar.CheckButton 1, True
    
    drawToolbar2.BarOrientation = tbHorizontal
    drawToolbar2.BuildToolbar PicTollBar2.Picture, vbButtonFace, 16, "NN|NNNN|C|NN"
    drawToolbar2.SetTooltips "Select all|UnselectAll|Bring to front|Send to back|Bring Forward|Send Backward|Lock|Group|Ungroup"
    For I = 3 To 8
       If I <> 7 Then
       drawToolbar2.EnableButton I, False
       End If
    Next
    
    drawToolbar1.BarOrientation = tbHorizontal
    drawToolbar1.BuildToolbar PicTollBar1.Picture, vbButtonFace, 16, "NNNNN|NNN|NN|N|NNNN"
    drawToolbar1.SetTooltips "New draw|Open draw|Save draw|Export draw|Print|Cut|Copy|Paste|Undo|Redo|Delete|Symbol|Pen|Fill|Transforming"
    For I = 6 To 11
       drawToolbar1.EnableButton I, False
    Next
    
    drawToolbar3.BarOrientation = tbHorizontal
    drawToolbar3.BuildToolbar PicTollBar3.Picture, vbButtonFace, 16, "NNN"
    drawToolbar3.SetTooltips "Zoom All|Zoom+|Zoom-"
    
    drawToolbar.BarEdge = True
    drawToolbar1.BarEdge = True
    drawToolbar2.BarEdge = True
    drawToolbar3.BarEdge = True
    
    drawToolbar2.EnableButton 8, False
    drawToolbar2.EnableButton 9, False
     
End Sub
   

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        DrawControl1.ZoomFactor = 1
    Case 2
        DrawControl1.ZoomFactor = DrawControl1.ZoomFactor - 0.1
    Case 3
        DrawControl1.ZoomFactor = DrawControl1.ZoomFactor + 0.1
    End Select
    ComboZoom.Text = Str(Round(DrawControl1.ZoomFactor * 100)) & " %"
    StatusBar1.Panels(1).Text = "Zoom (" & Round(DrawControl1.ZoomFactor * 100) & "%)"
End Sub

' Select the arrow tool.
Public Sub SelectArrowTool()

    ' Make sure the arrow button is pressed.
    'tbrTools.Buttons("Arrow").Value = tbrPressed
    drawToolbar.CheckButton 1, True
    ' Prepare to deal with this tool.
    DrawControl1.SelectTool 1 '"Arrow"
End Sub

Sub Open_Form(Index As Integer) ' nameform As String)
    DrawControl1.SelectTool Index ' nameform
    
    'If nameform = "Pen" Or nameform = "Fill" Then
    If Index = 12 Or Index = 13 Then
          SelectArrowTool
    End If
End Sub


