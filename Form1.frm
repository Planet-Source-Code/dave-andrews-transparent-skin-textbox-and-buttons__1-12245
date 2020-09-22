VERSION 5.00
Object = "*\ATextPic.vbp"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Journal"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TextPic.ucBox ucBox1 
      Height          =   3255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65535
      ScaleHeight     =   213
      ScaleMode       =   0
      ScaleWidth      =   413
      Text            =   "This is a textbox"
      Text            =   "This is a textbox"
   End
   Begin VB.PictureBox btnClear 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   615
      Left            =   360
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox btnExit 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Matisse ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6240
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim bExit As New clsButton
Dim bClear As New clsButton

Private Sub btnClear_Click()
ucBox1.Text = ""
ucBox1.PaintText
End Sub

Private Sub btnClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bClear.TriggerDown
End Sub


Private Sub btnClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bClear.TriggerUp
End Sub


Private Sub btnExit_Click()
Unload Me
End Sub


Private Sub btnExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bExit.TriggerDown
End Sub


Private Sub btnExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bExit.TriggerUp
End Sub


Private Sub Form_Load()
Me.Picture = LoadPicture("Default.jpg")
CutRR Me, 16
bClear.InitButton btnClear, "CLEAR", True, 0.4, 12, 4, True, 7
bExit.InitButton btnExit, "X", True, 0.4, 9, 9, True, 2
BitBlt ucBox1.hDC, 0, 0, ucBox1.Width, ucBox1.Height, Me.hDC, ucBox1.Left, ucBox1.Top, vbSrcCopy
Set ucBox1.Picture = ucBox1.Image
ucBox1.PaintText
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartMove Me
End Sub


