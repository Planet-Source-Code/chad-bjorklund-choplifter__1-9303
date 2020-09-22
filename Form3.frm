VERSION 5.00
Begin VB.Form CollisionTest 
   Caption         =   "Form3"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form3"
   ScaleHeight     =   677
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "3"
      Height          =   375
      Left            =   11280
      TabIndex        =   18
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2"
      Height          =   375
      Left            =   11280
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1"
      Height          =   375
      Left            =   11280
      TabIndex        =   16
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox BlankCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   10200
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   15
      Top             =   3720
      Width           =   90
   End
   Begin VB.PictureBox B 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   1200
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   1200
      Width           =   90
      Begin VB.Shape TB 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   90
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   255
      Left            =   8880
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Y2 
      Height          =   285
      Left            =   8400
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox X2 
      Height          =   285
      Left            =   7320
      TabIndex        =   8
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Y1 
      Height          =   285
      Left            =   8400
      TabIndex        =   7
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox X1 
      Height          =   285
      Left            =   7320
      TabIndex        =   6
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Check"
      Height          =   615
      Left            =   7440
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move"
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.PictureBox Blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1200
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   2
      Top             =   1440
      Width           =   90
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   3600
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   1200
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   1680
      Width           =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "Y2"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "X2"
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Y1"
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "X1"
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Collision"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "CollisionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BX As Single
Dim BWY As Single
Dim PX As Single
Dim PY As Single

Private Sub Command1_Click()
B.Left = X1.Text
B.Top = Y1.Text
P.Left = X2.Text
P.Top = Y2.Text
End Sub

Private Sub Command2_Click()
    'BitBlt BlankCheck.hdc, 0, 0, 6, 6, Blank.hdc, 0, 0, vbSrcCopy
    BX = X1.Text
    BWY = Y1.Text
    PX = X2.Text
    PY = Y2.Text
    If CollisionDetect(BX, BWY, B, PX, PY, P, Blank) Then
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If
    
    
    
    
End Sub

Private Sub Command3_Click()
    B.Left = 80
    B.Top = 80
    Blank.Left = 80
    Blank.Top = 96
    P.Left = 80
    P.Top = 112
End Sub

Private Sub Command4_Click()
BitBlt BlankCheck.hdc, 0, 0, BlankCheck.ScaleWidth, BlankCheck.ScaleHeight, BlankCheck.hdc, 0, 0, vbNotSrcCopy
End Sub

Private Sub Command5_Click()
BitBlt BlankCheck.hdc, 0, 0, 6, 6, B.hdc, 0, 0, vbSrcPaint
End Sub

Private Sub Command6_Click()
BitBlt BlankCheck.hdc, 0, 0, 200, 130, P.hdc, 184, 168, vbSrcPaint
End Sub

Private Sub Form_Load()
    B.Width = 6
    B.Height = 6
  
End Sub

