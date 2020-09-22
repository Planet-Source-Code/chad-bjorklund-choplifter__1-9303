VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   DrawMode        =   7  'Invert
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   864
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1152
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   13200
      Top             =   10560
   End
   Begin Project1.ProgBar PBar 
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      BackColour      =   16776960
      BarStartColour  =   16776960
      BarEndColour    =   255
      BorderStyle     =   0
      Max             =   25
      Message         =   "Hits Left"
      ShowMessage     =   -1  'True
      ShowValue       =   -1  'True
      BarStyle        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarEndColour    =   255
   End
   Begin VB.PictureBox B 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   750
      ScaleHeight     =   75
      ScaleWidth      =   90
      TabIndex        =   6
      Top             =   -1950
      Width           =   90
      Begin VB.Shape TB 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   90
         Index           =   51
         Left            =   0
         Shape           =   2  'Oval
         Top             =   0
         Width           =   90
      End
   End
   Begin VB.PictureBox Blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   4920
      ScaleHeight     =   90
      ScaleWidth      =   90
      TabIndex        =   5
      Top             =   -1950
      Width           =   90
   End
   Begin VB.PictureBox Pl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   1320
      Picture         =   "NewMethod.frx":0000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   -1950
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   17235
      TabIndex        =   3
      Top             =   12585
      Width           =   17295
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2160
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   -1080
      Width           =   330
   End
   Begin VB.PictureBox picRope 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   1680
      Picture         =   "NewMethod.frx":26E2
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   -1080
      Width           =   330
   End
   Begin VB.PictureBox picGuy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1200
      Picture         =   "NewMethod.frx":2844
      ScaleHeight     =   450
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   -1080
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   5280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dead"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   200.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   4815
      Left            =   3120
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   11295
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   50
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   49
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   48
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   47
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   46
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   45
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   44
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   43
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   42
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   41
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   40
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   39
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   38
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   37
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   36
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   35
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   34
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   33
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   32
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   31
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   30
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   29
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   28
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   27
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   26
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   25
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   24
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   23
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   22
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   21
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   20
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   19
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   18
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   17
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   16
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   15
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   14
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   13
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   12
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   11
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   10
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   9
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   8
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   7
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   6
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   5
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   4
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   3
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   2
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Shape TB 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   90
      Index           =   0
      Left            =   9720
      Shape           =   2  'Oval
      Top             =   9480
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   6
      Left            =   1.22025e5
      Top             =   11775
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   5
      Left            =   79200
      Top             =   11775
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   4
      Left            =   65025
      Top             =   11775
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   3
      Left            =   41550
      Top             =   11775
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   2
      Left            =   21225
      Top             =   11775
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   1
      Left            =   13350
      Top             =   11775
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image House 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   3450
      Top             =   11775
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   99
      Visible         =   0   'False
      X1              =   104
      X2              =   144
      Y1              =   520
      Y2              =   536
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   98
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   97
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   96
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   95
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   94
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   93
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   92
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   91
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   90
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   89
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   88
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   87
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   86
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   85
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   84
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   83
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   82
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   81
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   80
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   79
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   78
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   77
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   76
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   75
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   74
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   73
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   72
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   71
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   70
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   69
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   68
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   67
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   66
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   65
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   64
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   63
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   62
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   61
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   60
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   59
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   58
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   57
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   56
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   55
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   54
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   53
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   52
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   51
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   50
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   49
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   48
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   47
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   46
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   45
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   44
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   43
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   42
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   41
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   40
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   39
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   38
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   37
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   36
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   35
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   34
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   33
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   32
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   31
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   30
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   29
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   28
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   27
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   26
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   25
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   24
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   23
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   22
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   21
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   20
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   19
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   18
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   17
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   16
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   15
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   14
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   13
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   12
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   11
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   10
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Fifth 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   152
      X2              =   192
      Y1              =   576
      Y2              =   592
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   399
      Visible         =   0   'False
      X1              =   32
      X2              =   32
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   398
      Visible         =   0   'False
      X1              =   40
      X2              =   40
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   397
      Visible         =   0   'False
      X1              =   48
      X2              =   48
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   396
      Visible         =   0   'False
      X1              =   112
      X2              =   112
      Y1              =   688
      Y2              =   704
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   395
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   394
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   393
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   392
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   391
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   390
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   389
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   388
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   387
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   386
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   385
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   384
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   383
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   382
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   381
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   380
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   379
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   378
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   377
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   376
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   375
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   374
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   373
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   372
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   371
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   370
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   369
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   368
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   367
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   366
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   365
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   364
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   363
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   362
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   361
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   360
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   359
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   358
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   357
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   356
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   355
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   354
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   353
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   352
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   351
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   350
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   349
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   348
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   347
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   346
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   345
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   344
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   343
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   342
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   341
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   340
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   339
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   338
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   337
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   336
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   335
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   334
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   333
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   332
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   331
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   330
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   329
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   328
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   327
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   326
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   325
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   324
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   323
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   322
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   321
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   320
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   319
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   318
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   317
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   316
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   315
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   314
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   313
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   312
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   311
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   310
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   309
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   308
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   307
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   306
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   305
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   304
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   303
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   302
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   301
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   300
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   299
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   298
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   297
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   296
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   295
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   294
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   293
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   292
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   291
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   290
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   289
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   288
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   287
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   286
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   285
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   284
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   283
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   282
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   281
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   280
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   279
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   278
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   277
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   276
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   275
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   274
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   273
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   272
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   271
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   270
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   269
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   268
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   267
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   266
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   265
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   264
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   263
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   262
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   261
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   260
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   259
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   258
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   257
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   256
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   255
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   254
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   253
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   252
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   251
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   250
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   249
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   248
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   247
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   246
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   245
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   244
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   243
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   242
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   241
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   240
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   239
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   238
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   237
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   236
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   235
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   234
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   233
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   232
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   231
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   230
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   229
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   228
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   227
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   226
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   225
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   224
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   223
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   222
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   221
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   220
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   219
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   218
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   217
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   216
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   215
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   214
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   213
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   212
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   211
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   210
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   209
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   208
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   207
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   206
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   205
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   204
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   203
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   202
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   201
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   200
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   199
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   198
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   197
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   196
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   195
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   194
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   193
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   192
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   191
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   190
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   189
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   188
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   187
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   186
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   185
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   184
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   183
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   182
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   181
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   180
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   179
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   178
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   177
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   176
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   175
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   174
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   173
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   172
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   171
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   170
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   169
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   168
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   167
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   166
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   165
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   164
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   163
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   162
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   161
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   160
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   159
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   158
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   157
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   156
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   155
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   154
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   153
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   152
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   151
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   150
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   149
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   148
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   147
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   146
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   145
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   144
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   143
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   142
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   141
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   140
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   139
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   138
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   137
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   136
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   135
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   134
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   133
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   132
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   131
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   130
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   129
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   128
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   127
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   126
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   125
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   124
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   123
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   122
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   121
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   120
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   119
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   118
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   117
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   116
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   115
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   114
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   113
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   112
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   111
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   110
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   109
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   108
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   107
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   106
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   105
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   104
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   103
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   102
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   101
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   100
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   99
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   98
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   97
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   96
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   95
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   94
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   93
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   92
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   91
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   90
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   89
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   88
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   87
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   86
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   85
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   84
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   83
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   82
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   81
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   80
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   79
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   78
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   77
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   76
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   75
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   74
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   73
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   72
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   71
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   70
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   69
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   68
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   67
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   66
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   65
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   64
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   63
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   62
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   61
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   60
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   59
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   58
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   57
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   56
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   55
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   54
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   53
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   52
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   51
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   50
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   49
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   48
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   47
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   46
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   45
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   44
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   43
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   42
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   41
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   40
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   39
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   38
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   37
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   36
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   35
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   34
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   33
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   32
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   31
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   30
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   29
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   28
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   27
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   26
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   25
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   24
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   23
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   22
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   21
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   20
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   19
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   18
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   17
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   16
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   15
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   14
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   13
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   12
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   11
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   10
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   9
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   8
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   7
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   6
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   5
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   4
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   3
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   2
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   1
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   608
      Y2              =   624
   End
   Begin VB.Line Tally 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Index           =   0
      Visible         =   0   'False
      X1              =   24
      X2              =   24
      Y1              =   712
      Y2              =   728
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   20
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   19
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   18
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   17
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   16
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   15
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   14
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   13
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   12
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   11
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   10
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   9
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   8
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   7
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   6
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   5
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   4
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   3
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   2
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   1
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Bullet 
      BackColor       =   &H80000003&
      BorderStyle     =   3  'Dot
      DrawMode        =   2  'Blackness
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   45
      Index           =   0
      Left            =   120
      Shape           =   2  'Oval
      Top             =   480
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public X As Single

Private Sub Form_Load()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While RPlaneW <= 0                                      '
        RPlaneW = GenerateDC(App.Path & "\NHeliw.bmp")      '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While RPlaneB <= 0                                      '
        RPlaneB = GenerateDC(App.Path & "\NHeli.bmp")      '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While LPlaneW <= 0                                      '
        LPlaneW = GenerateDC(App.Path & "\NHeli2w.bmp")     '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While LPlaneB <= 0                                      '
        LPlaneB = GenerateDC(App.Path & "\NHeli2.bmp")     '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While TankW <= 0                                        '
        TankW = GenerateDC(App.Path & "\TankW.bmp")         '
    Wend                                                    '
    While TankB <= 0                                        '
        TankB = GenerateDC(App.Path & "\TankB.bmp")         '
    Wend                                                    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While BarrelW <= 0                                      '
        BarrelW = GenerateDC(App.Path & "\BarrelW.bmp")     '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While BarrelB <= 0                                      '
        BarrelB = GenerateDC(App.Path & "\BarrelB.bmp")     '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While SandB <= 0                                        '
        SandB = GenerateDC(App.Path & "\SandB2.bmp")        '
    Wend                                                    '
    While WaterB <= 0                                       '
        WaterB = GenerateDC(App.Path & "\WaterB2.bmp")      '
    Wend                                                    '
    While DirtB <= 0                                        '
        DirtB = GenerateDC(App.Path & "\DirtB2.bmp")        '
    Wend                                                    '
    While DirtW <= 0                                        '
        DirtW = GenerateDC(App.Path & "\DirtW2.bmp")        '
    Wend                                                    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While BoomW <= 0                                        '
        BoomW = GenerateDC(App.Path & "\BoomW.bmp")         '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While BoomB <= 0                                        '
        BoomB = GenerateDC(App.Path & "\BoomB.bmp")         '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
     While WBoomB <= 0                                      '
        WBoomB = GenerateDC(App.Path & "\WBoomB.bmp")       '
    Wend                                                    '
    While WBoomW <= 0                                       '
        WBoomW = GenerateDC(App.Path & "\WBoomW.bmp")       '
    Wend                                                    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While GuyB <= 0                                         '
        GuyB = GenerateDC(App.Path & "\GuyB.bmp")           '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While GuyW <= 0                                         '
        GuyW = GenerateDC(App.Path & "\GuyW.bmp")           '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While GraveBW <= 0                                      '
        GraveBW = GenerateDC(App.Path & "\Grave.bmp")       '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While RopeW <= 0                                        '
        RopeW = GenerateDC(App.Path & "\RopeW.bmp")         '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    While RopeB <= 0                                        '
        RopeB = GenerateDC(App.Path & "\RopeB.bmp")         '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While BG <= 0                                           '
        BG = GenerateDC(App.Path & "\bgbighsmall.bmp")      '
    Wend                                                    '
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 5   '
    AltStart.Refresh                                        '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PlaneW = RPlaneW
    PlaneB = RPlaneB
    PX = 270
    X = 500
    Y = 500
    Direction = 5
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    AltStart.Refresh
    For P = 0 To 6
        House(P).Left = House(P).Left + 30 + 1149
        House(P).Width = 20
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    AltStart.Refresh
    For G = 0 To 5
        GuyX(G) = House(0).Left + House(0).Width
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 3) - 4
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    AltStart.Refresh
    For G = 6 To 8
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
         If GuySpeed(G) >= 0 Then
                GuyX(G) = House(1).Left
            Else
                GuyX(G) = House(1).Left + House(1).Width
            End If
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 9 To 15
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
                GuyX(G) = House(2).Left
            Else
                GuyX(G) = House(2).Left + House(2).Width
            End If
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 16 To 26
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
                GuyX(G) = House(3).Left
            Else
                GuyX(G) = House(3).Left + House(3).Width
            End If
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 27 To 38
        GuySpeed(G) = Int(Rnd * 8) - 4
        GuyY(G) = 785
        If GuySpeed(G) >= 0 Then
                GuyX(G) = House(4).Left
            Else
                GuyX(G) = House(4).Left + House(4).Width
            End If
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 39 To 43
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
                GuyX(G) = House(5).Left
            Else
                GuyX(G) = House(5).Left + House(5).Width
            End If
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 44 To 50
        GuyX(G) = House(6).Left
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 4)
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
    For G = 0 To 50
        If GuySpeed(G) = 0 Then GuySpeed(G) = 1
        Guy(G) = True
    Next
    TalleyY = 10
    For u = 0 To 15
        TankX(u) = Int(Rnd * 10000) + 1000
        TankSpeed(u) = Int(Rnd * 6) - 3
        Tank(u) = True
    Next
    AltStart.ProgBar1.Value = AltStart.ProgBar1.Value + 3
End Sub

Public Function Run()
    Do
        Form1.Refresh
        
        Up = GetAsyncKeyState(vbKeyUp)
        Down = GetAsyncKeyState(vbKeyDown)
        Right = GetAsyncKeyState(vbKeyRight)
        Lef = GetAsyncKeyState(vbKeyLeft)
        STP = GetAsyncKeyState(vbKeyEscape)
        Space = GetAsyncKeyState(vbKeyControl)
        Shift = GetAsyncKeyState(vbKeyShift)
        
        Switch = True
        Direction = EndDirection
        
        If Shift <= -32767 Then
            Shoot = True
        Else
            Shoot = False
        End If
        If Up <= -32767 Then
            Y = Y - 5
            For m = 0 To 10
                If GuyClimb(m) = True Then
                    GuyY(m) = GuyY(m) - 5
                End If
            Next
            Switch = False
        ElseIf Down <= -32767 Then
            For n = 0 To 10
                If GuyClimb(n) = True Then
                    GuyY(n) = GuyY(n) + 5
                End If
            Next
        End If
        If Right <= -32767 Then
            If PX < 540 And SSlowdown = 6 Then
                PX = PX + 135
                SSlowdown = 0
            End If
            Switch = False
            If LRCounter = 2 Then
                PlaneW = RPlaneW
                PlaneB = RPlaneB
                LRCounter = 0
            End If
            If Direction < 10 Then
                Direction = Direction + 1
            End If
        End If
        If Down <= -32767 Then
            Y = Y + 5
            Switch = False
        End If
        If Lef <= -32767 Then
            If PX > 0 And SSlowdown = 6 Then
                PX = PX - 135
                SSlowdown = 0
            End If
            Switch = False
            If LRCounter = 2 Then
                PlaneW = LPlaneW
                PlaneB = LPlaneB
                LRCounter = 0
            End If
            If Direction > -10 Then
                Direction = Direction - 1
            End If
        End If
        If Space <= -32767 And slowdown = 5 Then
            slowdown = 0
            While ind < 5
                If Bomb(ind) = False Then
                    BarrelX(ind) = X + 100
                    BarrelY(ind) = Y + 100
                    Bomb(ind) = True
                    ind = 4
                End If
                ind = ind + 1
            Wend
            ind = 0
        End If
        If Switch = True And DSlowdown = 8 Then
            If PX > 270 Then
                PX = PX - 135
            End If
            If PX < 270 Then
                PX = PX + 135
            End If
            If Direction > 0 Then Direction = Direction - 1
            If Direction < 0 Then Direction = Direction + 1
            DSlowdown = 0
        End If
        If Direction = 0 Then
            LRCounter = LRCounter + 1
        Else
            LRCounter = 0
        End If
        SSlowdown = SSlowdown + 1
        DSlowdown = DSlowdown + 1
        slowdown = slowdown + 1
        If LRCounter = 3 Then LRCounter = 2
        If SSlowdown = 7 Then SSlowdown = 6
        If DSlowdown = 9 Then DSlowdown = 8
        If slowdown = 6 Then slowdown = 5
        If STP = -32767 Then
            DeleteGeneratedDC PlaneB
            DeleteGeneratedDC PlaneW
            DeleteGeneratedDC LPlaneB
            DeleteGeneratedDC LPlaneW
            DeleteGeneratedDC RPlaneB
            DeleteGeneratedDC RPlaneW
            DeleteGeneratedDC RopeB
            DeleteGeneratedDC RopeW
            DeleteGeneratedDC BarrelB
            DeleteGeneratedDC BarrelW
            DeleteGeneratedDC BoomB
            DeleteGeneratedDC BoomW
            DeleteGeneratedDC GuyB
            DeleteGeneratedDC GuyW
            DeleteGeneratedDC DirtB
            DeleteGeneratedDC DirtW
            DeleteGeneratedDC TankB
            DeleteGeneratedDC TankW
            DeleteGeneratedDC GraveBW
            DeleteGeneratedDC WBoomW
            DeleteGeneratedDC WBoomB
            DeleteGeneratedDC SandB
            DeleteGeneratedDC WaterB
            DeleteGeneratedDC BG
            
            Unload Me
            Set Form1 = Nothing
            End
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''
        BitBlt Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, BG, BGPosition, 0, vbSrcCopy
        EndDirection = Direction
        If BGPosition < 0 Then
            X = X + Direction
            If X >= 500 Then BGPosition = BGPosition + Direction
            Direction = 0
        ElseIf BGPosition > 10282 Then
            X = X + Direction
            If X <= 500 Then BGPosition = BGPosition + Direction
            Direction = 0
        Else
            BGPosition = BGPosition + Direction
        End If
        For f = 0 To 50
        If Guy(f) = True Then
            If GuyClimb(f) = True Then
                GuyY(f) = GuyY(f) - 1
                If RPlaneB = PlaneB Then GuyX(f) = X + 100
                If LPlaneB = PlaneB Then GuyX(f) = X + 35
                If Y + RopeP + 45 < GuyY(f) Then
                    Fall(f) = True
                    GuyClimb(f) = False
                End If
            Else
                If (GuyX(f) < House(0).Left + House(0).Width) Or (GuyX(f) > House(1).Left And GuyX(f) < House(1).Left + House(1).Width) Or (GuyX(f) > House(2).Left And GuyX(f) < House(2).Left + House(2).Width) Or (GuyX(f) > House(3).Left And GuyX(f) < House(3).Left + House(3).Width) Or (GuyX(f) > House(4).Left And GuyX(f) < House(4).Left + House(4).Width) Or (GuyX(f) > House(5).Left And GuyX(f) < House(5).Left + House(5).Width) Or (GuyX(f) > House(6).Left And GuyX(f) < House(6).Left + House(6).Width) Then
                    Men (f)
                End If
                GuyX(f) = GuyX(f) - GuySpeed(f) - Direction
            End If
            
            If Fall(f) = True Then
                GuyY(f) = GuyY(f) + 8
                If GuyY(f) > 780 Then Fall(f) = False
            End If
            If GuyY(f) < Y + 45 And GuyClimb(f) = True Then
                Guy(f) = False
                GuyClimb(f) = False
                Talley (&H80000007)
            End If
            
            GuySpeed(f) = GuySpeed(f)
            If BGPosition + GuyX(f) < BGPosition + Form1.ScaleWidth And BGPosition + GuyX(f) >= BGPosition Then
                BitBlt Form1.hdc, GuyX(f), GuyY(f), 20, 30, GuyW, GuyP(f), 0, vbSrcAnd
                BitBlt Form1.hdc, GuyX(f), GuyY(f), 20, 30, GuyB, GuyP(f), 0, vbSrcPaint
                GuySlowdown(f) = GuySlowdown(f) + 1
                If GuySlowdown(f) = 5 Then
                    GuyP(f) = GuyP(f) + 20
                    GuySlowdown(f) = 0
                End If
                If GuyP(f) = 40 Then GuyP(f) = 0
                For e = 0 To 7
                    If Dirt(e) = True And DirtX(e) < GuyX(f) And DirtX(e) + 10 > GuyX(f) Then
                        Guy(f) = False
                        Grave(f) = True
                        GraveX(f) = GuyX(f)
                        Talley (&HFF&)
                    End If
                Next
            End If
        ElseIf Grave(f) = True Then
            BitBlt Form1.hdc, GraveX(f), 800, 15, 12, GraveBW, 15, 0, vbSrcAnd
            BitBlt Form1.hdc, GraveX(f), 800, 15, 12, GraveBW, 0, 0, vbSrcPaint
            GraveX(f) = GraveX(f) - Direction
            'If GraveEnd(f) < BGPosition Then Grave(f) = False
        End If
        Next
        For P = 0 To 6
            House(P).Left = House(P).Left - Direction ''''''''''''''''''''''''''''''''
        Next
        ''''''''''''''
                     '
        ''''''       '
        Rope '       '
        '''''''      '
        Tanks '      '
        '''''''''''' '
        TankBullet ' '
        '''''''''''' '
        Bombs '      '
        '''''''''    '
        Bullets '    '
        '''''''''    '
        Dirts '      '
        '''''''      '
                     '
        ''''''''''''''
        If Dead = True Or Tally(50).Visible = True Then Exit Do
        BitBlt Pl.hdc, 0, 0, 135, 80, PlaneW, PX, 0, vbSrcCopy
        BitBlt Form1.hdc, X, Y, 135, 80, PlaneW, PX, 0, vbSrcAnd
        BitBlt Form1.hdc, X, Y, 135, 80, PlaneB, PX, 0, vbSrcPaint
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Loop
    '''''''''''''''''''''
    '''''''''''''''''''''
    '''''''''''''''''''''
    Timer2.Enabled = True
    temp = 0
    Label1.Visible = True
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Private Sub Timer1_Timer()
    If V < 1 Then Run
    If V < 6 Then Guy(V) = True
    If V < 3 Then Guy(V + 6) = True
    If V < 7 Then Guy(V + 9) = True
    If V < 11 Then Guy(V + 16) = True
    Guy(V + 27) = True
    If V < 5 Then Guy(V + 39) = True
    If V < 6 Then Guy(V + 44) = True
    V = V + 1
    If V >= 12 Then Timer1.Enabled = False
End Sub
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Bullets()
    While indb < 21
        If Bullet(indb).Visible = False And Shoot = True Then
            Bullet(indb).Visible = True
            If RPlaneB = PlaneB Then
                If PX = 0 Then
                    Bullet(indb).Left = X + 131
                    Bullet(indb).Top = Y + 12
                    BulletX(indb) = 20
                    angle(indb) = -14
                ElseIf PX = 135 Then
                    Bullet(indb).Left = X + 135
                    Bullet(indb).Top = Y + 22
                    BulletX(indb) = 25
                    angle(indb) = -12
                ElseIf PX = 270 Then
                    Bullet(indb).Left = X + 135
                    Bullet(indb).Top = Y + 40
                    BulletX(indb) = 30
                    angle(indb) = 0
                ElseIf PX = 405 Then
                    Bullet(indb).Left = X + 131
                    Bullet(indb).Top = Y + 60
                    BulletX(indb) = 25 ''''''''
                    angle(indb) = 12 '''''''''''
                ElseIf PX = 540 Then
                    Bullet(indb).Left = X + 127
                    Bullet(indb).Top = Y + 75
                    BulletX(indb) = 20
                    angle(indb) = 14
                End If
                indb = 20
            Else
                If PX = 0 Then
                    Bullet(indb).Left = X + 8
                    Bullet(indb).Top = Y + 75
                    BulletX(indb) = -20
                    angle(indb) = 14
                ElseIf PX = 135 Then
                    Bullet(indb).Left = X
                    Bullet(indb).Top = Y + 60
                    BulletX(indb) = -25 '''''''''
                    angle(indb) = 12 '''''''''''''
                ElseIf PX = 270 Then
                    Bullet(indb).Left = X
                    Bullet(indb).Top = Y + 40
                    BulletX(indb) = -30
                    angle(indb) = 0
                ElseIf PX = 405 Then
                    Bullet(indb).Left = X
                    Bullet(indb).Top = Y + 22
                    BulletX(indb) = -25 ''''''''''
                    angle(indb) = -12 '''''''''''''
                ElseIf PX = 540 Then
                    Bullet(indb).Left = X + 4
                    Bullet(indb).Top = Y + 12
                    BulletX(indb) = -20
                    angle(indb) = -14
                End If
                indb = 20
            End If
            'Bullet(indb).Visible = True
            'If Direction > 0 Then
            '    Bullet(indb).Left = X + 178
            '    BulletX(indb) = 21
            'Else
            '    Bullet(indb).Left = X + 22
            '    BulletX(indb) = -21
            'End If
            'Bullet(indb).Top = Y + 60
            'angle(indb) = 0
            'If Down <= -32767 Then angle(indb) = 13
            'If Up <= -32767 Then angle(indb) = -13
            'If angle(indb) < 0 Then
            '    Bullet(indb).Top = Y + 38
            '    If Right <= -32767 Then angle(indb) = -10
            'End If
            'If angle(indb) > 0 Then
            '    Bullet(indb).Top = Y + 86
            '    If Right <= -32767 Then angle(indb) = 10
            'End If
            'indb = 20
        End If
        If Bullet(indb).Visible = True Then
            Bullet(indb).Left = Bullet(indb).Left + BulletX(indb)
            Bullet(indb).Top = Bullet(indb).Top + angle(indb)
            If Bullet(indb).Top < 0 Then Bullet(indb).Visible = False
            If Bullet(indb).Left > Form1.ScaleWidth Or Bullet(indb).Left < 0 Then Bullet(indb).Visible = False
            If Bullet(indb).Top > 785 Then
                Dirt(DInd) = True
                DirtX(DInd) = Bullet(indb).Left
                DInd = DInd + 1
                If DInd = 8 Then DInd = 0
                Bullet(indb).Visible = False
            End If
        End If
        indb = indb + 1
    Wend
    indb = 0
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Dirts()
    For d = 0 To 7
        If Dirt(d) = True Then
            If DirtX(d) + BGPosition < 1059 Then
                BitBlt Form1.hdc, DirtX(d), 815, 10, 10, DirtW, DirtP(d), 0, vbSrcAnd
                BitBlt Form1.hdc, DirtX(d), 815, 10, 10, WaterB, DirtP(d), 0, vbSrcPaint
            ElseIf DirtX(d) + BGPosition < 1153 Then
                BitBlt Form1.hdc, DirtX(d), 815, 10, 10, DirtW, DirtP(d), 0, vbSrcAnd
                BitBlt Form1.hdc, DirtX(d), 815, 10, 10, SandB, DirtP(d), 0, vbSrcPaint
            Else
                BitBlt Form1.hdc, DirtX(d), 800, 10, 10, DirtW, DirtP(d), 0, vbSrcAnd
                BitBlt Form1.hdc, DirtX(d), 800, 10, 10, DirtB, DirtP(d), 0, vbSrcPaint
            End If
            If DirtP(d) = 10 Then
                DirtP(d) = 0
                Dirt(d) = False
                DirtX(d) = DirtX(d) - Direction
            Else
                DirtP(d) = 10
            End If
        End If
    Next
End Function
Public Function Rope()
    If Y >= 670 Then
        If RopeP < 70 Then RopeP = RopeP + 10
    ElseIf RopeP > 0 Then
        RopeP = RopeP - 10
    End If
    If RopeP > 0 Then
        If RPlaneB = PlaneB Then
            BitBlt Form1.hdc, X + 100, Y + 45, 20, RopeP, RopeW, 0, 0, vbSrcAnd
            BitBlt Form1.hdc, X + 100, Y + 45, 20, RopeP, RopeB, 0, 0, vbSrcPaint
            If Direction >= -4 Then
                For G = 0 To 50
                    If BGPosition + GuyX(G) >= BGPosition And BGPosition + GuyX(G) <= BGPosition + Form1.ScaleWidth Then
                        If GuyX(G) > X + 100 And GuyX(G) < X + 120 And Guy(G) = True Then
                            If (Rnd * 5) > 4 Then GuyClimb(G) = True
                            Exit For
                        End If
                    End If
                Next
            End If
        Else
            BitBlt Form1.hdc, X + 35, Y + 45, 20, RopeP, RopeW, 0, 0, vbSrcAnd
            BitBlt Form1.hdc, X + 35, Y + 45, 20, RopeP, RopeB, 0, 0, vbSrcPaint
            For G = 0 To 50
                If BGPosition + GuyX(G) >= BGPosition And BGPosition + GuyX(G) <= BGPosition + Form1.ScaleWidth Then
                    If GuyX(G) > X + 35 And GuyX(G) < X + 55 Then
                        If (Rnd * 5) > 4 Then GuyClimb(G) = True
                        Exit For
                    End If
                End If
            Next
        End If
    End If
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Men(G As Integer)
    If G >= 0 And G < 6 Then
        GuySpeed(G) = Int(Rnd * 3) - 4
        GuyY(G) = 785
        GuyX(G) = House(0).Left + House(0).Width
    End If
    If G > 5 And G < 9 Then
        GuySpeed(G) = Int(Rnd * 8) - 4
        GuyY(G) = 785
        If GuySpeed(G) >= 0 Then
            GuyX(G) = House(1).Left
        Else
            GuyX(G) = House(1).Left + House(1).Width
        End If
    End If
    If G > 8 And G < 16 Then
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
            GuyX(G) = House(2).Left
        Else
            GuyX(G) = House(2).Left + House(2).Width
        End If
    End If
    If G > 15 And G < 27 Then
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
            GuyX(G) = House(3).Left
        Else
            GuyX(G) = House(3).Left + House(3).Width
        End If
    End If
    If G > 26 And G < 39 Then
        GuySpeed(G) = Int(Rnd * 8) - 4
        GuyY(G) = 785
        If GuySpeed(G) >= 0 Then
            GuyX(G) = House(4).Left
        Else
            GuyX(G) = House(4).Left + House(4).Width
        End If
    End If
    If G > 38 And G < 44 Then
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 8) - 4
        If GuySpeed(G) >= 0 Then
            GuyX(G) = House(5).Left
        Else
            GuyX(G) = House(5).Left + House(5).Width
        End If
    End If
    If G > 43 And G < 51 Then
        GuyX(G) = House(6).Left
        GuyY(G) = 785
        GuySpeed(G) = Int(Rnd * 3) + 1
    End If
    For G = 0 To 50
        If GuySpeed(G) = 0 Then GuySpeed(G) = 1
    Next
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Talley(Colour As Variant)
                TalleyX = TalleyX + 8
                If TalleyX = 1096 Then
                    TalleyX = 8
                    TalleyY = TalleyY + 24
                End If
                TalleyCount = TalleyCount + 1
                If TalleyCount = 5 Then
                    Fifth((TalleyIndex / 4) - 1).X1 = TalleyX - 40
                    Fifth((TalleyIndex / 4) - 1).X2 = TalleyX
                    Fifth((TalleyIndex / 4) - 1).Y1 = TalleyY
                    Fifth((TalleyIndex / 4) - 1).Y2 = TalleyY + 12
                    Fifth((TalleyIndex / 4) - 1).BorderColor = Colour
                    Fifth((TalleyIndex / 4) - 1).Visible = True
                    TalleyCount = 0
                Else
                    Tally(TalleyIndex).X1 = TalleyX
                    Tally(TalleyIndex).X2 = TalleyX
                    Tally(TalleyIndex).Y1 = TalleyY
                    Tally(TalleyIndex).Y2 = TalleyY + 12
                    Tally(TalleyIndex).BorderColor = Colour
                    Tally(TalleyIndex).Visible = True
                End If
                TalleyIndex = TalleyIndex + 1
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Bombs()
    For t = 0 To 4
        If Bomb(t) = True Then
            BitBlt Form1.hdc, BarrelX(t), BarrelY(t), 20, 20, BarrelW, BarrelP(t), 0, vbSrcAnd
            BitBlt Form1.hdc, BarrelX(t), BarrelY(t), 20, 20, BarrelB, BarrelP(t), 0, vbSrcPaint
            BarrelX(t) = BarrelX(t) - Direction
            BarrelY(t) = BarrelY(t) * 1.04
            If BarrelY(t) < 0 Then Bomb(t) = False
            If BarrelY(t) > Form1.ScaleHeight - 80 Then
                Bomb(t) = False
                Boom(t) = True
                BoomX(t) = BarrelX(t)
            End If
            If SlowDownBarrel(t) = 2 Then
                BarrelP(t) = BarrelP(t) + 20
                SlowDownBarrel(t) = 0
            End If
            SlowDownBarrel(t) = SlowDownBarrel(t) + 1
            If SlowDownBarrel(t) = 3 Then SlowDownBarrel(t) = 2
            If BarrelP(t) = 80 Then
                BarrelP(t) = 0
                SlowDownBarrel(t) = 0
            End If
        End If
        If Boom(t) = True Then
            If BoomX(t) + BGPosition > 1059 Then
                BitBlt Form1.hdc, BoomX(t), 790, 20, 20, BoomW, BoomP(t), 0, vbSrcAnd
                BitBlt Form1.hdc, BoomX(t), 790, 20, 20, BoomB, BoomP(t), 0, vbSrcPaint
                BoomX(t) = BoomX(t) - Direction
                If SlowDownBoom(t) = 2 Then BoomP(t) = BoomP(t) + 20
                SlowDownBoom(t) = SlowDownBoom(t) + 1
                If SlowDownBoom(t) = 3 Then SlowDownBoom(t) = 2
                If BoomP(t) = 80 Then
                    Boom(t) = False
                    BoomP(t) = 0
                    SlowDownBoom(t) = 0
                End If
                For s = 0 To 50
                    If (BoomP(t) <= 20 And Boom(t) = True And BoomX(t) <= GuyX(s) + 20 And BoomX(t) + 20 >= GuyX(s) And Guy(s) = True) Or Death(s) = True Then
                        Death(s) = False
                        Guy(s) = False
                        Grave(s) = True
                        GraveX(s) = GuyX(s)
                        Talley (&HFF&)
                    End If
                Next
                For j = 0 To 15
                    If (BoomP(t) <= 20 And Boom(t) = True And BoomX(t) <= TankX(j) + 40 And BoomX(t) + 20 >= TankX(j) And Tank(j) = True) Then
                        TankHit(j) = TankHit(j) + 1
                        If TankHit(j) = 12 Then
                            Tank(j) = False
                            TankHit(j) = 0
                        End If
                    End If
                Next
            Else
                BitBlt Form1.hdc, BoomX(t), 810, 20, 20, WBoomW, BoomP(t), 0, vbSrcAnd
                BitBlt Form1.hdc, BoomX(t), 810, 20, 20, WBoomB, BoomP(t), 0, vbSrcPaint
                BoomX(t) = BoomX(t) - Direction
                If SlowDownBoom(t) = 2 Then BoomP(t) = BoomP(t) + 20
                SlowDownBoom(t) = SlowDownBoom(t) + 1
                If SlowDownBoom(t) = 3 Then SlowDownBoom(t) = 0
                If BoomP(t) = 100 Then
                    Boom(t) = False
                    BoomP(t) = 0
                    SlowDownBoom(t) = 0
                End If
            End If
        End If
    Next
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function Tanks()
For L = 0 To 15
    If Tank(L) = True Then
        For e = 0 To 7
            If Dirt(e) = True And DirtX(e) < TankX(L) + 40 And DirtX(e) + 10 > TankX(L) Then
                TankHit(L) = TankHit(L) + 1
                If TankHit(L) = 12 Then
                    Tank(L) = False
                    TankHit(L) = 0
                End If
            End If
        Next
        TankX(L) = TankX(L) - Direction + TankSpeed(L)
        For i = 0 To 6
            If TankX(L) > House(i).Left And TankX(L) < House(i).Left + House(i).Width Then
                TankSpeed(L) = Int(Rnd * 6) - 3
            End If
            If TankX(L) + BGPosition < 1153 Then TankSpeed(L) = -4
            If TankX(L) + BGPosition > 10282 Then TankSpeed(L) = 4
        Next
        If BGPosition + TankX(L) >= BGPosition And BGPosition + TankX(L) <= BGPosition + Form1.ScaleWidth Then
            If TankSlowDown(L) = 25 Then
                While TAind < 51
                    If TB(TAind).Visible = False Then
                        TB(TAind).Visible = True
                        TAngleX(TAind) = (X - TankX(L) + 20) / 20
                        TAngleY(TAind) = (Y - 790) / 20
                        TB(TAind).Left = TankX(L) + 20
                        TB(TAind).Top = 790
                        TAind = 50
                    End If
                    TAind = TAind + 1
                Wend
                TAind = 0
            End If
            TankSlowDown(L) = TankSlowDown(L) + 1
            If TankSlowDown(L) = 26 Then TankSlowDown(L) = 0
            BitBlt Form1.hdc, TankX(L), 785, 40, 30, TankW, 0, 0, vbSrcAnd
            BitBlt Form1.hdc, TankX(L), 785, 40, 30, TankB, 0, 0, vbSrcPaint
        End If
    Else
        TankX(L) = TankX(L) - Direction
        TankHit(L) = TankHit(L) + 1
        If BGPosition + TankX(L) >= BGPosition And BGPosition + TankX(L) <= BGPosition + Form1.ScaleWidth Then
            BitBlt Form1.hdc, TankX(L), 785, 40, 30, TankW, 0, 0, vbSrcAnd
            BitBlt Form1.hdc, TankX(L), 785, 40, 30, TankB, 0, 0, vbSrcPaint
        End If
        If TankHit(L) = 1000 Then
            Tank(L) = True
            TankHit(L) = 0
        End If
    End If
Next
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Public Function TankBullet()
For c = 0 To 50
    If TB(c).Visible = True Then
        If Direction < 0 Then
            TB(c).Left = TB(c).Left + TAngleX(c) - Direction
        Else
            TB(c).Left = TB(c).Left + TB(c).Width + TAngleX(c) - Direction
        End If
        TB(c).Top = TB(c).Top + TAngleY(c)
        If TB(c).Left < 0 Or TB(c).Left > Form1.ScaleWidth Then TB(c).Visible = False
        If TB(c).Top < 0 Then
            TB(c).Visible = False
        End If
        If TB(c).Left + 6 > X And TB(c).Left < X + 135 And TB(c).Top + 6 > Y And TB(c).Top < Y + 80 Then
            If CollisionDetect(X, Y, Pl, TB(c).Left, TB(c).Top, B, Blank) Then
                TB(c).Visible = False
                PBar.Value = PBar.Value + 1
                If PBar.Value = PBar.Max Then
                    Dead = True
                End If
            End If
        End If
    End If
Next c
End Function
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Private Sub Timer2_Timer()
    If temp = 0 Then Label1.ForeColor = &HC0C000
    If temp = 1 Then Label1.ForeColor = &H808000
    If temp = 2 Then Label1.ForeColor = &H404000
    If temp = 3 Then Label1.ForeColor = &H40&
    If temp = 4 Then Label1.ForeColor = &H80&
    If temp = 5 Then Label1.ForeColor = &HC0&
    If temp = 6 Then Label1.ForeColor = &HFF&
    If temp = 7 Then Label1.ForeColor = &H8080FF
    If temp = 8 Then Label1.ForeColor = &HC0C0FF
    If temp = 9 Then Label1.ForeColor = &HFFFFFF
    If temp = 50 Then
        Timer2.Enabled = False
        Form1.Visible = False
        Results.Visible = True
    End If
    temp = temp + 1
    
End Sub
