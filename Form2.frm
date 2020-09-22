VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   LinkTopic       =   "Form2"
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   8040
      Top             =   6240
   End
   Begin Project1.ProgBar ProgBar 
      Height          =   12960
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   22860
      BackColour      =   0
      BarStartColour  =   0
      BarEndColour    =   255
      BorderStyle     =   0
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
   Begin Project1.ProgBar ProgBar 
      Height          =   12960
      Index           =   1
      Left            =   4320
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   22860
      BackColour      =   0
      BorderStyle     =   0
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
   End
   Begin Project1.ProgBar ProgBar 
      Height          =   12960
      Index           =   2
      Left            =   12960
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   22860
      BackColour      =   0
      BarStartColour  =   0
      BarEndColour    =   255
      BorderStyle     =   0
      FillDirection   =   2
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
   Begin Project1.ProgBar ProgBar 
      Height          =   12960
      Index           =   3
      Left            =   8640
      Top             =   0
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   22860
      BackColour      =   0
      BorderStyle     =   0
      FillDirection   =   2
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
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        End
    End If
End Sub

Private Sub Form_Load()
    'ProgBar(0).BackColour = &H0&
    'ProgBar(1).BackColour = &H0&
    'ProgBar(0).BarStyle = pbGradient
    'ProgBar(1).BarStyle = pbGradient
    'ProgBar(0).FillDirection = pbRight
    'ProgBar(1).FillDirection = pbRight
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Form2.Picture = LoadPicture(App.Path & "\bgbighstartsmall.bmp")
    Start
End Sub
Public Function Start()
    For j = 0 To 1
        For i = ProgBar(j).Min To ProgBar(j).Max
            ProgBar(j).Value = i
            ProgBar(j + 2).Value = i
        Next i
    Next j
    For j = 0 To 1
        For i = ProgBar(j).Min To ProgBar(j).Max - 5
            'Form2.Refresh
            ProgBar(1 - j).Value = ProgBar(j).Max - i
            ProgBar(3 - j).Value = ProgBar(j).Max - i
            ProgBar(1 - j).Width = ProgBar(1 - j).Width - 43.2
            ProgBar(3 - j).Width = ProgBar(3 - j).Width - 43.2
            ProgBar(3 - j).Left = ProgBar(3 - j).Left + 43.2
            Form2.Refresh
            'ProgBar(1).Value = ProgBar(1).Max - i
            'ProgBar(3).Value = ProgBar(3).Max - i
            'ProgBar(1).Width = ProgBar(1).Width - 43.2
            'ProgBar(3).Width = ProgBar(3).Width - 43.2
            'ProgBar(3).Left = ProgBar(3).Left + 43.2
            'ProgBar(2).Value = ProgBar(2).Max - i
            'ProgBar(0).Value = ProgBar(0).Max - i
            'ProgBar(0).Width = ProgBar(0).Width - 43.2
            'ProgBar(2).Width = ProgBar(2).Width - 43.2
            'ProgBar(0).Left = ProgBar(1).Left - ProgBar(0).Width + 1
        Next i
        ProgBar(1).Visible = False
        ProgBar(3).Visible = False
    Next j
    AltStart.Visible = True
    'Form2.Visible = False
    'While ProgBar(1).Left + ProgBar(1).Width < Form2.Width
    '    For b = 0 To 1
    '        ProgBar(b).Width = ProgBar(b).Width + 100
    '    Next
    '    ProgBar(1).Left = ProgBar(0).Width
    'Wend
End Function
