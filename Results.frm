VERSION 5.00
Begin VB.Form Results 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4710
   ClientLeft      =   2520
   ClientTop       =   4110
   ClientWidth     =   12105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   85
      Left            =   5400
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   85
      Left            =   5400
      Top             =   2160
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      Picture         =   "Results.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      Picture         =   "Results.frx":0E52
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7920
      Picture         =   "Results.frx":12E4
      ScaleHeight     =   585
      ScaleWidth      =   705
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9600
      Picture         =   "Results.frx":2136
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      Picture         =   "Results.frx":3438
      ScaleHeight     =   345
      ScaleWidth      =   1305
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1200
      Picture         =   "Results.frx":473A
      ScaleHeight     =   1185
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "People Killed"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "People Saved"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "Results"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counts As Integer

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Counts < 51 Then
        If Form1.Tally(Counts).BorderColor = &H80000007 Then
            Label2.Caption = Label2.Caption + 1
        End If
        Counts = Counts + 1
    Else
        Timer1.Enabled = False
        Timer2.Enabled = True
        Counts = 0
    End If
End Sub

Private Sub Timer2_Timer()
    If Counts < 100 Then
        If Form1.Tally(Counts).BorderColor = &HFF& Then
            Label4.Caption = Label4.Caption + 1
        End If
        Counts = Counts + 1
    Else
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
End Sub

