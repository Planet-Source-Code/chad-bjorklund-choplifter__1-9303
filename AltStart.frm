VERSION 5.00
Begin VB.Form AltStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1215
   ClientLeft      =   4080
   ClientTop       =   5370
   ClientWidth     =   9615
   LinkTopic       =   "Form3"
   ScaleHeight     =   1215
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1200
      Top             =   120
   End
   Begin Project1.ProgBar ProgBar1 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2143
      BackColour      =   0
      BarStartColour  =   0
      BarEndColour    =   16777215
      BorderStyle     =   0
      FontColour      =   16777215
      Message         =   "LOADING,"
      ShowMessage     =   -1  'True
      ShowPercent     =   -1  'True
      BarStyle        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "UOP"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarEndColour    =   16777215
   End
End
Attribute VB_Name = "AltStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Begin()
    Form1.Visible = True
    AltStart.Visible = True
    Timer1.Enabled = False
End Function

Private Sub Timer1_Timer()
    Begin
End Sub
