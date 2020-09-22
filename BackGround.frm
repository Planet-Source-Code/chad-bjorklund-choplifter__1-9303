VERSION 5.00
Begin VB.Form BackGround 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9105
   LinkTopic       =   "Form4"
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "BackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        BGPosition = BGPosition + 5
    End If
    
    If KeyCode = vbKeyLeft Then
        BGPosition = BGPosition - 5
    End If
    
    If KeyCode = vbKeyEscape Then
        DeleteGeneratedDC (BG)
        Unload Me
        End
    End If
    
    BitBlt BackGround.hdc, 0, 0, BackGround.ScaleWidth, BackGround.ScaleHeight, BG, BGPosition, 0, vbSrcCopy
End Sub

Private Sub Form_Load()
    While BG <= 0                                           '
        BG = GenerateDC(App.Path & "\bgbigh.bmp")           '
    Wend
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox (BGPosition + X)
End Sub
