VERSION 5.00
Begin VB.UserControl ProgBar 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   ClipControls    =   0   'False
   DrawWidth       =   50
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   ToolboxBitmap   =   "ctlProgBar.ctx":0000
End
Attribute VB_Name = "ProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'= ProgBar Control V2.1                                    =
'= --------------------                                    =
'= (C)2000 NJE                                             =
'= NE94252@netscape.net                                    =
'=                                                         =
'= You may use this source code within your own            =
'= applications, just give me a shout.                     =
'= You may not distribute it on a website or ftp site      =
'= without my express permission.                          =
'===========================================================
'= Updates:                                                =
'= --------                                                =
'= V1.1          - Addition of the VerticalText property.  =
'=               - General code clean up.                  =
'=               - Addition of the ability to play a wav   =
'=                 file at 100%.                           =
'= V1.2          - Addition of gradient fill (BarStyle).   =
'=               - All bar and background drawing handled  =
'=                 by APIs to speed things up.             =
'=               - The ability to wait for the sound to    =
'=                 finish or not before releasing to code. =
'= V1.2.1        - Fixed a problem with the use of         =
'=                 reserved words.                         =
'= V2.0          - Used DCs to remove flicker.             =
'=               - Added the font choice.                  =
'= V2.0.1        - Removed the half developed shadow text  =
'=                 property I had been working on, oops!   =
'= V2.1          - Nightmare! Corrected a problem with the =
'=                 control not drawing correctly when      =
'=                 resized. This was due to the DC being   =
'=                 created for the initial size of the     =
'=                 user control and not changing for the   =
'=                 life of the user control.               =
'===========================================================
'= RunTime Properties: (Aphabetical order)                 =
'= -------------------                                     =
'= BackColour    - The back ground colour of the bar.      =
'=                 Standard colour range.                  =
'= BarEndColour  - The colour the bar fades into when the  =
'=                 'BarStyle' is Gradient.                 =
'=                 Standard colour range.                  =
'= BarStartColour- The colour the bar fades from or the    =
'=                 colour of the bar if the 'BarStyle' is  =
'=                 Solid. Standard colour range.           =
'= BarStyle      - The style of bar fill (gradient or      =
'=                 solid).                                 =
'=                 0 = Gradiant, 1 = Solid.                =
'= BorderStyle   - Standard border style.                  =
'=                 0 = Flat, 1 = ThreeD                    =
'= FillDirection - The direction the bar should fill.      =
'=                 0 = Up, 1 = Down, 2 = Left, 3 = Right   =
'= Font          - The font to use for the text.           =
'=                 Standard font dialog.                   =
'= FontColour    - The colour of the text displayed.       =
'=                 Standard colour range.                  =
'= Max           - The upper limit of the bar.             =
'=                 Long value, -2147483648 to 2147483647   =
'= Message       - The message to display in the bar.      =
'=                 String.                                 =
'= Min           - The lower limit of the progress bar.    =
'=                 Long value, -2147483648 to 2147483647   =
'= Percent       - The current bar percentage.             =
'=                 Byte value, 0 to 100 (obviously :))     =
'= PlaySound     - Flag to indicate the sound file         =
'=                 specified in the SoundToPlay property   =
'=                 sould be played when 100% is reached.   =
'=                 (TRUE, FALSE)                           =
'= ShowMessage   - Flag to indicate the message should be  =
'=                 shown. (TRUE, FALSE)                    =
'= ShowPercent   - Flag to incicate the current percentage =
'=                 should be shown. (TRUE, FALSE)          =
'= ShowValue     - Flag to indicate the current value      =
'=                 should be shown. (TRUE, FALSE)          =
'= SoundToPlay   - A string value holding the path and     =
'=                 name of the wav file to play @ 100%.    =
'= Value         - The current value of the progress bar.  =
'=                 Long value, -2147483648 to 2147483647   =
'= VerticalText  - Flag to indicate that the text should   =
'=                 be written top to bottom, useful for up =
'=                 or down progress bars. (TRUE, FALSE)    =
'= WaitForSound  - This flag indicates that the code will  =
'=                 susspend until the sound file played at =
'=                 100% has finished playing.  If one's    =
'=                 set to play that is. (TRUE, FALSE)      =
'===========================================================
'= Notes:                                                  =
'= ------                                                  =
'= 1. You can either show the percentage or value or       =
'=    neither.  You can't show both.  Setting one will     =
'=    disable the other.                                   =
'= 2. Setting the value above the 'Max' or below the 'Min' =
'=    will result in the value being set to the 'Max' or   =
'=    'Min'.                                               =
'= 3. Setting the percent above 100 or below 0 will result =
'=    in the percentage being changed to 100 or 0.         =
'= 4. Setting the 'Max' below the 'Min' will result in the =
'=    'Max' being set to the 'Min' + 1.                    =
'= 5. Setting the 'Min' below the 'Max' will result in the =
'=    'Min' being set to the 'Max' - 1.                    =
'= 6. Adjusting either the 'Max' or the 'Min' will cause   =
'=    the 'Value' to be recalculated.                      =
'= 7. If the 'BarStyle' is set to solid the colour of the  =
'=    bar is defined by the 'BarStartColour' property.     =
'= 8. If a sound is playing and the flag to play one at    =
'=    100% is set the currently playing file will stop and =
'=    the specified one will start.                        =
'===========================================================
'= Have fun! NJE                                           =
'===========================================================

Option Explicit

'API and constant to play wav file.
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1

'API's, types and constants for the bar fills and text generation.
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Private Type Size
    cx As Long
    cy As Long
End Type
Private Type RECT
    vLeft    As Long
    vTop     As Long
    vRight   As Long
    vBottom  As Long
End Type
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type
Private Const PLANES = 14
Private Const BITSPIXEL = 12
Private Const TRANSPARENT = 1

'Memory DC declarations.
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "OlePro32" (ByVal clrAny As Long, ByVal hPal As Long, ByRef clrConvertedOut As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long
Private Const SRCCOPY = &HCC0020
Private Const SRCINVERT = &H660046
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086

'Memory DC variables.
Private hMemoryDC As Long
Private hCompatibleBitmapTmp As Long
Private hCompatibleBitmap As Long

'Fill direction list.
Public Enum FillDirection
    pbUp
    pbDown
    pbLeft
    pbRight
End Enum

'Border style list.
Public Enum BorderStyles
    pbNone
    pbFixedSingle
End Enum

'Appearance style list.
Public Enum AppearanceStyles
    pbFlat
    pbThreeD
End Enum

'Bar style list.
Public Enum BarStyle
    pbGradient
    pbSolid
End Enum

'Local variables to hold property values.
Private mvarPercent As Byte
Private mvarMin As Long
Private mvarMax As Long
Private mvarValue As Long
Private mvarShowPercent As Boolean
Private mvarMessage As String
Private mvarShowMessage As Boolean
Private mvarBarStartColour As OLE_COLOR
Private mvarBarEndColour As OLE_COLOR
Private mvarShowValue As Boolean
Private mvarFillDirection As FillDirection
Private mvarBackColour As OLE_COLOR
Private mvarSoundToPlay As String
Private mvarPlaySound As Boolean
Private mvarVerticalText As Boolean
Private mvarBarStyle As BarStyle
Private mvarWaitForSound As Boolean
Private WithEvents mvarFont As StdFont
Attribute mvarFont.VB_VarHelpID = -1

'Default property values.
Const mdefPercent = 0               'Start percent.
Const mdefMin = 0                   'Lower limit.
Const mdefMax = 100                 'Upper limit.
Const mdefValue = 0                 'Start value.
Const mdefShowPercent = False       'Don't show the percentage.
Const mdefMessage = ""              'No start message.
Const mdefShowMessage = False       'Don't show the message.
Const mdefBarStartColour = &HFF     'Red bar colour start.
Const mdefBarEndColour = &H0        'Black bar colour end.
Const mdefShowValue = False         'Don't show the value.
Const mdefFillDirection = 3         'Right fill.
Const mdefBackColour = &HFFFFFF     'White background.
Const mdefBorderStyle = 1           'ThreeD border style.
Const mdefFontColour = &HFF0000     'Blue Text.
Const mdefVerticalText = False      'Normal left to right text.
Const mdefSoundToPlay = ""          'No initial sound.
Const mdefPlaySound = False         'Don't play sound.
Const mdefBarStyle = 1              'Solid.
Const mdefWaitForSound = False      'Don't bother waiting.

Public Property Let WaitForSound(ByVal vdata As Boolean)
    'Set the wait for sound property.
    mvarWaitForSound = vdata
    'Indicate a property change.
    PropertyChanged "WaitForSound"
End Property

Public Property Get WaitForSound() As Boolean
    'Get the current state of the wait for sound flag.
    WaitForSound = mvarWaitForSound
End Property

Public Property Let BarStyle(ByVal vdata As BarStyle)
    'Check the bar style chosen, if it's outside the available
    'settings set it to Solid.
    If vdata < 0 Or vdata > 1 Then vdata = 1
    'Set the bar style.
    mvarBarStyle = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "BarStyle"
End Property

Public Property Get BarStyle() As BarStyle
    'Get the current barstyle property value.
    BarStyle = mvarBarStyle
End Property

Public Property Let VerticalText(ByVal vdata As Boolean)
    'Set the vertical text flag.
    mvarVerticalText = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "VerticalText"
End Property

Public Property Get VerticalText() As Boolean
    'Get the state of the vertical text flag.
    VerticalText = mvarVerticalText
End Property

Public Property Let SoundToPlay(ByVal vdata As String)
    'Set the sound to play file string.
    mvarSoundToPlay = vdata
    'Indicate a property change.
    PropertyChanged "SoundToPlay"
End Property

Public Property Get SoundToPlay() As String
    'Get the surrent sound file string.
    SoundToPlay = mvarSoundToPlay
End Property

Public Property Let PlaySound(ByVal vdata As Boolean)
    'Set the play sound flag.
    mvarPlaySound = vdata
    'Indicate a property change.
    PropertyChanged "PlaySound"
End Property

Public Property Get PlaySound() As Boolean
    'Get the current play sound flag.
    PlaySound = mvarPlaySound
End Property

Public Property Let FontColour(ByVal vdata As OLE_COLOR)
    'Set the font colour by changing the forecolor.
    UserControl.ForeColor = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "FontColour"
End Property

Public Property Get FontColour() As OLE_COLOR
    'Get the current font colour.
    FontColour = UserControl.ForeColor
End Property

Public Property Let BorderStyle(ByVal vdata As BorderStyles)
    'Set the border style for the progress bar.
    If vdata < 0 Then
        vdata = 0
    ElseIf vdata > 1 Then
        vdata = 1
    End If
    UserControl.BorderStyle = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderStyle() As BorderStyles
    'Get the current border style.
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BarStartColour(ByVal vdata As OLE_COLOR)
    'Set the bar start colour value.
    mvarBarStartColour = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "BarStartColour"
End Property

Public Property Get BarStartColour() As OLE_COLOR
    'Return the start colour value.
    BarStartColour = mvarBarStartColour
End Property

Public Property Let BarEndColour(ByVal vdata As OLE_COLOR)
    'Set the bar end colour.
    mvarBarEndColour = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "BarEndColour"
End Property

Public Property Get BarEndColour() As OLE_COLOR
    'Return the end bar colour.
    BarEndColour = mvarBarEndColour
End Property

Public Property Let BackColour(ByVal vdata As OLE_COLOR)
    'Set the back colour.
    mvarBackColour = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "BackColour"
End Property

Public Property Get BackColour() As OLE_COLOR
    'Get the current back colour.
    BackColour = mvarBackColour
End Property

Public Property Let Value(ByVal vdata As Long)
Attribute Value.VB_Description = "Returns/sets the value on the progress bar."
    'Make sure the value chosen resides in the set range.
    If vdata < mvarMin Then
        vdata = mvarMin
    ElseIf vdata > mvarMax Then
        vdata = mvarMax
    End If
    'Set the current progress bar value.
    mvarValue = vdata
    'Calculate the percentage.
    mvarPercent = Int(((mvarValue - mvarMin) / (mvarMax - mvarMin)) * 100)
    'Update the control.
    UserControl_Paint
    'Indicate property changes.
    PropertyChanged "Value"
    PropertyChanged "Percent"
End Property

Public Property Get Value() As Long
    'Return the current value.
    Value = mvarValue
End Property

Public Property Let Min(ByVal vdata As Long)
Attribute Min.VB_Description = "Returns/sets the progress bars lower limit."
    'Check the min value is at least 1 less than
    'the max value
    If vdata >= mvarMax Then vdata = mvarMax - 1
    'Set the start value of the progress bar.
    mvarMin = vdata
    'Recalculate the value.
    mvarValue = Int(((mvarPercent / 100) * (mvarMax - mvarMin)) + mvarMin)
    'Update the control.
    UserControl_Paint
    'Indicate property changes.
    PropertyChanged "Min"
    PropertyChanged "Value"
End Property

Public Property Get Min() As Long
    'Return the value of the start.
    Min = mvarMin
End Property

Public Property Let ShowValue(ByVal vdata As Boolean)
Attribute ShowValue.VB_Description = "Returns/sets the flag to indicate the value should be shown."
    'Set the flag to indicate the value should be shown
    'in the progress bar.
    mvarShowValue = vdata
    'Check to see if the percentage is set to show in the
    'progress bar and disable it.
    If mvarShowValue = True Then
        mvarShowPercent = False
        'Indicate a property change.
        PropertyChanged "ShowPercent"
    End If
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "ShowValue"
End Property

Public Property Get ShowValue() As Boolean
    'Return the current state of the value show flag.
    ShowValue = mvarShowValue
End Property

Public Property Let ShowPercent(ByVal vdata As Boolean)
Attribute ShowPercent.VB_Description = "Returns/sets the flag to indicate the percentage should be shown."
    'Set the flag to indicate the percentage should be shown
    'in the progress bar.
    mvarShowPercent = vdata
    'Check to see if the value is set to be shown and
    'disable it.
    If mvarShowPercent = True Then
        mvarShowValue = False
        'Indicate a property change.
        PropertyChanged "ShowValue"
    End If
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "ShowPercent"
End Property

Public Property Get ShowPercent() As Boolean
    'Return the flag state for the percent showing.
    ShowPercent = mvarShowPercent
End Property

Public Property Let ShowMessage(ByVal vdata As Boolean)
Attribute ShowMessage.VB_Description = "Returns/sets the flag to indicate the message should be shown."
    'Set the flag to indicate the message should be shown.
    mvarShowMessage = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "ShowMessage"
End Property

Public Property Get ShowMessage() As Boolean
    'Return the show message flag.
    ShowMessage = mvarShowMessage
End Property

Public Property Let Percent(ByVal vdata As Byte)
Attribute Percent.VB_Description = "Returns/sets the percentage on the progress bar."
    'Ensure the percent chosen is between 0 and 100.
    If vdata < 0 Then
        vdata = 0
    ElseIf vdata > 100 Then
        vdata = 100
    End If
    'Set the percent property.
    mvarPercent = vdata
    'Calculate the value.
    mvarValue = Int(((mvarPercent / 100) * (mvarMax - mvarMin)) + mvarMin)
    'Update the control.
    UserControl_Paint
    'Indicate property changes.
    PropertyChanged "Percent"
    PropertyChanged "Value"
End Property

Public Property Get Percent() As Byte
    'Return the current percentage of the progress bar.
    Percent = mvarPercent
End Property

Public Property Let Message(ByVal vdata As String)
    'Set message to show in the progress bar.
    mvarMessage = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "Message"
End Property

Public Property Get Message() As String
    'Return the message to show.
    Message = mvarMessage
End Property

Public Property Let Max(ByVal vdata As Long)
    'Check that the max value is at least 1 higher than
    'the minimum value.
    If vdata <= mvarMin Then vdata = mvarMin + 1
    'Set the finish value for the progress bar.
    mvarMax = vdata
    'Recalculate the value.
    mvarValue = Int(((mvarPercent / 100) * (mvarMax - mvarMin)) + mvarMin)
    'Update the control.
    UserControl_Paint
    'Indicate property changes.
    PropertyChanged "Max"
    PropertyChanged "Value"
End Property

Public Property Get Max() As Long
    'Return the finish value.
    Max = mvarMax
End Property

Public Property Let FillDirection(ByVal vdata As FillDirection)
Attribute FillDirection.VB_Description = "Returns/sets the the fill direction of the progress bar."
    'Set the direction of the fill to right if it's invalid.
    If vdata < 0 Or vdata > 3 Then
        vdata = 3
    End If
    'Save the setting in the property variable.
    mvarFillDirection = vdata
    'Update the control.
    UserControl_Paint
    'Indicate a property change.
    PropertyChanged "FillDirection"
End Property

Public Property Get FillDirection() As FillDirection
    'Return the current fill direction.
    FillDirection = mvarFillDirection
End Property

Public Property Get Font() As Font
    'Return the current font.
    Set Font = mvarFont
End Property

Public Property Set Font(mnewFont As StdFont)
    'Set the current font.
    With mvarFont
        .Bold = mnewFont.Bold
        .Italic = mnewFont.Italic
        .Name = mnewFont.Name
        .Size = mnewFont.Size
        .Strikethrough = mnewFont.Strikethrough
        .Underline = mnewFont.Underline
    End With
    'Indicate a property change.
    PropertyChanged "Font"
End Property

Private Sub mvarFont_FontChanged(ByVal PropertyName As String)
   Set UserControl.Font = mvarFont
   Refresh
End Sub

Private Sub UserControl_Initialize()
   Set mvarFont = New StdFont
   Set UserControl.Font = mvarFont
End Sub

Private Sub UserControl_InitProperties()
    'Set the defaults.
    mvarFillDirection = mdefFillDirection
    mvarMin = mdefMin
    mvarMax = mdefMax
    mvarValue = mdefValue
    mvarPercent = mdefPercent
    mvarMessage = mdefMessage
    mvarShowMessage = mdefShowMessage
    mvarShowPercent = mdefShowPercent
    mvarShowValue = mdefShowValue
    UserControl.BorderStyle = mdefBorderStyle
    mvarBackColour = mdefBackColour
    mvarBarStartColour = mdefBarStartColour
    mvarBarEndColour = mdefBarEndColour
    UserControl.ForeColor = mdefFontColour
    mvarVerticalText = mdefVerticalText
    mvarSoundToPlay = mdefSoundToPlay
    mvarPlaySound = mdefPlaySound
    mvarBarStyle = mdefBarStyle
    mvarWaitForSound = mdefWaitForSound
End Sub

Private Sub UserControl_Paint()
    'Draw the bar.
    DrawBar
    'Play the wav file.
    PlayWav
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Restore the saved properties.
    mvarBackColour = PropBag.ReadProperty("BackColour", mdefBackColour)
    mvarBarStartColour = PropBag.ReadProperty("BarStartColour", mdefBarStartColour)
    mvarBarEndColour = PropBag.ReadProperty("BarEndColour", mdefBarEndColour)
    mvarFillDirection = PropBag.ReadProperty("FillDirection", mdefFillDirection)
    mvarMax = PropBag.ReadProperty("Max", mdefMax)
    mvarMessage = PropBag.ReadProperty("Message", mdefMessage)
    mvarMin = PropBag.ReadProperty("Min", mdefMin)
    mvarPercent = PropBag.ReadProperty("Percent", mdefPercent)
    mvarShowMessage = PropBag.ReadProperty("ShowMessage", mdefShowMessage)
    mvarShowPercent = PropBag.ReadProperty("ShowPercent", mdefShowPercent)
    mvarShowValue = PropBag.ReadProperty("ShowValue", mdefShowValue)
    mvarValue = PropBag.ReadProperty("Value", mdefValue)
    mvarVerticalText = PropBag.ReadProperty("VerticalText", mdefVerticalText)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", mdefBorderStyle)
    UserControl.ForeColor = PropBag.ReadProperty("FontColour", mdefFontColour)
    mvarSoundToPlay = PropBag.ReadProperty("SoundToPlay", mdefSoundToPlay)
    mvarPlaySound = PropBag.ReadProperty("PlaySound", mdefPlaySound)
    mvarBarStyle = PropBag.ReadProperty("BarStyle", mdefBarStyle)
    mvarWaitForSound = PropBag.ReadProperty("WaitForSound", mdefWaitForSound)
    Set mvarFont = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    'Recreate the memory DC seeing as the old one is now an incorrect size.
    'Delete the memory DC.
    SelectObject hMemoryDC, hCompatibleBitmap
    DeleteObject hCompatibleBitmapTmp
    DeleteDC hMemoryDC
    'Create a memory DC of the usercontrol.
    hMemoryDC = CreateCompatibleDC(UserControl.hdc)
    hCompatibleBitmapTmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
    hCompatibleBitmap = SelectObject(hMemoryDC, hCompatibleBitmapTmp)

    'Repaint the control.
    UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
    'Delete the memory DC.
    SelectObject hMemoryDC, hCompatibleBitmap
    DeleteObject hCompatibleBitmapTmp
    DeleteDC hMemoryDC
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Save the instances current properties.
    Call PropBag.WriteProperty("BackColour", mvarBackColour, mdefBackColour)
    Call PropBag.WriteProperty("BarStartColour", mvarBarStartColour, mdefBarStartColour)
    Call PropBag.WriteProperty("BarEndColour", mvarBarEndColour, mdefBarEndColour)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, mdefBorderStyle)
    Call PropBag.WriteProperty("FillDirection", mvarFillDirection, mdefFillDirection)
    Call PropBag.WriteProperty("FontColour", UserControl.ForeColor, mdefFontColour)
    Call PropBag.WriteProperty("Max", mvarMax, mdefMax)
    Call PropBag.WriteProperty("Message", mvarMessage, mdefMessage)
    Call PropBag.WriteProperty("Min", mvarMin, mdefMin)
    Call PropBag.WriteProperty("Percent", mvarPercent, mdefPercent)
    Call PropBag.WriteProperty("ShowMessage", mvarShowMessage, mdefShowMessage)
    Call PropBag.WriteProperty("ShowPercent", mvarShowPercent, mdefShowPercent)
    Call PropBag.WriteProperty("ShowValue", mvarShowValue, mdefShowValue)
    Call PropBag.WriteProperty("Value", mvarValue, mdefValue)
    Call PropBag.WriteProperty("VerticalText", mvarVerticalText, mdefVerticalText)
    Call PropBag.WriteProperty("SoundToPlay", mvarSoundToPlay, mdefSoundToPlay)
    Call PropBag.WriteProperty("PlaySound", mvarPlaySound, mdefPlaySound)
    Call PropBag.WriteProperty("BarStyle", mvarBarStyle, mdefBarStyle)
    Call PropBag.WriteProperty("WaitForSound", mvarWaitForSound, mdefWaitForSound)
    Call PropBag.WriteProperty("Font", mvarFont)
    Call PropBag.WriteProperty("BarEndColour", mvarBarEndColour, mdefBarEndColour)
End Sub

Private Sub DrawBar()
'================================================================================
'= Draw Bar
'================================================================================

    'Local variables for determining colour depth.
    Static intRgnCnt As Integer
    Dim lngBitsPerPixel As Long
    Dim lngNbrPlanes As Long
    Dim lngColourBits As Long
    
    'Local variables for drawing.
    Dim lngAreaHeight As Long
    Dim lngAreaWidth As Long
    Dim sngRedLevel As Single
    Dim sngGreenLevel As Single
    Dim sngBlueLevel As Single
    Dim sngRedColourVal As Single
    Dim sngGreenColourVal As Single
    Dim sngBlueColourVal As Single
    Dim dblIntervalY As Double
    Dim dblIntervalX As Double
    Dim dblCurrentY As Double
    Dim dblCurrentX As Double
    Dim i As Integer
    Dim FillArea As RECT
    Dim hBrush As Long
    
    If intRgnCnt = 0 Then
        'Determine number of color bits supported.
        lngBitsPerPixel = GetDeviceCaps(UserControl.hdc, BITSPIXEL)
        lngNbrPlanes = GetDeviceCaps(UserControl.hdc, PLANES)
        lngColourBits = (lngBitsPerPixel * lngNbrPlanes)
        'Calculate the number of regions that the screen will be divided into.
        'This is optimized for the current display's color depth.  Why waste
        'time rendering 256 shades if you can only discern 32 or 64 of them?
        Select Case lngColourBits
            Case 32:   intRgnCnt = 256     '16M colors:  8 bits for blue
            Case 24:   intRgnCnt = 256     '16M colors:  8 bits for blue
            Case 16:   intRgnCnt = 256     '64K colors:  5 bits for blue
            Case 15:   intRgnCnt = 32      '32K colors:  5 bits for blue
            Case 8:    intRgnCnt = 64      '256 colors:  64 dithered blues
            Case 4:    intRgnCnt = 64      '16 colors :  64 dithered blues
            Case Else: lngColourBits = 4
                intRgnCnt = 64      '16 colors assumed: 64 dithered blues
        End Select
        
        'Create a memory DC of the usercontrol.
        hMemoryDC = CreateCompatibleDC(UserControl.hdc)
        hCompatibleBitmapTmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
        hCompatibleBitmap = SelectObject(hMemoryDC, hCompatibleBitmapTmp)
    End If
    
    'Clear the memory DC with the specified background colour.
    FillArea.vLeft = 0
    FillArea.vTop = 0
    FillArea.vRight = UserControl.ScaleWidth
    FillArea.vBottom = UserControl.ScaleHeight
    hBrush = CreateSolidBrush(RGB(mvarBackColour And &HFF&, (mvarBackColour And &HFF00&) \ &H100&, (mvarBackColour And &HFF0000) \ &H10000))
    Call FillRect(hMemoryDC, FillArea, hBrush)
    Call DeleteObject(hBrush)
       
    'Get the current pixel sizes.
    lngAreaHeight = UserControl.ScaleHeight
    lngAreaWidth = UserControl.ScaleWidth
    
    'Determine start colour levels.
    sngRedLevel = mvarBarStartColour And &HFF&
    sngGreenLevel = (mvarBarStartColour And &HFF00&) \ &H100&
    sngBlueLevel = (mvarBarStartColour And &HFF0000) \ &H10000
       
    'Set the fill area to the entire bar.
    FillArea.vLeft = 0
    FillArea.vTop = 0
    FillArea.vRight = lngAreaWidth
    FillArea.vBottom = lngAreaHeight
    
    'If the bar is solid adjust the fill area to the current percentage.
    If mvarBarStyle = 1 Then
        Select Case mvarFillDirection
            Case 0 'UP
                FillArea.vTop = lngAreaHeight - ((lngAreaHeight / 100) * mvarPercent)
            Case 1 'DOWN
                FillArea.vBottom = (lngAreaHeight / 100) * mvarPercent
            Case 2 'LEFT
                FillArea.vLeft = lngAreaWidth - ((lngAreaWidth / 100) * mvarPercent)
            Case 3 'RIGHT
                FillArea.vRight = (lngAreaWidth / 100) * mvarPercent
        End Select
        'Fill the defined area with the bar start colour.
        hBrush = CreateSolidBrush(RGB(sngRedLevel, sngGreenLevel, sngBlueLevel))
        Call FillRect(hMemoryDC, FillArea, hBrush)
        Call DeleteObject(hBrush)
    'If it's a gradient fill run this code.
    Else
        'Number of pixels per region.
        dblIntervalY = lngAreaHeight / intRgnCnt
        dblIntervalX = lngAreaWidth / intRgnCnt
        'Colour difference between regions.
        sngRedColourVal = ((mvarBarEndColour And &HFF&) - sngRedLevel) / intRgnCnt
        sngGreenColourVal = (((mvarBarEndColour And &HFF00&) \ &H100&) - sngGreenLevel) / intRgnCnt
        sngBlueColourVal = (((mvarBarEndColour And &HFF0000) \ &H10000) - sngBlueLevel) / intRgnCnt
        'Work through the number of regions the bar has been split into.
        For i = 0 To intRgnCnt - 1
            'Create a brush of the appropriate colour.
            hBrush = CreateSolidBrush(RGB(Int(sngRedLevel), Int(sngGreenLevel), Int(sngBlueLevel)))
            'Select the appropriate fill direction.
            Select Case mvarFillDirection
                Case 0 'UP
                    'Adjust the fill area to the current region.
                    FillArea.vTop = lngAreaHeight - dblCurrentY - dblIntervalY
                    FillArea.vBottom = lngAreaHeight - dblCurrentY
                    'Fill this area if the area is shown, otherwise exit the loop.
                    If FillArea.vTop > lngAreaHeight - ((lngAreaHeight / 100) * mvarPercent) Then
                        Call FillRect(hMemoryDC, FillArea, hBrush)
                    Else
                        Exit For
                    End If
                Case 1 'DOWN
                    'Adjust the fill area to the current region.
                    FillArea.vTop = dblCurrentY
                    FillArea.vBottom = dblCurrentY + dblIntervalY
                    'Fill this area if the area is shown, otherwise exit the loop.
                    If FillArea.vBottom < (lngAreaHeight / 100) * mvarPercent Then
                        Call FillRect(hMemoryDC, FillArea, hBrush)
                    Else
                        Exit For
                    End If
                Case 2 'LEFT
                    'Adjust the fill area to the current region.
                    FillArea.vLeft = lngAreaWidth - dblCurrentX - dblIntervalX
                    FillArea.vRight = lngAreaWidth - dblCurrentX
                    'Fill this area if the area is shown, otherwise exit the loop.
                    If FillArea.vLeft > lngAreaWidth - ((lngAreaWidth / 100) * mvarPercent) Then
                        Call FillRect(hMemoryDC, FillArea, hBrush)
                    Else
                        Exit For
                    End If
                Case 3 'RIGHT
                    'Adjust the fill area to the current region.
                    FillArea.vLeft = dblCurrentX
                    FillArea.vRight = dblCurrentX + dblIntervalX
                    'Fill this area if the area is shown, otherwise exit the loop.
                    If FillArea.vRight < (lngAreaWidth / 100) * mvarPercent Then
                        Call FillRect(hMemoryDC, FillArea, hBrush)
                    Else
                        Exit For
                    End If
            End Select
            'Done with that brush, so delete it.
            Call DeleteObject(hBrush)
            'Increment the current X and Y locations.
            dblCurrentY = dblCurrentY + dblIntervalY
            dblCurrentX = dblCurrentX + dblIntervalX
            'Increment display colour depth.
            sngRedLevel = sngRedLevel + sngRedColourVal
            sngGreenLevel = sngGreenLevel + sngGreenColourVal
            sngBlueLevel = sngBlueLevel + sngBlueColourVal
        Next
        'Check to see if we bailed out of the for loop, if so
        'delete the brush.
        If i < intRgnCnt - 1 Then
            Call DeleteObject(hBrush)
        Else
            'If we're at the end of the bar.
            'Fill any of the remaining spaces with the bar end colour.
            Select Case mvarFillDirection
                Case 0 'UP
                    FillArea.vTop = 0
                    FillArea.vBottom = FillArea.vTop + dblIntervalY
                Case 1 'DOWN
                    FillArea.vBottom = lngAreaHeight
                    FillArea.vTop = FillArea.vBottom - dblIntervalY
                Case 2 'LEFT
                    FillArea.vLeft = 0
                    FillArea.vRight = FillArea.vLeft + dblIntervalX
                Case 3 'RIGHT
                    FillArea.vRight = lngAreaWidth
                    FillArea.vLeft = FillArea.vRight - dblIntervalX
            End Select
            hBrush = CreateSolidBrush(RGB(mvarBarEndColour And &HFF&, (mvarBarEndColour And &HFF00&) \ &H100&, (mvarBarEndColour And &HFF0000) \ &H10000))
            Call FillRect(hMemoryDC, FillArea, hBrush)
            Call DeleteObject(hBrush)
        End If
    End If
    
'================================================================================
'= Draw Text
'================================================================================
    Call DrawText
    
'================================================================================
'= Copy memory DC to control.
'================================================================================
    Call BitBlt(UserControl.hdc, 0, 0, UserControl.ScaleWidth, _
        UserControl.ScaleHeight, hMemoryDC, 0, 0, SRCCOPY)

End Sub

Private Sub DrawText()
    'Local variables.
    Dim strBarTxt As String
    Dim i As Integer
    Dim dblVertCurrentY As Double
    Dim typFont As LOGFONT
    Dim typTextMetric As TEXTMETRIC
    Dim typSize As Size
    Dim hPrevFont As Long, hFont As Long

'================================================================================
'= Generate text.
'================================================================================
    'If we want to show any text then draw it.
    If mvarShowMessage Or mvarShowPercent Or mvarShowValue Then
        'Set message if there's one flagged to show.
        If mvarShowMessage Then strBarTxt = mvarMessage
        'Add the percent or value if either are flagged to show.
        If mvarShowPercent Or mvarShowValue Then
            'Add a space if the percentage or value is to be shown and there is a message.
            If Len(strBarTxt) > 0 Then strBarTxt = strBarTxt & " "
            'Add the percentage if it's flagged to show.
            If mvarShowPercent Then
                strBarTxt = strBarTxt & Format$(mvarPercent, "##0") + "%"
            'Add the value if it's flagged to show.
            ElseIf mvarShowValue Then
                strBarTxt = strBarTxt & Trim(Str(mvarValue)) & "/" & Trim(Str(mvarMax))
            End If
        End If
'================================================================================
'= Draw text to DC.
'================================================================================
        'Set the font up.
        typFont.lfEscapement = 0
        typFont.lfFacename = mvarFont.Name & Chr$(0)
        typFont.lfHeight = (mvarFont.Size * -20) / Screen.TwipsPerPixelY
        typFont.lfItalic = mvarFont.Italic
        If mvarFont.Bold Then
            typFont.lfWeight = 700
        Else
            typFont.lfWeight = 400
        End If
        typFont.lfUnderline = mvarFont.Underline
        typFont.lfStrikeOut = mvarFont.Strikethrough
        
        'Create the font.
        hFont = CreateFontIndirect(typFont)
        hPrevFont = SelectObject(hMemoryDC, hFont)
        Call SetBkMode(hMemoryDC, TRANSPARENT)
        
        'Set the backcolour to the text colour.
        Call SetTextColor(hMemoryDC, UserControl.ForeColor)

        'Get the font metrics.
        Call GetTextMetrics(hMemoryDC, typTextMetric)
        
        'Draw the text vertically is so flagged.
        If mvarVerticalText Then
            'Calculate the total height of all the text.
            For i = 1 To Len(strBarTxt)
                Call GetTextExtentPoint32(hMemoryDC, Mid(strBarTxt, i, 1), Len(Mid(strBarTxt, i, 1)), typSize)
                dblVertCurrentY = dblVertCurrentY + typSize.cy
            Next i
            'Set the Y coord to the begining letter of the text.
            dblVertCurrentY = (UserControl.ScaleHeight - dblVertCurrentY) / 2
            'Work through each letter of the text and place it on the progress bar.
            For i = 1 To Len(strBarTxt)
                Call GetTextExtentPoint32(hMemoryDC, Mid(strBarTxt, i, 1), Len(Mid(strBarTxt, i, 1)), typSize)
                'Draw the letter to the memory DC.
                Call TextOut(hMemoryDC, _
                    (UserControl.ScaleWidth - typSize.cx) / 2, _
                    dblVertCurrentY, Mid(strBarTxt, i, 1), Len(Mid(strBarTxt, i, 1)))
                'Move the Y coord pointer for the next letter.
                dblVertCurrentY = dblVertCurrentY + typSize.cy
            Next i
        'Otherwise draw the text the normal left to right.
        Else
            Call GetTextExtentPoint32(hMemoryDC, strBarTxt, Len(strBarTxt), typSize)
            'Draw the text to the memory DC.
            Call TextOut(hMemoryDC, _
                (UserControl.ScaleWidth - typSize.cx) / 2, _
                (UserControl.ScaleHeight - typSize.cy) / 2, strBarTxt, Len(strBarTxt))
        End If
        'Reset the control's font settings.
        Call SelectObject(hMemoryDC, hPrevFont)
        Call DeleteObject(hFont)
    End If
End Sub

Private Sub PlayWav()
    'If the percentage has reached 100 and the flag to play
    'a sound is on, then play the wav file.
    If mvarPercent = 100 And mvarPlaySound Then
        'If the file can be found then play it.
        If Dir(mvarSoundToPlay) <> "" Then
            'If we're supposed to wait for the sound to finish
            'then play the sound sync'ed.
            If mvarWaitForSound Then
                sndPlaySound mvarSoundToPlay, SND_SYNC
            Else
                sndPlaySound mvarSoundToPlay, SND_ASYNC
            End If
        End If
    End If
End Sub

