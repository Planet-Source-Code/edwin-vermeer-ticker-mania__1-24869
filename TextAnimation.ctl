VERSION 5.00
Begin VB.UserControl TextAnimation 
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   106
   ToolboxBitmap   =   "TextAnimation.ctx":0000
   Begin VB.Timer ReDrawTimer 
      Interval        =   20
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   240
      ScaleHeight     =   840
      ScaleWidth      =   840
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "TextAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As Rect, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lprect As Rect) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lprect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As Rect) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Boolean
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (hwnd As Long, region As Rect, hRgn As Long, Flags As Integer) As Boolean
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const BLACKNESS = &H42              '(DWORD) dest = BLACK
Private Const NOTSRCCOPY = &H330008         '(DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6        '(DWORD) dest = (NOT src) AND (NOT dest)
Private Const SRCAND = &H8800C6             '(DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020            '(DWORD) dest = source
Private Const SRCERASE = &H440328           '(DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046          '(DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086           '(DWORD) dest = source OR dest
Private Const WHITENESS = &HFF0062          '(DWORD) dest = WHITE
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Private Const ANTIALIASED_QUALITY = 4 ' Ensure font edges are smoothed if system is set to smooth font edges
Private Const RDW_INTERNALPAINT = &H2

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
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type TextMessage
    MessageID As String
    MessageText As String
    MessageFontName As String
    MessageFontColorStart As OLE_COLOR
    MessageFontColorEnd As OLE_COLOR
    MessageFontSizeStart As Integer
    MessageFontSizeEnd As Integer
    MessageLeftStart As Integer
    MessageLeftEnd As Integer
    MessageTopStart As Integer
    MessageTopEnd As Integer
    MessageFontRotationStart As Integer
    MessageFontRotationEnd As Integer
    MessageIntervalStart As Long
    MessageIntervalCount As Long
End Type

Public Enum SPBorderStyle
    [None] = 0
    [Fixed Single] = 1
End Enum

Public Enum SPBackGroundStyle
    [Gradient] = 0
    [Picture] = 1
    [TransparentPicture] = 2
    [Transparent] = 3
    [FormTransparent] = 4
End Enum

Public Enum SPGradBlendMode
    [mRGB] = 0
    [mHSL] = 1
End Enum

Public Enum SPGradType
    [mNormal] = 0
    [mElliptical] = 1
    [mRectangular] = 2
End Enum

Private m_messages() As TextMessage
Dim m_counter As Long
Const m_def_counter = 0
Dim m_counterMax As Long
Const m_def_counterMax = 1800
Dim m_backcolorStart As OLE_COLOR
Const m_def_backcolorStart = 8388607  'RGB(255, 255, 127)
Dim m_backcolorEnd As OLE_COLOR
Const m_def_backcolorEnd = 16744319   'RGB(127, 127, 255)
Dim m_Border As Integer
Const m_def_Border = [None]
Dim m_BackGroundStyle As Integer
Const m_def_BackGroundStyle = [Gradient]
Dim m_BackGroundImage As String
Const m_def_BackGroundImage = ""
Dim m_Enabled As Boolean
Const m_def_Enabled = True
Dim m_Speed As Integer
Const m_def_Speed = 20
Dim m_AnimateInDesignmode As Boolean
Const m_def_AnimateInDesignmode = True
Dim m_TransparentColor As OLE_COLOR
Const m_def_TransparentColor = 16711935     'RGB(255, 0, 255)
Dim m_Angle As Integer
Const m_def_Angle = 0
Dim m_Repetitions As Integer
Const m_def_Repetitions = 1
Dim m_GradientType As Integer
Const m_def_GradientType = [mNormal]
Dim m_BlendMode As Integer
Const m_def_BlendMode = [mHSL]

Event BeforeDraw(PictureBuffer As PictureBox)
Event AfterDraw(PictureBuffer As PictureBox)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
Event MouseIn()
Event MouseOut()

Dim bMouseIn As Boolean
Dim ReEnter As Integer
Dim sButton As Integer
Dim sShift As Integer
Dim sX As Single
Dim SY As Single

Private Sub RedrawTimer_Timer()
On Error Resume Next
Dim RedFrom As Integer
Dim GreenFrom As Integer
Dim BlueFrom As Integer
Dim RedTo As Integer
Dim GreenTo As Integer
Dim BlueTo As Integer
Dim L As Long
Dim j As Long
Dim tLF As LOGFONT
Dim hFnt As Long
Dim hFntOld As Long
Dim lR As Long
Dim iChar As Integer
Dim rgn As Rect
Dim region As Long
Dim messages() As String

    ' Just some initialisation of the control
    ReEnter = ReEnter + 1
    If ReEnter > 1 Then If ReEnter < 100 Then Exit Sub
    m_counter = m_counter + 1
    If Counter > CounterMax Then Counter = 0
    picBuffer.BackColor = m_TransparentColor
    If m_BackGroundStyle = [Gradient] Or m_BackGroundStyle = [Picture] Or m_BackGroundStyle = [TransparentPicture] Then
      L = BitBlt(picBuffer.hdc, 0, picBackBuffer.ScaleTop, picBackBuffer.ScaleWidth, picBackBuffer.ScaleHeight, picBackBuffer.hdc, 0, 0, SRCCOPY)
    Else
      picBuffer.Cls
    End If
    RaiseEvent BeforeDraw(picBuffer)
    
    ' It is not an actual mousemove, but maybe the display underneath has changed.
    If bMouseIn Then
      BuildAray sX, SY, messages
      RaiseEvent MouseMove(sButton, sShift, sX, SY, messages)
    End If
    
    ' Now put in all the messages
    For j = 0 To MessageCount
      If m_messages(j).MessageIntervalStart <= m_counter And m_messages(j).MessageIntervalStart + m_messages(j).MessageIntervalCount > m_counter Then
        ' The text color
        RedFrom = m_messages(j).MessageFontColorStart And RGB(255, 0, 0)
        GreenFrom = (m_messages(j).MessageFontColorStart And RGB(0, 255, 0)) / 256
        BlueFrom = (m_messages(j).MessageFontColorStart And RGB(0, 0, 255)) / 65536
        RedTo = m_messages(j).MessageFontColorEnd And RGB(255, 0, 0)
        GreenTo = (m_messages(j).MessageFontColorEnd And RGB(0, 255, 0)) / 256
        BlueTo = (m_messages(j).MessageFontColorEnd And RGB(0, 0, 255)) / 65536
        picBuffer.ForeColor = RGB(RedFrom - (RedFrom - RedTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, GreenFrom - (GreenFrom - GreenTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, BlueFrom - (BlueFrom - BlueTo) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount)
        ' The text size
        tLF.lfHeight = MulDiv((m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount), (GetDeviceCaps(picBuffer.hdc, LOGPIXELSY)), 72)
        ' The rotation of the font
        tLF.lfEscapement = m_messages(j).MessageFontRotationStart - (m_messages(j).MessageFontRotationStart - m_messages(j).MessageFontRotationEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
        
        ' The text font
        For iChar = 1 To Len(m_messages(j).MessageFontName)
            tLF.lfFaceName(iChar - 1) = CByte(Asc(Mid$(m_messages(j).MessageFontName, iChar, 1)))
        Next iChar
        ' Other font properties (for now default)
        tLF.lfItalic = picBuffer.Font.Italic
        If (picBuffer.Font.Bold) Then
            tLF.lfWeight = FW_BOLD
        Else
            tLF.lfWeight = FW_NORMAL
        End If
        tLF.lfUnderline = picBuffer.Font.Underline
        tLF.lfStrikeOut = picBuffer.Font.Strikethrough
        tLF.lfCharSet = picBuffer.Font.Charset
        tLF.lfQuality = ANTIALIASED_QUALITY
        ' Print the text at the right location
        hFnt = CreateFontIndirect(tLF)
        If (hFnt <> 0) Then
          hFntOld = SelectObject(picBuffer.hdc, hFnt)
          lR = TextOut(picBuffer.hdc, m_messages(j).MessageLeftStart - (m_messages(j).MessageLeftStart - m_messages(j).MessageLeftEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, m_messages(j).MessageTopStart - (m_messages(j).MessageTopStart - m_messages(j).MessageTopEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount, m_messages(j).MessageText, lstrlen(m_messages(j).MessageText))
          SelectObject picBuffer.hdc, hFntOld
          DeleteObject hFnt
        End If
      End If
    Next j
    RaiseEvent AfterDraw(picBuffer)
    
    ' Now make the right things visible
    Select Case m_BackGroundStyle
    Case [Gradient], [Picture]
      UserControl.BackStyle = 1
      L = BitBlt(UserControl.hdc, 0, picBuffer.ScaleTop, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBuffer.hdc, 0, 0, SRCCOPY)
    Case [Transparent], [TransparentPicture]
      LockWindowUpdate UserControl.Parent.hwnd
      UserControl.MaskColor = m_TransparentColor
      UserControl.BackStyle = 0
      UserControl.Picture = picBuffer.Image
      UserControl.MaskPicture = picBuffer.Image
      LockWindowUpdate 0
    Case [FormTransparent]
      UserControl.BackStyle = 1
      UserControl.Picture = picBuffer.Image
      If UserControl.Ambient.UserMode Then
        SetAutoRgn UserControl.Parent, UserControl.Picture, m_TransparentColor
      End If
    End Select
    
    ' Because of this the animation in design mode will only stop after the first paint
    If Not UserControl.Ambient.UserMode Then UserControl!ReDrawTimer.Enabled = m_AnimateInDesignmode
    DoEvents
    ReEnter = 0
    
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim messages() As String
  BuildAray X, Y, messages
  RaiseEvent MouseDown(Button, Shift, X, Y, messages)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim messages() As String
  BuildAray X, Y, messages
  
  sButton = Button
  sShift = Shift
  sX = X
  SY = Y
  
  'If the mouse has left the region of the control then
  If (X < 0) Or (X > UserControl.Width / Screen.TwipsPerPixelX) Or (Y < 0) Or (Y > UserControl.Height / Screen.TwipsPerPixelY) Then
    ReleaseCapture
    RaiseEvent MouseMove(Button, Shift, X, Y, messages)
    RaiseEvent MouseOut
    bMouseIn = False
  Else
    SetCapture UserControl.hwnd 'Capture the mouse to the control so as to track it's movement
    If Not bMouseIn Then RaiseEvent MouseIn
    RaiseEvent MouseMove(Button, Shift, X, Y, messages)
    bMouseIn = True
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim messages() As String
  BuildAray X, Y, messages
  ReleaseCapture
  RaiseEvent MouseUp(Button, Shift, X, Y, messages)
End Sub

Private Sub BuildAray(X As Single, Y As Single, messages() As String)
Dim j As Integer
Dim t As Integer

  For j = 0 To MessageCount
    If m_messages(j).MessageIntervalStart <= m_counter And m_messages(j).MessageIntervalStart + m_messages(j).MessageIntervalCount > m_counter Then
      t = m_messages(j).MessageLeftStart - (m_messages(j).MessageLeftStart - m_messages(j).MessageLeftEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
      If m_messages(j).MessageFontRotationStart = 0 And m_messages(j).MessageFontRotationEnd = 0 Then
        If X >= t And X <= t + MessageWidth(j) Then
          t = m_messages(j).MessageTopStart - (m_messages(j).MessageTopStart - m_messages(j).MessageTopEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount
          If Y >= t And Y <= t + MessageHeight(j) Then
            On Error Resume Next
            ReDim Preserve messages(UBound(messages) + 1) As String
            If Err.Number = 9 Then ReDim Preserve messages(0)
            On Error GoTo 0
            messages(UBound(messages)) = MessageID(j)
          End If
        End If
      End If
    End If
  Next j
  
End Sub


Private Function PaintBackgroundBuffer(picbox As PictureBox)

  picBackBuffer.BackColor = m_TransparentColor
  Select Case m_BackGroundStyle
  Case [Gradient]
    GradientBackground picbox
  Case [Picture], [TransparentPicture]
'    It's already in the control itself. Else you would need:
'    picbox.Picture = LoadPicture(m_BackGroundImage)
  Case [Transparent], [FormTransparent]
    ' do nothing. The form will be used
  End Select
  
End Function


Private Function GradientBackground(picbox As PictureBox)
Dim grad As New Gradient
  grad.Color1 = m_backcolorStart
  grad.Color2 = m_backcolorEnd
  grad.Angle = m_Angle
  grad.Repetitions = m_Repetitions
  grad.GradientType = m_GradientType
  grad.BlendMode = m_BlendMode
  grad.Draw picbox
    
End Function

Private Sub rePaint()
    
  UserControl_Resize
  If UserControl!ReDrawTimer.Enabled = False Then
    m_counter = m_counter - 1
    If m_counter < 0 Then m_counter = CounterMax
    RedrawTimer_Timer
  End If

End Sub



'---------------------------------------------------------------------------
' Usercontrol events
'---------------------------------------------------------------------------

Private Sub UserControl_Initialize()
On Error Resume Next
Dim iLine As Integer
Dim X As Variant
    
    UserControl.ScaleMode = vbPixels
    bMouseIn = False
    ReEnter = 0
    picBuffer.ScaleMode = vbPixels
    picBuffer.ForeColor = vbWhite
    picBuffer.BackColor = vbBlack
    picBuffer.AutoRedraw = True
    ReDrawTimer.Enabled = True
    On Error Resume Next
    If Not UserControl.Ambient.UserMode Then
      AddMessage "design5", "If you speak Dutch, then please visit my homepage at www.beursmonitor.com", "Arial", RGB(150, 255, 150), RGB(255, 255, 255), 32, 32, 600, -1200, 0, 0, 0, 0, , 0, 1800
    Else
      UserControl!ReDrawTimer.Enabled = True
    End If
    
End Sub


Private Sub UserControl_Show()
    
    PaintBackgroundBuffer picBackBuffer

End Sub



Private Sub UserControl_Resize()
    
    picBackBuffer.Left = 0
    picBackBuffer.Top = 0
    picBackBuffer.Height = UserControl.ScaleHeight
    picBackBuffer.Width = UserControl.ScaleWidth
    
    picBuffer.Left = 0
    picBuffer.Top = 0
    picBuffer.Height = UserControl.ScaleHeight
    picBuffer.Width = UserControl.ScaleWidth
      
    PaintBackgroundBuffer picBackBuffer
    
End Sub



'---------------------------------------------------------------------------
' Executing Methods
'---------------------------------------------------------------------------

Public Sub AddMessage( _
       ByVal MessageID As String, _
       Optional ByVal MessageText As String, _
       Optional ByVal MessageFontName As String, _
       Optional ByVal MessageFontColorStart As OLE_COLOR, _
       Optional ByVal MessageFontColorEnd As OLE_COLOR, _
       Optional ByVal MessageFontSizeStart As Integer, _
       Optional ByVal MessageFontSizeEnd As Integer, _
       Optional ByVal MessageLeftStart As Integer, _
       Optional ByVal MessageLeftEnd As Integer, _
       Optional ByVal MessageTopStart As Integer, _
       Optional ByVal MessageTopEnd As Integer, _
       Optional ByVal MessageFontRotationStart As Integer, _
       Optional ByVal MessageFontRotationEnd As Integer, _
       Optional ByVal BeforeMessageID As Variant, _
       Optional ByVal MessageIntervalStart As Long = 0, _
       Optional ByVal MessageIntervalCount As Long = 0 _
       )

Dim iM As Long
Dim i As Long


   If IsMissing(MessageText) Then MessageText = "Edwin Vermeer"
   If IsMissing(MessageFontName) Then MessageFontName = "Ariel"
   If IsMissing(MessageFontColorStart) Then MessageFontColorStart = vbBlue
   If IsMissing(MessageFontColorEnd) Then MessageFontColorEnd = vbWhite
   If IsMissing(MessageFontSizeStart) Then MessageFontSizeStart = 8
   If IsMissing(MessageFontSizeEnd) Then MessageFontSizeEnd = 16
   If IsMissing(MessageLeftStart) Then MessageLeftStart = picBuffer.ScaleWidth
   If IsMissing(MessageLeftEnd) Then MessageLeftEnd = 0
   If IsMissing(MessageTopStart) Then MessageTopStart = picBuffer.ScaleHeight
   If IsMissing(MessageTopEnd) Then MessageTopEnd = 0
   If IsMissing(MessageFontRotationStart) Then MessageFontRotationStart = 0
   If IsMissing(MessageFontRotationEnd) Then MessageFontRotationEnd = 0
   If IsMissing(MessageIntervalStart) Then MessageIntervalStart = 0
   If IsMissing(MessageIntervalCount) Then MessageIntervalCount = CounterMax
   
   ReDim Preserve m_messages(0 To MessageCount + 1) As TextMessage
   If Not (IsMissing(BeforeMessageID)) Then
      iM = MessageIndex(BeforeMessageID)
      If (iM > -1) Then ' insert
         For i = MessageCount To iM + 1 Step -1
            LSet m_messages(i) = m_messages(i - 1)
         Next i
      End If
    Else
      iM = MessageCount
    End If
    With m_messages(iM)
       .MessageID = MessageID
       .MessageText = MessageText
       .MessageFontName = MessageFontName
       .MessageFontColorStart = MessageFontColorStart
       .MessageFontColorEnd = MessageFontColorEnd
       .MessageFontSizeStart = MessageFontSizeStart
       .MessageFontSizeEnd = MessageFontSizeEnd
       .MessageLeftStart = MessageLeftStart
       .MessageLeftEnd = MessageLeftEnd
       .MessageTopStart = MessageTopStart
       .MessageTopEnd = MessageTopEnd
       .MessageFontRotationStart = MessageFontRotationStart * 10
       .MessageFontRotationEnd = MessageFontRotationEnd * 10
       .MessageIntervalStart = MessageIntervalStart
       .MessageIntervalCount = MessageIntervalCount
    End With

End Sub


Public Sub RemoveMessage(ByVal MessageID As Variant)
Dim iM As Integer
Dim i As Long
   
   iM = MessageIndex(MessageID)
   If (iM > -1) Then
      If MessageCount > 0 Then
         For i = iM To MessageCount - 1
             LSet m_messages(i) = m_messages(i + 1)
         Next i
         ReDim Preserve m_messages(0 To MessageCount - 1) As TextMessage
      End If
   End If
   
End Sub


Public Sub RemoveAllMessages()
  ReDim m_messages(0) As TextMessage
End Sub


Public Sub Draw()
    
  UserControl_Resize
  m_counter = m_counter - 1
  If m_counter < 0 Then m_counter = CounterMax
  RedrawTimer_Timer

End Sub


'---------------------------------------------------------------------------
' Getting and Setting the properties
'---------------------------------------------------------------------------
Private Sub UserControl_InitProperties()

    m_backcolorStart = m_def_backcolorStart
    m_backcolorEnd = m_def_backcolorEnd
    m_Border = m_def_Border
    m_Enabled = m_def_Enabled
    m_counter = m_def_counter
    m_counterMax = m_def_counterMax
    m_Speed = m_def_Speed
    m_BackGroundStyle = m_def_BackGroundStyle
    m_AnimateInDesignmode = m_def_AnimateInDesignmode
    m_BackGroundImage = m_def_BackGroundImage
    m_TransparentColor = m_def_TransparentColor
    m_Angle = m_def_Angle
    m_Repetitions = m_def_Repetitions
    m_GradientType = m_def_GradientType
    m_BlendMode = m_BlendMode

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_backcolorStart = PropBag.ReadProperty("BackColorStart", m_def_backcolorStart)
    m_backcolorEnd = PropBag.ReadProperty("BackColorEnd", m_def_backcolorEnd)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_counter = PropBag.ReadProperty("Counter", m_def_counter)
    m_counterMax = PropBag.ReadProperty("CounterMax", m_def_counterMax)
    m_Speed = PropBag.ReadProperty("Speed", m_def_Speed)
    m_BackGroundStyle = PropBag.ReadProperty("BackGroundStyle", m_def_BackGroundStyle)
    m_AnimateInDesignmode = PropBag.ReadProperty("AnimateInDesignmode", m_def_AnimateInDesignmode)
    m_BackGroundImage = PropBag.ReadProperty("BackGroundImage", m_def_BackGroundImage)
    m_TransparentColor = PropBag.ReadProperty("TransparentColor", m_def_TransparentColor)
    m_Angle = PropBag.ReadProperty("Angle", m_def_Angle)
    m_Repetitions = PropBag.ReadProperty("Repetitions", m_def_Repetitions)
    m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)
    m_BlendMode = PropBag.ReadProperty("BlendMode", m_BlendMode)
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColorStart", m_backcolorStart, m_def_backcolorStart)
    Call PropBag.WriteProperty("BackColorEnd", m_backcolorEnd, m_def_backcolorEnd)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Counter", m_counter, m_def_counter)
    Call PropBag.WriteProperty("CounterMax", m_counterMax, m_def_counterMax)
    Call PropBag.WriteProperty("Speed", m_Speed, m_def_Speed)
    Call PropBag.WriteProperty("BackGroundStyle", m_BackGroundStyle, m_def_BackGroundStyle)
    Call PropBag.WriteProperty("AnimateInDesignmode", m_AnimateInDesignmode, m_def_AnimateInDesignmode)
    Call PropBag.WriteProperty("BackGroundImage", m_BackGroundImage, m_def_BackGroundImage)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
    Call PropBag.WriteProperty("Angle", m_Angle, m_def_Angle)
    Call PropBag.WriteProperty("Repetitions", m_Repetitions, m_def_Repetitions)
    Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)
    Call PropBag.WriteProperty("BlendMode", m_BlendMode, m_BlendMode)

End Sub

' .Hwnd
Public Property Get hwndX() As Long
    hwndX = UserControl.hwnd
End Property


' .Counter
Public Property Get Counter() As Long
Attribute Counter.VB_ProcData.VB_Invoke_Property = "General"
    Counter = m_counter
End Property
Public Property Let Counter(ByVal New_Counter As Long)
    m_counter = New_Counter
    PropertyChanged "Counter"
End Property


' .CounterMax
Public Property Get CounterMax() As Long
Attribute CounterMax.VB_ProcData.VB_Invoke_Property = "General"
    CounterMax = m_counterMax
End Property
Public Property Let CounterMax(ByVal New_CounterMax As Long)
    m_counterMax = New_CounterMax
    PropertyChanged "CounterMax"
End Property


' .BackGroundStyle
Public Property Get BackGroundStyle() As SPBackGroundStyle
    BackGroundStyle = m_BackGroundStyle
End Property
Public Property Let BackGroundStyle(ByVal New_BackGroundStyle As SPBackGroundStyle)
    m_BackGroundStyle = New_BackGroundStyle
    PropertyChanged "BackGroundStyle"
    Draw
End Property


' .BackColorStart
Public Property Get BackColorStart() As OLE_COLOR
    BackColorStart = m_backcolorStart
End Property
Public Property Let BackColorStart(ByVal New_BackColorStart As OLE_COLOR)
    m_backcolorStart = New_BackColorStart
    PropertyChanged "BackColorStart"
    Draw
End Property


' .Angle
Public Property Let Angle(ByVal fData As Double)
'Angles are counter-clockwise and may be
'any Single value from 0 to 359.999999999.

' 135  90 45
'    \ | /
'180 --o-- 0
'    / | \
' 235 270 315

    'Correct angle to ensure between 0 and 359.999999999
    m_Angle = fData Mod 360
    PropertyChanged "Angle"
    Draw
End Property
Public Property Get Angle() As Double
    Angle = m_Angle
End Property


' .Repetitions
Public Property Let Repetitions(ByVal fData As Double)
    m_Repetitions = Abs(fData)
    If m_Repetitions <= 0 Then m_Repetitions = 1
    PropertyChanged "Repetitions"
    Draw
End Property
Public Property Get Repetitions() As Double
    Repetitions = m_Repetitions
End Property


' .GradientType
Public Property Let GradientType(ByVal eData As SPGradType)
    m_GradientType = eData
    PropertyChanged "GradientType"
    Draw
End Property
Public Property Get GradientType() As SPGradType
    GradientType = m_GradientType
End Property


' .BlendMode
Public Property Let BlendMode(ByVal eData As SPGradBlendMode)
    m_BlendMode = eData
    PropertyChanged "BlendMode"
    Draw
End Property
Public Property Get BlendMode() As SPGradBlendMode
    BlendMode = m_BlendMode
End Property


' .BackColorEnd
Public Property Get BackColorEnd() As OLE_COLOR
    BackColorEnd = m_backcolorEnd
End Property
Public Property Let BackColorEnd(ByVal New_BackColorEnd As OLE_COLOR)
    m_backcolorEnd = New_BackColorEnd
    PropertyChanged "BackColorEnd"
    Draw
End Property


' .BackGroundImage
Public Property Get BackGroundImage() As String
    BackGroundImage = m_BackGroundImage
End Property
Public Property Let BackGroundImage(ByVal New_BackGroundImage As String)
    m_BackGroundImage = New_BackGroundImage
    PropertyChanged "BackGroundImage"
    Draw
End Property


' .TransparentColor
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = m_TransparentColor
End Property
Public Property Let TransparentColor(ByVal New_TransparentColor As OLE_COLOR)
    m_TransparentColor = New_TransparentColor
    PropertyChanged "TransparentColor"
    Draw
End Property


' . Border
Public Property Get bOrder() As SPBorderStyle
    bOrder = m_Border
End Property
Public Property Let bOrder(ByVal New_Border As SPBorderStyle)
    m_Border = New_Border
    PropertyChanged "Border"
    UserControl.BorderStyle = m_Border
End Property


' .Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    ReDrawTimer = m_Enabled
End Property


' .AnimateInDesignmode
Public Property Get AnimateInDesignmode() As Boolean
Attribute AnimateInDesignmode.VB_ProcData.VB_Invoke_Property = "General"
    AnimateInDesignmode = m_AnimateInDesignmode
End Property
Public Property Let AnimateInDesignmode(ByVal New_AnimateInDesignmode As Boolean)
    m_AnimateInDesignmode = New_AnimateInDesignmode
    PropertyChanged "AnimateInDesignmode"
    If Not UserControl.Ambient.UserMode Then UserControl!ReDrawTimer.Enabled = m_AnimateInDesignmode Else UserControl!ReDrawTimer.Enabled = True
End Property


' .Speed
Public Property Get Speed() As Long
Attribute Speed.VB_ProcData.VB_Invoke_Property = "General"
    Speed = m_Speed
End Property
Public Property Let Speed(ByVal New_Speed As Long)
    m_Speed = New_Speed
    PropertyChanged "Speed"
    ReDrawTimer.Interval = m_Speed
End Property


' .MessageCount
Public Property Get MessageCount() As Long
On Error Resume Next
    MessageCount = UBound(m_messages)
    If Err.Number <> 0 Then MessageCount = -1
End Property


' .MessageIndex(MessageID)
Public Property Get MessageIndex(ByVal MessageID As Variant) As Integer
Dim iM As Integer
Dim iIndex As Integer
    
    iIndex = -1
    If (IsNumeric(MessageID)) Then
        iIndex = CInt(MessageID)
    Else
        If MessageCount > 0 Then
           For iM = 0 To MessageCount
              If (m_messages(iM).MessageID = MessageID) Then
                  iIndex = iM
                  Exit For
              End If
           Next iM
        Else
           MessageIndex = -1
        End If
    End If
    If (iIndex > -1) And (iIndex <= MessageCount) Then
        MessageIndex = iIndex
    Else
        MessageIndex = -1
    End If
    
End Property


' .MessageID(MessageIndex)
Public Property Get MessageID(ByVal iMessage As Long) As String
   If (iMessage > -1) And (iMessage <= MessageCount) Then
      MessageID = m_messages(iMessage).MessageID
   End If
End Property


' .MessageWidth(MessageID)
Public Property Get MessageWidth(ByVal MessageID As Variant) As Integer
Dim j As Integer
Dim w As Integer

    j = MessageIndex(MessageID)
    If j < 0 Then
      MessageWidth = 0
    Else
      picBuffer.FontName = m_messages(j).MessageFontName
      w = (m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount) - 2
      If w < 1 Then w = 1
      picBuffer.FontSize = w
      MessageWidth = picBuffer.TextWidth(m_messages(j).MessageText)
    End If
    
End Property

' .MessageHeight(MessageID)
Public Property Get MessageHeight(ByVal MessageID As Variant) As Integer
Dim j As Integer
    
    j = MessageIndex(MessageID)
    picBuffer.FontName = m_messages(j).MessageFontName
    picBuffer.FontSize = (m_messages(j).MessageFontSizeStart - (m_messages(j).MessageFontSizeStart - m_messages(j).MessageFontSizeEnd) * (m_counter - m_messages(j).MessageIntervalStart) / m_messages(j).MessageIntervalCount)
    MessageHeight = picBuffer.TextHeight(m_messages(j).MessageText)
    
End Property


' .MessageText (MessageID)
Public Property Get MessageText(ByVal MessageID As Variant) As String
Attribute MessageText.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageText = m_messages(j).MessageText
End Property
Public Property Let MessageText(ByVal MessageID As Variant, ByVal New_MessageText As String)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageText = New_MessageText
End Property


' .MessageFontName (MessageID)
Public Property Get MessageFontName(ByVal MessageID As Variant) As String
Attribute MessageFontName.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontName = m_messages(j).MessageFontName
End Property
Public Property Let MessageFontName(ByVal MessageID As Variant, ByVal New_MessageFontName As String)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontName = New_MessageFontName
End Property


' .MessageFontColorStart (MessageID)
Public Property Get MessageFontColorStart(ByVal MessageID As Variant) As OLE_COLOR
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontColorStart = m_messages(j).MessageFontColorStart
End Property
Public Property Let MessageFontColorStart(ByVal MessageID As Variant, ByVal New_MessageFontColorStart As OLE_COLOR)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontColorStart = New_MessageFontColorStart
End Property


' .MessageFontColorEnd (MessageID)
Public Property Get MessageFontColorEnd(ByVal MessageID As Variant) As OLE_COLOR
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontColorEnd = m_messages(j).MessageFontColorEnd
End Property
Public Property Let MessageFontColorEnd(ByVal MessageID As Variant, ByVal New_MessageFontColorEnd As OLE_COLOR)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontColorEnd = New_MessageFontColorEnd
End Property


' .MessageFontSizeStart (MessageID)
Public Property Get MessageFontSizeStart(ByVal MessageID As Variant) As Integer
Attribute MessageFontSizeStart.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontSizeStart = m_messages(j).MessageFontSizeStart
End Property
Public Property Let MessageFontSizeStart(ByVal MessageID As Variant, ByVal New_MessageFontSizeStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontSizeStart = New_MessageFontSizeStart
End Property


' .MessageFontSizeEnd (MessageID)
Public Property Get MessageFontSizeEnd(ByVal MessageID As Variant) As Integer
Attribute MessageFontSizeEnd.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontSizeEnd = m_messages(j).MessageFontSizeEnd
End Property
Public Property Let MessageFontSizeEnd(ByVal MessageID As Variant, ByVal New_MessageFontSizeEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontSizeEnd = New_MessageFontSizeEnd
End Property


' .MessageLeftStart (MessageID)
Public Property Get MessageLeftStart(ByVal MessageID As Variant) As Integer
Attribute MessageLeftStart.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageLeftStart = m_messages(j).MessageLeftStart
End Property
Public Property Let MessageLeftStart(ByVal MessageID As Variant, ByVal New_MessageLeftStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageLeftStart = New_MessageLeftStart
End Property


' .MessageLeftEnd (MessageID)
Public Property Get MessageLeftEnd(ByVal MessageID As Variant) As Integer
Attribute MessageLeftEnd.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageLeftEnd = m_messages(j).MessageLeftEnd
End Property
Public Property Let MessageLeftEnd(ByVal MessageID As Variant, ByVal New_MessageLeftEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageLeftEnd = New_MessageLeftEnd
End Property


' .MessageTopStart (MessageID)
Public Property Get MessageTopStart(ByVal MessageID As Variant) As Integer
Attribute MessageTopStart.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageTopStart = m_messages(j).MessageTopStart
End Property
Public Property Let MessageTopStart(ByVal MessageID As Variant, ByVal New_MessageTopStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageTopStart = New_MessageTopStart
End Property


' .MessageTopEnd (MessageID)
Public Property Get MessageTopEnd(ByVal MessageID As Variant) As Integer
Attribute MessageTopEnd.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageTopEnd = m_messages(j).MessageTopEnd
End Property
Public Property Let MessageTopEnd(ByVal MessageID As Variant, ByVal New_MessageTopEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageTopEnd = New_MessageTopEnd
End Property


' .MessageFontRotationStart (MessageID)
Public Property Get MessageFontRotationStart(ByVal MessageID As Variant) As Integer
Attribute MessageFontRotationStart.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontRotationStart = m_messages(j).MessageFontRotationStart
End Property
Public Property Let MessageFontRotationStart(ByVal MessageID As Variant, ByVal New_MessageFontRotationStart As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontRotationStart = New_MessageFontRotationStart
End Property


' .MessageFontRotationEnd (MessageID)
Public Property Get MessageFontRotationEnd(ByVal MessageID As Variant) As Integer
Attribute MessageFontRotationEnd.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageFontRotationEnd = m_messages(j).MessageFontRotationEnd
End Property
Public Property Let MessageFontRotationEnd(ByVal MessageID As Variant, ByVal New_MessageFontRotationEnd As Integer)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageFontRotationEnd = New_MessageFontRotationEnd
End Property


' .MessageIntervalStart (MessageID)
Public Property Get MessageIntervalStart(ByVal MessageID As Variant) As Long
Attribute MessageIntervalStart.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageIntervalStart = m_messages(j).MessageIntervalStart
End Property
Public Property Let MessageIntervalStart(ByVal MessageID As Variant, ByVal New_MessageIntervalStart As Long)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageIntervalStart = New_MessageIntervalStart
End Property


' .MessageIntervalCount (MessageID)
Public Property Get MessageIntervalCount(ByVal MessageID As Variant) As Long
Attribute MessageIntervalCount.VB_ProcData.VB_Invoke_Property = "Messages"
Dim j As Integer
    j = MessageIndex(MessageID)
    MessageIntervalCount = m_messages(j).MessageIntervalCount
End Property
Public Property Let MessageIntervalCount(ByVal MessageID As Variant, ByVal New_MessageIntervalCount As Long)
Dim j As Integer
    j = MessageIndex(MessageID)
    m_messages(j).MessageIntervalCount = New_MessageIntervalCount
End Property


