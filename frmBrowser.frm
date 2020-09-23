VERSION 5.00
Begin VB.Form frmBrowser 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Hmm. the titelbar overlay should not get the focus!"
   ClientHeight    =   375
   ClientLeft      =   3975
   ClientTop       =   2835
   ClientWidth     =   4215
   ControlBox      =   0   'False
   FillColor       =   &H80000002&
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000002&
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   Begin TextAnimationDemo.GradientButton btnHelp 
      Height          =   315
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      AutoSize        =   1
      BackPicture     =   "frmBrowser.frx":044A
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownPicture     =   "frmBrowser.frx":09DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverMode       =   2
      HoverPicture    =   "frmBrowser.frx":0F6E
      Picture         =   "frmBrowser.frx":1500
      PictureCushion  =   0
      Style           =   1
   End
   Begin TextAnimationDemo.GradientButton btnRestore 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      AutoSize        =   1
      BackPicture     =   "frmBrowser.frx":1A92
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownPicture     =   "frmBrowser.frx":2024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverMode       =   2
      HoverPicture    =   "frmBrowser.frx":25B6
      Picture         =   "frmBrowser.frx":2B48
      PictureCushion  =   0
      Style           =   1
   End
   Begin TextAnimationDemo.GradientButton btnMaximize 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      AutoSize        =   1
      BackPicture     =   "frmBrowser.frx":30DA
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownPicture     =   "frmBrowser.frx":366C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverMode       =   2
      HoverPicture    =   "frmBrowser.frx":3BFE
      Picture         =   "frmBrowser.frx":4190
      PictureCushion  =   0
      Style           =   1
   End
   Begin TextAnimationDemo.GradientButton btnMinimize 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      AutoSize        =   1
      BackPicture     =   "frmBrowser.frx":4722
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownPicture     =   "frmBrowser.frx":4CB4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverMode       =   2
      HoverPicture    =   "frmBrowser.frx":5246
      Picture         =   "frmBrowser.frx":57D8
      PictureCushion  =   0
      Style           =   1
   End
   Begin TextAnimationDemo.GradientButton btnClose 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      AutoSize        =   1
      BackPicture     =   "frmBrowser.frx":5D6A
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownPicture     =   "frmBrowser.frx":62FC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverMode       =   2
      HoverPicture    =   "frmBrowser.frx":688E
      Picture         =   "frmBrowser.frx":6E20
      PictureCushion  =   0
      Style           =   1
   End
   Begin VB.Timer timTimer 
      Interval        =   20
      Left            =   1080
      Top             =   0
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      BackColorStart  =   4194304
      BackColorEnd    =   16744576
      Counter         =   600
      AnimateInDesignmode=   0   'False
      TransparentColor=   14737632
      Angle           =   85
      Repetitions     =   2
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Private Declare Function GetClassInfoEx Lib "user32" Alias "GetClassInfoExA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASSEX) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As Rect) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Boolean
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Boolean
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Boolean

Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_CHILD = &H40000000
Const WS_CHILDWINDOW = &H40000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_DLGFRAME = &H400000
Const WS_GROUP = &H20000
Const WS_HSCROLL = &H100000
Const WS_ICONIC = &H20000000
Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZE = &H20000000
Const WS_MINIMIZEBOX = &H20000
Const WS_OVERLAPPED = &H0
Const WS_OVERLAPPEDWINDOW = &HCF0000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = &H80880000
Const WS_SIZEBOX = &H40000
Const WS_SYSMENU = &H80000
Const WS_TABSTOP = &H10000
Const WS_THICKFRAME = &H40000
Const WS_TILED = &H0
Const WS_TILEDWINDOW = &HCF0000
Const WS_VISIBLE = &H10000000
Const WS_VSCROLL = &H200000
Const SM_CYCAPTION = 4
Const SM_CYSMCAPTION = 51
Const SM_CYEDGE = 46
Const SM_CXEDGE = 45
Const SM_CYBORDER = 6
Const SM_CXBORDER = 5
Const SM_CYSMSIZE = 53
Const SM_CXSMSIZE = 52
Const SM_CYSIZEFRAME = 33
Const SM_CXSIZEFRAME = 32
Const SM_CYSIZE = 31
Const SM_CXSIZE = 30
Const WS_EX_WINDOWEDGE = &H100  '       0x00000100L
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_STATICEDGE = &H20000
Const WS_EX_CLIENTEDGE = &H200
Const WS_EX_CONTEXTHELP = &H400
Const WS_EX_TOPMOST = &H8
Const DI_MASK = 1
Const DI_IMAGE = 2
Const DI_NORMAL = 3
Const DI_COMPAT = 4
Const DI_DEFAULTSIZE = 8
Const HWND_BOTTOM = 1
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_NOOWNERZORDER = &H200
Const GWL_HINSTANCE = -6 ' For GetWindowLong(..)
Const GCL_HICON = -14 ' For GetClassLong(..)
Const SW_SHOW = 5
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_SHOWMAXIMIZED = 3
Const SW_MINIMIZE = 6
Const HTCAPTION = 2
Const WM_CLOSE = &H10
Const WM_SETTEXT = &HC
Const WM_GETICON = &H7F
Const WM_NCLBUTTONDOWN = &HA1
Const WM_NCMBUTTONDBLCLK = &HA9
Const WM_LBUTTONDOWN = &H201
Const WM_HELP = 53

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type WNDCLASSEX ' Same as WNDCLASS but has a few advanced values
    cbSize As Long
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long ' Handle to large icon (Alt-Tab icon)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long ' Handle to Small icon (Top Left Icon/Taskbar Icon)
End Type

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long ' Handle to icon (only 1 size)
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
    

Dim lastHwnd As Long        ' This program will be on top of the program with this handle
Dim dragApp As Boolean      ' Bussy dragging the application
Dim dragDx As Long
Dim dragDy As Long
Dim ParrentHwnd As Long
Dim ParrentRect As Rect
Dim myrect As Rect
Dim bCaptionSize As Long
Dim xStyle As Long          ' Style
Dim ExStyle As Long         ' Style EX
Dim ReEnter As Boolean
Dim bOrder As Integer

Private Sub btnMinimize_Click()
On Error Resume Next
  Me.Visible = False
  ShowWindow lastHwnd, SW_MINIMIZE
End Sub
Private Sub btnMaximize_Click()
Dim X As Long
Dim Y As Long
On Error Resume Next
  Y = lastHwnd
  If IsZoomed(Y) Then
    ShowWindow Y, SW_SHOWNORMAL
  Else
    ShowWindow Y, SW_SHOWMAXIMIZED
  End If
  X = GetWindowRect(Y, ParrentRect)
  SetWindowPos Y, HWND_TOP, ParrentRect.Left, ParrentRect.Top, ParrentRect.Right - ParrentRect.Left, ParrentRect.Bottom - ParrentRect.Top, SWP_SHOWWINDOW
  timTimer_Timer
End Sub
Private Sub btnRestore_Click()
Dim X As Long
Dim Y As Long
On Error Resume Next
  Y = lastHwnd
  If IsZoomed(Y) Then
    ShowWindow Y, SW_SHOWNORMAL
  Else
    ShowWindow Y, SW_SHOWMAXIMIZED
  End If
  X = GetWindowRect(Y, ParrentRect)
  SetWindowPos Y, HWND_TOP, ParrentRect.Left, ParrentRect.Top, ParrentRect.Right - ParrentRect.Left, ParrentRect.Bottom - ParrentRect.Top, SWP_SHOWWINDOW
  timTimer_Timer
End Sub
Private Sub btnClose_Click()
On Error Resume Next
  Me.Visible = False
  PostMessage lastHwnd, WM_CLOSE, 0, 0
  lastHwnd = 0
  timTimer_Timer
End Sub
Private Sub btnHelp_Click()
On Error Resume Next
  ' Hmm.. This one does not work. does anybody know how to solve this?
  MsgBox "Hmm... This does not work:" & vbCrLf & "SendMessage xxHwnd, WM_HELP, 0, 0" & vbCrLf & "Does anybody know the solution for this?"
  SendMessage lastHwnd, WM_HELP, 0, 0
End Sub


Private Sub Form_Load()
On Error Resume Next
  
  ReEnter = False
  Me!TextAnimation1.Width = Me.ScaleWidth
  
  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "ticker", App.Title & " " & App.Major & "." & App.Minor, "Arial", RGB(63, 200, 63), vbWhite, 14, 14, Me.ScaleWidth, -Me.ScaleWidth, -1, -1, 0, 0, , 0, Me.ScaleWidth * 2
  TextAnimation1.MessageLeftEnd("ticker") = -TextAnimation1.MessageWidth("ticker")
  TextAnimation1.MessageIntervalCount("ticker") = Me.ScaleWidth + TextAnimation1.MessageWidth("ticker")
  TextAnimation1.Counter = 0
  TextAnimation1.CounterMax = TextAnimation1.MessageIntervalCount("ticker")
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.bOrder = None
  
  'the form must be fully visible before calling Shell_NotifyIcon
  Me.Show
  lastHwnd = GetForegroundWindow
  dragApp = False
  
  BuildTickerTape

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.timTimer.Interval = 30000
  DoEvents
  Unload Me
End Sub


Public Sub Form_Resize()
On Error Resume Next
Dim pc As Integer

  ' Reset the tickertape
  For pc = 1 To TextAnimation1.MessageCount
    TextAnimation1.MessageLeftStart(pc) = Me.ScaleWidth
    TextAnimation1.MessageIntervalCount(pc) = Me.ScaleWidth + TextAnimation1.MessageWidth(pc)
    TextAnimation1.CounterMax = TextAnimation1.MessageIntervalStart(pc) + TextAnimation1.MessageIntervalCount(pc)
  Next pc
  'this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then Me.Hide
    
End Sub


Private Sub TextAnimation1_BeforeDraw(PictureBuffer As PictureBox)
Dim windowname As String * 100
Dim TheIcon As Long
On Error Resume Next

  PictureBuffer.ForeColor = vbWhite
  PictureBuffer.CurrentX = 24
  PictureBuffer.CurrentY = 1
  PictureBuffer.FontSize = bCaptionSize / 2
  GetWindowText lastHwnd, windowname, 100
  PictureBuffer.Print Left(windowname, InStr(1, windowname, Chr(0)) - 1)
  TheIcon = GetIconHandle(lastHwnd)
  DrawIconEx PictureBuffer.hdc, 2, 2, TheIcon, Me.ScaleHeight - 2, Me.ScaleHeight - 2, ByVal 0&, ByVal 0&, DI_NORMAL

End Sub


Private Sub timTimer_Timer()
On Error Resume Next
  ' Change the screen position of the application
  If ReEnter Then Exit Sub
  ReEnter = True
  Reposition
  ReEnter = False
End Sub


Private Sub Reposition()
On Error Resume Next
Dim X As Variant

    ' What is the current active window
    ParrentHwnd = GetForegroundWindow
    
    ' Make sure you have the top parent of the active window
'    While GetParent(ParrentHwnd) <> 0
'      X = GetParent(ParrentHwnd)
'      ParrentHwnd = X
'    Wend
   
    ' The overlay can't be the foreground window
    If ParrentHwnd = Me.hwnd Then
      SetForegroundWindow lastHwnd
      Exit Sub
    End If
    
    ' Only if the form is visible
    X = GetWindowLong(ParrentHwnd, GWL_STYLE)
    If (X And WS_VISIBLE) <> WS_VISIBLE Then
      Me.Visible = False
      Exit Sub
    End If
    
    ' Do we realy have a window
    If fGetClassName(ParrentHwnd) & "" = "" Then
      Me.Visible = False
      Exit Sub
    End If
    
    ' Now reposition this form over the window
    lastHwnd = ParrentHwnd
  
  GetAndSetPos lastHwnd, Me

End Sub


Public Function fGetClassName(hwnd As Long)
On Error Resume Next
' Retrieve the class name of a window
Dim strBuffer As String
Dim lngRet As Long
  strBuffer = String$(32, 0)
  lngRet = apiGetClassName(hwnd, strBuffer, Len(strBuffer))
  If lngRet > 0 Then fGetClassName = Left$(strBuffer, lngRet)
End Function


' Nuts ! ... The Post (or Send) Message does not work for windows in an other thread.
'Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
'On Error Resume Next
'  ReleaseCapture
'  SetForegroundWindow lastHwnd
'  PostMessage lastHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End Sub

Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  If Button = vbLeftButton Then
    dragApp = True
    dragDx = X
    dragDy = Y
  End If
End Sub

Private Sub TextAnimation1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  If dragApp Then
    ' Phew... This one took me about a day to get it working. But the result was worth the efford.
    ' The order of the commands below are verry strict. Even a small modification can scr.. things up.
    GetWindowRect lastHwnd, ParrentRect
    SetWindowPos lastHwnd, HWND_TOP, ParrentRect.Left + X - dragDx, ParrentRect.Top + Y - dragDy, ParrentRect.Right - ParrentRect.Left, ParrentRect.Bottom - ParrentRect.Top, SWP_NOACTIVATE
    SetForegroundWindow lastHwnd
    GetAndSetPos lastHwnd, Me
    DoEvents
    SetCapture TextAnimation1.hwndX
  End If
End Sub

Private Sub TextAnimation1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  dragApp = False
  SetForegroundWindow lastHwnd
  ReleaseCapture
End Sub


Private Sub Textanimation1_AnimationRestart()

  BuildTickerTape

End Sub


Public Sub BuildTickerTape()
On Error Resume Next
Dim d As String
Dim x1 As Long
Dim x2 As Long
Dim m As String
Dim procent As Double

  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "ticker", App.Title & " " & App.Major & "." & App.Minor & "  (c) 2000, Vermeer Automatisering", "Arial", RGB(63, 200, 63), vbWhite, 12, 12, Me.ScaleWidth + 32, -Me.ScaleWidth + 32, -1, -1, 0, 0, , 0, Me.ScaleWidth * 2 + 64
  TextAnimation1.MessageLeftEnd("ticker") = -TextAnimation1.MessageWidth("ticker")
  TextAnimation1.MessageIntervalCount("ticker") = Me.ScaleWidth + 32 + TextAnimation1.MessageWidth("ticker")
  TextAnimation1.CounterMax = TextAnimation1.MessageIntervalCount("ticker")
  
  TextAnimation1.Counter = 0
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.bOrder = None

End Sub


Sub GetAndSetPos(hwnd As Long, vbForm As Form)
' This routine positions the form and resizes it to correct size
On Error Resume Next

    Dim xStyle As Long
    Dim ExStyle As Long
    Dim BHeight As Long
    Dim ButWidth As Long
    Dim ButHeight As Long
    Dim BLeft As Long
    Dim BTop As Long
    Dim hWndRECT As Rect
    Dim NewRECT As Rect
    
    xStyle = GetWindowLong(hwnd, GWL_STYLE)
    ExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
 
    'Get The Boarder Size
    BHeight = GetBoarderSize(xStyle, ExStyle)
    If BHeight = 0 Then vbForm.Visible = False: Exit Sub
    
    'Get Caption Size + Button Width/Height
    bCaptionSize = GetCaptionSize(xStyle, ExStyle, ButWidth, ButHeight)
    If bCaptionSize = 0 Then vbForm.Visible = False:  Exit Sub

    'Find the X location of the button
    BLeft = GetLeftPos(hwnd, ButWidth)
    If BLeft = 0 Then vbForm.Visible = False: Exit Sub

    'Find the Y location of the button
    GetWindowRect hwnd, hWndRECT
    BTop = hWndRECT.Top + (BHeight + ((bCaptionSize - ButHeight) / 2))
    
    Dim fs As Integer
    fs = 0  ' Value of free space around the button (normally should be 2)
    'Which buttons should be visible and what is the position
    With NewRECT
        .Left = hWndRECT.Left + BHeight
        .Top = BTop - BHeight + 2
        .Right = hWndRECT.Right - BHeight 'BLeft + ButWidth - 1
        .Bottom = .Top + ButHeight + BHeight - 1
        If (xStyle And WS_SYSMENU) = WS_SYSMENU Then
          ' Position the close button
          Me!btnClose.Left = .Right - .Left - (0 + ButWidth)
          Me!btnClose.Top = fs
          Me!btnClose.Width = ButWidth - fs
          Me!btnClose.Height = ButWidth - fs
          Me!btnClose.Visible = True
          bOrder = 2
          ' Position the maximize button
          If (IsZoomed(lastHwnd) = False) And (xStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Then
            Me!btnMaximize.Left = .Right - .Left - (0 + ButWidth) * bOrder
            Me!btnMaximize.Top = fs
            Me!btnMaximize.Width = ButWidth - fs
            Me!btnMaximize.Height = ButWidth - fs
            Me!btnMaximize.Visible = True
            bOrder = bOrder + 1
          Else
            Me!btnMaximize.Visible = False
          End If
          ' Position the restore button
          If (IsZoomed(lastHwnd) And (xStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX) Or (IsIconic(lastHwnd) And (xStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX) Then
            Me!btnRestore.Left = .Right - .Left - (0 + ButWidth) * bOrder
            Me!btnRestore.Top = fs
            Me!btnRestore.Width = ButWidth - fs
            Me!btnRestore.Height = ButWidth - fs
            Me!btnRestore.Visible = True
            bOrder = bOrder + 1
          Else
            Me!btnRestore.Visible = False
          End If
          ' Position of the minimize button
          If (IsIconic(lastHwnd) = False) And (xStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
            Me!btnMinimize.Left = .Right - .Left - (0 + ButWidth) * bOrder
            Me!btnMinimize.Top = fs
            Me!btnMinimize.Width = ButWidth - fs
            Me!btnMinimize.Height = ButWidth - fs
            Me!btnMinimize.Visible = True
            bOrder = bOrder + 1
          Else
            Me!btnMinimize.Visible = False
          End If
        Else
          ' None of the buttons are visible
          Me!btnMaximize.Visible = False
          Me!btnMinimize.Visible = False
          Me!btnClose.Visible = False
          Me!btnRestore.Visible = False
        End If
        ' The help button
        If (ExStyle And WS_EX_CONTEXTHELP) = WS_EX_CONTEXTHELP Then
          Me!btnHelp.Left = .Right - .Left - (0 + ButWidth) * bOrder
          Me!btnHelp.Top = 2
          Me!btnHelp.Width = ButWidth - 2
          Me!btnHelp.Height = ButWidth - 2
          Me!btnHelp.Visible = True
          bOrder = bOrder + 1
        Else
          Me!btnHelp.Visible = False
        End If

        Me!TextAnimation1.Width = .Right - .Left - (0 + ButWidth) * (bOrder - 1)

        ' now set the form at the right position
        vbForm.Visible = True
        SetWindowPos vbForm.hwnd, HWND_TOPMOST, .Left, .Top, .Right - .Left, .Bottom - .Top, SWP_NOACTIVATE + SWP_NOOWNERZORDER
    End With

   
End Sub



Function GetBoarderSize(xStyle As Long, ExStyle As Long) As Long
    If (xStyle And WS_THICKFRAME) = WS_THICKFRAME And (ExStyle And WS_EX_TOOLWINDOW) <> WS_EX_TOOLWINDOW Then
            ' Re-Sizeable Window
            GetBoarderSize = GetSystemMetrics(SM_CYSIZEFRAME)
    
    ElseIf (ExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then
            ' Normal Window
            GetBoarderSize = GetSystemMetrics(SM_CYEDGE) + 1
    
    ElseIf (xStyle And WS_BORDER) = WS_BORDER Then
            ' Single Boarder, Will fail next routine in 99% of cases
            GetBoarderSize = GetSystemMetrics(SM_CYBORDER)
    
    Else
            ' No Boarder, Exit Function
            GetBoarderSize = 0
            Exit Function
    End If


End Function



Function GetCaptionSize(xStyle As Long, ExStyle As Long, ByRef ButWidth As Long, ByRef ButHeight As Long) As Long
    ' Valid Options:
    '  Small Caption    (Tool Windows Etc)  WS_EX_TOOLWINDOW
    '  Large Caption    (Normal Windows)    WS_CAPTION or WS_OVERLAPPEDWINDOW
    '  No Caption
    
    If (ExStyle And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW Then
            ' Tool Bar Window
            ' Get Height of Caption
            GetCaptionSize = GetSystemMetrics(SM_CYSMCAPTION)
            ButHeight = GetSystemMetrics(SM_CYSMSIZE) - 3
            ButWidth = GetSystemMetrics(SM_CXSMSIZE) - 1
            
    ElseIf (xStyle And WS_CAPTION) = WS_CAPTION Or (xStyle And WS_OVERLAPPEDWINDOW) = WS_OVERLAPPEDWINDOW Or (xStyle And WS_TILEDWINDOW) = WS_TILEDWINDOW Then
            'Normal Caption
            GetCaptionSize = GetSystemMetrics(SM_CYCAPTION)
            ButHeight = GetSystemMetrics(SM_CYSIZE) - 3
            ButWidth = GetSystemMetrics(SM_CXSIZE) - 1
            
    Else
            'No Caption, Abort
            GetCaptionSize = 0
            Exit Function
    End If

End Function


Function GetLeftPos(hwnd As Long, ButWidth As Long)
    ' This gets the Windows Long Style and checks for boxes already visible.
    Dim xRECT As Rect           ' Windows X,Y
    Dim BoarderSize As Long     ' Right boarder
    Dim X As Long               ' Temp X for ret value
    
    xStyle = GetWindowLong(hwnd, GWL_STYLE)
    ExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    GetWindowRect hwnd, xRECT
    
    ' Cool.. now first, work out the Right most side.
    
    If (xStyle And WS_THICKFRAME) = WS_THICKFRAME Then
            ' Re-Sizeable Window
            BoarderSize = GetSystemMetrics(SM_CXSIZEFRAME)
    
    ElseIf (ExStyle And WS_EX_WINDOWEDGE) = WS_EX_WINDOWEDGE Then
            ' Normal Window
            BoarderSize = GetSystemMetrics(SM_CXEDGE)
    
    ElseIf (xStyle And WS_BORDER) = WS_BORDER Then
            ' Single Boarder, Will fail next routine in 99% of cases
            BoarderSize = GetSystemMetrics(SM_CXBORDER)
    
    Else
            ' No Boarder, Exit Function
            GetLeftPos = 0
            Exit Function
    End If
    
    
    ' OK, so now we have the boarder size.
    X = BoarderSize - 2     ' 2 Pixels left is the first one.
    
    X = xRECT.Right - X     ' Now we should have X = right side of First button
    
    If (xStyle And WS_SYSMENU) = WS_SYSMENU Then
            ' X is there
            X = X - ButWidth - 2        ' X has 2 pixels on each side
    Else
            ' NO SYS MENU!!! Return ZERO
            ' If a form does not have a system menu, they do not want a min to tray button!
            ' IE GAMES, Taskbars.. They have borders but no buttons.
            GetLeftPos = 0
            Exit Function
    End If
    
    If (xStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX Or (xStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX Then
            ' Either MAX/RESIZE or MIN button is there and enabled.
            ' (or both.. but can't have 1 without the other.. 1 is just enabled)
            X = X - (ButWidth * 2)
    ElseIf (ExStyle And WS_EX_CONTEXTHELP) = WS_EX_CONTEXTHELP Then
            ' CANNOT HAVE MAX/MIN AND ? AT SAME TIME :)
            ' Same as Max/Min box but only one of them
            X = X - ButWidth
    End If
    
    ' Cool, that is all of them. Now take away 2 pixels for the gap
    X = X - 4
    
    ' Then take away another Width for our button
    X = X - ButWidth
    
    GetLeftPos = X  ' simple as that
            
End Function


Function IsOnTop(hwnd As Long) As Boolean
    Dim X As Long
    X = GetWindowLong(hwnd, GWL_EXSTYLE)
    If (X And WS_EX_TOPMOST) = WS_EX_TOPMOST Then IsOnTop = True Else IsOnTop = False
End Function



Sub hWndontop(hwnd As Long, OnTop As Boolean)
' Sets Z of window to foreground (topmost)
    On Error Resume Next
   Dim Flags As Long
   Const SWP_NOMOVE = &H2
   Const SWP_NOSIZE = &H1
   Flags = SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos hwnd, IIf(OnTop = True, -1, -2), 0, 0, 0, 0, Flags
End Sub


Public Function GetIconHandle(hwnd As Long) As Long
    ' OK, This function is confusing
    ' Many windows have different ways of ha
    '     ndling Icons.
    '---------------------------
    '1. All VB apps use SendMessage(..WM_GET
    '     ICON..) to get the Icon (not only VB app
    '     s)
    '2. Other Programs like GetClassInfoEx(.
    '     .)(most non-SendMessage Apps)
    '3. Others Like GetClassInfo(..)(Very Ra
    '     re)
    '4. And the rest like GetClassLong(GCL_H
    '     ICON) (The rest.)
    '----------------------------
    ' Any program that doesn't work with the
    '     se 4 methods have issues.
    '
    ' All apps I have tried work fine with t
    '     hese 4 methods.. one or the other.
    '
    
    '*************************************
    'Method: SendMessage (Small Icon)
    '*************************************
    Dim hIcon As Long
    ' First, Try for the small icon. This wo
    '     uld be nice.
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(0), CLng(0))
    
    If hIcon > 0 Then GetIconHandle = hIcon: Exit Function ' found it
    ' Nope, keep trying
    
    
    '*************************************
    'Method: SendMessage (Large Icon)
    '*************************************
    ' Hmm.. No small Icon, Try LARGE icon.
    hIcon = SendMessage(hwnd, WM_GETICON, CLng(1), CLng(0))
    
    If hIcon > 0 Then GetIconHandle = hIcon: Exit Function ' found it
    ' Nope, keep trying
    
    
    '*************************************
    'Method: GetClassInfoEx (Small or Large with Small Pref.)
    '*************************************
        
        Dim ClassName As String
        Dim WCX As WNDCLASSEX
        Dim hInstance As Long
        
        ' First, get the Instance of the Class v
        '     ia GetWindowLong
        hInstance = GetWindowLong(hwnd, GWL_HINSTANCE)
        
        ' Now set the Size Value of WndClassEx
        WCX.cbSize = Len(WCX)
        
        ' Set The ClassName variable to 255 spac
        '     es (max len of the class name)
        ClassName = Space(255)
        
        Dim X As Long ' temp variable
        ' Get the Classname of hWnd and put into
        '     ClassName (max 255 chars)
        X = GetClassName(hwnd, ClassName, 255)
        
        ' Now Trim the Classname and add a NullC
        '     har to the end (reqd. for GetClassInfoEx
        '     )
        ClassName = Left$(ClassName, X) & vbNullChar
        
        ' Now, if GetClassInfoEx(..) Returns 0,
        '     their was an error. >0 = No probs
        X = GetClassInfoEx(hInstance, ClassName, WCX)


        If X > 0 Then
            ' Returned True
            ' So we should now have both WCX.hIcon a
            '     nd WCX.hIconSm


            If WCX.hIconSm = 0 Then 'No small icon
                hIcon = WCX.hIcon ' No small icon.. Windows should have given default.. weird
            Else
                hIcon = WCX.hIconSm ' Small Icon is better
            End If
            GetIconHandle = hIcon ' found it =]
            Exit Function
        End If
        
        
        '*************************************
        'Method: GetClassInfo (Large Icon)
        '*************************************
        ' Hmm.. ClassInfoEX failed, Try ClassInf
        '     o
        Dim WC As WNDCLASS
        X = GetClassInfo(hInstance, ClassName, WC)


        If X > 0 Then
            ' Woohoo.. dunno why but it liked that
            hIcon = WC.hIcon
            GetIconHandle = hIcon: Exit Function ' Found it
        End If
        '*************************************
        'Method: GetClassLong (Large Icon)
        '*************************************
        ' Hmm.. One more try
        X = GetClassLong(hwnd, GCL_HICON)


        If X > 0 Then
            ' Yay, about time.. annoying windows.. E
            '     xample: NOTEPAD
            hIcon = X
        Else
            ' This is most prob a Icon-less window.
            hIcon = 0
        End If
        If hIcon < 0 Then hIcon = 0 ' Handles must be > 0
        GetIconHandle = hIcon
    End Function

       

