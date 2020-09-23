VERSION 5.00
Begin VB.Form frmSystemTrayTicker 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   6480
   ClientTop       =   7050
   ClientWidth     =   2565
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
   Icon            =   "frmSystemTrayTicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   ShowInTaskbar   =   0   'False
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColorStart  =   14737632
      BackColorEnd    =   4210752
      Counter         =   508
      AnimateInDesignmode=   0   'False
      TransparentColor=   14737632
      Angle           =   170
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1920
      Top             =   0
   End
End
Attribute VB_Name = "frmSystemTrayTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As Rect) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Dim VisibleAfter As Date


Private Sub Form_Load()
On Error Resume Next
  
  VisibleAfter = Now
  
  SetWindowOverTray Me
  
  BuildTickerTape
  
End Sub


Public Sub Form_Resize()
  Form_Paint
End Sub


Public Sub Form_Paint()
On Error Resume Next
Dim pc As Integer

  '  hu ?
  Me!TextAnimation1.Width = Me.ScaleWidth * 2 '/ Screen.TwipsPerPixelX
  Me!TextAnimation1.Height = Me.ScaleHeight / Screen.TwipsPerPixelY
  Me!TextAnimation1.Left = 0
  Me!TextAnimation1.Top = 0
  
  For pc = 1 To TextAnimation1.MessageCount
    TextAnimation1.MessageLeftStart(pc) = Me.ScaleWidth
    TextAnimation1.MessageIntervalCount(pc) = Me.ScaleWidth + TextAnimation1.MessageWidth(pc)
    TextAnimation1.CounterMax = TextAnimation1.MessageIntervalStart(pc) + TextAnimation1.MessageIntervalCount(pc)
  Next pc

End Sub


Private Sub TextAnimation1_BeforeDraw(PictureBuffer As PictureBox)
Dim TickerBackground As String
On Error Resume Next

  PictureBuffer.CurrentX = 10
  PictureBuffer.CurrentY = -1
  PictureBuffer.FontName = "Arial"
  PictureBuffer.ForeColor = RGB(31, 61, 127)
  PictureBuffer.FontSize = 9
  PictureBuffer.Print TickerBackground

  PictureBuffer.ForeColor = vbWhite
  PictureBuffer.CurrentY = 4
  PictureBuffer.Print "  " & Format(Now, "dd mmm yyyy hh:nn")
  
End Sub





Public Sub BuildTickerTape()
'On Error Resume Next

  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "ticker", App.Title & " " & App.Major & "." & App.Minor & "  (c) " & Year(Now) & ", Vermeer Automatisering", "Arial", RGB(63, 200, 63), vbWhite, 14, 14, Me.ScaleWidth + 32, -Me.ScaleWidth + 32, -1, -1, 0, 0, , 0, Me.ScaleWidth * 2 + 64
  TextAnimation1.MessageLeftEnd("ticker") = -TextAnimation1.MessageWidth("ticker")
  TextAnimation1.MessageIntervalCount("ticker") = Me.ScaleWidth + 32 + TextAnimation1.MessageWidth("ticker")
  TextAnimation1.CounterMax = TextAnimation1.MessageIntervalCount("ticker")
  TextAnimation1.Counter = 0
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.Border = None

End Sub

Private Sub TextAnimation1_MouseMovesimple(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
  VisibleAfter = DateAdd("s", 4, Now)
  Me.Visible = False

End Sub

Private Sub Textanimation1_AnimationRestart()

  BuildTickerTape

End Sub

Private Sub Timer1_Timer()
  
  If Now > VisibleAfter Then Me.Visible = True

End Sub



Private Sub SetWindowOverTray(TrayForm As Form)
Dim fWidth As Integer
Dim fHeight As Integer

  Dim TrayWindowRect As Rect
  GetWindowRect GetTrayhWnd, TrayWindowRect
  fWidth = TrayWindowRect.Right - TrayWindowRect.Left
  fHeight = TrayWindowRect.Bottom - TrayWindowRect.Top
  With TrayWindowRect
    .Top = 0
    .Left = 0
    .Right = fWidth
  End With
  MoveWindow TrayForm.hwnd, TrayWindowRect.Left, TrayWindowRect.Top, TrayWindowRect.Right, TrayWindowRect.Bottom, 1
  SetWindowPos TrayForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  SetParent TrayForm.hwnd, GetTrayhWnd

End Sub

Private Function GetTrayhWnd() As Long
Dim OurParent As Long
Dim OurHandle As Long

  OurParent = FindWindow("Shell_TrayWnd", "")
  OurHandle = FindWindowEx(OurParent&, 0, "TrayNotifyWnd", vbNullString)
  GetTrayhWnd = OurHandle

End Function



