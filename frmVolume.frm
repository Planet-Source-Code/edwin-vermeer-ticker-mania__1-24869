VERSION 5.00
Begin VB.Form frmVolume 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   2085
   ClientTop       =   2895
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ShowInTaskbar   =   0   'False
   Begin TextAnimationDemo.FormTransparancy FormTransparancy1 
      Left            =   480
      Top             =   0
      _ExtentX        =   1085
      _ExtentY        =   1085
      TransparencyLevel=   150
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   0
      Top             =   0
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4048
      Counter         =   225
      BackGroundStyle =   4
      AnimateInDesignmode=   0   'False
      TransparentColor=   130817
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As Rect) As Long

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOOWNERZORDER = &H200


Private Sub Form_Load()
On Error Resume Next
Dim lprect As Rect

  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "Volume", "Volume", "Arial", RGB(0, 255, 0), RGB(100, 200, 100), 72, 72, 0, 0, 0, 0, 0, 0, , 0, 100000000
  TextAnimation1.AddMessage "Level", "|||||||||", "Arial", RGB(0, 255, 0), RGB(100, 200, 100), 72, 72, 0, 0, 80, 80, 0, 0, , 0, 100000000
  TextAnimation1.Counter = 0
  TextAnimation1.CounterMax = 100000000
  TextAnimation1.Speed = 100  'Don't need animation
  TextAnimation1.bOrder = None
  GetWindowRect Me.hwnd, lprect
  SetWindowPos Me.hwnd, HWND_TOPMOST, lprect.Left, lprect.Top, lprect.Right - lprect.Left, lprect.Bottom - lprect.Top, SWP_NOACTIVATE + SWP_NOOWNERZORDER
  
End Sub


Public Sub setVolume(level As Integer)
On Error Resume Next
Dim lprect As Rect
  TextAnimation1.MessageText("Level") = Left("||||||||||||", level)
  TextAnimation1.Draw
  Me!Timer1.Enabled = False
  Me!Timer1.Enabled = True
  Me.Show
  
End Sub


Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  ReleaseCapture
  PostMessage Me.hwnd, &HA1, 2, 0&
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
  Unload Me
End Sub
