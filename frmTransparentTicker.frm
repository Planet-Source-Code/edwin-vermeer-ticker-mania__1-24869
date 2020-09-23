VERSION 5.00
Begin VB.Form frmTransparentTicker 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   735
   ClientLeft      =   1230
   ClientTop       =   7305
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   Begin TextAnimationDemo.FormTransparancy FormTransparancy1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1085
      _ExtentY        =   1085
      TransparencyLevel=   150
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      Counter         =   9
      BackGroundStyle =   4
      AnimateInDesignmode=   0   'False
      TransparentColor=   65280
   End
End
Attribute VB_Name = "frmTransparentTicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
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
  TextAnimation1.AddMessage "design5", "If you speak Dutch, then please visit my homepage at www.beursmonitor.com", "Arial", RGB(150, 255, 150), RGB(255, 255, 255), 32, 32, Me.Width / Screen.TwipsPerPixelX, -1000, 0, 0, 0, 0, , 0, 1000 + Me.Width / Screen.TwipsPerPixelX
  GetWindowRect Me.hwnd, lprect
  SetWindowPos Me.hwnd, HWND_TOPMOST, lprect.Left, lprect.Top, lprect.Right - lprect.Left, lprect.Bottom - lprect.Top, SWP_NOACTIVATE + SWP_NOOWNERZORDER
End Sub


Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  ReleaseCapture
  SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
