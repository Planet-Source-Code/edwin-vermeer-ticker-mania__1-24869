VERSION 5.00
Begin VB.Form frmDockBrowser 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   330
   ClientLeft      =   6630
   ClientTop       =   2160
   ClientWidth     =   3570
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
   Icon            =   "frmDockBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      BackColorStart  =   16744576
      BackColorEnd    =   12583104
      Counter         =   73
      AnimateInDesignmode=   0   'False
      TransparentColor=   14737632
      Angle           =   80
   End
   Begin VB.CommandButton btnProperties 
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FF00FF&
      Picture         =   "frmDockBrowser.frx":247A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Properties"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
End
Attribute VB_Name = "frmDockBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public AppBar As New TAppBar

Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Sub Form_Load()
On Error Resume Next

  With frmDockBrowser.AppBar
    .Extends frmDockBrowser
    .AlwaysOnTop = True
    .AutoHide = False
    .FloatTop = True
    .FloatLeft = False
    .FloatRight = False
    .FloatBottom = False
    .SlideEffect = False
    .HorzDockSize = 22
    .Flags = abfAllowTop
    .Edge = abeTop
    DoEvents
    .UpdateBar
    DoEvents
    .UpdateBar
  End With
  Me.Visible = True
  Me.Show
  
  Me!Close.Left = Me.ScaleWidth - Me!Close.Width + 2
  Me!TextAnimation1.Width = Me.ScaleWidth
  
  BuildTickerTape

End Sub


Private Sub Textanimation1_AnimationRestart()

  BuildTickerTape

End Sub


Private Sub TextAnimation1_BeforeDraw(PictureBuffer As PictureBox)
  PictureBuffer.ForeColor = vbWhite
  PictureBuffer.CurrentY = 2
  PictureBuffer.Print "   " & Format(Now, "dd mmm yyyy hh:nn")
End Sub

Public Sub BuildTickerTape()
On Error Resume Next

  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "ticker", App.Title & " " & App.Major & "." & App.Minor & "  (c) " & Year(Now) & ", Vermeer Automatisering", "Arial", RGB(63, 200, 63), vbWhite, 14, 14, Me.ScaleWidth + 32, -Me.ScaleWidth + 32, -1, -1, 0, 0, , 0, Me.ScaleWidth * 2 + 64
  TextAnimation1.MessageLeftEnd("ticker") = -TextAnimation1.MessageWidth("ticker")
  TextAnimation1.MessageIntervalCount("ticker") = Me.ScaleWidth + 32 + TextAnimation1.MessageWidth("ticker")
  TextAnimation1.CounterMax = TextAnimation1.MessageIntervalCount("ticker")
  
  TextAnimation1.Counter = 0
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.Border = None

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

   AppBar.Detach
   Set AppBar = Nothing

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
  Me!Close.Left = Me.ScaleWidth - Me!Close.Width + 2
  Me!TextAnimation1.Width = Me!Close.Left - 1
  'this is necessary to assure that the minimized window is hidden
  If Me.WindowState = vbMinimized Then Me.Hide
    

End Sub



