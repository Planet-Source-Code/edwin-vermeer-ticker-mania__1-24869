VERSION 5.00
Begin VB.Form frmTextAnimation 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text animation demo"
   ClientHeight    =   7080
   ClientLeft      =   6585
   ClientTop       =   3765
   ClientWidth     =   7920
   Icon            =   "frmTextAnimation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   ShowInTaskbar   =   0   'False
   Begin TextAnimationDemo.VerticalTitleBar VerticalTitleBar2 
      Height          =   7095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   12515
   End
   Begin TextAnimationDemo.GradientButton cmdClear 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BevelWidth      =   5
      Caption         =   "Clear"
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownFontEnabled =   -1  'True
      DownForeColor   =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      GradientAngle   =   150
      GradientBlendMode=   1
      GradientColor1  =   32768
      GradientColor2  =   128
      GradientRepetitions=   2
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverFontEnabled=   -1  'True
      HoverForeColor  =   -2147483624
      HoverMode       =   2
      Style           =   2
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      BackColorStart  =   12640511
      BackColorEnd    =   16736448
      Counter         =   1337
      AnimateInDesignmode=   0   'False
      BackGroundImage =   "Textanimation.BMP"
      TransparentColor=   12648384
      Angle           =   200
      Repetitions     =   3
      GradientType    =   1
   End
   Begin VB.TextBox MessageClick 
      Height          =   5535
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox MouseDown 
      Height          =   2175
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4200
      Width           =   3855
   End
   Begin TextAnimationDemo.GradientButton cmdReset 
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BevelWidth      =   5
      Caption         =   "Reset"
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownFontEnabled =   -1  'True
      DownForeColor   =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      GradientAngle   =   150
      GradientBlendMode=   1
      GradientColor1  =   32768
      GradientColor2  =   128
      GradientRepetitions=   2
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverFontEnabled=   -1  'True
      HoverForeColor  =   -2147483624
      HoverMode       =   2
      Style           =   2
   End
   Begin TextAnimationDemo.GradientButton cmdOK 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BevelWidth      =   5
      Caption         =   "Exit"
      BeginProperty DisabledFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DownFontEnabled =   -1  'True
      DownForeColor   =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      GradientAngle   =   150
      GradientBlendMode=   1
      GradientColor1  =   32768
      GradientColor2  =   128
      GradientRepetitions=   2
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverFontEnabled=   -1  'True
      HoverForeColor  =   -2147483624
      HoverMode       =   2
      Style           =   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTextAnimation.frx":044A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Clicked on:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Message clicked:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTextAnimation.frx":0545
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1080
      Left            =   495
      TabIndex        =   10
      Top             =   135
      Width           =   7320
   End
End
Attribute VB_Name = "frmTextAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim grad As New Gradient
Dim sc As String
Dim sc1 As Long
Dim sc2 As Long

Private Sub cmdOk_Click()
  Unload Me
End Sub


Private Sub cmdClear_Click()
  TextAnimation1.RemoveAllMessages
End Sub


Private Sub cmdReset_Click()
  InitializeMessages
End Sub


Private Sub Form_Load()
  InitializeMessages
  grad.Color1 = RGB(15, 21, 15)
  grad.Color2 = RGB(150, 150, 210)
  grad.Angle = 120
  grad.Repetitions = 1.8
  grad.GradientType = 1
End Sub


Private Sub Form_Resize()
On Error Resume Next
  grad.Draw Me               'Actually draws the gradient on the picture box
End Sub


Private Sub Form_Unload(Cancel As Integer)
  frmMainMenu.cmdAnimationOpen_Click (1)
End Sub


Private Sub TextAnimation1_BeforeDraw(PictureBuffer As PictureBox)
' This will be printed behind the text animation
' All message properties, methods and events will not aply to this text.
  PictureBuffer.CurrentX = (PictureBuffer.ScaleWidth / 2) - 95
  PictureBuffer.CurrentY = 50
  PictureBuffer.FontName = "Arial"
  PictureBuffer.ForeColor = vbWhite
  PictureBuffer.FontSize = 32
  PictureBuffer.Print "Edwin"

End Sub

Private Sub TextAnimation1_AfterDraw(PictureBuffer As PictureBox)
' This will be printed in front of the text animation
' All message properties, methods and events will not aply to this text.
  PictureBuffer.CurrentX = (PictureBuffer.ScaleWidth / 2) - 50
  PictureBuffer.CurrentY = 66
  PictureBuffer.FontName = "Arial"
  PictureBuffer.ForeColor = vbRed
  PictureBuffer.FontSize = 32
  PictureBuffer.Print "Vermeer"

End Sub


Private Sub TextAnimation1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
' Just do the same as mouse down
  TextAnimation1_MouseDown Button, Shift, X, Y, messages
End Sub


Private Sub TextAnimation1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
' I disabled the mousemove event because this is slowing down the animation considerably as soon as you move the mouse.
' If you want to enable mousemove, then just convert the two comment lines to statements (around line 164)
  TextAnimation1_MouseDown Button, Shift, X, Y, messages
End Sub


Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
Dim i As Integer
Dim m As Integer
Dim c As String

  ' Display the normal mouse event parameters
  MouseDown = "Button :  " & Button & vbCrLf & _
              "Shift :   " & Shift & vbCrLf & _
              "X :       " & X & vbCrLf & _
              "Y :       " & Y & vbCrLf
  
  ' How many messages where unther the mouse cursor?
  On Error Resume Next
  m = UBound(messages)
  If Err.Number = 9 Then
    c = ""
  Else
    c = messages(UBound(messages))
  End If
  
  ' Change the color of the clicked message
  ' Exept for MessageHeight, MessageID, MessageIndex, and MessageWidth you can also change all the other message properties
  If sc <> c Then
    TextAnimation1.MessageFontColorStart(sc) = sc1
    TextAnimation1.MessageFontColorEnd(sc) = sc2
    sc = c
    If c <> "" Then
      sc1 = TextAnimation1.MessageFontColorStart(c)
      sc2 = TextAnimation1.MessageFontColorEnd(c)
      TextAnimation1.MessageFontColorStart(c) = vbRed
      TextAnimation1.MessageFontColorEnd(c) = vbRed
    End If
  End If
  
  m = UBound(messages)
  If Err.Number = 9 Then
    ' No messages under the mouse cursor
    On Error GoTo 0
    MouseDown = MouseDown & "No messages here"
    MessageClick = ""
  Else
    ' Display all the message ID's that are under the mouse cursor
    On Error GoTo 0
    MouseDown = MouseDown & "Messages here : " & m + 1 & " out of " & TextAnimation1.MessageCount & vbCrLf
    For i = 0 To m
      MouseDown = MouseDown & "Message " & i & " : " & messages(i) & vbCrLf
    Next i
    ' The last message in the passed array will be the topmost. (the one you actually clicked on)
    ' Display all properties of this message.
    MessageClick = "MessageFontColorEnd : " & sc1 & vbCrLf & _
                   "MessageFontColorStart : " & sc2 & vbCrLf & _
                   "MessageFontName : " & TextAnimation1.MessageFontName(c) & vbCrLf & _
                   "MessageFontRotationEnd : " & TextAnimation1.MessageFontRotationEnd(c) & vbCrLf & _
                   "MessageFontRotationStart : " & TextAnimation1.MessageFontRotationStart(c) & vbCrLf & _
                   "MessageFontSizeEnd : " & TextAnimation1.MessageFontSizeEnd(c) & vbCrLf & _
                   "MessageFontSizeStart : " & TextAnimation1.MessageFontSizeStart(c) & vbCrLf & _
                   "MessageHeight : " & TextAnimation1.MessageHeight(c) & vbCrLf & _
                   "MessageID : " & TextAnimation1.MessageID(0) & vbCrLf & _
                   "MessageIndex : " & TextAnimation1.MessageIndex(c) & vbCrLf & _
                   "MessageIntervalCount : " & TextAnimation1.MessageIntervalCount(c) & vbCrLf & _
                   "MessageIntervalStart : " & TextAnimation1.MessageIntervalStart(c) & vbCrLf & _
                   "MessageLeftEnd : " & TextAnimation1.MessageLeftEnd(c) & vbCrLf & _
                   "MessageLeftStart : " & TextAnimation1.MessageLeftStart(c) & vbCrLf & _
                   "MessageText : " & TextAnimation1.MessageText(c) & vbCrLf & _
                   "MessageTopEnd : " & TextAnimation1.MessageTopEnd(c) & vbCrLf & _
                   "MessageTopStart : " & TextAnimation1.MessageTopStart(c) & vbCrLf & _
                   "MessageWidth : " & TextAnimation1.MessageWidth(c) & vbCrLf
  End If

End Sub


Public Sub InitializeMessages()
Dim i As Integer
Dim j As Integer
Dim s As Integer

  TextAnimation1.RemoveAllMessages
  
  ' Animate the 4 messages
  TextAnimation1.AddMessage "test00", "This is test message 1", "Arial", vbBlue, vbWhite, 16, 16, 200, 0, 0, 0, 0, 0, , 0, 300
  TextAnimation1.AddMessage "test01", "This is test message 2", "Arial", vbBlue, vbRed, 16, 16, 200, 0, 200, 0, 0, 0, , 0, 300
  TextAnimation1.AddMessage "test02", "Turn !", "Arial", vbBlue, vbYellow, 24, 24, 100, 100, 100, 100, 0, 360, , 0, 300
  TextAnimation1.AddMessage "test03", "Zoom", "Arial", vbBlue, vbGreen, 1, 100, 100, 0, 0, 170, 0, 0, , 0, 300
  
  ' Followed by the referce animation
  TextAnimation1.AddMessage "test04", "This is test message 1", "Arial", vbWhite, vbBlue, 16, 16, 0, 200, 0, 0, 0, 0, , 300, 300
  TextAnimation1.AddMessage "test05", "This is test message 2", "Arial", vbRed, vbBlue, 16, 16, 0, 200, 0, 200, 0, 0, , 300, 300
  TextAnimation1.AddMessage "test06", "Turn !", "Arial", vbYellow, vbBlue, 24, 24, 100, 100, 100, 100, 0, 360, , 300, 300
  TextAnimation1.AddMessage "test07", "Zoom", "Arial", vbGreen, vbBlue, 100, 1, 0, 100, 170, 0, 0, 0, , 300, 300
  
  ' And another animation cyclus for those 4 messages
  TextAnimation1.AddMessage "test08", "This is test message 1", "Arial", vbBlue, vbWhite, 16, 16, 200, 0, 0, 0, 0, 0, , 600, 300
  TextAnimation1.AddMessage "test09", "This is test message 2", "Arial", vbBlue, vbRed, 16, 16, 200, 0, 200, 0, 0, 0, , 600, 300
  TextAnimation1.AddMessage "test10", "Turn !", "Arial", vbBlue, vbYellow, 24, 24, 100, 100, 100, 100, 0, 360, , 600, 300
  TextAnimation1.AddMessage "test11", "Zoom", "Arial", vbBlue, vbGreen, 1, 100, 100, 0, 0, 170, 0, 0, , 600, 300
  
  ' Followed by the referce animation
  TextAnimation1.AddMessage "test12", "This is test message 1", "Arial", vbWhite, vbBlue, 16, 16, 0, 200, 0, 0, 0, 0, , 900, 300
  TextAnimation1.AddMessage "test13", "This is test message 2", "Arial", vbRed, vbBlue, 16, 16, 0, 200, 0, 200, 0, 0, , 900, 300
  TextAnimation1.AddMessage "test14", "Turn !", "Arial", vbYellow, vbBlue, 24, 24, 100, 100, 100, 100, 0, 360, , 900, 300
  TextAnimation1.AddMessage "test15", "Zoom", "Arial", vbGreen, vbBlue, 100, 1, 0, 100, 170, 0, 0, 0, , 900, 300
  
  ' Just put in 14 * 6 = 84 different messages
  For i = 0 To 13
    s = 0
    For j = 0 To 5
      TextAnimation1.AddMessage "text" & Trim(Str(i)) & "-" & Trim(Str(j)), Left(Trim(Str(i)) & "-" & Trim(Str(j)) & "ABCDEFGHIJKLMNOPQRSTUVWXYZ", i + j + 3), "Arial", RGB(0, 180 + i * j, 0), RGB(0, 0, 180 + i * j), 12, 12, 300, -200, i * 16, i * 16, 0, 0, , s, 500
      s = s + TextAnimation1.MessageWidth("text" & Trim(Str(i)) & "-" & Trim(Str(j)))
    Next j
  Next i
  
  ' Set the general animation properties
  TextAnimation1.Counter = 0
  TextAnimation1.CounterMax = 1200
  TextAnimation1.Speed = 20  'This is equal to a refresh rate of 50 frames per second. Making this number smaller will only slow down your application. If you want more speed, then make the all the MessageIntervalCount parameters smaller.
  TextAnimation1.Border = None
  
End Sub

