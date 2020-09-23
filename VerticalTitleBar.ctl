VERSION 5.00
Begin VB.UserControl VerticalTitleBar 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   ScaleHeight     =   3165
   ScaleWidth      =   375
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5530
      BackColorStart  =   0
      BackColorEnd    =   16711680
      Counter         =   311
      Speed           =   30000
      AnimateInDesignmode=   0   'False
      TransparentColor=   12632256
      Angle           =   20
      Repetitions     =   2
   End
End
Attribute VB_Name = "VerticalTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SW_SHOWNORMAL = 1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long


Private Sub TextAnimation1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  TextAnimation1.MessageFontColorStart("title") = RGB(255, 100, 100)
  TextAnimation1.MessageFontColorEnd("title") = RGB(255, 100, 100)
End Sub

Private Sub TextAnimation1_MouseIn()
On Error Resume Next
  TextAnimation1.MessageFontColorStart("title") = RGB(100, 100, 255)
  TextAnimation1.MessageFontColorEnd("title") = RGB(100, 100, 255)
  UserControl.MousePointer = vbArrowQuestion
End Sub

Private Sub TextAnimation1_MouseOut()
On Error Resume Next
  UserControl.MousePointer = vbDefault
  TextAnimation1.MessageFontColorStart("title") = vbWhite
  TextAnimation1.MessageFontColorEnd("title") = vbWhite
End Sub

Private Sub TextAnimation1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, messages() As String)
On Error Resume Next
  UserControl.MousePointer = vbDefault
  TextAnimation1.MessageFontColorStart("title") = vbWhite
  TextAnimation1.MessageFontColorEnd("title") = vbWhite
  RunThisURL "http://www.beursmonitor.com"
End Sub


Private Sub UserControl_Initialize()
  ' Setting the vertical title bar
  TextAnimation1.Height = UserControl.Height
  TextAnimation1.Width = UserControl.Width
  TextAnimation1.RemoveAllMessages
  TextAnimation1.AddMessage "title", App.Title & " " & App.Major & "." & App.Minor & "  " & App.LegalCopyright, "Arial", vbWhite, vbWhite, 12, 12, 0, 0, TextAnimation1.Height / Screen.TwipsPerPixelY - 10, TextAnimation1.Height / Screen.TwipsPerPixelY - 10, 90, 90, , 0, 100000000
  TextAnimation1.Counter = 0
  TextAnimation1.CounterMax = 100000000
  TextAnimation1.Border = None
  TextAnimation1.Draw
End Sub

Private Sub UserControl_Resize()
  UserControl_Initialize
End Sub

Public Sub RunThisURL(myURL As String)
On Error Resume Next
Dim sFileName    As String
Dim sDummy       As String
Dim sBrowserExec As String * 255
Dim lRetVal      As Long
Dim iFileNumber  As Integer

  ' Create a temporary HTM file
  sBrowserExec = Space(255)
  sFileName = "testapp.htm"
  iFileNumber = FreeFile
  Open sFileName For Output As #iFileNumber
    Write #iFileNumber, "<HTML> <\HTML>"
  Close #iFileNumber

  ' Find the default browser.
  lRetVal = FindExecutable(sFileName, sDummy, sBrowserExec)
  sBrowserExec = Trim$(sBrowserExec)

  ' If an application is found, launch it!
  If lRetVal <= 32 Or IsEmpty(sBrowserExec) Then
    MsgBox "Could not find your Browser", vbExclamation, "Browser Not Found"
  Else
    lRetVal = ShellExecute(UserControl.hwnd, "open", sBrowserExec, myURL, sDummy, SW_SHOWNORMAL)
    If lRetVal <= 32 Then
      MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
    End If
  End If
  
  ' remove the temporary file
  Kill sFileName
  
End Sub

Private Sub UserControl_Show()
  On Error Resume Next
  If Not UserControl.Ambient.UserMode Then TextAnimation1.Speed = 30000 Else TextAnimation1.Speed = 50
End Sub
