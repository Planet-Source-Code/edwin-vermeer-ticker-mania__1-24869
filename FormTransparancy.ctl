VERSION 5.00
Begin VB.UserControl FormTransparancy 
   BackColor       =   &H80000001&
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1095
   ScaleWidth      =   1350
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   600
   End
End
Attribute VB_Name = "FormTransparancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Const m_def_TransparencyLevel = 0
Private Const m_def_TransparencyDirection = 0
Dim m_TransparencyLevel As Integer
Dim m_TransparencyDirection As Integer

Private Sub UserControl_Initialize()
    m_TransparencyLevel = 0
    m_TransparencyDirection = 0
End Sub

Private Sub UserControl_InitProperties()
    m_TransparencyLevel = m_def_TransparencyLevel
    m_TransparencyDirection = m_def_TransparencyDirection
    If Ambient.UserMode = True Then Timer1.Enabled = True
    MakeTransparent UserControl.Parent.hwnd, m_TransparencyLevel
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TransparencyLevel = PropBag.ReadProperty("TransparencyLevel", m_def_TransparencyLevel)
    m_TransparencyDirection = PropBag.ReadProperty("TransparencyDirection", m_def_TransparencyDirection)
    If Ambient.UserMode = True Then Timer1.Enabled = True
    MakeTransparent UserControl.Parent.hwnd, m_TransparencyLevel
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TransparencyLevel", m_TransparencyLevel, m_def_TransparencyLevel)
    Call PropBag.WriteProperty("TransparencyDirection", m_TransparencyDirection, m_def_TransparencyDirection)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim retVal As Long

  If m_TransparencyDirection <> 0 Then
    m_TransparencyLevel = m_TransparencyLevel + m_TransparencyDirection
    If m_TransparencyLevel < Abs(m_TransparencyDirection) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 0
    End If
    If m_TransparencyLevel > (255 - Abs(m_TransparencyDirection)) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 255
    End If
    retVal = MakeTransparent(UserControl.Parent.hwnd, m_TransparencyLevel)
    Select Case retVal
    Case 1
      UserControl.Parent.Visible = False
    Case 2
      If m_TransparencyDirection < 0 Then
        UserControl.Parent.Visible = False
        m_TransparencyDirection = 0
        m_TransparencyLevel = 0
      Else
        UserControl.Parent.Visible = True
        m_TransparencyDirection = 0
        m_TransparencyLevel = 255
      End If
    Case Else
      UserControl.Parent.Visible = True
    End Select
  End If
  
End Sub


Public Property Get TransparencyDirection() As Long
    TransparencyDirection = m_TransparencyDirection
End Property

Public Property Let TransparencyDirection(ByVal New_TransparencyDirection As Long)
    m_TransparencyDirection = New_TransparencyDirection
    PropertyChanged "TransparencyDirection"
End Property


Public Property Get TransparencyLevel() As Long
    TransparencyLevel = m_TransparencyLevel
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As Long)
    m_TransparencyLevel = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
End Property

Public Function MakeVisible() As Variant
    m_TransparencyDirection = 4
    MakeTransparent UserControl.Parent.hwnd, m_TransparencyLevel
    UserControl.Parent.SetFocus
End Function

Public Function MakeInVisible() As Variant
    m_TransparencyDirection = -4
    MakeTransparent UserControl.Parent.hwnd, m_TransparencyLevel
End Function


Private Function isTransparent(ByVal hwnd As Long) As Boolean
On Error Resume Next
Dim Msg As Long
  
  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then isTransparent = True Else isTransparent = False
  If Err Then isTransparent = False
End Function


Private Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next

  If Not Ambient.UserMode Then Exit Function
  If Perc < 0 Or Perc > 255 Then
    MakeTransparent = 1
  Else
    Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
    MakeTransparent = 0
  End If
  If Err Then MakeTransparent = 2
End Function


Private Function MakeOpaque(ByVal hwnd As Long) As Long
Dim Msg As Long
On Error Resume Next

  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  Msg = Msg And Not WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
  MakeOpaque = 0
  If Err Then MakeOpaque = 2
End Function




