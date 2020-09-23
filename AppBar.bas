Attribute VB_Name = "CBAppBar"
Option Explicit

' SetWindowLong selectors
Const GWL_WNDPROC = -4&

' Windows messages
Const WM_ACTIVATE = &H6
Const WM_GETMINMAXINFO = &H24
Const WM_ENTERSIZEMOVE = &H231
Const WM_EXITSIZEMOVE = &H232
Const WM_MOVING = &H216
Const WM_NCHITTEST = &H84
Const WM_NCMOUSEMOVE = &HA0
Const WM_SIZING = &H214
Const WM_TIMER = &H113
Const WM_WINDOWPOSCHANGED = &H47

' AppBar's user notification message
Const WM_USER = &H400
Const WM_APPBARNOTIFY = WM_USER + 100

' Subclassing function default result
Const INHERIT_DEFAULT_CALLBACK = -1

Private ghWnd As Long
Private gAppBar As TAppBar
Private gpcbOldWindowProc As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Any, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
                 

' LinkCallback ----------------------------------------------------------------
Public Function LinkCallback(ByVal frmInstance As Form, _
                             ByVal clsInstance As TAppBar)
  ' Store the calling window
  ghWnd = frmInstance.hwnd
  
  ' Store the AppBar class instance
  Set gAppBar = clsInstance
  
  ' Subclass the window procedure
  gpcbOldWindowProc = SetWindowLong(ghWnd, _
                                    GWL_WNDPROC, _
                                    AddressOf lfnAppBarCallback)
End Function

' DetachCallback --------------------------------------------------------------
Public Function DetachCallback()
  
  ' Restore the original window procedure
  SetWindowLong ghWnd, GWL_WNDPROC, gpcbOldWindowProc

End Function

' AppBar Callback function ----------------------------------------------------
Private Function lfnAppBarCallback(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
  
  ' Message Result to be returned by the Callback
  Dim Result As Long
  
  ' Set the standard return value
  Result = INHERIT_DEFAULT_CALLBACK
  
  ' Subclass some events BEFORE the default window procedure
  Select Case uMsg
    
    Case WM_APPBARNOTIFY
      Result = gAppBar.OnAppBarCallbackMsg(wParam, lParam)
    
    Case WM_ENTERSIZEMOVE
      Result = gAppBar.OnEnterSizeMove
    
    Case WM_EXITSIZEMOVE
      Result = gAppBar.OnExitSizeMove
      
    Case WM_GETMINMAXINFO
      Result = gAppBar.OnGetMinMaxInfo(lParam)
    
    Case WM_MOVING
      Result = gAppBar.OnMoving(lParam)
      
    Case WM_NCMOUSEMOVE
      gAppBar.OnNcMouseMove
      
    Case WM_SIZING
      Result = gAppBar.OnSizing(wParam, lParam)
      
    Case WM_TIMER
      gAppBar.OnAppBarTimer
  
  End Select
  
  ' If the subclassing function did not provide a return value
  ' or wants to inherit the default procedure
  If Result = INHERIT_DEFAULT_CALLBACK Then
    ' Call the default window procedure
    Result = CallWindowProc(gpcbOldWindowProc, hwnd, uMsg, wParam, lParam)
  End If
  
  ' Subclass some events AFTER the default window procedure
  Select Case uMsg
    
    Case WM_ACTIVATE
      gAppBar.OnActivate wParam
    
    Case WM_NCHITTEST
      gAppBar.OnNcHitTest lParam, Result
    
    Case WM_WINDOWPOSCHANGED
      gAppBar.OnWindowPosChanged
  
  End Select
  
  ' Return the value set by the subclassing function or by the default proc
  lfnAppBarCallback = Result

End Function
