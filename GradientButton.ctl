VERSION 5.00
Begin VB.UserControl GradientButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   DefaultCancel   =   -1  'True
   ScaleHeight     =   915
   ScaleWidth      =   1575
   Begin VB.PictureBox picGradient 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "GradientButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
'Declarations
'**************************************************************************

'Private Control Constants (Used with AboutBox)
Private Const AppName       As String = "Gradient Button"
Private Const Version       As String = "1.3.0"
Private Const Author        As String = "Stephen Kent"
Private Const SpecialThanks As String = "Special thanks to the people/groups from which I used code from: (In no particular order)" & vbCrLf & _
                                        "Edwin Vermeer, Kath-Rock Software, Night Wolf, Nightshadow, Microsoft, Stuart Pennington, and Ulli"
Private Const SupInfo       As String = "I have tried to make this as bug free as possible, but I can't guarantee that it is bug free.  If you do find a bug please send me an e-mail with any information you have on it.  SFalcon@Softhome.net" & vbCrLf & _
                                        "NOTE: This has only been tested in the Visual Basic environment and no guarantees are made concerning other environments."

'API Declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lprect As RectAPI, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lprect As RectAPI, ByVal hBrush As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lprect As RectAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'API Type Definitions
Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type TOOLINFO   'API ToolTip Info Structure
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uID As Long
    rct As RectAPI
    hInst As Long
    TipText As String
    lParam As Long
End Type

'Public Enumerations (For Properties)
Public Enum gbAlignment
    gbaLeftTop = 0
    gbaLeftMiddle = 1
    gbaLeftBottom = 2
    gbaRightTop = 3
    gbaRightMiddle = 4
    gbaRightBottom = 5
    gbaCenterTop = 6
    gbaCenterMiddle = 7
    gbaCenterBottom = 8
End Enum

Public Enum gbAppearance
    gbaFlat = 0
    gba3D = 1
    gbaEtched = 2
    gbaBevel = 3
End Enum

Public Enum gbAutoSize
    gbasNone = 0
    gbasPictureToControl = 1
    gbasControlToPicture = 2
End Enum

Public Enum gbStyle
    gbsStandard = 0
    gbsGraphical = 1
    gbsGradient = 2
    gbsGraphicalGradient = 3
    gbsPicture = 4
    gbsGraphicalPicture = 5
End Enum

Public Enum gbCaptionStyle
    gbcStandard = 0
    gbcInsetLight = 1
    gbcInsetHeavy = 2
    gbcRaisedLight = 3
    gbcRaisedHeavy = 4
    gbcDropShadow = 5
End Enum

Public Enum gbHoverMode
    gbhFullHover = 0
    gbhBorderOnly = 1
    gbhAllButBorder = 2
End Enum

Public Enum gbType
    gbtStandardButton = 0
    gbtStateButton = 1
    gbtOptionButton = 2
End Enum

'Private Enumerations (Internal use only)
Private Enum gbState        'Button states
    gbsDefault = 0          'Special state only used as default value of optional parameters of this type to indicate no value was passed
    gbsDisable = 99         'Button is disabled
    gbsDown = 2             'Button is depressed
    gbsMouseOff = 4         'Mouse is not over the button (graphic/effects depend on whether or not hover mode is enabled [if hover disabled then same as gbsMouseOver])
    gbsMouseOver = 3        'Mouse is currently over the button
    gbsUp = 1               'State used in Design mode to force button to be placed in the "Up" or "Mouse Over" state
End Enum

'CreateWindow Constants
Private Const CW_USEDEFAULT = &H80000000
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const TTS_ALWAYSTIP = &H1
Private Const TTS_NOPREFIX = &H2
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_POPUP = &H80000000

'SetWindowPos Constants
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

'ToolTip Info Constants
Private Const TTF_SUBCLASS = &H10

'SendMessage Constants
Private Const WM_USER = &H400
Private Const TTM_ADDTOOL = (WM_USER + 4)           'ANSI
Private Const TTM_NEWTOOLRECT = (WM_USER + 6)       'ANSI
Private Const TTM_UPDATETIPTEXT = (WM_USER + 12)    'ANSI

'DrawText API Constants
Private Const DT_CALCRECT   As Long = &H400     'Used to adjust the bottom of the rectangle to account for all text
Private Const DT_CENTER     As Long = &H1       'Used to center the text
Private Const DT_LEFT       As Long = &H0       'Used to left justify the text
Private Const DT_RIGHT      As Long = &H2       'Used to right justify the text

'Raster Operation Codes
Private Const DSna = &H220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines

'Color Adjustment Constants
Private Const AmountLighten     As Long = 48
Private Const AmountDarken      As Long = -80

'Default Control Constants
Private Const DefHeight         As Long = 510   'Default initialization height for the control
Private Const DefWidth          As Long = 1230  'Default initialization width for the control

'Local Variables:
Private bAutoSizing             As Boolean      'Used to prevent unnecessary iterations of the Resize event while autosizing.
Private bInitializing           As Boolean      'Used to prevent excess redraws on initialize
Private bInitiating             As Boolean      'Used to prevent Value Change while resetting the values of other buttons in group in option button mode
Private bInside                 As Boolean      'Variable used for Mouse Enter/Exit
Private BorderDark              As Long         'Variable that holds the Dark Border Color
Private BorderDarkest           As Long         'Variable that holds the Darkest Border Color
Private BorderLight             As Long         'Variable that holds the Light Border Color
Private BorderLightest          As Long         'Variable that holds the Lightest Border Color
Private bParentAvailable        As Boolean      'Variable that holds whether the parent object is available with all necessary properties
Private bToolTipAvailable       As Boolean      'Variable that holds whether the extender with ToolTip info is present
Private CurrentCaptionStyle     As Long         'Variable to hold the style to draw the caption with currently
Private grad                    As Gradient     'Variable draws all of the gradients (holds instance of class Gradient)
Private GradPic                 As Picture      'Variable to hold the Gradient after it has been drawn
Private hWndToolTip             As Long         'Variable to hold the reference to the ToolTip control
Private rctToolTip              As RectAPI      'Variable to hold the Rectanglular co-ordinates of the button
Private State                   As gbState      'Internal variable that contains the current state of the button
Private tiToolTip               As TOOLINFO     'Variable to hold the information about the ToolTip
Private WrappedCaption          As String       'Variable to hold a pre-prepared caption

'Default Property Values:
Private Const m_def_Alignment           As Long = gbaCenterMiddle   'Default Alignment: True Center
Private Const m_def_AlignmentCushion    As Long = 3                 'Default Alignment Cushion: 3 Pixels
Private Const m_def_Appearance          As Integer = 1              'Default Appearance: 3D
Private Const m_def_AutoSize            As Long = gbasNone          'Default AutoSizing: None
Private Const m_def_BevelIntensity      As Long = 20                'Default Bevel Intensity: 20
Private Const m_def_BevelWidth          As Byte = 3                 'Default Bevel Width: 3 Pixels
Private Const m_def_BorderColor         As Long = vbButtonFace      'Default Border Color: System Button Face Color
Private Const m_def_ButtonType          As Long = gbtStandardButton 'Default Button Type: Standard Button
Private Const m_def_Caption             As String = vbNullString    'Default Caption: Empty
Private Const m_def_CaptionStyle        As Long = 0                 'Default Caption Style: Standard
Private Const m_def_DisabledCaptionStyle As Long = 0                'Default DisabledCaption Style: Standard
Private Const m_def_DisabledFontEnabled As Boolean = False          'Default Disabled Font: Use Regular Font
Private Const m_def_DisabledForeColor   As Long = vbGrayText        'Default Disabled Text Color: System Button Text
Private Const m_def_DisabledMousePointer As Long = vbDefault        'Default Disabled Mouse Pointer: Default Cursor
Private Const m_def_DownCaptionStyle    As Long = 0                 'Default Down Caption Style: Standard
Private Const m_def_DownFontEnabled     As Boolean = False          'Default Down Font: Use Regular Font
Private Const m_def_DownForeColor       As Long = vbButtonText      'Default Down Text Color: System Button Text
Private Const m_def_DownMousePointer    As Long = vbDefault         'Default Down Mouse Pointer: Default Cursor
Private Const m_def_Enabled             As Boolean = True           'Default Enable: True (Events are Fired)
Private Const m_def_ForeColor           As Long = vbButtonText      'Default Text Color: System Button Text
Private Const m_def_GradientAngle       As Double = 0               'Default Gradient Angle: 0 (Color Fade Color2 - to - Color1)
Private Const m_def_GradientBlendMode   As Long = gbmRGB            'Default Gradient Blend Mode: RGB Colors
Private Const m_def_GradientColor1      As Long = vbButtonFace      'Default Gradient Color1: System Button Face
Private Const m_def_GradientColor2      As Long = vbButtonFace      'Default Gradient Color2: System Button Face
Private Const m_def_GradientRepetitions As Double = 1               'Default Gradient Repetitions: 1
Private Const m_def_GradientType        As Long = gtNormal          'Default Gradient Type: Normal (Lines)
Private Const m_def_HoverCaptionStyle   As Long = 0                 'Default Hover Caption Style: Standard
Private Const m_def_HoverFontEnabled    As Boolean = False          'Default Hover Font: Use Regular Font
Private Const m_def_HoverForeColor      As Long = vbButtonText      'Default Hover Text Color: System Button Text
Private Const m_def_HoverMode           As Long = gbhFullHover      'Default Hover Mode: Full Hover Effects
Private Const m_def_HoverMousePointer   As Long = vbDefault         'Default Hover Mouse Pointer: Default Cursor
Private Const m_def_MousePointer        As Long = vbDefault         'Default Mouse Pointer: Default Cursor
Private Const m_def_PictureAlignment    As Long = gbaCenterMiddle   'Default Picture Alignment: True Center
Private Const m_def_PictureCushion      As Long = 3                 'Default Picture Cushion: 3 Pixels
Private Const m_def_Style               As Long = 0                 'Default Style: Standard
Private Const m_def_ToolTipText         As String = vbNullString    'Default Tool Tip Text: Null String
Private Const m_def_UseClassicBorders   As Boolean = False          'Default Use Classic Borders: False (Use New Borders)
Private Const m_def_UseHover            As Boolean = True           'Default Hover Mode: Enabled
Private Const m_def_Value               As Boolean = False          'Default Check Value: False (Not Depressed)

'Property Variables:
Private m_Alignment             As Long             'Local Variable holds caption alignment
Private m_AlignmentCushion      As Long             'Local Variable holds the cushion between the edge of the control and the Caption for non-centered alignments
Private m_Appearance            As Integer          'Local Variable holds Drawing Style (Flat[0]/3D[1]) [Enumeration not necessary for local]
Private m_AutoSize              As gbAutoSize       'Local Variable to hold the autosizing mode used by the control/Background Picture [Used only for Picture and Graphical Picture Modes]
Private m_BackPicture           As Picture          'Local Variable to hold the picture to use as the back of the button.  [Used only for Picture and Graphical Picture Modes]
Private m_BevelIntensity        As Byte             'Local Variable to hold the intensity adjustment of the Bevel border
Private m_BevelWidth            As Long             'Local Variable to hold the with of the Bevel Style Border
Private m_BorderColor           As Long             'Local Variable holds color to use when drawing the border [Save resources by using Long instead of OLE_COLOR]
Private m_ButtonType            As Long             'Local Variable to hold the type of button that the control is.
Private m_Caption               As String           'Local Variable to hold the caption of the button.
Private m_CaptionStyle          As gbCaptionStyle   'Local Variable to hold the Caption display style
Private m_DisabledCaptionStyle  As gbCaptionStyle   'Local Variable to hold the Caption display style for the Disabled state
Private m_DisabledFont          As Font             'Local Variable holds font to use when button is disabled [Font is an object so always use set or individual properties]
Private m_DisabledFontEnabled   As Boolean          'Local Variable to tell whether to use regular or disabled font when button is disabled
Private m_DisabledForeColor     As Long             'Local Variable holds color to use for text color when button is disabled [Save resources by using Long instead of OLE_COLOR]
Private m_DisabledMouseIcon     As Picture          'Local Variable holds the picture to use as the custom mouse icon in the disabled state
Private m_DisabledMousePointer  As Long             'Local Variable to hold the pointer to use for the mouse in the disabled state
Private m_DisabledPicture       As Picture          'Local Variable holds picture to use when button is disabled (if nothing use regular picture {only applies to Graphical modes}) [Picture is an object so always use set or individual properties]
Private m_DownCaptionStyle      As gbCaptionStyle   'Local Variable to hold the Caption display style for the Depressed state
Private m_DownFont              As Font             'Local Variable holds font to use when button is down [Font is an object so always use set or individual properties]
Private m_DownFontEnabled       As Boolean          'Local Variable to tell whether to use regular or down font when button is down
Private m_DownForeColor         As Long             'Local Variable holds color to use for text color when button is down [Save resources by using Long instead of OLE_COLOR]
Private m_DownMouseIcon         As Picture          'Local Variable holds the picture to use as the custom mouse icon in the down state
Private m_DownMousePointer      As Long             'Local Variable to hold the pointer to use for the mouse in the down state
Private m_DownPicture           As Picture          'Local Variable holds picture to use when button is down (if nothing use regular picture {only applies to Graphical modes}) [Picture is an object so always use set or individual properties]
Private m_Enabled               As Boolean          'Local Variable to control whether events are fired or not
Private m_Font                  As Font             'Local Variable holds the regular font for the button and default if any other font is disabled [Font is an object so always use set or individual properties]
Private m_ForeColor             As Long             'Local Variable holds the regular font Fore Color and default if any other font is disabled [Save resources by using Long instead of OLE_COLOR]
Private m_GradientAngle         As Double           'Local Variable holds the angle to draw the gradient background at (Only applies to Gradient modes)
Private m_GradientBlendMode     As GradBlendMode    'Local Variable hold the blending mode to use for gradients
Private m_GradientColor1        As Long             'Local Variable holds the first color code to use when drawing the Gradient [Save resources by using Long instead of OLE_COLOR]
Private m_GradientColor2        As Long             'Local Variable holds the second color code to use when drawing the Gradient [Save resources by using Long instead of OLE_COLOR]
Private m_GradientRepetitions   As Double           'Local Variable holds the number of times to repeat the gradient across the button
Private m_GradientType          As GradType         'Local Variable holds the type of gradient to draw. (Only used for Gradient Modes)
Private m_HoverCaptionStyle     As gbCaptionStyle   'Local Variable to hold the Caption display style for the Hover state
Private m_HoverFont             As Font             'Local Variable holds font to use when button is in hover mode (Only if Hover Mode is enabled) [Font is an object so always use set or individual properties]
Private m_HoverFontEnabled      As Boolean          'Local Variable to tell whether to use regular or disabled font when button is in hover mode (Only if Hover Mode is enabled)
Private m_HoverForeColor        As Long             'Local Variable holds color to use for text color when button is in hover mode (Only if Hover Mode is enabled) [Save resources by using Long instead of OLE_COLOR]
Private m_HoverMode             As gbHoverMode      'Local Variable to hold the Hover mode which is used to determine which elements use hover effects if hover mode is on.
Private m_HoverMouseIcon        As Picture          'Local Variable holds the picture to use as the custom mouse icon in the mouse over state
Private m_HoverMousePointer     As Long             'Local Variable to hold the pointer to use for the mouse in the mouse over state
Private m_HoverPicture          As Picture          'Local Variable holds picture to use when button is in hover mode {Only if Hover Mode is enabled} (if nothing use regular picture {only applies to Graphical modes}) [Picture is an object so always use set or individual properties]
Private m_MouseIcon             As Picture          'Local Variable holds the picture to use as the custom mouse icon in the Up/Default state
Private m_MousePointer          As Long             'Local Variable to hold the pointer to use for the mouse in the Up/Default state
Private m_Picture               As Picture          'Local Variable holds regular picture and default for any other picture that is not present (Only applies to Graphical Modes) [Picture is an object so always use set or individual properties]
Private m_PictureAlignment      As gbAlignment      'Local Variable holds the alignment at which to draw the picture in graphical modes.  [only applies to Graphical modes]
Private m_PictureCushion        As Long             'Local Variable to hold the number of pixels between the picture and the edge of the control.
Private m_Style                 As Integer          'Local Variable holds button style (Standard[0]/Graphical[1]/Gradient[2]/Graphical Gradient[3]) [Enumeration not necessary for local]
Private m_ToolTipText           As String           'Local Variable to hold ToolTipText if Extender ToolTip Information is not available
Private m_UseClassicBorders     As Boolean          'Local Variable to indicate whether we are using the new or classic border styles.
Private m_UseHover              As Boolean          'Local Variable to tell whether Hover mode is enabled or disabled.
Private m_Value                 As Boolean          'Local Variable to hold what the current value of the button is if it is acting as a state button

'**************************************************************************
'Events
'**************************************************************************

'Event Declarations:
Public Event Click()        'Default Interface Event (Fired on MouseUp but before MouseUp event also fired if user presses space/enter or one of the Access keys)
Attribute Click.VB_Description = "Event raised whenever the button is clicked."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)      'KeyDown event fired when any key is pressed while the button has focus
Attribute KeyDown.VB_Description = "Event fired whenever a key is released while the button has focus."
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(KeyAscii As Integer)      'KeyPress event fired when any key is pressed while the button has focus
Attribute KeyPress.VB_Description = "Event fired whenever a key is pressed while the button has focus."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)        'KeyUp event fired when any key is released while the button has focus
Attribute KeyUp.VB_Description = "Event fired whenever a key is depressed while the button has focus."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)       'MouseDown event fired when any button of the mouse is pressed while over the button
Attribute MouseDown.VB_Description = "Event fired when any Mouse button is pressed while over the button."
Public Event MouseEnter()       'MouseEnter event fired when the mouse cursor enters the button area
Attribute MouseEnter.VB_Description = "Event fired when the mouse enters the area over the button."
Public Event MouseExit()        'MouseExit event fired when the mouse cursor leaves the button area
Attribute MouseExit.VB_Description = "Event fired when the mouse leaves the area over the button."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)       'MouseMove event fired when the mouse moves in the button area
Attribute MouseMove.VB_Description = "Event fired when the Mouse is moved while in the area of the button."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)         'MouseUp event fired when any button of the mouse is released in the button area (Excluding times when all mouse activity is captured by another control or object)
Attribute MouseUp.VB_Description = "Event fired when any Mouse button is released while in the area over the button."
Public Event OLECompleteDrag(Effect As Long)        'OLECompleteDrag Occurs after OLE object is dropped on button
Attribute OLECompleteDrag.VB_Description = "Event fired when an OLE drag operation completes."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)     'OLEDragDrop occurs when Source determins if a drop can occur
Attribute OLEDragDrop.VB_Description = "Event fired when an OLE object is actually dropped on the button."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)       'OLEDragOver occurs when OLE object is dragged over button
Attribute OLEDragOver.VB_Description = "Event fired when an OLE object is dragged over the button."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)     'OLEGiveFeedback occurs after OLEDragOver to provide user some visual feedback
Attribute OLEGiveFeedback.VB_Description = "Event used to provide feedback to the source of an OLE operation."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)          'OLESetData occurs when target does OLEGetData but data has not yet been set
Attribute OLESetData.VB_Description = "Event fired when the target of an OLE operation requests a format that data has not been set for in the dragged object."
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)       'OLEStartDrag occurs when button initiates OLE Drag/Drop Operation
Attribute OLEStartDrag.VB_Description = "Event fired after an OLE Drag is started for the current button."
Public Event ValueChanged(New_Value As Boolean)     'ValueChanged is fired whenever the value of the button changes while in State or Option Button mode
Attribute ValueChanged.VB_Description = "Event fired whenever the Value of the button is changed.  (Only fired when ButtonType is StateButton or OptionButton)"

'**************************************************************************
'Properties & Methods
'**************************************************************************

'Sub-Procedure to display information about the control.
Public Sub About()
Attribute About.VB_Description = "Displays an about box giving information about the control."
Attribute About.VB_UserMemId = -552
On Error Resume Next
    MsgBox "Control : " & AppName & vbCrLf & _
     "Version : " & Version & vbCrLf & _
     "Author : " & Author & vbCrLf & _
     "SpecialThanks : " & SpecialThanks & vbCrLf & _
     "SupInfo : " & SupInfo & vbCrLf
End Sub

'Property set for Alignment of Caption (LeftTop/LeftMiddle/LeftBottom/RightTop/RightMiddle/RightBottom/CenterTop/CenterMiddle/CenterBottom)
Public Property Get Alignment() As gbAlignment
Attribute Alignment.VB_Description = "Returns/sets the caption text alignment (LeftTop / LeftMiddle / LeftBottom / RightTop / RightMiddle / RightBottom / CenterTop / CenterMiddle / CenterBottom)."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As gbAlignment)
On Error Resume Next
    If (New_Alignment < gbaLeftTop) Or (New_Alignment > gbaCenterBottom) Then Exit Property   'If new value isn't valid then exit the property doing nothing
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint for the caption to be redrawn with the new alignment
End Property
'End Alignment Property set

'Property set for Alignment cushion for the Caption
Public Property Get AlignmentCushion() As Long
Attribute AlignmentCushion.VB_Description = "Returns/sets the number of pixels between the Caption and the edge of the button as a cushion zone."
Attribute AlignmentCushion.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    AlignmentCushion = m_AlignmentCushion
End Property

Public Property Let AlignmentCushion(ByVal New_AlignmentCushion As Long)
On Error Resume Next
    If New_AlignmentCushion < 0 Then Exit Property  'Make sure we have a valid cushion
    m_AlignmentCushion = New_AlignmentCushion
    PropertyChanged "AlignmentCushion"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint for the caption to be redrawn with the new alignment
End Property
'End AlignmentCushion Property set

'Property set for button display type (Flat/3D)
Public Property Get Appearance() As gbAppearance
Attribute Appearance.VB_Description = "Returns/sets the button's border appearance and behavior (Flat / 3D / Etched / Bevel)."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
On Error Resume Next
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As gbAppearance)
On Error Resume Next
    If (New_Appearance < gbaFlat) Or (New_Appearance > gbaBevel) Then Exit Property     'If new value isn't valid then exit the property doing nothing
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Execute a fast repaint of the control to reflect new appearance (Gradient need not be redrawn)
End Property
'End Appearance Property set

'Property set for the autosizing property of the control
Public Property Get AutoSize() As gbAutoSize
Attribute AutoSize.VB_Description = "Returns/sets what AutoSizing method to use for the background picture / control.  (Only used for Picture and Graphical Picture modes) [None / Picture to Control / Control to Picture]."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As gbAutoSize)
On Error Resume Next
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    UserControl_Resize  'Call the resize event so that the control is resized if necessary
End Property
'End BackColor Property set

'Property set for the back color of the button in Standard and Graphical Modes (Not applicable to any Gradient Mode)
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color of the button.  (Used only in Standard and Graphical Modes)"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
On Error Resume Next
    BackColor = picGradient.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next
    picGradient.BackColor() = New_BackColor
    PropertyChanged "BackColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint because changing the back color redraws the control
End Property
'End BackColor Property set

'Property set for the Back Picture of the button (Picture and Graphical Picture styles only)
Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "Returns/set the Picture to use as the background of the button.  (Used only in Picture and Graphical Picture Modes)"
Attribute BackPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    Set BackPicture = m_BackPicture
End Property

Public Property Set BackPicture(New_BackPicture As Picture)
On Error Resume Next
    Set m_BackPicture = New_BackPicture
    PropertyChanged "BackPicture"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gbsPicture) Or (m_Style = gbsGraphicalPicture) Then PaintFast     'If in one of the picture style modes then do a fast paint to update the control.
End Property
'End BackPicture Property set

'Property set for the Bevel Border Intensity used only for Bevel Appearance
Public Property Get BevelIntensity() As Long
Attribute BevelIntensity.VB_Description = "Returns/sets the Bevel Intensity of the Beveled Border (Bevel Appearance Only)"
Attribute BevelIntensity.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    BevelIntensity = m_BevelIntensity
End Property

Public Property Let BevelIntensity(ByVal New_BevelIntensity As Long)
On Error Resume Next
    If (New_BevelIntensity < 0) Or (New_BevelIntensity > 255) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_BevelIntensity = New_BevelIntensity
    PropertyChanged "BevelIntensity"    'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If m_Appearance = gbaBevel Then PaintFast   'If appearance is beveled then do a fast paint to update control
End Property
'End BevelIntensity Property set

'Property set for the Bevel Border Width used only for Bevel Appearance
Public Property Get BevelWidth() As Long
Attribute BevelWidth.VB_Description = "Returns/Sets the Bevel Width of the button.  (Bevel Appearance Only)"
Attribute BevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    BevelWidth = m_BevelWidth
End Property

Public Property Let BevelWidth(ByVal New_BevelWidth As Long)
On Error Resume Next
    If (New_BevelWidth < 1) Or (New_BevelWidth > (ScaleX(ScaleWidth, ScaleMode, vbPixels) / 2)) Or (New_BevelWidth > (ScaleY(ScaleHeight, ScaleMode, vbPixels) / 2)) Then Exit Property     'If new value isn't valid then exit the property doing nothing
    m_BevelWidth = New_BevelWidth
    PropertyChanged "BevelWidth"    'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If m_Appearance = gbaBevel Then PaintFast   'If appearance is beveled then do a fast paint to update control
End Property
'End BevelWidth Property set

'Property set for Button Border color in Gradient Modes only
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of the border.  (Used only when UseClassicBorders is set to True, Only for Gradient, Graphical Gradient, Picture, and Graphical Picture modes)"
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderColor.VB_UserMemId = -503
On Error Resume Next
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
On Error Resume Next
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gbsGradient) Or (m_Style = gbsGraphicalGradient) Then     'If in a gradient mode then
        SetColors       'Execute a set colors so that the colors will correctly update
        PaintFast       'Execute a fast repaint of the control to reflect new border color
    End If
End Property
'End BorderColor Property set

'Property set for Button Type
Public Property Get ButtonType() As gbType
Attribute ButtonType.VB_Description = "Returns/sets the type of button that the control will act as.  (Standard Button / State Button / Option Button)"
Attribute ButtonType.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    ButtonType = m_ButtonType
End Property

Public Property Let ButtonType(ByVal New_ButtonType As gbType)
On Error Resume Next
    If (New_ButtonType < gbtStandardButton) Or (New_ButtonType > gbtOptionButton) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_ButtonType = New_ButtonType
    PropertyChanged "ButtonType"    'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint to update the appearance of the button because of changing type modes
End Property
'End ButtonType Property set

'Property set for button caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the caption that is to be used for the button."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
On Error Resume Next
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
On Error Resume Next
    m_Caption = New_Caption
    SetAccessKeys       'Set the access keys for the control
    PropertyChanged "Caption"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a Paint All so that the caption is re-wrapped and displayed correctly.
End Property
'End Caption Property set

'Property set for Caption Style
Public Property Get CaptionStyle() As gbCaptionStyle
Attribute CaptionStyle.VB_Description = "Returns/sets the Caption Style for the button.  (Standard / Light Inset / Heavy Inset / Light Raised / Heavy Raised / Drop Shadow)"
Attribute CaptionStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As gbCaptionStyle)
On Error Resume Next
    If (New_CaptionStyle < gbcStandard) Or (New_CaptionStyle > gbcDropShadow) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint to refresh the caption with the new style
End Property
'End CaptionStyle Property set

'Property set for Disabled Caption Style
Public Property Get DisabledCaptionStyle() As gbCaptionStyle
Attribute DisabledCaptionStyle.VB_Description = "Returns/sets the Caption Style for the button in the disabled state.  (Standard / Light Inset / Heavy Inset / Light Raised / Heavy Raised / Drop Shadow)"
Attribute DisabledCaptionStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    DisabledCaptionStyle = m_DisabledCaptionStyle
End Property

Public Property Let DisabledCaptionStyle(ByVal New_DisabledCaptionStyle As gbCaptionStyle)
On Error Resume Next
    If (New_DisabledCaptionStyle < gbcStandard) Or (New_DisabledCaptionStyle > gbcDropShadow) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_DisabledCaptionStyle = New_DisabledCaptionStyle
    PropertyChanged "DisabledCaptionStyle"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledCaptionStyle Property set

'Property set for Disabled State Font
Public Property Get DisabledFont() As Font
Attribute DisabledFont.VB_Description = "Returns/sets the font to use for the caption when the button is disabled.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFont.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    Set DisabledFont = m_DisabledFont
End Property

Public Property Set DisabledFont(ByVal New_DisabledFont As Font)
On Error Resume Next
    Set m_DisabledFont = New_DisabledFont
    PropertyChanged "DisabledFont"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFont Property set

'Property set for Disabled Font: Bold state
Public Property Get DisabledFontBold() As Boolean
Attribute DisabledFontBold.VB_Description = "Returns/sets the bold attribute for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontBold.VB_MemberFlags = "400"
On Error Resume Next
    DisabledFontBold = m_DisabledFont.Bold
End Property

Public Property Let DisabledFontBold(ByVal New_DisabledFontBold As Boolean)
On Error Resume Next
    m_DisabledFont.Bold = New_DisabledFontBold
    PropertyChanged "DisabledFontBold"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontBold Property set

'Property set for Enabling/Disabling of the Disabled state font
Public Property Get DisabledFontEnabled() As Boolean
Attribute DisabledFontEnabled.VB_Description = "Returns/sets whether or not the DisabledFont is used when the button is disabled."
Attribute DisabledFontEnabled.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    DisabledFontEnabled = m_DisabledFontEnabled
End Property

Public Property Let DisabledFontEnabled(ByVal New_DisabledFontEnabled As Boolean)
On Error Resume Next
    m_DisabledFontEnabled = New_DisabledFontEnabled
    PropertyChanged "DisabledFontEnabled"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontEnabled Property set

'Property set for Disabled Font: Italic state
Public Property Get DisabledFontItalic() As Boolean
Attribute DisabledFontItalic.VB_Description = "Returns/sets the italic attribute for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontItalic.VB_MemberFlags = "400"
On Error Resume Next
    DisabledFontItalic = m_DisabledFont.Italic
End Property

Public Property Let DisabledFontItalic(ByVal New_DisabledFontItalic As Boolean)
On Error Resume Next
    m_DisabledFont.Italic = New_DisabledFontItalic
    PropertyChanged "DisabledFontItalic"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontItalic Property set

'Property set for Disabled Font: Name
Public Property Get DisabledFontName() As String
Attribute DisabledFontName.VB_Description = "Returns/sets the name of the font to use for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontName.VB_MemberFlags = "400"
On Error Resume Next
    DisabledFontName = m_DisabledFont.Name
End Property

Public Property Let DisabledFontName(ByVal New_DisabledFontName As String)
On Error Resume Next
    m_DisabledFont.Name = New_DisabledFontName
    PropertyChanged "DisabledFontName"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontName Property set

'Property set for Disabled Font: Size setting
Public Property Get DisabledFontSize() As Single
Attribute DisabledFontSize.VB_Description = "Returns/sets the size attribute for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontSize.VB_MemberFlags = "400"
On Error Resume Next
    DisabledFontSize = m_DisabledFont.Size
End Property

Public Property Let DisabledFontSize(ByVal New_DisabledFontSize As Single)
On Error Resume Next
    m_DisabledFont.Size = New_DisabledFontSize
    PropertyChanged "DisabledFontSize"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontSize Property set

'Property set for Disabled Font: Strike Through state
Public Property Get DisabledFontStrikethrough() As Boolean
Attribute DisabledFontStrikethrough.VB_Description = "Returns/sets the strike through attribute for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontStrikethrough.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontStrikethrough.VB_MemberFlags = "400"
On Error Resume Next
 DisabledFontStrikethrough = m_DisabledFont.Strikethrough
End Property

Public Property Let DisabledFontStrikethrough(ByVal New_DisabledFontStrikethrough As Boolean)
On Error Resume Next
    m_DisabledFont.Strikethrough = New_DisabledFontStrikethrough
    PropertyChanged "DisabledFontStrikethrough"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontStrikethrough Property set

'Property set for Disabled Font: Underline state
Public Property Get DisabledFontUnderline() As Boolean
Attribute DisabledFontUnderline.VB_Description = "Returns/sets the underline attribute for the disabled font.  (Only used if DisabledFontEnabled is set to True)"
Attribute DisabledFontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DisabledFontUnderline.VB_MemberFlags = "400"
On Error Resume Next
    DisabledFontUnderline = m_DisabledFont.Underline
End Property

Public Property Let DisabledFontUnderline(ByVal New_DisabledFontUnderline As Boolean)
On Error Resume Next
    m_DisabledFont.Underline = New_DisabledFontUnderline
    PropertyChanged "DisabledFontUnderline"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledFontUnderline Property set

'Property set for the Disabled state's Text color
Public Property Get DisabledForeColor() As OLE_COLOR
Attribute DisabledForeColor.VB_Description = "Returns/sets the color used for the font when the button is disabled."
Attribute DisabledForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    DisabledForeColor = m_DisabledForeColor
End Property

Public Property Let DisabledForeColor(ByVal New_DisabledForeColor As OLE_COLOR)
On Error Resume Next
    m_DisabledForeColor = New_DisabledForeColor
    PropertyChanged "DisabledForeColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast        'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledForeColor Property set

'Property set for the Disabled Mouse Icon
Public Property Get DisabledMouseIcon() As Picture
On Error Resume Next
    Set DisabledMouseIcon = m_DisabledMouseIcon
End Property

Public Property Set DisabledMouseIcon(ByVal New_DisabledMouseIcon As Picture)
On Error Resume Next
    Set m_DisabledMouseIcon = New_DisabledMouseIcon
    PropertyChanged "DisabledMouseIcon"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End DisabledMouseIcon Property set

'Property set for the visible Disabled Mouse Pointer
Public Property Get DisabledMousePointer() As MousePointerConstants
On Error Resume Next
    DisabledMousePointer = m_DisabledMousePointer
End Property

Public Property Let DisabledMousePointer(ByVal New_DisabledMousePointer As MousePointerConstants)
On Error Resume Next
    If (New_DisabledMousePointer < vbDefault) Or ((New_DisabledMousePointer > vbSizeAll) And Not (New_DisabledMousePointer = vbCustom)) Then Exit Property
    m_DisabledMousePointer = New_DisabledMousePointer
    PropertyChanged "DisabledMousePointer"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End DisabledMousePointer Property set

'Property set for the picture to display in the disabled state
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets the picture for the button to use while the button is disabled.  (Used only for Graphical,  Graphical Gradient, and Graphical Picture modes)"
Attribute DisabledPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    Set DisabledPicture = m_DisabledPicture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
On Error Resume Next
    Set m_DisabledPicture = New_DisabledPicture
    PropertyChanged "DisabledPicture"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDisable Then PaintFast    'If button is disabled then do a fast repaint to show changes
End Property
'End DisabledPicture Property set

'Property set for Down Caption Style
Public Property Get DownCaptionStyle() As gbCaptionStyle
Attribute DownCaptionStyle.VB_Description = "Returns/sets the Caption Style for the button in the depressed state.  (Standard / Light Inset / Heavy Inset / Light Raised / Heavy Raised / Drop Shadow)"
Attribute DownCaptionStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    DownCaptionStyle = m_DownCaptionStyle
End Property

Public Property Let DownCaptionStyle(ByVal New_DownCaptionStyle As gbCaptionStyle)
On Error Resume Next
    If (New_DownCaptionStyle < gbcStandard) Or (New_DownCaptionStyle > gbcDropShadow) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_DownCaptionStyle = New_DownCaptionStyle
    PropertyChanged "DownCaptionStyle"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownCaptionStyle Property set

'Property set for the Font in the Down state
Public Property Get DownFont() As Font
Attribute DownFont.VB_Description = "Returns/sets the font to use for the caption when the button is depressed.  (Only used if DownFontEnabled is set to True)"
Attribute DownFont.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    Set DownFont = m_DownFont
End Property

Public Property Set DownFont(ByVal New_DownFont As Font)
On Error Resume Next
    Set m_DownFont = New_DownFont
    PropertyChanged "DownFont"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFont Property set

'Property set for Down Font: Bold state
Public Property Get DownFontBold() As Boolean
Attribute DownFontBold.VB_Description = "Returns/sets the bold attribute for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontBold.VB_MemberFlags = "400"
On Error Resume Next
    DownFontBold = m_DownFont.Bold
End Property

Public Property Let DownFontBold(ByVal New_DownFontBold As Boolean)
On Error Resume Next
    m_DownFont.Bold = New_DownFontBold
    PropertyChanged "DownFontBold"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontBold Property set

'Property set for Enabling/Disabling of the Down state font
Public Property Get DownFontEnabled() As Boolean
Attribute DownFontEnabled.VB_Description = "Returns/sets whether or not the DownFont is used when the button is depressed."
Attribute DownFontEnabled.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    DownFontEnabled = m_DownFontEnabled
End Property

Public Property Let DownFontEnabled(ByVal New_DownFontEnabled As Boolean)
On Error Resume Next
    m_DownFontEnabled = New_DownFontEnabled
    PropertyChanged "DownFontEnabled"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontEnabled Property set

'Property set for Down Font: Italic state
Public Property Get DownFontItalic() As Boolean
Attribute DownFontItalic.VB_Description = "Returns/sets the italic attribute for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontItalic.VB_MemberFlags = "400"
On Error Resume Next
    DownFontItalic = m_DownFont.Italic
End Property

Public Property Let DownFontItalic(ByVal New_DownFontItalic As Boolean)
On Error Resume Next
    m_DownFont.Italic = New_DownFontItalic
    PropertyChanged "DownFontItalic"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontItalic Property set

'Property set for Down Font: Name
Public Property Get DownFontName() As String
Attribute DownFontName.VB_Description = "Returns/sets the name of the font to use for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontName.VB_MemberFlags = "400"
On Error Resume Next
    DownFontName = m_DownFont.Name
End Property

Public Property Let DownFontName(ByVal New_DownFontName As String)
On Error Resume Next
    m_DownFont.Name = New_DownFontName
    PropertyChanged "DownFontName"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontName Property set

'Property set for Down Font: Size setting
Public Property Get DownFontSize() As Single
Attribute DownFontSize.VB_Description = "Returns/sets the size attribute for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontSize.VB_MemberFlags = "400"
On Error Resume Next
    DownFontSize = m_DownFont.Size
End Property

Public Property Let DownFontSize(ByVal New_DownFontSize As Single)
On Error Resume Next
    m_DownFont.Size = New_DownFontSize
    PropertyChanged "DownFontSize"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontSize Property set

'Property set for Down Font: Strike through state
Public Property Get DownFontStrikethrough() As Boolean
Attribute DownFontStrikethrough.VB_Description = "Returns/sets the strike through attribute for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontStrikethrough.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontStrikethrough.VB_MemberFlags = "400"
On Error Resume Next
    DownFontStrikethrough = m_DownFont.Strikethrough
End Property

Public Property Let DownFontStrikethrough(ByVal New_DownFontStrikethrough As Boolean)
On Error Resume Next
    m_DownFont.Strikethrough = New_DownFontStrikethrough
    PropertyChanged "DownFontStrikeStrikethrough"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontStrikethrough Property set

'Property set for Down Font: Underline state
Public Property Get DownFontUnderline() As Boolean
Attribute DownFontUnderline.VB_Description = "Returns/sets the underline attribute for the down font.  (Only used if DownFontEnabled is set to True)"
Attribute DownFontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute DownFontUnderline.VB_MemberFlags = "400"
On Error Resume Next
    DownFontUnderline = m_DownFont.Underline
End Property

Public Property Let DownFontUnderline(ByVal New_DownFontUnderline As Boolean)
On Error Resume Next
    m_DownFont.Underline = New_DownFontUnderline
    PropertyChanged "DownFontUnderline"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownFontUnderline Property set

'Property set for the Down state's Text color
Public Property Get DownForeColor() As OLE_COLOR
Attribute DownForeColor.VB_Description = "Returns/sets the color used for the font when the button is depressed."
Attribute DownForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    DownForeColor = m_DownForeColor
End Property

Public Property Let DownForeColor(ByVal New_DownForeColor As OLE_COLOR)
On Error Resume Next
    m_DownForeColor = New_DownForeColor
    PropertyChanged "DownForeColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownForeColor Property set

'Property set for the Down Mouse Icon
Public Property Get DownMouseIcon() As Picture
On Error Resume Next
    Set DownMouseIcon = m_DownMouseIcon
End Property

Public Property Set DownMouseIcon(ByVal New_DownMouseIcon As Picture)
On Error Resume Next
    Set m_DownMouseIcon = New_DownMouseIcon
    PropertyChanged "DownMouseIcon"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End DownMouseIcon Property set

'Property set for the visible Down Mouse Pointer
Public Property Get DownMousePointer() As MousePointerConstants
On Error Resume Next
    DownMousePointer = m_DownMousePointer
End Property

Public Property Let DownMousePointer(ByVal New_DownMousePointer As MousePointerConstants)
On Error Resume Next
    If (New_DownMousePointer < vbDefault) Or ((New_DownMousePointer > vbSizeAll) And Not (New_DownMousePointer = vbCustom)) Then Exit Property
    m_DownMousePointer = New_DownMousePointer
    PropertyChanged "DownMousePointer"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End DownMousePointer Property set

'Property set for the picture to display in the down state
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets the picture for the button to use while the button is depressed.  (Used only for Graphical,  Graphical Gradient, and Graphical Picture modes)"
Attribute DownPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    Set DownPicture = m_DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
On Error Resume Next
    Set m_DownPicture = New_DownPicture
    PropertyChanged "DownPicture"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsDown Then PaintFast       'If button is down then do a fast repaint to show changes
End Property
'End DownPicture Property set

'Property set for the Enabling/Disabling of the button
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Enables/Disables the button and all associated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
On Error Resume Next
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next
    If New_Enabled Then     'If enabled then
        State = gbsMouseOff     'Set state to Mouse not Over
    Else                    'Not enabled
        State = gbsDisable      'Set state to disabled for correct drawing
    End If
    m_Enabled = New_Enabled       'Disable the control
    PropertyChanged "Enabled"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast               'Do a fast repaint on the control to account for different Disabled settings
End Property
'End Enabled Property set

'Property set for the Regular Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets the font to use for the caption while the mouse is not over the button or in design mode."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
On Error Resume Next
    Set m_Font = New_Font
    PropertyChanged "Font"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End Font Property set

'Property set for Font: Bold state
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets the bold attribute for the regular font."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
On Error Resume Next
    FontBold = m_Font.Bold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
On Error Resume Next
    m_Font.Bold = New_FontBold
    PropertyChanged "FontBold"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontBold Property set

'Property set for Font: Italic state
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets the italic attribute for the regular font."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
On Error Resume Next
    FontItalic = m_Font.Italic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
On Error Resume Next
    m_Font.Italic = New_FontItalic
    PropertyChanged "FontItalic"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontItalic Property set

'Property set for Font: Name
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Returns/sets the font name for the regular font."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
On Error Resume Next
    FontName = m_Font.Name
End Property

Public Property Let FontName(ByVal New_FontName As String)
On Error Resume Next
    m_Font.Name = New_FontName
    PropertyChanged "FontName"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontName Property set

'Property set for Font: Size setting
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Returns/sets the size of the regular font."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
On Error Resume Next
    FontSize = m_Font.Size
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
On Error Resume Next
    m_Font.Size = New_FontSize
    PropertyChanged "FontSize"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontSize Property set

'Property set for Font: Strike through state
Public Property Get FontStrikethrough() As Boolean
Attribute FontStrikethrough.VB_Description = "Returns/sets the strike through attribute for the regular font."
Attribute FontStrikethrough.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethrough.VB_MemberFlags = "400"
On Error Resume Next
    FontStrikethrough = m_Font.Strikethrough
End Property

Public Property Let FontStrikethrough(ByVal New_FontStrikethrough As Boolean)
On Error Resume Next
    m_Font.Strikethrough = New_FontStrikethrough
    PropertyChanged "FontStrikethrough"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontStrikethrough Property set

'Property set for Font: Underline state
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets the underline attribute for the regular font."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
On Error Resume Next
    FontUnderline = m_Font.Underline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
On Error Resume Next
    m_Font.Underline = New_FontUnderline
    PropertyChanged "FontUnderline"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the font changes
End Property
'End FontUnderline Property set

'Property set for Text Color
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute ForeColor.VB_UserMemId = -513
On Error Resume Next
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
On Error Resume Next
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast repaint to reflect the text color changes
End Property
'End ForeColor Property set

'Property set for the Angle which the Gradient is drawn at
Public Property Get GradientAngle() As Double
Attribute GradientAngle.VB_Description = "Returns/sets the angle at which the gradient should be drawn on the button.  (Used only for Gradient and Graphical Gradient modes)"
Attribute GradientAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Double)
On Error Resume Next
    New_GradientAngle = New_GradientAngle Mod 360   'Get value to within -359.999999999 to 359.999999999
    If (New_GradientAngle < 0) Then     'If angle is negative then
        'Add 360 to get it into valid range
        m_GradientAngle = 360# + New_GradientAngle
    Else
        m_GradientAngle = New_GradientAngle
    End If
    PropertyChanged "GradientAngle"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientAngle Property set

'Property set for the Gradient Blending Mode
Public Property Get GradientBlendMode() As GradBlendMode
Attribute GradientBlendMode.VB_Description = "Returns/sets the blending mode to use with the gradient modes."
Attribute GradientBlendMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientBlendMode = m_GradientBlendMode
End Property

Public Property Let GradientBlendMode(ByVal New_GradientBlendMode As GradBlendMode)
On Error Resume Next
    If (New_GradientBlendMode < gbmRGB) Or (New_GradientBlendMode > gbmHSL) Then Exit Property  'Make sure we have a valid value
    m_GradientBlendMode = New_GradientBlendMode
    PropertyChanged "GradientBlendMode"
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientBlendMode Property set

'Property set for the First Color of the Gradient
Public Property Get GradientColor1() As OLE_COLOR
Attribute GradientColor1.VB_Description = "Returns/sets the start color of the gradient.  (Used only for Gradient and Graphical Gradient modes)"
Attribute GradientColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
On Error Resume Next
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientColor1 Property set

'Property set for the Second Color of the Gradient
Public Property Get GradientColor2() As OLE_COLOR
Attribute GradientColor2.VB_Description = "Returns/sets the end color of the gradient.  (Used only for Gradient and Graphical Gradient Modes)"
Attribute GradientColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
On Error Resume Next
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientColor2 Property set

'Property set for number of times to repeat the gradient
Public Property Get GradientRepetitions() As Double
Attribute GradientRepetitions.VB_Description = "Returns/sets the number times to show the gradient.  Valid values are 1 to 45.  (Used only for Gradient and Graphical Gradient Modes)"
Attribute GradientRepetitions.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientRepetitions = m_GradientRepetitions
End Property

Public Property Let GradientRepetitions(ByVal New_GradientRepetitions As Double)
On Error Resume Next
    If (New_GradientRepetitions < 1) Or (New_GradientRepetitions > 45) Then Exit Property   'Make sure within valid range
    m_GradientRepetitions = New_GradientRepetitions
    PropertyChanged "GradientRepetitions"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientRepetitions Property set

'Property set for the type of gradient
Public Property Get GradientType() As GradType
Attribute GradientType.VB_Description = "Returns/sets the type of Gradient to Draw (Normal / Elliptical / Rectangular)"
Attribute GradientType.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    GradientType = m_GradientType
End Property

Public Property Let GradientType(ByVal New_GradientType As GradType)
On Error Resume Next
    If (New_GradientType < gtNormal) Or (New_GradientType > gtRectangular) Then Exit Property   'Make sure we have a valid value
    m_GradientType = New_GradientType
    PropertyChanged "GradientType"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientType Property set

'Read Only property of the Control's Device Context
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns the current device context of the button."
Attribute hdc.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute hdc.VB_MemberFlags = "400"
On Error Resume Next
    hdc = UserControl.hdc
End Property

'Property set for Hover Caption Style
Public Property Get HoverCaptionStyle() As gbCaptionStyle
Attribute HoverCaptionStyle.VB_Description = "Returns/sets the Caption Style for the button while the mouse is over the button.  (Only if UseHover is set to True) (Standard / Light Inset / Heavy Inset / Light Raised / Heavy Raised / Drop Shadow)"
Attribute HoverCaptionStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    HoverCaptionStyle = m_HoverCaptionStyle
End Property

Public Property Let HoverCaptionStyle(ByVal New_HoverCaptionStyle As gbCaptionStyle)
On Error Resume Next
    If (New_HoverCaptionStyle < gbcStandard) Or (New_HoverCaptionStyle > gbcDropShadow) Then Exit Property  'If new value isn't valid then exit the property doing nothing
    m_HoverCaptionStyle = New_HoverCaptionStyle
    PropertyChanged "HoverCaptionStyle"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverCaptionStyle Property set

'Property set for the Font in the Hover state
Public Property Get HoverFont() As Font
Attribute HoverFont.VB_Description = "Returns/sets the font to use for the caption when the Mouse is over the button.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFont.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    Set HoverFont = m_HoverFont
End Property

Public Property Set HoverFont(ByVal New_HoverFont As Font)
On Error Resume Next
    Set m_HoverFont = New_HoverFont
    PropertyChanged "HoverFont"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFont Property set

'Property set for Hover Font: Bold state
Public Property Get HoverFontBold() As Boolean
Attribute HoverFontBold.VB_Description = "Returns/sets the bold attribute for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontBold.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontBold = m_HoverFont.Bold
End Property

Public Property Let HoverFontBold(ByVal New_HoverFontBold As Boolean)
On Error Resume Next
    m_HoverFont.Bold = New_HoverFontBold
    PropertyChanged "HoverFontBold"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontBold Property set

'Property set for Enabling/Disabling of the Hover state font
Public Property Get HoverFontEnabled() As Boolean
Attribute HoverFontEnabled.VB_Description = "Returns/sets whether or not the HoverFont is used when the mouse is over the button.  (Only applicable if UseHover is set to True)"
Attribute HoverFontEnabled.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    HoverFontEnabled = m_HoverFontEnabled
End Property

Public Property Let HoverFontEnabled(ByVal New_HoverFontEnabled As Boolean)
On Error Resume Next
    m_HoverFontEnabled = New_HoverFontEnabled
    PropertyChanged "HoverFontEnabled"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontEnabled Property set

'Property set for Hover Font: Italic state
Public Property Get HoverFontItalic() As Boolean
Attribute HoverFontItalic.VB_Description = "Returns/sets the italic attribute for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontItalic.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontItalic = m_HoverFont.Italic
End Property

Public Property Let HoverFontItalic(ByVal New_HoverFontItalic As Boolean)
On Error Resume Next
    m_HoverFont.Italic = New_HoverFontItalic
    PropertyChanged "HoverFontItalic"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontItalic Property set

'Property set for Hover Font: Name
Public Property Get HoverFontName() As String
Attribute HoverFontName.VB_Description = "Returns/sets the name of the font to use for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontName.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontName = m_HoverFont.Name
End Property

Public Property Let HoverFontName(ByVal New_HoverFontName As String)
On Error Resume Next
    m_HoverFont.Name = New_HoverFontName
    PropertyChanged "HoverFontName"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontName Property set

'Property set for Hover Font: Size setting
Public Property Get HoverFontSize() As Single
Attribute HoverFontSize.VB_Description = "Returns/sets the size attribute for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontSize.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontSize = m_HoverFont.Size
End Property

Public Property Let HoverFontSize(ByVal New_HoverFontSize As Single)
On Error Resume Next
    m_HoverFont.Size = New_HoverFontSize
    PropertyChanged "HoverFontSize"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontSize Property set

'Property set for Hover Font: Strike through state
Public Property Get HoverFontStrikethrough() As Boolean
Attribute HoverFontStrikethrough.VB_Description = "Returns/sets the strike through attribute for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontStrikethrough.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontStrikethrough.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontStrikethrough = m_HoverFont.Strikethrough
End Property

Public Property Let HoverFontStrikethrough(ByVal New_HoverFontStrikethrough As Boolean)
On Error Resume Next
    m_HoverFont.Strikethrough = New_HoverFontStrikethrough
    PropertyChanged "HoverFontStrikeStrikethrough"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontStrikethrough Property set

'Property set for Hover Font: Underline state
Public Property Get HoverFontUnderline() As Boolean
Attribute HoverFontUnderline.VB_Description = "Returns/sets the underline attribute for the hover font.  (Only used if HoverFontEnabled is set to True and UseHover is set to True)"
Attribute HoverFontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute HoverFontUnderline.VB_MemberFlags = "400"
On Error Resume Next
    HoverFontUnderline = m_HoverFont.Underline
End Property

Public Property Let HoverFontUnderline(ByVal New_HoverFontUnderline As Boolean)
On Error Resume Next
    m_HoverFont.Underline = New_HoverFontUnderline
    PropertyChanged "HoverFontUnderline"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverFontUnderline Property set

'Property set for the Hover state's Text color
Public Property Get HoverForeColor() As OLE_COLOR
Attribute HoverForeColor.VB_Description = "Returns/sets the color used for the font when the mouse is over the button.  (Used only if UseHover is set to True)"
Attribute HoverForeColor.VB_ProcData.VB_Invoke_Property = ";Font"
On Error Resume Next
    HoverForeColor = m_HoverForeColor
End Property

Public Property Let HoverForeColor(ByVal New_HoverForeColor As OLE_COLOR)
On Error Resume Next
    m_HoverForeColor = New_HoverForeColor
    PropertyChanged "HoverForeColor"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverForeColor Property set

'Property set for the Hovermode to use for the control
Public Property Get HoverMode() As gbHoverMode
Attribute HoverMode.VB_Description = "Returns/sets the parts of the control to hover."
Attribute HoverMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    HoverMode = m_HoverMode
End Property

Public Property Let HoverMode(ByVal New_HoverMode As gbHoverMode)
On Error Resume Next
    m_HoverMode = New_HoverMode
    PropertyChanged "HoverMode"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint to update the control to match the new value
End Property
'End HoverMode Property set

'Property set for the Hover Mouse Icon
Public Property Get HoverMouseIcon() As Picture
On Error Resume Next
    Set HoverMouseIcon = m_HoverMouseIcon
End Property

Public Property Set HoverMouseIcon(ByVal New_HoverMouseIcon As Picture)
On Error Resume Next
    Set m_HoverMouseIcon = New_HoverMouseIcon
    PropertyChanged "HoverMouseIcon"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End HoverMouseIcon Property set

'Property set for the visible Hover Mouse Pointer
Public Property Get HoverMousePointer() As MousePointerConstants
On Error Resume Next
    HoverMousePointer = m_HoverMousePointer
End Property

Public Property Let HoverMousePointer(ByVal New_HoverMousePointer As MousePointerConstants)
On Error Resume Next
    If (New_HoverMousePointer < vbDefault) Or ((New_HoverMousePointer > vbSizeAll) And Not (New_HoverMousePointer = vbCustom)) Then Exit Property
    m_HoverMousePointer = New_HoverMousePointer
    PropertyChanged "HoverMousePointer"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End HoverMousePointer Property set

'Property set for the picture to display in the hover state
Public Property Get HoverPicture() As Picture
Attribute HoverPicture.VB_Description = "Returns/sets the picture for the button to use while the mouse is over the button.  (Used only for Graphical,  Graphical Gradient, and Graphical Picture modes, Only used if UseHover is set to True)"
Attribute HoverPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    Set HoverPicture = m_HoverPicture
End Property

Public Property Set HoverPicture(ByVal New_HoverPicture As Picture)
On Error Resume Next
    Set m_HoverPicture = New_HoverPicture
    PropertyChanged "HoverPicture"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If State = gbsMouseOver Then PaintFast      'If Mouse cursor is over then do a fast paint to reflect changes
End Property
'End HoverPicture Property set

'Read Only Property of the control's Window Handle
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the current window handle of the button."
Attribute hwnd.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute hwnd.VB_UserMemId = -515
Attribute hwnd.VB_MemberFlags = "400"
On Error Resume Next
    hwnd = UserControl.hwnd
End Property

'Property set for the Picture Mask Color
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color to consider a transparency when painting pictures onto the button face.  (Used only in Graphical and Graphical Gradient modes)"
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
On Error Resume Next
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint just in case there would be any visual changes
End Property
'End MaskColor Property set

'Property set for the Mouse Icon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets the picture to use for the mouse while it is in the area over the button."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
On Error Resume Next
    Set m_MouseIcon = New_MouseIcon       'Set the Control's Mouse Icon (Change will be automatic)
    PropertyChanged "MouseIcon"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End MouseIcon Property set

'Property set for the visible Mouse Pointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the Mouse Pointer to use while the mouse is over the button."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
On Error Resume Next
    If (New_MousePointer < vbDefault) Or ((New_MousePointer > vbSizeAll) And Not (New_MousePointer = vbCustom)) Then Exit Property
    m_MousePointer = New_MousePointer   'Set the Control's Mouse Pointer (Change will be automatic)
    PropertyChanged "MousePointer"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    'SetMousePointer
End Property
'End MousePointer Property set

'Sub-procedure to start an OLE Drag operation on the Button
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Method to initiate an OLE drag operation."
On Error Resume Next
    UserControl.OLEDrag
End Sub

'Property set for OLE Drop Mode when dealing with Drag/Drop Operations
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets the way the button will behave when an OLE object is dropped on it."
Attribute OLEDropMode.VB_ProcData.VB_Invoke_Property = ";Data"
On Error Resume Next
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
On Error Resume Next
    UserControl.OLEDropMode() = New_OLEDropMode     'Set the Control's Drop Mode
    PropertyChanged "OLEDropMode"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End OLEDropMode Property set

'Property set for the standard picture to use with graphical modes
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets the picture for the button to use while the mouse is not over the button or in design mode.  (Used only for Graphical,  Graphical Gradient, and Graphical Picture modes)"
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
On Error Resume Next
    Set m_Picture = New_Picture
    PropertyChanged "Picture"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint so as to update the control with the changes
End Property
'End Picture Property set

'Property set for the standard picture alignment to use with graphical modes
Public Property Get PictureAlignment() As gbAlignment
Attribute PictureAlignment.VB_Description = "Returns/sets the alignment of the fore ground picture(s).  (Used only for Graphical modes) [LeftTop / LeftMiddle / LeftBottom / RightTop / RightMiddle / RightBottom / CenterTop / CenterMiddle / CenterBottom]"
Attribute PictureAlignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    PictureAlignment = m_PictureAlignment
End Property

Public Property Let PictureAlignment(ByVal New_PictureAlignment As gbAlignment)
On Error Resume Next
    If (New_PictureAlignment < gbaLeftTop) Or (New_PictureAlignment > gbaCenterBottom) Then Exit Property   'Make sure we have a valid alignment
    m_PictureAlignment = New_PictureAlignment
    PropertyChanged "PictureAlignment"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint so as to update the control with the changes
End Property
'End PictureAlignment Property set

'Property set for the standard picture cushion to use with graphical modes
Public Property Get PictureCushion() As Long
Attribute PictureCushion.VB_Description = "Returns/sets the number of pixels between the Picture and the edge of the button as a cushion zone."
Attribute PictureCushion.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    PictureCushion = m_PictureCushion
End Property

Public Property Let PictureCushion(ByVal New_PictureCushion As Long)
On Error Resume Next
    If New_PictureCushion < 0 Then Exit Property    'Make sure we have a valid cushion
    m_PictureCushion = New_PictureCushion
    PropertyChanged "PictureCushion"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint so as to update the control with the changes
End Property
'End PictureCushion Property set

'Sub-procedure to allow the user to force a full repaint on the control
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
On Error Resume Next
    PaintAll
End Sub

'Property set for the Style of the button (Standard/Graphical/Gradient/Graphical Gradient)
Public Property Get style() As gbStyle
Attribute style.VB_Description = "Returns/sets the display style of the button.  (Standard / Graphical / Gradient / Graphical Gradient / Picture / Graphical Picture)"
Attribute style.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    style = m_Style
End Property

Public Property Let style(ByVal New_Style As gbStyle)
On Error Resume Next
    If (New_Style < gbsStandard) Or (New_Style > gbsGraphicalPicture) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_Style = New_Style
    PropertyChanged "Style"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all with the new style to make sure all colors and other settings are correct
End Property
'End Style Property set

'Property set for the ToolTip Text
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text to use for the control's ToolTip."
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    If bToolTipAvailable Then   'If we have extender info then
        ToolTipText = Extender.ToolTipText  'Get the Current ToolTip Text from VB
    Else
        ToolTipText = m_ToolTipText
    End If
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
On Error Resume Next
    If bToolTipAvailable Then   'If we have extender info then
        Extender.ToolTipText = New_ToolTipText  'Assign the New ToolTipText back to VB's Extender
    Else
        m_ToolTipText = New_ToolTipText
    End If
    PropertyChanged "ToolTipText"   'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End ToolTipText Property set

'Property set for Using the classic border styles or not (Does not apply to Bevel)
Public Property Get UseClassicBorders() As Boolean
Attribute UseClassicBorders.VB_Description = "Returns/sets whether or not Classic Broders are to be used instead of the newer blending borders."
Attribute UseClassicBorders.VB_ProcData.VB_Invoke_Property = ";Appearance"
On Error Resume Next
    UseClassicBorders = m_UseClassicBorders
End Property

Public Property Let UseClassicBorders(ByVal New_UseClassicBorders As Boolean)
On Error Resume Next
    m_UseClassicBorders = New_UseClassicBorders
    PropertyChanged "UseClassicBorders"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint because it may change the current display
End Property
'End UseClassicBorders Property set

'Property set for Turning on/off Hover Mode
Public Property Get UseHover() As Boolean
Attribute UseHover.VB_Description = "Returns/sets whether the mouse will change it's appearance when the mouse is over the button or not."
Attribute UseHover.VB_ProcData.VB_Invoke_Property = ";Behavior"
On Error Resume Next
    UseHover = m_UseHover
End Property

Public Property Let UseHover(ByVal New_UseHover As Boolean)
On Error Resume Next
    m_UseHover = New_UseHover
    PropertyChanged "UseHover"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint because it may change the current display
End Property
'End UseHover Property set

'Property set for State/Option button Value setting
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value of the button True / False (Up or Down, Only used if ButtonType is either StateButton or OptionButton)"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
On Error Resume Next
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
On Error Resume Next
    Dim Changed As Boolean      'Variable to temporarily hold whether the Value of the button has changed or not

    If Not bInitiating Then     'If we didn't initiate an Option button reset then
        Changed = False     'Initialize temporary holder just to make sure we don't have any slip-ups
        If m_Value <> New_Value Then Changed = True     'If value is different then set flag signaling that value has changed
        If (New_Value) And (m_ButtonType = gbtOptionButton) Then    'If the new value is True and we are an Option button then
            ResetAll        'Call the reset procedure to reset the values of the other option buttons within this group.
        End If
        m_Value = New_Value     'Make the assignment whether value has changed or not in case of unforseen error
        PropertyChanged "Value"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
        If (m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton) Then PaintFast 'If in state or option button mode then do a fast paint to update appearance
        If (Changed) And ((m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton)) And (m_Enabled) Then RaiseEvent ValueChanged(m_Value) 'If changed flag is set and in State Button mode then raise the ValueChanged event.
    End If
End Property
'End Value Property set



'**************************************************************************
'Internal Event Coding & Implementation
'**************************************************************************

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    UserControl.SetFocus    'Make sure that this button has focus from the access key press.
    If (m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton) Then ChangeValue: PaintFast  'If in State or Option button mode then change button's value then perform a fast paint to update control
    RaiseEvent Click        'Pass the access key through as a click event to the developer
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
On Error Resume Next
    PaintAll            'Do a full RePaint just in case anything unusual changed (though most likely just Focus)
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    Set grad = New Gradient     'Initialize the Gradient object along with the control
    State = gbsMouseOff         'Initialize the control's state
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    'Load the default values (special values may have added line comments)
    bInitializing = True
    m_Alignment = m_def_Alignment
    m_AlignmentCushion = m_def_AlignmentCushion
    m_Appearance = m_def_Appearance
    m_AutoSize = m_def_AutoSize
    m_BevelIntensity = m_def_BevelIntensity
    m_BevelWidth = m_def_BevelWidth
    m_BorderColor = m_def_BorderColor
    m_ButtonType = m_def_ButtonType
    m_Caption = Ambient.DisplayName             'Pull the correct value from VB as to what it should be named initially
    m_CaptionStyle = m_def_CaptionStyle
    m_DisabledCaptionStyle = m_def_DisabledCaptionStyle
    Set m_DisabledFont = Ambient.Font           'Pull the ambient font to use as the default font
    m_DisabledFontEnabled = m_def_DisabledFontEnabled
    m_DisabledForeColor = m_def_DisabledForeColor
    Set m_DisabledMouseIcon = Nothing
    m_DisabledMousePointer = m_def_DisabledMousePointer
    Set m_DisabledPicture = Nothing
    m_DownCaptionStyle = m_def_DownCaptionStyle
    Set m_DownFont = Ambient.Font               'Pull the ambient font to use as the default font
    m_DownFontEnabled = m_def_DownFontEnabled
    m_DownForeColor = m_def_DownForeColor
    Set m_DownMouseIcon = Nothing
    m_DownMousePointer = m_def_DownMousePointer
    Set m_DownPicture = Nothing
    m_Enabled = m_def_Enabled
    m_ForeColor = m_def_ForeColor
    m_GradientAngle = m_def_GradientAngle
    m_GradientBlendMode = m_def_GradientBlendMode
    m_GradientColor1 = m_def_GradientColor1
    m_GradientColor2 = m_def_GradientColor2
    m_GradientRepetitions = m_def_GradientRepetitions
    m_GradientType = m_def_GradientType
    Set m_Font = Ambient.Font                   'Pull the ambient font to use as the default font
    m_HoverCaptionStyle = m_def_HoverCaptionStyle
    Set m_HoverFont = Ambient.Font              'Pull the ambient font to use as the default font
    m_HoverFontEnabled = m_def_HoverFontEnabled
    m_HoverForeColor = m_def_HoverForeColor
    m_HoverMode = m_def_HoverMode
    Set m_HoverMouseIcon = Nothing
    m_HoverMousePointer = m_def_HoverMousePointer
    Set m_HoverPicture = Nothing
    Set m_MouseIcon = Nothing
    m_MousePointer = m_def_MousePointer
    Set m_Picture = Nothing
    m_PictureAlignment = m_def_PictureAlignment
    m_PictureCushion = m_def_PictureCushion
    m_Style = m_def_Style
    m_ToolTipText = m_def_ToolTipText
    m_UseClassicBorders = m_def_UseClassicBorders
    m_UseHover = m_def_UseHover
    m_Value = m_def_Value
    Height = DefHeight      'This should give an actual default height of 525 twips
    Width = DefWidth        'This should give an actual default width of 1245 twips
    bInside = False         'Assume that Mouse is not over the control, if it is over it will be captured correctly on move
    State = gbsMouseOff     'Set the state assuming that the mouse is not over the control
    CheckParent             'Check to make sure we have parent information available
    CheckToolTip            'Check to make sure we have ToolTip information from the extender.
    bInitializing = False
    PaintAll                'Do a first paint of the control
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    If KeyCode = 32 Then        'If the user pressed spacebar then
        State = gbsDown         'Set the control state as depressed
        PaintFast               'Do a fast paint to reflect new state
    End If
    RaiseEvent KeyDown(KeyCode, Shift)      'Pass the event along to the developer
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    RaiseEvent KeyPress(KeyAscii)       'Pass the event along to the developer
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
                                Shift As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    If KeyCode = 32 Then        'If the user pressed spacebar then
        UserControl.SetFocus    'Give the control focus
        State = gbsMouseOff     'Not necessarily using mouse
        If (m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton) Then ChangeValue 'If in State or Option button mode then change button's value
        PaintFast               'Do a fast paint to show changed state
        RaiseEvent Click        'Raise the click event
    End If
    RaiseEvent KeyUp(KeyCode, Shift)        'Pass the event along to the developer
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
'    ReleaseCapture              'Release the capture on the mouse so that VB can temporarily take control and resolve any possible mouse ownership problems
    SetCapture UserControl.hwnd 'Capture the mouse to the control so as to track it's movement
    If (Button And vbLeftButton) Then    'If the Left Mouse button is pressed then
        State = gbsDown         'Set the state to depressed
        PaintFast               'Re-Paint the control to relect the new state
    End If
    'Raise an event to the Calling program as well as rescale the co-ordinates to match the calling program
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode))
    Else
        RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips))
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
On Error Resume Next
    Dim PassX As Single     'Temporary variable to hold the translated X co-ordinate
    Dim PassY As Single     'Temporary variable to hold the translated Y co-ordinate
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    If bParentAvailable Then    'If parent information is available then
        PassX = ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode)      'Translate the X co-ordinate to match the parent's scale mode
        PassY = ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode)      'Translate the Y co-ordinate to match the parent's scale mode
    Else
        PassX = ScaleX(X, ScaleMode, vbTwips)   'Translate the X co-ordinate to twips
        PassY = ScaleY(Y, ScaleMode, vbTwips)   'Translate the Y co-ordinate to twips
    End If
    If (Me.ToolTipText <> tiToolTip.TipText) Then     'If the ToolTipText Doesn't match what is currently set then
        ModifyToolTipText       'Update the ToolTip's Text
    End If
    'If the mouse has left the region of the control then
    If (X < 0) Or (X > ScaleWidth) Or (Y < 0) Or (Y > ScaleHeight) Then
        bInside = False         'Mouse is no long in the control area
        ReleaseCapture          'Release this control capture on it
        State = gbsMouseOff     'Change the state to reflect the mouse has left
        PaintFast               'Do a fast paint on the control to reflect the new state
        RaiseEvent MouseExit    'Raise the Mouse Exit Event
        Exit Sub
    ElseIf bInside = False Then     'Mouse is still in the control area, so If it wasn't before then
        bInside = True                  'Set the variable to say the mouse is in the control area
        ReleaseCapture          'Release this control capture on it
        SetCapture UserControl.hwnd     'Make sure that we own the movement of the mouse
        If (Button And &H1) Then        'If the Left mouse button is down then
            State = gbsDown             'change the button state to depressed
        Else                            'Otherwise
            State = gbsMouseOver        'change the state to hover
        End If
        PaintFast                   'Do a fast paint to update the button with the new state
        RaiseEvent MouseEnter       'Raise the Mouse Enter event
    End If
    RaiseEvent MouseMove(Button, Shift, PassX, PassY)   'Now raise the MouseMove event
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
On Error Resume Next
    Dim PassX As Single     'Temporary variable to hold the translated X co-ordinate
    Dim PassY As Single     'Temporary variable to hold the translated Y co-ordinate
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    ReleaseCapture
    If bParentAvailable Then    'If parent information is available then
        PassX = ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode)  'Translate the X co-ordinate to match the parent's scale mode
        PassY = ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode)  'Translate the Y co-ordinate to match the parent's scale mode
    Else
        PassX = ScaleX(X, ScaleMode, vbTwips)   'Translate the X co-ordinate to twips
        PassY = ScaleY(Y, ScaleMode, vbTwips)   'Translate the Y co-ordinate to twips
    End If
'    SetCapture UserControl.hwnd     'Capture the mouse movement to this control
    If (Button And &H1) Then        'If Left mouse button is down then
'        UserControl.SetFocus        'Give this control focus
        State = gbsMouseOver        'Change the current state to Mouse Over
        If (m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton) Then ChangeValue 'If in State or Option button mode then change button's value
        PaintFast                   'Do a fast paint to update the control's display
        RaiseEvent Click            'Raise the click event because the user finalized his/her decision on this control
    End If
    RaiseEvent MouseUp(Button, Shift, PassX, PassY)     'Raise the MouseUp event
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    RaiseEvent OLECompleteDrag(Effect)      'Pass the event along to the developer
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode))     'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    Else
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips))   'Pass the event along to the developer (After rescale to twips)
    End If
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single, _
                                State As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode), State)      'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    Else
        RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips), State)    'Pass the event along to the developer (After rescale to twips)
    End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, _
                                DefaultCursors As Boolean)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)      'Pass the event along to the developer
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, _
                                DataFormat As Integer)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    RaiseEvent OLESetData(Data, DataFormat)     'Pass the event along to the developer
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, _
                                AllowedEffects As Long)
On Error Resume Next
    If Not (m_Enabled) Then Exit Sub    'If control is disabled then don't process event
    RaiseEvent OLEStartDrag(Data, AllowedEffects)       'Pass the event along to the developer
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next        'If an error occurs skip that line and continue on
    'Read properties from property bag (Special cases will have additional line comments)
    m_Alignment = PropBag.ReadProperty("Alignment", gbaCenterMiddle)
    m_AlignmentCushion = PropBag.ReadProperty("AlignmentCushion", m_def_AlignmentCushion)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    picGradient.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)     'Default is System's Button Face color
    Set m_BackPicture = PropBag.ReadProperty("BackPicture", Nothing)
    m_BevelIntensity = PropBag.ReadProperty("BevelIntensity", m_def_BevelIntensity)
    m_BevelWidth = PropBag.ReadProperty("BevelWidth", m_def_BevelWidth)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_ButtonType = PropBag.ReadProperty("ButtonType", m_def_ButtonType)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    SetAccessKeys       'Set the access keys for the control
    m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    m_DisabledCaptionStyle = PropBag.ReadProperty("DisabledCaptionStyle", m_def_DisabledCaptionStyle)
    Set m_DisabledFont = PropBag.ReadProperty("DisabledFont", Ambient.Font)     'Default is the Default Font
    m_DisabledFontEnabled = PropBag.ReadProperty("DisabledFontEnabled", m_def_DisabledFontEnabled)
    m_DisabledForeColor = PropBag.ReadProperty("DisabledForeColor", m_def_DisabledForeColor)
    Set m_DisabledMouseIcon = PropBag.ReadProperty("DisabledMouseIcon", Nothing)
    m_DisabledMousePointer = PropBag.ReadProperty("DisabledMousePointer", m_def_DisabledMousePointer)
    Set m_DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    m_DownCaptionStyle = PropBag.ReadProperty("DownCaptionStyle", m_def_DownCaptionStyle)
    Set m_DownFont = PropBag.ReadProperty("DownFont", Ambient.Font)     'Default is the Default Font
    m_DownFontEnabled = PropBag.ReadProperty("DownFontEnabled", m_def_DownFontEnabled)
    m_DownForeColor = PropBag.ReadProperty("DownForeColor", m_def_DownForeColor)
    Set m_DownMouseIcon = PropBag.ReadProperty("DownMouseIcon", Nothing)
    m_DownMousePointer = PropBag.ReadProperty("DownMousePointer", m_def_DownMousePointer)
    Set m_DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Me.Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)     'Default is the Default Font
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_GradientAngle = PropBag.ReadProperty("GradientAngle", m_def_GradientAngle)
    m_GradientBlendMode = PropBag.ReadProperty("GradientBlendMode", m_def_GradientBlendMode)
    m_GradientColor1 = PropBag.ReadProperty("GradientColor1", m_def_GradientColor1)
    m_GradientColor2 = PropBag.ReadProperty("GradientColor2", m_def_GradientColor2)
    m_GradientRepetitions = PropBag.ReadProperty("GradientRepetitions", m_def_GradientRepetitions)
    m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)
    m_HoverCaptionStyle = PropBag.ReadProperty("HoverCaptionStyle", m_def_HoverCaptionStyle)
    Set m_HoverFont = PropBag.ReadProperty("HoverFont", Ambient.Font)       'Default is the Default Font
    m_HoverFontEnabled = PropBag.ReadProperty("HoverFontEnabled", m_def_HoverFontEnabled)
    m_HoverForeColor = PropBag.ReadProperty("HoverForeColor", m_def_HoverForeColor)
    m_HoverMode = PropBag.ReadProperty("HoverMode", m_def_HoverMode)
    Set m_HoverMouseIcon = PropBag.ReadProperty("HoverMouseIcon", Nothing)
    m_HoverMousePointer = PropBag.ReadProperty("HoverMousePointer", m_def_HoverMousePointer)
    Set m_HoverPicture = PropBag.ReadProperty("HoverPicture", Nothing)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", vbButtonFace) 'Default is System's Button Face color
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", vbOLEDropNone)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_PictureAlignment = PropBag.ReadProperty("PictureAlignment", m_def_PictureAlignment)
    m_PictureCushion = PropBag.ReadProperty("PictureCushion", m_def_PictureCushion)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_UseClassicBorders = PropBag.ReadProperty("UseClassicBorders", m_def_UseClassicBorders)
    m_UseHover = PropBag.ReadProperty("UseHover", m_def_UseHover)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    CheckParent             'Check to make sure we have parent information available
    CheckToolTip            'Check to make sure we have ToolTip information from the extender.
    PaintAll        'Do a paint all to update all of the control's display
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    If Not (bAutoSizing) Then
        If ((m_Style = gbsPicture) Or (m_Style = gbsGraphicalPicture)) And (m_AutoSize = gbasControlToPicture) Then  'If it is a Picture mode and we are autosizing control to picture then
            If Not (m_BackPicture Is Nothing) Then
                bAutoSizing = True
                Width = ScaleX(m_BackPicture.Width, vbHimetric, vbTwips)
                Height = ScaleY(m_BackPicture.Height, vbHimetric, vbTwips)
                bAutoSizing = False
            End If
        End If
        'We need to execute a resize to make sure the control looks correct
        picGradient.Height = ScaleHeight    'Make the Picturebox the same size as the control so that when the gradient is draw it is drawn at the correct size
        picGradient.Width = ScaleWidth
        ResizeToolTip       'Resize the ToolTip to match the new Button Size
        If Not bInitializing Then   'If we're not initializing the control then
            PaintAll        'Do a paint all on the control to resize the caption as well as redraw any necessary borders and positioning of the picture (if any)
        End If
    End If
End Sub

Private Sub UserControl_Show()
On Error Resume Next
    CreateToolTip       'Create the ToolTip Control
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
    Set grad = Nothing      'Release the reference to the gradient class
    ReleaseCapture          'Release any capture the control had on the mouse
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next        'If an error occurs skip that line and continue on
    'Write all of the properties to the property bag for retrieval later
    Call PropBag.WriteProperty("Alignment", m_Alignment, gbaCenterMiddle)
    Call PropBag.WriteProperty("AlignmentCushion", m_AlignmentCushion, m_def_AlignmentCushion)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("BackColor", picGradient.BackColor, vbButtonFace)
    Call PropBag.WriteProperty("BackPicture", m_BackPicture, Nothing)
    Call PropBag.WriteProperty("BevelIntensity", m_BevelIntensity, m_def_BevelIntensity)
    Call PropBag.WriteProperty("BevelWidth", m_BevelWidth, m_def_BevelWidth)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("ButtonType", m_ButtonType, m_def_ButtonType)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("DisabledCaptionStyle", m_DisabledCaptionStyle, m_def_DisabledCaptionStyle)
    Call PropBag.WriteProperty("DisabledFont", m_DisabledFont, Ambient.Font)
    Call PropBag.WriteProperty("DisabledFontEnabled", m_DisabledFontEnabled, m_def_DisabledFontEnabled)
    Call PropBag.WriteProperty("DisabledForeColor", m_DisabledForeColor, m_def_DisabledForeColor)
    Call PropBag.WriteProperty("DisabledMouseIcon", m_DisabledMouseIcon, Nothing)
    Call PropBag.WriteProperty("DisabledMousePointer", m_DisabledMousePointer, m_def_DisabledMousePointer)
    Call PropBag.WriteProperty("DisabledPicture", m_DisabledPicture, Nothing)
    Call PropBag.WriteProperty("DownCaptionStyle", m_DownCaptionStyle, m_def_DownCaptionStyle)
    Call PropBag.WriteProperty("DownFont", m_DownFont, Ambient.Font)
    Call PropBag.WriteProperty("DownFontEnabled", m_DownFontEnabled, m_def_DownFontEnabled)
    Call PropBag.WriteProperty("DownForeColor", m_DownForeColor, m_def_DownForeColor)
    Call PropBag.WriteProperty("DownMouseIcon", m_DownMouseIcon, Nothing)
    Call PropBag.WriteProperty("DownMousePointer", m_DownMousePointer, m_def_DownMousePointer)
    Call PropBag.WriteProperty("DownPicture", m_DownPicture, Nothing)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("GradientAngle", m_GradientAngle, m_def_GradientAngle)
    Call PropBag.WriteProperty("GradientBlendMode", m_GradientBlendMode, m_def_GradientBlendMode)
    Call PropBag.WriteProperty("GradientColor1", m_GradientColor1, m_def_GradientColor1)
    Call PropBag.WriteProperty("GradientColor2", m_GradientColor2, m_def_GradientColor2)
    Call PropBag.WriteProperty("GradientRepetitions", m_GradientRepetitions, m_def_GradientRepetitions)
    Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)
    Call PropBag.WriteProperty("HoverCaptionStyle", m_HoverCaptionStyle, m_def_HoverCaptionStyle)
    Call PropBag.WriteProperty("HoverFont", m_HoverFont, Ambient.Font)
    Call PropBag.WriteProperty("HoverFontEnabled", m_HoverFontEnabled, m_def_HoverFontEnabled)
    Call PropBag.WriteProperty("HoverForeColor", m_HoverForeColor, m_def_HoverForeColor)
    Call PropBag.WriteProperty("HoverMode", m_HoverMode, m_def_HoverMode)
    Call PropBag.WriteProperty("HoverMouseIcon", m_HoverMouseIcon, Nothing)
    Call PropBag.WriteProperty("HoverMousePointer", m_HoverMousePointer, m_def_HoverMousePointer)
    Call PropBag.WriteProperty("HoverPicture", m_HoverPicture, Nothing)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, vbButtonFace)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, vbOLEDropNone)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureAlignment", m_PictureAlignment, m_def_PictureAlignment)
    Call PropBag.WriteProperty("PictureCushion", m_PictureCushion, m_def_PictureCushion)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("UseClassicBorders", m_UseClassicBorders, m_def_UseClassicBorders)
    Call PropBag.WriteProperty("UseHover", m_UseHover, m_def_UseHover)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

'**************************************************************************
'Internal Routines
'**************************************************************************

Private Function AdjustColor(ByVal RGBColor As Long, _
                            ByVal Amount As Long) As Long
On Error Resume Next
    Dim Blue As Long    'Variable to hold the Blue Value while in this procedure
    Dim Green As Long   'Variable to hold the Green Value while in this procedure
    Dim Red As Long     'Variable to hold the Red Value while in this procedure

    If (RGBColor = (RGBColor Or &H80000000)) Then RGBColor = GetSysColor(RGBColor Xor &H80000000)   'If working with a system color get the RGB equivilent
    RGBColor = Abs(RGBColor)        'Make sure working with a positive number
    Blue = RGBColor \ 65536         'Seperate out the Blue
    RGBColor = RGBColor Mod 65536   'Remove Blue from Value
    Green = RGBColor \ 256          'Seperate out the Green
    Red = RGBColor Mod 256          'Remove Green from Value and place in Red
    Red = Red + Amount              'Change Red by Amount
    If Red < 0 Then Red = 0         'If less than 0 then make it 0
    If Red > 255 Then Red = 255     'If greater than 255 then make it 255
    Green = Green + Amount          'Change Green by Amount
    If Green < 0 Then Green = 0     'If less than 0 then make it 0
    If Green > 255 Then Green = 255 'If greater than 255 then make it 255
    Blue = Blue + Amount            'Change Blue by Amount
    If Blue < 0 Then Blue = 0       'If less than 0 then make it 0
    If Blue > 255 Then Blue = 255   'If greater than 255 then make it 255
    AdjustColor = RGB(Red, Green, Blue)  'Combine the colors and pass back the value
End Function

Private Sub Appearance3D(ByVal DefaultOffset As Long, _
                                Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    If InState = gbsDefault Then InState = State    'If state was not passed in then use
    Select Case InState     'Select correct state so it can draw the correct borders
        Case gbsDisable, gbsMouseOff
            PositionCaption         'Position Label as if in Up State
            If Not m_UseHover Or (m_HoverMode = gbhAllButBorder) Then   'If Using Hover Mode
                Border3D DefaultOffset  'Draw Border as if in Up state
            End If
        Case gbsDown
            PositionCaption True    'Position Label as if Depressed
            Border3D DefaultOffset, True    'Draw Border as if depressed
        Case gbsMouseOver, gbsUp
            PositionCaption         'Position Label as if in Up State
            Border3D DefaultOffset      'Draw Border as if in Up state
        Case Else   'Invalid state only Position label
            PositionCaption         'Position Label as if in Up State
    End Select
End Sub

Private Sub AppearanceBevel(ByVal DefaultOffset As Long, _
                                ByVal BeveledWidth As Long, _
                                Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    If InState = gbsDefault Then InState = State    'If state was not passed in then use
    Select Case InState     'Select correct state so it can draw the correct borders
        Case gbsDisable, gbsMouseOff
            PositionCaption         'Position Label as if in Up State
            If Not m_UseHover Or (m_HoverMode = gbhAllButBorder) Then  'If Using Hover Mode
                BorderBevel DefaultOffset, BeveledWidth     'Draw Border as if in Up state
            End If
        Case gbsDown
            PositionCaption True    'Position Label as if Depressed
            BorderBevel DefaultOffset, BeveledWidth, True   'Draw Border as if depressed
        Case gbsMouseOver, gbsUp
            PositionCaption         'Position Label as if in Up State
            BorderBevel DefaultOffset, BeveledWidth         'Draw Border as if in Up state
        Case Else   'Invalid state only Position label
            PositionCaption         'Position Label as if in Up State
    End Select
End Sub

Private Sub AppearanceEtched(ByVal DefaultOffset As Long, _
                                Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    DefaultOffset = 0       'Etched Scheme doesn't look good if shifted so ignore all shifts
    If InState = gbsDefault Then InState = State    'If state was not passed in then use
    Select Case InState     'Select correct state so it can draw the correct borders
        Case gbsDisable, gbsMouseOff
            PositionCaption         'Position Label as if in Up State
            If Not m_UseHover Or (m_HoverMode = gbhAllButBorder) Then      'If Using Hover Mode
                BorderEtched DefaultOffset      'Draw Border as if in Up state
            End If
        Case gbsDown
            PositionCaption True    'Position Label as if Depressed
            BorderEtched DefaultOffset, True    'Draw Border as if depressed
        Case gbsMouseOver, gbsUp
            PositionCaption         'Position Label as if in Up State
            BorderEtched DefaultOffset      'Draw Border as if in Up state
        Case Else   'Invalid state only Position label
            PositionCaption         'Position Label as if in Up State
    End Select
End Sub

Private Sub AppearanceFlat(ByVal DefaultOffset As Long, _
                                Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    If InState = gbsDefault Then InState = State    'If state was not passed in then use
    Select Case InState     'Select correct state so it can draw the correct borders
        Case gbsDisable, gbsMouseOff
            PositionCaption         'Position Label as if in Up State
            If Not m_UseHover Or (m_HoverMode = gbhAllButBorder) Then      'If Using Hover Mode
                BorderFlat DefaultOffset    'Draw Border as if in Up state
            End If
        Case gbsDown
            PositionCaption True    'Position Label as if Depressed
            BorderFlat DefaultOffset, True  'Draw Border as if in Up state
        Case gbsMouseOver, gbsUp
            PositionCaption         'Position Label as if in Up State
            BorderFlat DefaultOffset        'Draw Border as if in Up state
        Case Else   'Invalid state only Position label
            PositionCaption         'Position Label as if in Up State
    End Select
End Sub

'This is a template for additional Appearance schemes do note that if you add a scheme you must also add it
'to the Appearance Enumeration and to the DrawBorders sub-procedure.  For the Border Drawing code you may split
'that out into a seperate sub-procedure to maintain code readability, like I have done with 3D, Flat, Etched,
'and Bevel appearances.
'Private Sub AppearanceTemplate(ByVal DefaultOffset As Long, _
'                                Optional ByVal InState As gbState = gbsDefault)
'    If InState = gbsDefault Then InState = State    'If state was not passed in then use
'    Select Case InState     'Select correct state so it can draw the correct borders
'        Case gbsDisable
'            PositionCaption         'Position Label as if in Up State
'            If Not m_UseHover Or (m_HoverMode = gbhAllButBorder) Then      'If Using Hover Mode
'               'Code to draw the border in the disabled state goes here
'            End If
'        Case gbsDown
'            PositionCaption True    'Position Label as if Depressed
'            'Code to draw the border in the down state goes here
'        Case gbsMouseOff
'            PositionCaption         'Position Label as if in Up State
'            If Not m_UseHover Then      'If Using Hover Mode
'               'Code to draw the border in the Mouse Off (Up) state goes here
'            End If
'        Case gbsMouseOver
'            PositionCaption         'Position Label as if in Up State
'            'Code to draw the border in the Mouse Over (Hover) state goes here
'        Case gbsUp
'            PositionCaption         'Position Label as if in Up State
'            'Code to draw the border in design mode goes here
'        Case Else   'Invalid state
'            PositionCaption         'Position Label as if in Up State
'            'Code to draw the border for any invalid/default state goes here (if any)
'    End Select
'End Sub

Private Sub Border3D(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If m_UseClassicBorders Then     'If we're using the old style borders then
        Border3DOld DefaultOffset, Depressed    'Call Old Border Draw Routine
    Else                            'Otherwise
        Border3DNew DefaultOffset, Depressed    'Call New Border Draw Routine
    End If
End Sub

Private Sub Border3DNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picGradient.hdc          'Get the Device Context handle and store it in a local variable for quick access
    With picGradient
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountDarken)
        Next
    End If
End Sub

Private Sub Border3DOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If Depressed Then   'If button is depressed then
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight, B               'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest      'Draw two lines that will be the Top and Left Outside border lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDark     'Draw two lines that will be the Top and Left Inside border lines
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark
    Else        'Button should be drawn up
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B         'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderLightest     'Draw two lines that will be the Top and Left Outside border lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight    'Draw two lines that will be the Top and Left Inside border lines
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLight
    End If
End Sub

Private Sub BorderBevel(ByVal DefaultOffset As Long, _
                                ByVal BeveledWidth As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels
    Dim ColorAdjust As Long     'Variable containing the amount to change the color of each pixel
    Dim modDepress As Integer   'Variable to hold a modifier determining if the border should be drawn up or depressed (color adjusting)

    ColorAdjust = m_BevelIntensity  'Initialize the color adjustment (Intensity)
    hPic = picGradient.hdc          'Get the Device Context handle and store it in a local variable for quick access
    With picGradient
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then       'If border should be draw as depressed then
        modDepress = -1     'Set a negative modifier to reverse the color adjustment
    Else                    'Otherwise
        modDepress = 1      'Set a positive modifier to leave color adjustment alone
    End If
    Do      'Loop through the following block until we've finished the bevel (BeveledWidth = -1) moving from the inside > outside
        'Start a loop through the Horizontal points from Left to Right on the y positions of the Bevel
        For i = (BeveledWidth + DefaultOffset) To (picWidth - BeveledWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, BeveledWidth + DefaultOffset)  'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the left side
            SetPixelV hPic, i, BeveledWidth + DefaultOffset, AdjustColor(CurColor, ColorAdjust * modDepress)
            CurColor = GetPixel(hPic, i, picHeight - BeveledWidth - DefaultOffset - 1)  'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, i, picHeight - BeveledWidth - DefaultOffset - 1, AdjustColor(CurColor, -ColorAdjust * modDepress)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the x positions of the Bevel
        For i = (BeveledWidth + DefaultOffset + 1) To (picHeight - BeveledWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, BeveledWidth + DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the top
            SetPixelV hPic, BeveledWidth + DefaultOffset, i, AdjustColor(CurColor, ColorAdjust * modDepress)
            CurColor = GetPixel(hPic, picWidth - BeveledWidth - DefaultOffset - 1, i)   'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the bottom
            SetPixelV hPic, picWidth - BeveledWidth - DefaultOffset - 1, i, AdjustColor(CurColor, -ColorAdjust * modDepress)
        Next
        BeveledWidth = BeveledWidth - 1     'Reduce the bevel width left to draw
        ColorAdjust = ColorAdjust + m_BevelIntensity    'Increase the Intensity of the color change
    Loop Until BeveledWidth = -1
End Sub

Private Sub BorderEtched(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If m_UseClassicBorders Then     'If we're using the old style borders then
        BorderEtchedOld DefaultOffset, Depressed    'Call Old Border Draw Routine
    Else                            'Otherwise
        BorderEtchedNew DefaultOffset, Depressed    'Call New Border Draw Routine
    End If
End Sub

Private Sub BorderEtchedNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picGradient.hdc          'Get the Device Context handle and store it in a local variable for quick access
    With picGradient
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            If Not (i = (picWidth - DefaultOffset - 1)) Then
                SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            Else
                SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            End If
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            If Not (i = (picWidth - DefaultOffset - 2)) Then
                SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, 2 * AmountLighten)
            Else
                SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, 2 * AmountDarken)
            End If
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
    End If
End Sub

Private Sub BorderEtchedOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If Depressed Then   'If button is depressed then
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight, B               'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest          'Draw two lines that will be the Top and Left Outside border lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDark     'Draw two lines that will be the Top and Left Inside border lines
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark
    Else        'Button should be drawn up
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B
        picGradient.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B
    End If
End Sub

Private Sub BorderFlat(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If m_UseClassicBorders Then     'If we're using the old style borders then
        BorderFlatOld DefaultOffset, Depressed  'Call Old Border Draw Routine
    Else                            'Otherwise
        BorderFlatNew DefaultOffset, Depressed  'Call New Border Draw Routine
    End If
End Sub

Private Sub BorderFlatNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picGradient.hdc          'Get the Device Context handle and store it in a local variable for quick access
    With picGradient
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
    End If
End Sub

Private Sub BorderFlatOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If Depressed Then   'If button is depressed then
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box on the border of the control that will become the Bottom and Right border lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest     'Draw the Top and Left Border Lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
    Else        'Button should be drawn up
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B     'Draw a box on the border of the control that will become the Bottom and Right border lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest    'Draw the Top and Left Border Lines
        picGradient.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderLightest
    End If
End Sub

Private Sub ChangeValue()       'This sub-procedure will only occur while button is in state mode
On Error Resume Next
    m_Value = m_Value Xor True          'Toggle the value of m_value
    If (m_Value) And (m_ButtonType = gbtOptionButton) Then      'If the value is now true and we're an option button then
        ResetAll        'Reset the values of the other buttons to False
    End If
    If Not (m_Enabled) Then Exit Sub    'Make sure control is enabled before throwing event
    RaiseEvent ValueChanged(m_Value)    'Signal to user that the value has changed
End Sub

Private Sub CheckParent()
    On Error GoTo NoParent      'Start error handling in case the parent object doesn't support what we need
    If Not (UserControl.Parent.ScaleMode = vbUser) Then 'If the scalemode is not User defined then
        bParentAvailable = True 'Indicate we have parent available and it's not in user mode
    End If
    Exit Sub    'Finish this sub procedure

NoParent:
    bParentAvailable = False    'We don't have a parent with the scalemode property so indicate such to the control.
End Sub

Private Sub CheckToolTip()
    On Error GoTo NoExtender    'Start error handling in case the extender object doesn't support what we need
    If Not (UserControl.Extender.ToolTipText = vbNullString) Then
        'Do nothing no error
    End If
    bToolTipAvailable = True 'Indicate we have extender available
    Exit Sub    'Finish this sub procedure

NoExtender:
    bToolTipAvailable = False    'We don't have an extender available with ToolTip Info so indicate such to the control.
End Sub

Private Sub CopyGradient()
On Error Resume Next
    If (m_Style = gbsGradient) Or (m_Style = gbsGraphicalGradient) Then     'If it is a Gradient mode then
        PaintTransparentStdPic picGradient.hdc, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), GradPic, 0, 0, picGradient.BackColor   'Copy the gradient from the picture to the drawing area
        picGradient.Refresh         'Refresh so that gradient will show up
    End If
End Sub

Private Sub CopyPicture(Optional ByVal InState As gbState = gbsDefault, _
                                Optional ByVal XPos As Long = -1, _
                                Optional ByVal YPos As Long = -1)
    Dim TempPic As Picture      'Variable to hold the picture to place onto the control
    Dim defXPos As Boolean      'Variable to hold whether an X-Position has been passed to the procedure
    Dim defYPos As Boolean      'Variable to hold whether a Y-Position has been passed to the procedure

    On Error Resume Next
    If (m_Style = gbsGraphical) Or (m_Style = gbsGraphicalGradient) Or (m_Style = gbsGraphicalPicture) Then 'If it is a Graphical mode then
        If InState = gbsDefault Then InState = State        'If a State wasn't passed then assign the current state of the control
        Select Case State       'Select which state we are in
            Case gbsDisable     'Disabled
                Set TempPic = m_DisabledPicture     'Assign the disabled picture to the Temp variable
            Case gbsDown        'Down
                Set TempPic = m_DownPicture         'Assign the down picture to the Temp variable
            Case gbsMouseOver   'Mouse is over the control
                If m_UseHover And Not (m_HoverMode = gbhBorderOnly) Then    'If using Hover mode then
                    Set TempPic = m_HoverPicture    'Assign the hover picture to the Temp variable
                Else        'Otherwise
                    Set TempPic = m_Picture         'Assign the standard picture to the Temp variable
                End If
            Case gbsUp, gbsMouseOff     'Mouse is either not over control or we're in design mode
                Set TempPic = m_Picture             'Assign the standard picture to the Temp variable
            Case Else       'some invalid state
                'Invalid State set picture to nothing
                Set TempPic = Nothing
        End Select
        If TempPic Is Nothing Then      'If the temp variable contains nothing then
            Set TempPic = m_Picture     'Assign the standard picture to the Temp Variable
        End If
        'If Center[-1] was used for any co-ordinate then calculate the co-ordinate that would place the picture dead center on that axis
        If XPos = -1 Then       'If we did not receive an X-Coordinate then
            'Calculate the x position that would center the graphic
            XPos = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - ScaleX(TempPic.Width, vbHimetric, vbPixels)) / 2)
        Else                    'Otherwise we received the coordinate
            defXPos = True      'Indicate that we received it so that we don't erase it by aligning
        End If
        If YPos = -1 Then       'If we did not receive an X-Coordinate then
            'Calculate the y position that would center the graphic
            YPos = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - ScaleY(TempPic.Height, vbHimetric, vbPixels)) / 2)
        Else                    'Otherwise we received the coordinate
            defYPos = True      'Indicate that we received it so that we don't erase it by aligning
        End If

        ' Aligning the picture
        If Not (defXPos) Then
            Select Case m_PictureAlignment
                Case gbaLeftTop, gbaLeftMiddle, gbaLeftBottom
                    XPos = CLng(m_PictureCushion)
                Case gbaCenterTop, gbaCenterMiddle, gbaCenterBottom
                    XPos = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - ScaleX(TempPic.Width, vbHimetric, vbPixels)) / 2)
                Case gbaRightTop, gbaRightMiddle, gbaRightBottom
                    XPos = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - ScaleX(TempPic.Width, vbHimetric, vbPixels)) - m_PictureCushion)
            End Select
        End If
        If Not (defYPos) Then
            Select Case m_PictureAlignment
                Case gbaLeftTop, gbaCenterTop, gbaRightTop
                    YPos = CLng(m_PictureCushion)
                Case gbaLeftMiddle, gbaCenterMiddle, gbaRightMiddle
                    YPos = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - ScaleY(TempPic.Height, vbHimetric, vbPixels)) / 2)
                Case gbaLeftBottom, gbaCenterBottom, gbaRightBottom
                    YPos = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - ScaleY(TempPic.Height, vbHimetric, vbPixels)) - m_PictureCushion)
            End Select
        End If

        'Copy the picture using the mask color to the proper co-ordinates on the control
        PaintTransparentStdPic picGradient.hdc, XPos, YPos, ScaleX(TempPic.Width, vbHimetric, vbPixels), ScaleY(TempPic.Height, vbHimetric, vbPixels), TempPic, 0, 0, UserControl.MaskColor
        picGradient.Refresh     'Refresh the graphics now that the picture has been copied
    End If
End Sub

Private Sub CreateToolTip()
On Error Resume Next
    If (hWndToolTip = 0) And (Ambient.UserMode) Then 'If the ToolTip Control doesn't exist
        'Create the control that will hold the Tool
        hWndToolTip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, 0, WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, UserControl.hwnd, 0, App.hInstance, 0)
        'Make sure the window will appear on top
        SetWindowPos hWndToolTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        GetClientRect UserControl.hwnd, rctToolTip  'The the Rectagular Area of the button
        With tiToolTip      'Execute the following statements with the ToolTip Info Variable as the default
            .cbSize = Len(tiToolTip)    'Assign the Length of the structure to the Info structure
            .uFlags = TTF_SUBCLASS      'Make sure that the control will SubClass itself
            .hwnd = UserControl.hwnd    'Assign the Control's window handle as the parent
            .hInst = App.hInstance      'Assign this instance of the Button as the owner
            .uID = 0                    'No ID is needed so set to 0
            .TipText = Me.ToolTipText   'Assign the ToolTip Text to the Info Structure
            .rct = rctToolTip           'Assign the Rectangular area of the control to the Info Structure
        End With
        'Send the message to the ToolTip control to add in this tool
        SendMessage hWndToolTip, TTM_ADDTOOL, 0, tiToolTip
    End If
End Sub

Private Sub DrawBorders(Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    Dim DefaultOffset As Long           'Variable to hold any offset for the Default Button Border

    If Ambient.DisplayAsDefault Then    'If control has focus then
        DefaultOffset = 1               'Set the Offset equal to 1 indicating a 1-pixel offset (To prevent the border from shifting in 1 pixel on focus comment this line out)
    End If
    If Not Ambient.UserMode Then        'If in design mode then
        If m_Enabled Then               'If the control is enabled then
            InState = gbsUp             'Set state to UP in design mode
        Else                            'Otherwise
            InState = gbsDisable        'Set state to disabled
        End If
    End If
    If ((m_ButtonType = gbtStateButton) Or (m_ButtonType = gbtOptionButton)) And (m_Value) Then InState = gbsDown 'If button is a state or option button and value is true then draw as if Down
    If InState = gbsDefault Then InState = State        'If no state was passed to this routine then use the current Control state
    CopyPicture         'Copy the correct picture to the button
    SetCaptionFont      'Set the Font & Fore Color of the caption to the appropriate one
    Select Case m_Appearance        'Select the correct Appearance Scheme
        Case gbaFlat        'Flat
            AppearanceFlat DefaultOffset, InState       'Use the Flat Border Scheme
        Case gba3D          '3D
            Appearance3D DefaultOffset, InState         'Use the 3D Border Scheme
        Case gbaEtched      'Etched
            AppearanceEtched DefaultOffset, InState     'Use the Etched Border Scheme
        Case gbaBevel       'Bevel
            AppearanceBevel DefaultOffset, m_BevelWidth, InState
        Case Else           'Invalid appearance so use the default 3D
            Appearance3D DefaultOffset, InState         'Use the 3D Border Scheme because of invalid appearance
    End Select
    If Ambient.DisplayAsDefault Then    'If control has focus then
        'Draw the Black Focus border around the control
        picGradient.Line (0, 0)-(ScaleWidth - ScaleX(1, vbPixels, ScaleMode), ScaleHeight - ScaleY(1, vbPixels, ScaleMode)), 0, B
    End If
End Sub

Private Sub DrawCaption(ByVal Caption As String, _
                        ByRef Target As Control, _
                        Optional ByVal Alignment As gbAlignment = gbaCenterMiddle, _
                        Optional ByVal CaptionStyle As gbCaptionStyle = gbcStandard, _
                        Optional ByVal LeftMargin As Long = 3, _
                        Optional ByVal TopMargin As Long = 0, _
                        Optional ByVal RightMargin As Long = 3, _
                        Optional ByVal BottomMargin As Long = 0)
    'Top Margin is used for Vertical centering and Top alignment only
    'Bottom Margin is only used for Bottom alignment
    'Primitives
On Error Resume Next
    Dim Flags As Long               'To hold alignment flags for drawing the text
    Dim Line As String              'The current working line
    Dim OrigColor As Long           'Variable to hold the original text color
    Dim PosY As Single              'The Y position of where the Top of the line should be (Only used at the end when printing to the Target)
    Dim ret As Long                 'To hold the DrawText return value (Text Height)
    'Structures
    Dim rct As RectAPI              'Structure to hold the rectangle information about where to draw the Caption
    Dim rctAdjust As RectAPI        'Structure to hold the rectangle information after adjustments for effects.

    Line = Caption
    rct.Top = TopMargin     'Initialize the Rect variable with the top, left, and right boundaries
    rct.Left = LeftMargin
    rct.Right = Target.ScaleX(Target.ScaleWidth, Target.ScaleMode, vbPixels) - (RightMargin)
    ret = DrawText(Target.hdc, Line, Len(Line), rct, DT_CALCRECT)       'Have the API calculate the bottom boundary
    rct.Right = Target.ScaleX(Target.ScaleWidth, Target.ScaleMode, vbPixels) - (RightMargin)    'ReAdjust right side for alignment purposes
    Select Case Alignment   'Select alignment
        Case gbaLeftTop, gbaCenterTop, gbaRightTop  'Align caption to top
            rct.Top = TopMargin             'Align caption to top margin
            rct.Bottom = rct.Top + ret      'Adjust the bottom because the top was adjusted
        Case gbaLeftBottom, gbaCenterBottom, gbaRightBottom 'Align caption to Bottom
            rct.Top = Target.ScaleY(Target.ScaleHeight, Target.ScaleMode, vbPixels) - (ret + BottomMargin)      'Align caption top to height of target - (height of caption plus margin)
            rct.Bottom = rct.Top + ret      'Adjust the bottom because the top was adjusted
        Case Else           'Center caption vertically
            TopMargin = TopMargin - m_AlignmentCushion
            rct.Top = ((Target.ScaleY(Target.ScaleHeight, Target.ScaleMode, vbPixels) - ret) / 2) + TopMargin       'Center caption on target vertically
            rct.Bottom = rct.Top + ret      'Adjust the bottom because the top was adjusted
    End Select
    Select Case Alignment           'Select how the text lines should be aligned
        Case gbaLeftTop, gbaLeftMiddle, gbaLeftBottom   'Justified along the left side
            Flags = DT_LEFT         'Set the flag for Left alignment
        Case gbaRightTop, gbaRightMiddle, gbaRightBottom    'Justified along the right side
            Flags = DT_RIGHT        'Set the flag for Right alignment
        Case Else                   'Each line centered (also covers invalid alignments by defaulting to centered)
            Flags = DT_CENTER       'Set the flag for Center alignment
    End Select
    OrigColor = Target.ForeColor
    Select Case CaptionStyle        'Choose the caption style
        Case gbcRaisedLight         'Light Inset Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gbcRaisedHeavy          'Heavy Inset Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 2         'Move rectangle down two pixels
                .Bottom = .Bottom + 2
                .Left = .Left + 2       'Move rectangle right two pixels
                .Right = .Right + 2
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gbcInsetLight         'Light Raised Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 1         'Move rectangle down one pixel
                .Bottom = .Bottom + 1
                .Left = .Left + 1       'Move rectangle right one pixel
                .Right = .Right + 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gbcInsetHeavy         'Heavy Raised Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 2         'Move rectangle down two pixels
                .Bottom = .Bottom + 2
                .Left = .Left + 2       'Move rectangle right two pixels
                .Right = .Right + 2
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gbcDropShadow          'Drop Shadow Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 1         'Move rectangle down one pixel
                .Bottom = .Bottom + 1
                .Left = .Left + 1       'Move rectangle right one pixel
                .Right = .Right + 1
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hdc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case Else                   'No Effect/Invalid Effect
            'Do nothing
    End Select
    ret = DrawText(Target.hdc, Line, Len(Line), rct, Flags)     'Actually Draw on the Caption
End Sub

Private Sub ModifyToolTipText()
On Error Resume Next
    If (hWndToolTip <> 0) Then      'If ToolTip exists then
        tiToolTip.TipText = Me.ToolTipText      'Get the correct ToolTip Text
        SendMessage hWndToolTip, TTM_UPDATETIPTEXT, 0, tiToolTip    'Send the message to update the ToolTipText
    End If
End Sub

Private Sub PaintAll()
On Error Resume Next
    picGradient.Cls 'Clear the Temporary graphic holder
    SetColors       'Set the colors just in case any color corruption occured on screen
    PaintGradient   'Paint the gradient into a seperate picture box
    CopyGradient    'Copy the gradient onto the control
    SetBackPicture  'Paint on the Back Picture if appropriate
    WrapCaption m_Caption, picGradient, True    'Pre-Prepare the caption to be drawn on the control.
    DrawBorders     'Draw the borders (which will also copy pictures)
    PaintToControl  'Paint the drawn control face onto the control
    SetMousePointer 'Set the mouse pointer information that the control should use
End Sub

Private Sub PaintFast()
On Error Resume Next
    picGradient.Cls 'Clear the Temporary graphic holder
    CopyGradient    'Copy the gradient onto the control
    SetBackPicture  'Paint on the Back Picture if appropriate
    DrawBorders     'Draw the borders (which will also copy pictures, reposition caption, and reconfiguring of the font)
    PaintToControl  'Paint the drawn control face onto the control
    SetMousePointer 'Set the mouse pointer information that the control should use
End Sub

Private Sub PaintGradient()
On Error Resume Next
    If (m_Style = gbsGradient) Or (m_Style = gbsGraphicalGradient) Then     'If it is a Gradient mode then
        'Draws the Gradient onto a PictureBox using the gradient class
        grad.Color1 = m_GradientColor1      'Configure 1st color
        grad.Color2 = m_GradientColor2      'Configure 2nd color
        grad.Angle = m_GradientAngle        'Set the angle to draw at
        grad.BlendMode = m_GradientBlendMode    'Set the blending mode
        grad.Repetitions = m_GradientRepetitions    'Set the number of gradient repetitions
        grad.GradientType = m_GradientType  'Set the gradient type to draw
        grad.Draw picGradient               'Actually draws the gradient on the picture box
        picGradient.Picture = picGradient.Image     'Move the picture from a temporary position into the picture so that it can be copied and used as a standard picture
        Set GradPic = picGradient.Picture   'Copy the gradient into the picture variable
    Else
        picGradient.Picture = LoadPicture() 'Remove any trace gradient elements
    End If
End Sub

Private Sub PaintToControl()
On Error Resume Next
BitBlt UserControl.hdc, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), picGradient.hdc, 0, 0, vbSrcCopy       'Copy pre-painted control picture onto control
' Set control.backstyle to transparent and enable the next 2 lines and disable te previous
' UserControl.Picture = picGradient.Image
' UserControl.MaskPicture = UserControl.Picture
UserControl.Refresh     'Do a refresh on the control so that what we just copied will show up
End Sub

'Provided with comments by Microsoft
Private Sub PaintStretchedStdPic(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal tWidth As Long, _
                                    ByVal tHeight As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal sWidth As Long, _
                                    ByVal sHeight As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
On Error Resume Next
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RectAPI
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintStretchedStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            StretchBlt hdcDest, xDest, yDest, tWidth, tHeight, hdcSrc, xSrc, ySrc, sWidth, sHeight, vbSrcCopy
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Transparent image
            StretchBlt hdcDest, xDest, yDest, tWidth, tHeight, hdcSrc, xSrc, ySrc, sWidth, sHeight, vbSrcCopy
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintStretchedStdPic_InvalidParam
    End Select
    Exit Sub

PaintStretchedStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Private Sub PaintTransparentDC(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
On Error Resume Next
    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hBmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long

    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        'Create halftone palette
        hPal = CreateHalftonePalette(hdcScreen)
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcDest, xDest, yDest, vbSrcCopy

    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hBmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hBmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hdcDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    DeleteObject hPal
End Sub

Private Sub PaintTransparentStdPic(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
On Error Resume Next
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RectAPI
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Transparent image
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub

PaintTransparentStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Private Sub PositionCaption(Optional ByVal Depressed As Boolean = False)
On Error Resume Next
    If Depressed Then           'If the button is depressed then
        DrawCaption WrappedCaption, picGradient, m_Alignment, CurrentCaptionStyle, m_AlignmentCushion + 1, m_AlignmentCushion + 1, m_AlignmentCushion - 1, m_AlignmentCushion - 1
    Else        'Button is to be put in the Up or Hover position
        DrawCaption WrappedCaption, picGradient, m_Alignment, CurrentCaptionStyle, m_AlignmentCushion, m_AlignmentCushion, m_AlignmentCushion, m_AlignmentCushion
    End If
End Sub

Private Sub ResetAll()
On Error Resume Next
    Dim ctl As Object       'Dimension a variable to hold each control as I'm looking (defined as object so it won't crash on the form which isn't a control)
    Dim MyName As String    'The returned stripped name of this button

    bInitiating = True      'Set the Initiating value to true so that we don't change this button just the others
    MyName = StripName(Ambient.DisplayName)     'Strip the name of this button to get the real name minus any indexes
    For Each ctl In UserControl.ParentControls  'Loop through all of the parent controls
        If (ctl.Name = MyName) Then     'Check to see if the control has MyName if so then
            If ctl.ButtonType = gbtOptionButton Then    'Check to see if it's an Option button we don't need to mess with it if it isn't
                ctl.Value = False       'Set It's Value to False
            End If
        End If
    Next
    bInitiating = False     'Reset the initiating value because we've completed the task
End Sub

Private Sub ResizeToolTip()
On Error Resume Next
    If (hWndToolTip <> 0) Then      'If ToolTip exists then
        GetClientRect Me.hwnd, rctToolTip   'Get the client rectangle (Size)
        tiToolTip.rct = rctToolTip  'Assign the size to the infostructure
        SendMessage hWndToolTip, TTM_NEWTOOLRECT, 0, tiToolTip  'Send the message to update the size
    End If
End Sub

Private Sub SetAccessKeys()
On Error Resume Next
    Dim lPlace As Long  'Variable to hold the current position in the string

    'set access key
    UserControl.AccessKeys = vbNullString       'Initialize AccessKeys so that there are no duplicates or wrong keys left over
    lPlace = 0      'Initialize Place Holder
    lPlace = InStr(lPlace + 1, m_Caption, "&", vbTextCompare)     'Locate first ampersand (&) and save it's location
    Do While lPlace <> 0        'While we still know where an ampersand is do
        If Mid$(m_Caption, lPlace + 1, 1) <> "&" Then     'If the next character is not an ampersand then
            UserControl.AccessKeys = UserControl.AccessKeys & Mid$(m_Caption, lPlace + 1, 1)      'Add the character to the accesskeys
        Else        'Otherwise
            lPlace = lPlace + 1     'Skip over the ampersand as it is not an access character
        End If
        lPlace = InStr(lPlace + 1, m_Caption, "&", vbTextCompare)     'Find the next ampersand in the caption
    Loop
End Sub

Private Sub SetBackPicture()
On Error Resume Next
    If ((m_Style = gbsPicture) Or (m_Style = gbsGraphicalPicture)) And Not (m_BackPicture Is Nothing) Then  'If it is in one of the Picture styles and the picture isn't nothing then
        'Paint the picture into the picturebox to prepare for overlay onto the button
        If m_AutoSize = gbasPictureToControl Then
            PaintStretchedStdPic picGradient.hdc, 0, 0, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels), m_BackPicture, 0, 0, ScaleX(m_BackPicture.Width, vbHimetric, vbPixels), ScaleY(m_BackPicture.Height, vbHimetric, vbPixels), picGradient.BackColor
        Else
            PaintTransparentStdPic picGradient.hdc, 0, 0, ScaleX(m_BackPicture.Width, vbHimetric, vbPixels), ScaleY(m_BackPicture.Height, vbHimetric, vbPixels), m_BackPicture, 0, 0, picGradient.BackColor
        End If
        picGradient.Refresh
    End If
End Sub

Private Sub SetCaptionFont(Optional ByVal InState As gbState = gbsDefault)
On Error Resume Next
    Dim TempFont As Font

    If InState = gbsDefault Then InState = State        'If no state was passed then use the current state of the control
    Select Case InState         'Select the correct state
        Case gbsDisable     'Disabled
            If m_DisabledFontEnabled Then                   'If Using font for disabled state then
                CurrentCaptionStyle = m_DisabledCaptionStyle    'Set the caption style
                Set TempFont = m_DisabledFont               'Set the Font to the disabled font
            Else                                            'Otherwise
                CurrentCaptionStyle = m_CaptionStyle        'Set the caption style
                Set TempFont = m_Font                       'Use standard Font
            End If
            picGradient.ForeColor = m_DisabledForeColor     'Set the Disabled text color
        Case gbsDown        'Down
            If m_DownFontEnabled Then                       'If Using font for Down state then
                CurrentCaptionStyle = m_DownCaptionStyle    'Set the caption style
                Set TempFont = m_DownFont                   'Set the Font to the down font
            Else                                            'Otherwise
                CurrentCaptionStyle = m_CaptionStyle        'Set the caption style
                Set TempFont = m_Font                       'Use standard Font
            End If
            picGradient.ForeColor = m_DownForeColor         'Set the Down text color
        Case gbsMouseOver   'Mouse is over Control
            If m_UseHover And Not (m_HoverMode = gbhBorderOnly) Then    'If Hover mode is enabled then
                If m_HoverFontEnabled Then                  'If Using Hover State font then
                    CurrentCaptionStyle = m_HoverCaptionStyle   'Set the caption style
                    Set TempFont = m_HoverFont              'Set the Font to the hover font
                Else
                    CurrentCaptionStyle = m_CaptionStyle    'Set the caption style
                    Set TempFont = m_Font                   'Use standard Font
                End If
                picGradient.ForeColor = m_HoverForeColor    'Set the Hover text color
            Else                                            'Otherwise
                CurrentCaptionStyle = m_CaptionStyle        'Set the caption style
                Set TempFont = m_Font                       'Use standard Font
                picGradient.ForeColor = m_ForeColor         'Set standard text color
            End If
        Case gbsUp, gbsMouseOff     'Mouse is not over control or in Design Mode
            CurrentCaptionStyle = m_CaptionStyle            'Set the caption style
            Set TempFont = m_Font                           'Set standard Font
            picGradient.ForeColor = m_ForeColor             'Set standard text color
        Case Else           'Invalid state
            CurrentCaptionStyle = m_CaptionStyle    'Set the caption style
            Set TempFont = m_Font                   'Set standard Font
            picGradient.ForeColor = m_ForeColor     'Set standard text color
    End Select
    If Not (TempFont Is Nothing) Then
        Set picGradient.Font = TempFont          'Perform the Font change
    End If
End Sub

Private Sub SetColors()
On Error Resume Next
    Select Case m_Style        'Select the style so we know what colors to load
        Case gbsStandard, gbsGraphical      'Standard modes like the standard command button so use system colors
            BorderDark = GetSysColor(vb3DShadow Xor &H80000000)         'Dark
            BorderDarkest = GetSysColor(vb3DDKShadow Xor &H80000000)    'Darkest
            BorderLight = GetSysColor(vb3DLight Xor &H80000000)         'Light
            BorderLightest = GetSysColor(vb3DHighlight Xor &H80000000)  'Lightest
        Case gbsGradient, gbsGraphicalGradient, gbsPicture, gbsGraphicalPicture     'New Modes use the color that was selected by the Developer
            BorderDark = AdjustColor(m_BorderColor, AmountDarken)           'Take the user selected color and Darken it
            BorderDarkest = AdjustColor(m_BorderColor, 2 * AmountDarken)    'Take the user selected color and double Darken it
            BorderLight = AdjustColor(m_BorderColor, AmountLighten)         'Take the user selected color and Lighten it
            BorderLightest = AdjustColor(m_BorderColor, 2 * AmountLighten)  'Take the user selected color and double Lighten it
        Case Else       'Invalid style use system colors
            BorderDark = GetSysColor(vb3DShadow Xor &H80000000)         'Dark
            BorderDarkest = GetSysColor(vb3DDKShadow Xor &H80000000)    'Darkest
            BorderLight = GetSysColor(vb3DLight Xor &H80000000)         'Light
            BorderLightest = GetSysColor(vb3DHighlight Xor &H80000000)  'Lightest
    End Select
End Sub

Private Sub SetMousePointer(Optional ByVal InState As gbState)
On Error Resume Next
    Dim TempIcon As Picture
    Dim TempPointer As Long

    If InState = gbsDefault Then InState = State        'If no state was passed then use the current state of the control
    Select Case InState
        Case gbsDisable     'Disabled
            TempPointer = m_DisabledMousePointer
            Set TempIcon = m_DisabledMouseIcon
        Case gbsDown        'Down
            TempPointer = m_DownMousePointer
            Set TempIcon = m_DownMouseIcon
        Case gbsMouseOver   'Mouse is over Control
            If m_UseHover And Not (m_HoverMode = gbhBorderOnly) Then    'If Hover mode is enabled then
                TempPointer = m_HoverMousePointer
                Set TempIcon = m_HoverMouseIcon
            End If
        Case gbsUp, gbsMouseOff     'Mouse is not over control or in Design Mode
            TempPointer = m_MousePointer
            Set TempIcon = m_MouseIcon
        Case Else           'Invalid state
            TempPointer = m_MousePointer
            Set TempIcon = m_MouseIcon
    End Select
    UserControl.MousePointer = TempPointer  'Assign the pointer that the control should use
    Set UserControl.MouseIcon = TempIcon    'Assign the icon that the control should use for custom pointers
End Sub

Private Function StripName(ByVal MyName As String)
On Error Resume Next
    Dim Work As String      'Temporary variable to hold the name I'm working with
    Dim i As Long           'a variable to hold the position of the left parenthesis "(" in the string if any.

    Work = MyName           'Assign the input name into the working variable
    i = InStr(1, Work, "(", vbTextCompare)  'Look for a left parenthesis
    If i <> 0 Then          'If we found a left parenthesis then
        Work = Mid$(Work, 1, i - 1)         'Assign The name minus the parenthesis and everything after
    End If
    StripName = Work        'Return the stripped down name
End Function

Private Sub WrapCaption(ByVal Caption As String, _
                        ByRef Target As Control, _
                        Optional ByVal RemovePadding As Boolean = True, _
                        Optional ByVal LeftMargin As Long = -1, _
                        Optional ByVal RightMargin As Long = -1)
On Error Resume Next
    'Primitives
    Dim Counter As Long             'A counter for splitting long "Words" to multiple lines
    Dim Index As Long               'Just a counter for the "For - Next" loops through the arrays
    Dim Line As String              'The current working line
    Dim ScaledWidth As Single       'The width after being scaled to include the margins if any
    'Arrays
    Dim Lines() As String           'Array to hold the lines that are constructed from the passed Caption
    Dim Words() As String           'Array holding the words in the passed Caption

    If LeftMargin = -1 Then LeftMargin = m_AlignmentCushion
    If RightMargin = -1 Then RightMargin = m_AlignmentCushion
    Line = vbNullString         'Initialize Line to null string
    ScaledWidth = Target.ScaleWidth - (Target.ScaleX(LeftMargin, vbPixels, Target.ScaleMode) + Target.ScaleX(RightMargin, vbPixels, Target.ScaleMode))  'Calculate the adjusted/scaled width for each line
    Words = Split(Caption, " ")     'Create words array from the caption with the delimeter as space all captions should have all words seperated by spaces
    ReDim Lines(0)      'Initialize the Lines array to one element to prevent errors on element value checks
    For Index = 0 To UBound(Words)  'Begin the loop through the words of the caption
        If Not (RemovePadding) Or (Words(Index) <> vbNullString) Then       'If not an empty line or leaving padding then execute code otherwise skip line
            If Len(Line) = 0 Then       'If the line is empty then (length = 0 characters)
                If Target.TextWidth(Words(Index)) <= ScaledWidth Then       'If the length of the "word" at the current index is within the adjusted width then
                    Line = Words(Index) 'Make the line that word at the moment
                Else                    'Otherwise
                    For Counter = 1 To Len(Words(Index))            'Start looping through each character in the long "word"
                        If (Target.TextWidth(Line & Mid$(Words(Index), Counter, 1)) > ScaledWidth) And (Line <> vbNullString) Then      'If the length of the line is greater than the adjusted width and the line isn't empty then
                            If Len(Lines(UBound(Lines))) = 0 Then   'If the current line in the Lines array is empty then (length = 0 characters)
                                Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                            Else                    'Otherwise
                                ReDim Preserve Lines(UBound(Lines) + 1)     'Add one more element to the lines array
                                Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                            End If
                            Line = Mid$(Words(Index), Counter, 1)   'Re-initialize string with the character that triggered the new line so the no excess is carried over
                        Else
                            Line = Line & Mid$(Words(Index), Counter, 1)    'Add current character to the line
                        End If
                    Next Counter        'Increment counter and continue the loop
                End If
            Else                        'Otherwise
                If Target.TextWidth(Line & " " & Words(Index)) <= ScaledWidth Then      'If the width of the current line and the current word is within the adjusted width then
                    Line = Line & " " & Words(Index)                'Add word to line
                ElseIf Target.TextWidth(Words(Index)) > ScaledWidth Then    'If the current word width is greater than the adjusted width then
                    If Len(Lines(UBound(Lines))) = 0 Then           'If the uppermost Line in the Lines collection is empty then
                        Lines(UBound(Lines)) = Line                 'Assign the current line to the uppermost element in the lines collection
                        Line = vbNullString         'Re-initialize string so the no excess is carried over
                        For Counter = 1 To Len(Words(Index))            'Start looping through each character in the long "word"
                            If (Target.TextWidth(Line & Mid$(Words(Index), Counter, 1)) > ScaledWidth) And (Line <> vbNullString) Then      'If the length of the line is greater than the adjusted width and the line isn't empty then
                                If Len(Lines(UBound(Lines))) = 0 Then   'If the current line in the Lines array is empty then (length = 0 characters)
                                    Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                                Else                    'Otherwise
                                    ReDim Preserve Lines(UBound(Lines) + 1)     'Add one more element to the lines array
                                    Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                                End If
                                Line = Mid$(Words(Index), Counter, 1)   'Re-initialize string with the character that triggered the new line so the no excess is carried over
                            Else
                                Line = Line & Mid$(Words(Index), Counter, 1)    'Add current character to the line
                            End If
                        Next Counter        'Increment counter and continue the loop
                    Else                'Otherwise
                        ReDim Preserve Lines(UBound(Lines) + 1)     'Add an element to the lines array
                        Lines(UBound(Lines)) = Line                 'Assign the current line to the uppermost element in the lines collection
                        Line = vbNullString             'Re-initialize string so the no excess is carried over
                        For Counter = 1 To Len(Words(Index))            'Start looping through each character in the long "word"
                            If (Target.TextWidth(Line & Mid$(Words(Index), Counter, 1)) > ScaledWidth) And (Line <> vbNullString) Then      'If the length of the line is greater than the adjusted width and the line isn't empty then
                                If Len(Lines(UBound(Lines))) = 0 Then   'If the current line in the Lines array is empty then (length = 0 characters)
                                    Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                                Else                    'Otherwise
                                    ReDim Preserve Lines(UBound(Lines) + 1)     'Add one more element to the lines array
                                    Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                                End If
                                Line = Mid$(Words(Index), Counter, 1)   'Re-initialize string with the character that triggered the new line so the no excess is carried over
                            Else
                                Line = Line & Mid$(Words(Index), Counter, 1)    'Add current character to the line
                            End If
                        Next Counter        'Increment counter and continue the loop
                    End If
                Else                    'Otherwise
                    If Len(Lines(UBound(Lines))) = 0 Then   'If uppermost element is empty then (length = 0 characters)
                        Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                    Else
                        ReDim Preserve Lines(UBound(Lines) + 1)         'Add one more element to the lines array
                        Lines(UBound(Lines)) = Line         'Assign the current line to the uppermost element in the lines collection
                    End If
                    Line = Words(Index)     'Assign current word to be the new line
                End If
            End If
        End If
    Next Index      'Increment the word and continue the loop
    Erase Words     'Free up memory from the word array as it is no longer needed
    If Len(Line) <> 0 Then      'If current line is not empty then
        If Len(Lines(UBound(Lines))) = 0 Then       'If the uppermost element in the lines array is empty then (length = 0 characters)
            Lines(UBound(Lines)) = Line     'Assign the current line to the uppermost element in the lines collection
        Else        'Otherwise
            ReDim Preserve Lines(UBound(Lines) + 1) 'Add one more element to the lines array
            Lines(UBound(Lines)) = Line     'Assign the current line to the uppermost element in the lines collection
        End If
    End If
    WrappedCaption = Join(Lines, vbCrLf)      'Assign the joined lines together into a single variable so that the height can be calculated
    Erase Lines     'Free memory by erasing the lines array because it is no longer necessary
End Sub
