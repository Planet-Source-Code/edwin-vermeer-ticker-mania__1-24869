VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'API Declarations
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'API Data Structures
Private Type POINTAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Type PointSng   'Internal Point structure
    X   As Double       'Uses Doubles for more precision and speed.
    Y   As Double
End Type

Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Drawing Constants
Private Const PS_SOLID As Long = 0  'Solid Pen Style (Used for CreatePen())

'Conversion Constants
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

'HSL Types & constants
Const RGBMAX    As Long = 255       'Note since RGMMAX=HSLMAX I'll simplify the math in the HSL routines
Const HSLMAX    As Long = RGBMAX 'with Windows HSLMAX is 240 to make it dividable by 6
                                 'and still fit in a byte; we're using floating point
                                 'arithmetic so there's no need for that
Private Type tHSL
    h As Double: s As Double: L As Double   'Hue Saturation Luminance
End Type '                                   -   -          -

Private Type tRGB
    R As Long: g As Long: b As Long         'Red Green Blue
End Type '                                   -   -     -

Public Enum GradBlendMode
    gbmRGB = 0
    gbmHSL = 1
End Enum

Public Enum GradType
    gtNormal = 0
    gtElliptical = 1
    gtRectangular = 2
End Enum

'Property Storage Variables
Private mlColor1        As Long
Private mlColor2        As Long
Private mfAngle         As Double
Private mfRepetitions   As Double
Private m_GradientType  As GradType
Private m_BlendMode     As GradBlendMode

'Property Default Constants
Private Const mlDefColor1       As Long = &HFFFFFF  'White
Private Const mlDefColor2       As Long = &H0       'Black
Private Const mfDefAngle        As Double = 0       'Right to Left
Private Const mfDefRepetitions  As Double = 1
Private Const m_DefGradientType As Long = gtNormal
Private Const m_DefBlendMode    As Long = gbmRGB

Public Function Draw(picObj As Object) As Boolean
'Note: This class uses API functions to draw. If the
'      destination object is in AutoRedraw mode, the
'      Refresh method for that object must be invoked.

'picObj can be a Form or PictureBox.

    Dim lRet    As Long
    Dim lIdx    As Long
'    Dim lTime   As Long
    Dim uRect   As RectAPI

'    lTime = GetTickCount()

    On Error GoTo LocalError

    'Stop the window from updating until we're finished.
    lRet = LockWindowUpdate(picObj.hwnd)

    'Get the client rect in pixels
    lRet = GetClientRect(picObj.hwnd, uRect)

    'Test for possible errors (GetClientRect failure or Rect < 2 pixels)
    If lRet <> 0 Then
        If uRect.Right > 1 And uRect.Bottom > 1 Then
            lIdx = DrawGradient(picObj.hdc, uRect.Right, uRect.Bottom)
            Draw = (lIdx > 0)
        End If
    End If

    'My P3-500 took 99 millisecs (.099 secs) to create and draw 2554 diagonal
    'lines at 315 degrees. That was frmDemo maximized on a 1280 x 1024 screen.
    'At this speed I can redraw an entire 1280px. screen over 10 times per second.

    'Same size rect at a 0 degree angle took 48 millisecs (.048 secs) to create and
    'draw 1278 lines. This speed can redraw a 1280px. screen 20 times per second.

    'Uncomment the two lines below and the lTime line at the top
    'of this function to test the times on your PC.

'    lTime = GetTickCount() - lTime
'    Debug.Print CStr(lIdx / 2) & " lines drawn in " & CStr(lTime) & " milliseconds"

NormalExit:
    'Unlock the window to allow it to update now.
    lRet = LockWindowUpdate(0)
    Exit Function

LocalError:
    Err.Raise Err.Number
    Resume NormalExit
End Function

Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, fRepetitions As Double, laRetColors() As Long) As Long
'Creates an array of colors blending from
'Color1 to Color2 in lSteps number of steps.
'Returns the count and fills the laRetColors() array.

    Dim lIdx    As Long
    Dim lIdx2   As Long
    Dim lRed    As Long
    Dim lGrn    As Long
    Dim lBlu    As Long
    Dim fRedStp As Double
    Dim fGrnStp As Double
    Dim fBluStp As Double
    Dim sHSL1 As tHSL
    Dim sHSL2 As tHSL
    Dim vHSL  As tHSL

    'Stop possible error
    If lSteps < 2 Then lSteps = 2

    'Create the colors
    ReDim laRetColors(lSteps * 2)
    Select Case m_BlendMode
        Case gbmRGB
            'Extract Red, Blue and Green values from the start and end colors.
            lRed = (lColor1 And &HFF&)
            lGrn = (lColor1 And &HFF00&) / &H100
            lBlu = (lColor1 And &HFF0000) / &H10000

            'Find the amount of change for each color element per color change.
            fRedStp = Div((lColor2 And &HFF&) - lRed, lSteps / fRepetitions)
            fGrnStp = Div(((lColor2 And &HFF00&) / &H100&) - lGrn, lSteps / fRepetitions)
            fBluStp = Div(((lColor2 And &HFF0000) / &H10000) - lBlu, lSteps / fRepetitions)

            laRetColors(0) = lColor1    'First Color
            laRetColors(Int(lSteps / fRepetitions)) = lColor2        'Last Color
            laRetColors(Int(lSteps / fRepetitions) + 1) = lColor2    'Last Color
            For lIdx = 1 To Int(lSteps / fRepetitions) - 1           'All Colors between
                laRetColors(lIdx) = CLng(lRed + (fRedStp * lIdx)) + _
                    (CLng(lGrn + (fGrnStp * lIdx)) * &H100&) + _
                    (CLng(lBlu + (fBluStp * lIdx)) * &H10000)
            Next lIdx
        Case gbmHSL
            ' A HUE gradient
            sHSL1 = RGBtoHSL(lColor1)
            sHSL2 = RGBtoHSL(lColor2)
            lIdx = sHSL2.h - sHSL1.h
            If Abs(lIdx) > HSLMAX / 2 Then
                lIdx = (HSLMAX - Abs(lIdx)) * -Sgn(lIdx)
            End If
            vHSL.h = (lIdx) / Int(lSteps / fRepetitions)
            vHSL.s = (sHSL2.s - sHSL1.s) / Int(lSteps / fRepetitions)
            vHSL.L = (sHSL2.L - sHSL1.L) / Int(lSteps / fRepetitions)
            For lIdx = 0 To Int(lSteps / fRepetitions) + 1
                sHSL1.h = sHSL1.h + vHSL.h
                Select Case sHSL1.h
                Case Is < 0
                    sHSL1.h = sHSL1.h + HSLMAX
                Case Is > HSLMAX
                    sHSL1.h = sHSL1.h - HSLMAX
                End Select
                sHSL1.s = sHSL1.s + vHSL.s
                sHSL1.L = sHSL1.L + vHSL.L
                laRetColors(lIdx) = HSLtoRGB(sHSL1)
            Next lIdx
    End Select

    ' first the complete color repetitions
    If Int(fRepetitions) >= 1 Then
      For lIdx2 = 1 To Int(fRepetitions) + 1
        If lIdx2 / 2 = Int(lIdx2 / 2) Then
          ' reverse gradient
          For lIdx = 0 To Int(lSteps / fRepetitions)
            laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors((lSteps / fRepetitions) - lIdx)
          Next lIdx
        Else
          ' normal gradient
          For lIdx = 0 To Int(lSteps / fRepetitions)
            laRetColors(((lIdx2 - 1) * Int(lSteps / fRepetitions)) + lIdx) = laRetColors(lIdx)
          Next lIdx
        End If
      Next lIdx2
    End If

    'Return number of colors in array
    BlendColors = lSteps
End Function

Private Function DrawGradient(ByVal hdc As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
    Dim bDone       As Boolean
    Dim iIncX       As Integer
    Dim iIncY       As Integer
    Dim lIdx        As Long
    Dim lRet        As Long
    Dim hPen        As Long
    Dim hOldPen     As Long
    Dim lPointCnt   As Long
    Dim laColors()  As Long
    Dim fMovX       As Double
    Dim fMovY       As Double
    Dim fDist       As Double
    Dim fAngle      As Double
    Dim fLongSide   As Double
    Dim uTmpPt      As POINTAPI
    Dim uaPts()     As POINTAPI
    Dim uaTmpPts()  As PointSng
    Dim MinDist     As Long

    On Error GoTo LocalError

    'Start with center of rect
    ReDim uaTmpPts(2)
    uaTmpPts(2).X = Int(lWidth / 2)
    uaTmpPts(2).Y = Int(lHeight / 2)

    'Calc distance to furthest edge as if rect were square
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2

    'Create points to the left and the right at a 0º angle (horizontal)
    uaTmpPts(0).X = uaTmpPts(2).X - fDist
    uaTmpPts(0).Y = uaTmpPts(2).Y
    uaTmpPts(1).X = uaTmpPts(2).X + fDist
    uaTmpPts(1).Y = uaTmpPts(2).Y

    'Lines will be drawn perpendicular to mfAngle so
    'add 90º and correct for 360º wrap
    fAngle = (mfAngle + 90) Mod 360

    'Rotate second and third points to fAngle
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)

    'We now have a line that crosses the center and
    'two sides of the rect at the correct angle.

    'Calc the starting quadrant, direction of and amount of first move
    '(fMovX, fMovY moves line from center to starting edge)
    'and direction of each incremental move (iIncX, iIncY).
    Select Case mfAngle
        Case 0 To 90
            'Left Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 90 To 180
            'Right Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 180 To 270
            'Right Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
        Case Else   '(270 to 360)
            'Left Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
    End Select

    'At this point we could calculate where the lines will cross the rect edges, but
    'this would slow things down. The picObj clipping region will take care of this.

    'Start with 1000 points and add more if needed. This increases
    'speed by not re-dimming the array in each loop.
    ReDim uaPts(999)

    'Set the first two points in the array
    uaPts(0).X = uaTmpPts(0).X + fMovX
    uaPts(0).Y = uaTmpPts(0).Y + fMovY
    uaPts(1).X = uaTmpPts(1).X + fMovX
    uaPts(1).Y = uaTmpPts(1).Y + fMovY

    lIdx = 2
    'Create the rest of the points by incrementing both points
    'on each line iIncX, iIncY from the previous line's points.
    'Where we stop depends on the direction of travel.
    'We'll continue until both points in a set reach the end.
    While Not bDone
        uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
        uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
        lIdx = lIdx + 1
        Select Case True
            Case iIncX > 0  'Moving Left to Right
                bDone = uaPts(lIdx - 1).X > lWidth And uaPts(lIdx - 2).X > lWidth
            Case iIncX < 0  'Moving Right to Left
                bDone = uaPts(lIdx - 1).X < 0 And uaPts(lIdx - 2).X < 0
            Case iIncY > 0  'Moving Top to Bottom
                bDone = uaPts(lIdx - 1).Y > lHeight And uaPts(lIdx - 2).Y > lHeight
            Case iIncY < 0  'Moving Bottom to Top
                bDone = uaPts(lIdx - 1).Y < 0 And uaPts(lIdx - 2).Y < 0
        End Select
        If (lIdx Mod 1000) = 0 Then
            ReDim Preserve uaPts(UBound(uaPts) + 1000)
        End If
    Wend

    'Free excess memory (may have 1001 points dimmed to 2000)
    ReDim Preserve uaPts(lIdx - 1)

    'Create the array of colors blending from mlColor1 to mlColor2
    lRet = BlendColors(mlColor1, mlColor2, lIdx / 2, mfRepetitions, laColors)

    Select Case m_GradientType
        Case gtNormal
            'Now draw each line in it's own color
            For lIdx = 0 To UBound(uaPts) - 1 Step 2
                'Move to next point
                lRet = MoveToEx(hdc, uaPts(lIdx).X, uaPts(lIdx).Y, uTmpPt)
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
                hOldPen = SelectObject(hdc, hPen)
                'Draw the line
                lRet = LineTo(hdc, uaPts(lIdx + 1).X, uaPts(lIdx + 1).Y)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hdc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
            DrawGradient = lIdx
        Case gtElliptical
            MinDist = IIf(lWidth < lHeight, Int(lWidth / 2), Int(lHeight / 2))  'Calculate the smallest distance to use as a buffer
            lRet = IIf(lWidth < lHeight, lWidth, lHeight) 'Calculate how many times we will iterate through the loop
            For lIdx = UBound(laColors) To 0 Step -1    'Make sure we don't end up with a black hole in the middle
                If laColors(lIdx) <> 0 Then
                    lIdx = lIdx - 1
                    Exit For
                End If
            Next lIdx
            lPointCnt = Int(lIdx / lRet)    'Calculate the color step to use for Ellipses (different from lines)
            For lIdx = 0 To lRet - 1
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx * lPointCnt)))
                hOldPen = SelectObject(hdc, hPen)
                'Draw the ellipse
                lRet = Ellipse(hdc, -MinDist + lIdx, -MinDist + lIdx, (lWidth + MinDist) - lIdx, (lHeight + MinDist) - lIdx)
                'The next ellipses are to make sure we don't have any gaps
                lRet = Ellipse(hdc, (-MinDist + lIdx) + 1, -MinDist + lIdx, ((lWidth + MinDist) - lIdx) - 1, (lHeight + MinDist) - lIdx)
                lRet = Ellipse(hdc, (-MinDist + lIdx) + 1, (-MinDist + lIdx) + 1, (lWidth + MinDist) - lIdx, (lHeight + MinDist) - lIdx)
                lRet = Ellipse(hdc, -MinDist + lIdx, (-MinDist + lIdx) + 1, (lWidth + MinDist) - lIdx, ((lHeight + MinDist) - lIdx) - 1)
                lRet = Ellipse(hdc, -MinDist + lIdx, -MinDist + lIdx, ((lWidth + MinDist) - lIdx) - 1, ((lHeight + MinDist) - lIdx) - 1)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hdc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
        Case gtRectangular
            MinDist = IIf(lWidth < lHeight, lWidth, lHeight)    'Calculate the smallest distance (number of rectangles we are going to draw * 2)
            For lIdx = UBound(laColors) To 0 Step -1    'Make sure we don't end up with a black hole in the middle
                If laColors(lIdx) <> 0 Then
                    lIdx = lIdx - 1
                    Exit For
                End If
            Next lIdx
            lPointCnt = Int(lIdx / (MinDist / 2)) 'Calculate the color step to use for rectangles (different from lines)
            For lIdx = 0 To (MinDist / 2)
                'Create the colored pen and select it into the DC
                hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx * lPointCnt)))
                hOldPen = SelectObject(hdc, hPen)
                'Draw the rectangle
                lRet = Rectangle(hdc, lIdx, lIdx, lWidth - lIdx, lHeight - lIdx)
                'Get the pen back out of the DC and destroy it
                lRet = SelectObject(hdc, hOldPen)
                lRet = DeleteObject(hPen)
            Next lIdx
    End Select

NormalExit:
    'Free the memory
    Erase laColors
    Erase uaPts
    Erase uaTmpPts
    Exit Function

LocalError:
    'MsgBox Err.Description, vbExclamation, "GradientRect.cls"
    DrawGradient = 0
    Resume NormalExit
End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Double)
    Dim fDX         As Double
    Dim fDY         As Double
    Dim fRadians    As Double

    fRadians = fDegrees * RADS
    fDX = uRotatePt.X - uAxisPt.X
    fDY = uRotatePt.Y - uAxisPt.Y
    uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
End Sub

Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
'Divides dNumer by dDenom if dDenom <> 0
'Auto-Handles 'Division By Zero' error.

    On Error Resume Next
    Div = dNumer / dDenom
End Function

Public Property Let Color1(ByVal lData As Long)
    Dim lIdx As Long

    mlColor1 = lData
    If mlColor1 < 0 Then
        lIdx = (mlColor1 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor1 = GetSysColor(lIdx)
        End If
    End If
End Property

Public Property Get Color1() As Long
    Color1 = mlColor1
End Property

Public Property Let Color2(ByVal lData As Long)
    Dim lIdx As Long

    mlColor2 = lData
    If mlColor2 < 0 Then
        lIdx = (mlColor2 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mlColor2 = GetSysColor(lIdx)
        End If
    End If
End Property

Public Property Get Color2() As Long
    Color2 = mlColor2
End Property

Public Property Let Angle(ByVal fData As Double)
'Angles are counter-clockwise and may be
'any Single value from 0 to 359.999999999.

' 135  90 45
'    \ | /
'180 --o-- 0
'    / | \
' 235 270 315

    'Correct angle to ensure between 0 and 359.999999999
    mfAngle = fData Mod 360
End Property

Public Property Get Angle() As Double
    Angle = mfAngle
End Property

Public Property Let Repetitions(ByVal fData As Double)
    mfRepetitions = Abs(fData)
    If mfRepetitions = 0 Then mfRepetitions = 1
End Property

Public Property Get Repetitions() As Double
    Repetitions = mfRepetitions
End Property

Public Property Let GradientType(ByVal eData As GradType)
    m_GradientType = eData
End Property

Public Property Get GradientType() As GradType
    GradientType = m_GradientType
End Property

Public Property Let BlendMode(ByVal eData As GradBlendMode)
    m_BlendMode = eData
End Property

Public Property Get BlendMode() As GradBlendMode
    BlendMode = m_BlendMode
End Property

Private Function Max(i1 As Variant, i2 As Variant) As Variant
    Max = IIf(i1 > i2, i1, i2)
End Function

Private Function Min(i1 As Variant, i2 As Variant) As Variant
    Min = IIf(i1 < i2, i1, i2)
End Function

Private Function RGBtoHSL(Color As Long) As tHSL
    Dim sRGB      As tRGB
    Dim ColorMax  As Long
    Dim ColorMin  As Long
    Dim MmM       As Long
    Dim MpM       As Long
    Dim RD        As Double
    Dim GD        As Double
    Dim BD        As Double

'   untangle colors
    sRGB.R = (Color And &HFF&)
    sRGB.g = (Color And &HFF00&) / &H100
    sRGB.b = (Color And &HFF0000) / &H10000

    ColorMax = Max(Max(sRGB.R, sRGB.g), sRGB.b)
    ColorMin = Min(Min(sRGB.R, sRGB.g), sRGB.b)
    MmM = ColorMax - ColorMin
    MpM = ColorMax + ColorMin
    RGBtoHSL.L = MpM / 2    'MpM * HSLMAX / (RGBMAX + RGBMAX) -> MpM / 2
    If ColorMax = ColorMin Then       'all color components equal -> gray
        RGBtoHSL.s = 0                'gray -> no saturation
        RGBtoHSL.h = HSLMAX * 2 / 3   'default blue for achromatic (same as Windows)
    Else
        If RGBtoHSL.L <= HSLMAX / 2 Then
            RGBtoHSL.s = MmM * HSLMAX / MpM
        Else
            RGBtoHSL.s = MmM * HSLMAX / (RGBMAX + RGBMAX - MpM)
        End If
        RD = (ColorMax - sRGB.R) * HSLMAX / MmM / 6
        GD = (ColorMax - sRGB.g) * HSLMAX / MmM / 6
        BD = (ColorMax - sRGB.b) * HSLMAX / MmM / 6
        Select Case ColorMax
            Case sRGB.R
                RGBtoHSL.h = BD - GD
            Case sRGB.g
                RGBtoHSL.h = HSLMAX / 3 + RD - BD
            Case sRGB.b
                RGBtoHSL.h = HSLMAX * 2 / 3 + GD - RD
        End Select
        If RGBtoHSL.h < 0 Then
            RGBtoHSL.h = RGBtoHSL.h + HSLMAX
        End If
    End If
End Function

Private Function HSLtoRGB(HSL As tHSL) As Long
    Dim FactA     As Double
    Dim FactB     As Double
    Dim sRGB      As tRGB

    HSL.h = Max(0, Min(HSL.h, HSLMAX))
    HSL.s = Max(0, Min(HSL.s, HSLMAX))
    HSL.L = Max(0, Min(HSL.L, HSLMAX))
    If HSL.s = 0 Then                     'no saturation -> gray
        sRGB.R = HSL.L '* RGBMAX / HSLMAX -> * 1  'luminance to all color components
        sRGB.g = sRGB.R
        sRGB.b = sRGB.R
    Else
        If HSL.L <= HSLMAX / 2 Then
            FactA = HSL.L * (HSLMAX + HSL.s) / HSLMAX
        Else
            FactA = HSL.L + HSL.s - HSL.L * HSL.s / HSLMAX
        End If
        FactB = HSL.L + HSL.L - FactA
        sRGB.R = HUEtoRGB(FactA, FactB, HSL.h + HSLMAX / 3) '* RGBMAX / HSLMAX -> * 1
        sRGB.g = HUEtoRGB(FactA, FactB, HSL.h) '* RGBMAX / HSLMAX -> * 1
        sRGB.b = HUEtoRGB(FactA, FactB, HSL.h - HSLMAX / 3) '* RGBMAX / HSLMAX -> * 1
    End If
    HSLtoRGB = RGB(sRGB.R, sRGB.g, sRGB.b)
End Function

Function HUEtoRGB(FA As Double, FB As Double, ByVal h As Double) As Long
    Select Case h
        Case Is < 0
            h = h + HSLMAX
        Case Is > HSLMAX
            h = h - HSLMAX
    End Select
    Select Case True
        Case h < HSLMAX / 6
            HUEtoRGB = FB + 6 * (FA - FB) * h / HSLMAX
        Case h < HSLMAX / 2
            HUEtoRGB = FA
        Case h < HSLMAX * 2 / 3
            HUEtoRGB = FB + 6 * (FA - FB) * (HSLMAX * 2 / 3 - h) / HSLMAX
        Case Else
            HUEtoRGB = FB
    End Select
End Function

Private Sub Class_Initialize()
    mlColor1 = mlDefColor1
    mlColor2 = mlDefColor2
    mfAngle = mfDefAngle
    mfRepetitions = mfDefRepetitions
    m_GradientType = m_DefGradientType
    m_BlendMode = m_DefBlendMode
End Sub
