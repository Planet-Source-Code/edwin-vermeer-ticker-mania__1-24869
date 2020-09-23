VERSION 5.00
Begin VB.Form frmMainMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text animation and tickertape demo"
   ClientHeight    =   6975
   ClientLeft      =   3000
   ClientTop       =   1995
   ClientWidth     =   7230
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   Begin TextAnimationDemo.ctlGlobe ctlGlobe1 
      Height          =   495
      Left            =   6600
      TabIndex        =   17
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TextAnimationDemo.VerticalTitleBar VerticalTitleBar1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   11668
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   5040
      Max             =   10
      Min             =   1
      TabIndex        =   6
      Top             =   6600
      Value           =   10
      Width           =   2055
   End
   Begin TextAnimationDemo.GradientButton cmdAnimationOpen 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Open"
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
      ForeColor       =   16777215
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
   Begin TextAnimationDemo.GradientButton cmdAnimationOpen 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Close"
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
      Enabled         =   0   'False
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
      Value           =   -1  'True
   End
   Begin TextAnimationDemo.GradientButton cmdSystemtray 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   8
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Open"
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
      ForeColor       =   16777215
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
   Begin TextAnimationDemo.GradientButton cmdSystemtray 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   9
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Close"
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
      Enabled         =   0   'False
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
      Value           =   -1  'True
   End
   Begin TextAnimationDemo.GradientButton cmdDockedWindow 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   11
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Open"
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
      ForeColor       =   16777215
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
   Begin TextAnimationDemo.GradientButton cmdDockedWindow 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   12
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Close"
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
      Enabled         =   0   'False
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
      Value           =   -1  'True
   End
   Begin TextAnimationDemo.GradientButton cmdtitlebar 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   14
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Open"
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
      ForeColor       =   16777215
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
   Begin TextAnimationDemo.GradientButton cmdtitlebar 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   15
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Close"
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
      Enabled         =   0   'False
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
      Value           =   -1  'True
   End
   Begin TextAnimationDemo.TextAnimation TextAnimation1 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1296
      BackColorStart  =   4194304
      BackColorEnd    =   8388608
      Counter         =   2
      CounterMax      =   1400
      Speed           =   50
      AnimateInDesignmode=   0   'False
      TransparentColor=   0
      Angle           =   90
      Repetitions     =   2
   End
   Begin TextAnimationDemo.GradientButton cmdTransparent 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   19
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Open"
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
      ForeColor       =   16777215
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
   Begin TextAnimationDemo.GradientButton cmdTransparent 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   20
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Close"
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
      Enabled         =   0   'False
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
      Value           =   -1  'True
   End
   Begin TextAnimationDemo.GradientButton cmdVote 
      Height          =   375
      Left            =   300
      TabIndex        =   22
      Top             =   0
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   661
      Appearance      =   3
      ButtonType      =   2
      Caption         =   "Click here to vote for this PSC contribution !!!"
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
      ForeColor       =   0
      GradientAngle   =   150
      GradientBlendMode=   1
      GradientColor1  =   12648384
      GradientColor2  =   12632319
      GradientRepetitions=   8
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
      HoverForeColor  =   49152
      HoverMode       =   2
      Style           =   2
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Transparent form tickertape demo. Try dragging this. Quality is affected by the redraw of the application underneath."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   480
      TabIndex        =   21
      Top             =   5520
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Active titlebar tickertape demo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   16
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Docked window tickertape demo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Systemtray tickertape demo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Demonstrate how a volume control interface could look."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMainMenu.frx":044A
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
      Height          =   1455
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Open the form with a demonstration of most of the TextAnimation control options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMainMenu.frx":0595
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
      Height          =   1455
      Left            =   495
      TabIndex        =   5
      Top             =   615
      Width           =   6015
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grad As New Gradient
Const PSCid = 24869


' Of course these are the 4 most important commands of this code ;-)
Private Sub cmdVote_Click()
On Error Resume Next
  Me!VerticalTitleBar1.RunThisURL "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=" & Trim(Str(PSCid)) & "&optCodeRatingValue=5"
End Sub


Public Sub cmdAnimationOpen_Click(Index As Integer)
On Error Resume Next
  cmdAnimationOpen(1 - Index).Enabled = True
  cmdAnimationOpen(Index).Enabled = False
  Select Case Index
  Case 0
    frmTextAnimation.Show
  Case 1
    Unload frmTextAnimation
  End Select
  DoEvents
  Me.SetFocus
End Sub

Private Sub cmdDockedWindow_Click(Index As Integer)
On Error Resume Next
  cmdDockedWindow(1 - Index).Enabled = True
  cmdDockedWindow(Index).Enabled = False
  Select Case Index
  Case 0
    frmDockBrowser.Show
  Case 1
    Unload frmDockBrowser
  End Select
  DoEvents
  Me.SetFocus

End Sub

Private Sub cmdSystemtray_Click(Index As Integer)
On Error Resume Next
  cmdSystemtray(1 - Index).Enabled = True
  cmdSystemtray(Index).Enabled = False
  Select Case Index
  Case 0
    frmSystemTrayTicker.Show
  Case 1
    Unload frmSystemTrayTicker
  End Select
  DoEvents
  Me.SetFocus
End Sub

Private Sub cmdtitlebar_Click(Index As Integer)
On Error Resume Next
  cmdtitlebar(1 - Index).Enabled = True
  cmdtitlebar(Index).Enabled = False
  Select Case Index
  Case 0
    frmBrowser.Show
  Case 1
    Unload frmBrowser
  End Select
  DoEvents
  Me.SetFocus
End Sub

Private Sub cmdTransparent_Click(Index As Integer)
On Error Resume Next
  cmdTransparent(1 - Index).Enabled = True
  cmdTransparent(Index).Enabled = False
  Select Case Index
  Case 0
    Load frmTransparentTicker
    DoEvents
    frmTransparentTicker.Show
  Case 1
    Unload frmTransparentTicker
  End Select
  DoEvents
  Me.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
  ctlGlobe1.Start
  grad.Color1 = RGB(0, 0, 0)
  grad.Color2 = RGB(100, 100, 140)
  grad.Angle = 35
  grad.Repetitions = 2.5
  grad.GradientType = 0

End Sub

Private Sub Form_Resize()
On Error Resume Next
  grad.Draw Me               'Actually draws the gradient on the picture box
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  Unload frmTextAnimation
  Unload frmVolume
  Unload frmBrowser
  Unload frmDockBrowser
  Unload frmSystemTrayTicker
  Unload frmBrowser
  Unload frmTransparentTicker
  End
  
End Sub

Private Sub HScroll_Change()
On Error Resume Next
  frmVolume.setVolume HScroll
End Sub


