VERSION 5.00
Object = "{74D127CD-2669-4E53-B9AE-DEBF1399A699}#4.0#0"; "SliderOCX.ocx"
Begin VB.Form TestSlider 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Test Sliders"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5175
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picRainbow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   -1155
      Picture         =   "SlidTest.frx":0000
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   474
      TabIndex        =   9
      Top             =   3105
      Visible         =   0   'False
      Width           =   7110
   End
   Begin SliderOCX.Slider Slider6 
      Height          =   225
      Left            =   1245
      TabIndex        =   6
      Top             =   2550
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   397
      PictureBody     =   "SlidTest.frx":3D72
      PictureThumb    =   "SlidTest.frx":3D8E
      PictureLeft     =   "SlidTest.frx":3E5C
      PictureRite     =   "SlidTest.frx":3E78
      Notches         =   9
      BackColor       =   -2147483633
      ForeColor       =   16711680
      Caption         =   "Intensity"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SliderOCX.Slider Slider4 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1395
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   529
      PictureBody     =   "SlidTest.frx":3E94
      PictureThumb    =   "SlidTest.frx":3EB0
      PictureLeft     =   "SlidTest.frx":4042
      PictureRite     =   "SlidTest.frx":405E
      Value           =   50
      BackColor       =   16777215
      Caption         =   "min <--     Speed    --> max"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
      BorderStyle     =   1
   End
   Begin SliderOCX.Slider Slider3 
      Height          =   210
      Left            =   2700
      TabIndex        =   4
      Top             =   1980
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   370
      PictureBody     =   "SlidTest.frx":407A
      PictureThumb    =   "SlidTest.frx":4096
      PictureLeft     =   "SlidTest.frx":4190
      PictureRite     =   "SlidTest.frx":41AC
      Notches         =   6
      VMax            =   5
      BackColor       =   12632256
      ForeColor       =   255
      Caption         =   "0     1     2     3     4     5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
   End
   Begin SliderOCX.Slider Slider2 
      Height          =   315
      Left            =   3225
      TabIndex        =   1
      ToolTipText     =   "On or Off"
      Top             =   1395
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      PictureBody     =   "SlidTest.frx":41C8
      PictureThumb    =   "SlidTest.frx":41E4
      PictureLeft     =   "SlidTest.frx":44CA
      PictureRite     =   "SlidTest.frx":44E6
      Notches         =   2
      VMax            =   -1
      Caption         =   "On  Off"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
      BorderStyle     =   1
   End
   Begin SliderOCX.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   450
      PictureBody     =   "SlidTest.frx":4502
      PictureThumb    =   "SlidTest.frx":8284
      PictureLeft     =   "SlidTest.frx":84DE
      PictureRite     =   "SlidTest.frx":85D8
      Caption         =   "Rainbow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
      BorderStyle     =   1
   End
   Begin SliderOCX.Slider Slider5 
      Height          =   270
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   476
      PictureBody     =   "SlidTest.frx":86D2
      PictureThumb    =   "SlidTest.frx":B724
      PictureLeft     =   "SlidTest.frx":B97E
      PictureRite     =   "SlidTest.frx":B99A
      VMin            =   100
      VMax            =   0
      Value           =   50
      BackColor       =   16777215
      Caption         =   "Miles / Gallon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   0   'False
      BorderStyle     =   1
   End
   Begin VB.Label lb 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "This is a demonstration of what you can do with the SliderOCX.ocx Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   337
      TabIndex        =   8
      Top             =   150
      Width           =   4500
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00808080&
      Height          =   465
      Index           =   4
      Left            =   1110
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2490
      Width           =   2565
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00808080&
      Height          =   375
      Index           =   3
      Left            =   2640
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1905
      Width           =   2040
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00808080&
      Height          =   465
      Index           =   2
      Left            =   3105
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00808080&
      Height          =   960
      Index           =   1
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1320
      Width           =   2190
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Gerundetes Rechteck
      Top             =   720
      Width           =   4545
   End
   Begin VB.Label lbOnOff 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   285
      Left            =   3975
      TabIndex        =   3
      Top             =   1410
      Width           =   570
   End
   Begin VB.Label lbVal 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lpScale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ÿ ž Ÿ ž Ÿ ž Ÿ ž Ÿ"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1260
      TabIndex        =   7
      Top             =   2745
      Width           =   2250
   End
End
Attribute VB_Name = "TestSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()

    Slider1_Scroll Slider1.Value
    Slider2_Scroll Slider2.Value

End Sub

Private Sub Slider1_Change(ByVal Value As Single)

    Slider1_Scroll Value

End Sub

Private Sub Slider1_Scroll(ByVal Value As Single)

    lbVal = Format$(Value, "#0.0")
    lbVal.BackColor = picRainbow.Point((picRainbow.ScaleWidth - 1) * Value / 100, 0)

End Sub

Private Sub Slider2_Change(ByVal Value As Single)

    Slider2_Scroll Value

End Sub

Private Sub Slider2_Scroll(ByVal Value As Single)

    If Value Then
        Slider2.BackColor = vbRed
        lbOnOff = "True"
      Else 'VALUE = FALSE
        Slider2.BackColor = vbGreen
        lbOnOff = "False"
    End If

End Sub

Private Sub Slider4_Change(ByVal Value As Single)

    Slider4_Scroll Value

End Sub

Private Sub Slider4_Scroll(ByVal Value As Single)

    Slider5 = Slider4

End Sub

Private Sub Slider5_Change(ByVal Value As Single)

    Slider5_Scroll Value

End Sub

Private Sub Slider5_Scroll(ByVal Value As Single)

    Slider4 = Slider5

End Sub

':) Ulli's VB Code Formatter V2.10.8 (12.03.2002 21:19:19) 1 + 64 = 65 Lines
