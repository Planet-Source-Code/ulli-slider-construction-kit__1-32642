VERSION 5.00
Begin VB.UserControl Slider 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   KeyPreview      =   -1  'True
   ScaleHeight     =   255
   ScaleWidth      =   2205
   ToolboxBitmap   =   "Slider.ctx":0000
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   840
      Top             =   -45
   End
   Begin VB.PictureBox picTemp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   60
      Picture         =   "Slider.ctx":00FA
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgRight 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image imgLeft 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
   Begin VB.Image imgThumb 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblCapt 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1155
      TabIndex        =   1
      Top             =   75
      Width           =   45
   End
   Begin VB.Image imgBody 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefLng A-Z

'Version 2
''''''''''
'new Notch algorithm
'added Events Click,DblClick,KeyDown -Press -Up
'added Resolution property for easier sizing
'added Focus indication
'fixed some property dependencies (hope found all)
'fixed a few minor quirks
'code cosmetics

'Version 1
''''''''''
'Prototye

Private myIntervals '    that's one less than Notches
Private myValue          As Single
Private myMin            As Single
Private myMax            As Single
Private myAlignment      As Align
Private myFocus          As Boolean

Private Sliding          As Boolean
Private Changed          As Boolean
Private ReadingProp      As Boolean
Private NoSizing         As Boolean
Private LegalClick       As Boolean
Private WasUnderlined    As Boolean
Private XMin
Private XMax
Private OldPosn
Private DownX
Private Direction
Private Accu
Private TpP
Private TpP2             As Single
Private NotchDistance    As Single
Private OrgThumbWidth    As Single
Private OrgLeftWidth     As Single
Private OrgRightWidth    As Single
Private CurrentHeight    As Single

Public Enum Borderstyle
    BorderNone
    Border3D
End Enum

Public Enum Align
    [Align Left]
    [Align Center]
    [Align Right]
End Enum

Public Enum Component
    Background
    Body
    BtnThumb
    BtnLeft
    BtnRight
End Enum

Public Event Change(ByVal Value As Single)
Public Event Scroll(ByVal Value As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Click(ByVal Component As Component)
Public Event DblClick(ByVal Component As Component)
Private Const SM_CYHSCROLL As Long = 3
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Property Let Alignment(nwAlign As Align)
Attribute Alignment.VB_Description = "Sets/Returns the caption alignment."
Attribute Alignment.VB_HelpID = 10001

    If nwAlign = [Align Center] Or nwAlign = [Align Left] Or nwAlign = [Align Right] Then
        myAlignment = nwAlign
        With lblCapt
            .Top = (ScaleHeight - .Height) / 2
            Select Case myAlignment
              Case [Align Center]
                .Left = (ScaleWidth - .Width) / 2
              Case [Align Left]
                .Left = imgLeft.Width + TpP + TpP
              Case Else
                .Left = ScaleWidth - .Width - imgRight.Width - TpP - TpP
            End Select
        End With 'LBLCAPT
        PropertyChanged "Alignment"
      Else 'NOT NWALIGN...
        Err.Raise 380, UserControl.Name
    End If

End Property

Public Property Get Alignment() As Align

    Alignment = myAlignment

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets/Returns the Control's BackColor. If a PictureBody is used then this property is hidden underneath that picture. "
Attribute BackColor.VB_HelpID = 10002
Attribute BackColor.VB_UserMemId = -501

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(nwColor As OLE_COLOR)

    UserControl.BackColor = nwColor
    PropertyChanged "BackColor"

End Property

Public Property Let Borderstyle(nwStyle As Borderstyle)
Attribute Borderstyle.VB_Description = "Sets/Returns the borderstyle."
Attribute Borderstyle.VB_HelpID = 10003
Attribute Borderstyle.VB_UserMemId = -504

    If Not Sliding Then
        NoSizing = True
        If nwStyle = BorderNone Then
            Width = ScaleWidth
            Height = ScaleHeight
            UserControl.Borderstyle = BorderNone
            PropertyChanged "BorderStyle"
          ElseIf nwStyle = Border3D Then 'NOT NWSTYLE...
            Width = ScaleWidth + TpP * 4
            Height = ScaleHeight + TpP * 4
            UserControl.Borderstyle = Border3D
            PropertyChanged "BorderStyle"
          Else 'NOT NWSTYLE...
            Err.Raise 380, UserControl.Name
        End If
        NoSizing = False
    End If

End Property

Public Property Get Borderstyle() As Borderstyle

    Borderstyle = UserControl.Borderstyle

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the caption in the Control. The caption is merged into the PictureBody."
Attribute Caption.VB_HelpID = 10004
Attribute Caption.VB_UserMemId = -518

    Caption = Trim$(lblCapt)

End Property

Public Property Let Caption(nwCaption As String)

    If Len(Trim$(nwCaption)) Then
        lblCapt = nwCaption
      Else 'LEN(TRIM$(NWCAPTION)) = 0'LEN(TRIM$(NWCAPTION)) = FALSE
        lblCapt = Space$(8)
    End If
    PropertyChanged "Caption"
    Alignment = myAlignment

End Property

Public Property Set Font(ByVal nwFont As Font)
Attribute Font.VB_Description = "Sets/returns the font object for the Control."
Attribute Font.VB_HelpID = 10008
Attribute Font.VB_UserMemId = -512

    Set lblCapt.Font = nwFont
    Alignment = myAlignment
    WasUnderlined = lblCapt.Font.Underline

End Property

Public Property Get Font() As Font

    Set Font = lblCapt.Font

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets/returns the Control's foreground color."
Attribute ForeColor.VB_HelpID = 10009
Attribute ForeColor.VB_UserMemId = -513

    ForeColor = lblCapt.ForeColor

End Property

Public Property Let ForeColor(nwColor As OLE_COLOR)

    lblCapt.ForeColor = nwColor
    PropertyChanged "ForeColor"

End Property

Private Sub imgBody_Click()

    If imgBody = 0 Then
        UserControl_Click
      Else 'NOT IMGBODY...
        RaiseEvent Click(Body)
    End If

End Sub

Private Sub imgBody_DblClick()

    If imgBody = 0 Then
        UserControl_DblClick
      Else 'NOT IMGBODY...
        RaiseEvent DblClick(Body)
    End If

End Sub

Private Sub imgBody_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Value = (myMax - myMin) * X / imgBody.Width + myMin
    RaiseEvent Change(Value)

End Sub

Private Sub imgLeft_Click()

    If imgLeft = 0 Then
        imgBody_Click
      Else 'NOT IMGLEFT...
        If LegalClick Then
            RaiseEvent Click(BtnLeft)
        End If
    End If

End Sub

Private Sub imgLeft_DblClick()

    If imgLeft = 0 Then
        imgBody_DblClick
      Else 'NOT IMGLEFT...
        If LegalClick Then
            RaiseEvent DblClick(BtnLeft)
        End If
    End If

End Sub

Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If imgLeft <> 0 Then
        LegalClick = (myValue = myMin)
        imgThumb_MouseDown vbLeftButton, 0, imgThumb.Width / 2, 0
        Direction = -TpP
        tmrSlide.Enabled = True
    End If

End Sub

Private Sub imgLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If imgLeft <> 0 Then
        tmrSlide.Enabled = False
        imgThumb_MouseUp 0, 0, 0, 0
    End If

End Sub

Private Sub imgRight_Click()

    If imgRight = 0 Then
        imgBody_Click
      Else 'NOT IMGRIGHT...
        If LegalClick Then
            RaiseEvent Click(BtnRight)
        End If
    End If

End Sub

Private Sub imgRight_DblClick()

    If imgRight = 0 Then
        imgBody_DblClick
      Else 'NOT IMGRIGHT...
        If LegalClick Then
            RaiseEvent DblClick(BtnRight)
        End If
    End If

End Sub

Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If imgRight <> 0 Then
        LegalClick = (myValue = myMax)
        imgThumb_MouseDown vbLeftButton, 0, imgThumb.Width / 2, 0
        Direction = TpP
        tmrSlide.Enabled = True
    End If

End Sub

Private Sub imgRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If imgRight <> 0 Then
        tmrSlide.Enabled = False
        imgThumb_MouseUp 0, 0, 0, 0
    End If

End Sub

Private Sub imgThumb_Click()

    If Not Changed Then
        RaiseEvent Click(BtnThumb)
    End If

End Sub

Private Sub imgThumb_DblClick()

    If Not Changed Then
        RaiseEvent DblClick(BtnThumb)
    End If

End Sub

Private Sub imgThumb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Sliding = True
    Changed = False
    DownX = X

End Sub

Private Sub imgThumb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim NewPosn

    If Sliding Then
        NewPosn = imgThumb.Left + X - DownX
        If Abs(NewPosn - OldPosn) < NotchDistance / 2 Then
            NewPosn = OldPosn
          Else 'NOT ABS(NEWPOSN...
            If NotchDistance <> 0 Then
                NewPosn = imgThumb.Left + NotchDistance * Sgn(NewPosn - OldPosn)
            End If
        End If
        If NewPosn > XMax Then
            NewPosn = XMax
        End If
        If NewPosn < XMin Then
            NewPosn = XMin
        End If
        If NewPosn <> OldPosn Then
            imgThumb.Move NewPosn, imgThumb.Top
            myValue = ((NewPosn - XMin) / (XMax - XMin)) * (myMax - myMin) + myMin
            RaiseEvent Scroll(myValue)
            Changed = True
            Accu = 0
            OldPosn = NewPosn
        End If

    End If

End Sub

Private Sub imgThumb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Sliding = False
    If Changed Then
        RaiseEvent Change(myValue)
    End If

End Sub

Private Sub lblCapt_Click()

    imgBody_Click

End Sub

Private Sub lblCapt_DblClick()

    imgBody_DblClick

End Sub

Private Sub lblCapt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgBody_MouseDown Button, Shift, lblCapt.Left + X - imgLeft.Width, lblCapt.Top + Y

End Sub

Public Property Get Max() As Single
Attribute Max.VB_Description = "Value at right hand end of the slider."
Attribute Max.VB_HelpID = 10013

    Max = myMax

End Property

Public Property Let Max(nwMax As Single)

    If Not Sliding Then
        myMax = nwMax
        If ValueIsOut(myValue) Then
            Value = myMax
        End If
        Notches = myIntervals + 1
        Value = myValue
        PropertyChanged "VMax"
    End If

End Property

Public Property Get Min() As Single
Attribute Min.VB_Description = "Value at left hand end of the slider."
Attribute Min.VB_HelpID = 10014

    Min = myMin

End Property

Public Property Let Min(nwMin As Single)

    If Not Sliding Then
        myMin = nwMin
        If ValueIsOut(myValue) Then
            Value = myMin
        End If
        Notches = myIntervals + 1
        Value = myValue
        PropertyChanged "VMin"
    End If

End Property

Public Property Get Notches()
Attribute Notches.VB_Description = "Sets/returns the number of notches. Notches make the thumb move in distinct steps; eg for ten steps you need 11 notches - one at either end and nine in between."
Attribute Notches.VB_HelpID = 10015

    Notches = myIntervals + 1

End Property

Public Property Let Notches(nwNotches)

  Dim tmp1, tmp2

    If Not Sliding Then
        myIntervals = nwNotches - 1
        tmp1 = (XMax - XMin)
        tmp2 = tmp1 / TpP / 2
        If myIntervals > tmp2 Then
            myIntervals = tmp2
        End If
        If myIntervals < 1 Then
            myIntervals = -1
            NotchDistance = 0
          Else 'NOT MYINTERVALS...
            NotchDistance = tmp1 / myIntervals
        End If
        Value = myValue
        PropertyChanged "Notches"
    End If

End Property

Public Property Get PictureBody() As Picture
Attribute PictureBody.VB_Description = "Picture to be used as the body of the slider."
Attribute PictureBody.VB_HelpID = 10016

    Set PictureBody = imgBody

End Property

Public Property Set PictureBody(ByVal nwPic As Picture)

    If Not Ambient.UserMode Or ReadingProp Then
        Set imgBody = nwPic
        PropertyChanged "PictureBody"
    End If

End Property

Public Property Get PictureLeft() As Picture

    Set PictureLeft = imgLeft

End Property

Public Property Set PictureLeft(ByVal nwPic As Picture)

  Dim t As Boolean

    If Not Ambient.UserMode Or ReadingProp Then
        Set picTemp = nwPic
        t = (nwPic Is Nothing)
        If Not t Then
            t = (nwPic = 0)
        End If
        If t Then
            OrgLeftWidth = 0
          Else 'T = 0'T = FALSE
            OrgLeftWidth = picTemp.Width * ScaleHeight / picTemp.Height
        End If
        imgLeft.Width = Int((OrgLeftWidth + TpP2) / TpP) * TpP
        XMin = OrgLeftWidth
        Set imgLeft = nwPic
        UserControl_Resize
        PropertyChanged "PictureLeft"
    End If

End Property

Public Property Get PictureRight() As Picture

    Set PictureRight = imgRight

End Property

Public Property Set PictureRight(ByVal nwPic As Picture)

  Dim t As Boolean

    If Not Ambient.UserMode Or ReadingProp Then
        Set picTemp = nwPic
        t = (nwPic Is Nothing)
        If Not t Then
            t = (nwPic = 0)
        End If
        If t Then
            OrgRightWidth = 0
          Else 'T = 0'T = FALSE
            OrgRightWidth = picTemp.Width * ScaleHeight / picTemp.Height
        End If
        imgRight.Width = Int((OrgRightWidth + TpP2) / TpP) * TpP
        Set imgRight = nwPic
        UserControl_Resize
        PropertyChanged "PictureRight"
    End If

End Property

Public Property Get PictureThumb() As Picture
Attribute PictureThumb.VB_Description = "Picture to be used as the thumb."
Attribute PictureThumb.VB_HelpID = 10019

    Set PictureThumb = imgThumb

End Property

Public Property Set PictureThumb(ByVal nwPic As Picture)

    If Not Ambient.UserMode Or ReadingProp Then
        Set picTemp = nwPic
        OrgThumbWidth = picTemp.Width * ScaleHeight / picTemp.Height
        imgThumb.Width = Int((OrgThumbWidth + TpP2) / TpP) * TpP
        Set imgThumb = nwPic
        UserControl_Resize
        PropertyChanged "PictureThumb"
    End If

End Property

Public Property Get Resolution() As Single
Attribute Resolution.VB_Description = "Read only; returns the delta(value)  for one pixel thumb movement."
Attribute Resolution.VB_HelpID = 10020

    On Error Resume Next
      Resolution = Abs(myMax - myMin) / ((XMax - XMin) / TpP)
    On Error GoTo 0

End Property

Public Property Let Resolution(NwDist As Single)

  '   empty

End Property

Public Property Get ShowFocus() As Boolean

    ShowFocus = myFocus

End Property

Public Property Let ShowFocus(nwFocus As Boolean)

    myFocus = (nwFocus = True)
    If Ambient.UserMode And Not ReadingProp Then
        UserControl_EnterFocus
    End If
    PropertyChanged "ShowFocus"

End Property

Private Sub tmrSlide_Timer()

    Accu = Accu + Direction
    imgThumb_MouseMove vbLeftButton, 0, Accu + DownX, 0

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click(Background)

End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick(Background)

End Sub

Private Sub UserControl_EnterFocus()

    lblCapt.Font.Underline = myFocus Or WasUnderlined

End Sub

Private Sub UserControl_ExitFocus()

    lblCapt.Font.Underline = WasUnderlined

End Sub

Private Sub UserControl_Initialize()

    TpP = Screen.TwipsPerPixelX
    TpP2 = TpP * 0.499

End Sub

Private Sub UserControl_InitProperties()

    imgThumb.Left = 0
    myValue = 0
    myMin = 0
    myMax = 100
    Borderstyle = Border3D
    myAlignment = [Align Center]
    Set Font = Ambient.Font
    lblCapt = Ambient.DisplayName
    NoSizing = True
    Width = picTemp.Width + 104 * TpP
    Height = (GetSystemMetrics(SM_CYHSCROLL) + 5) * TpP
    NoSizing = False
    Set PictureThumb = picTemp.Picture

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ReadingProp = True
    With PropBag
        Set PictureBody = .ReadProperty("PictureBody", Nothing)
        Set PictureLeft = .ReadProperty("PictureLeft", Nothing)
        Set PictureThumb = .ReadProperty("PictureThumb", Nothing)
        Set PictureRight = .ReadProperty("PictureRite", Nothing)
        Min = .ReadProperty("VMin", 0)
        Max = .ReadProperty("VMax", 100)
        Notches = .ReadProperty("Notches", 0)
        Value = .ReadProperty("Value", 0)
        BackColor = .ReadProperty("BackColor", &HE0E0E0)
        ForeColor = .ReadProperty("ForeColor", &H80000008)
        Set lblCapt.Font = .ReadProperty("Font", Ambient.Font)
        WasUnderlined = lblCapt.Font.Underline
        Caption = .ReadProperty("Caption", "")
        Alignment = .ReadProperty("Alignment", [Align Center])
        ShowFocus = .ReadProperty("ShowFocus", True)
        UserControl.Borderstyle = .ReadProperty("BorderStyle", BorderNone)
    End With 'PROPBAG
    ReadingProp = False

End Sub

Private Sub UserControl_Resize()

  Dim sizeFactor As Single

    If Not NoSizing Then
        If Height < 210 Then
            Size Width, 210
        End If
        If CurrentHeight = 0 Then
            CurrentHeight = ScaleHeight
        End If
        If ScaleHeight <> CurrentHeight Then
            sizeFactor = ScaleHeight / CurrentHeight
            OrgThumbWidth = OrgThumbWidth * sizeFactor
            OrgLeftWidth = OrgLeftWidth * sizeFactor
            OrgRightWidth = OrgRightWidth * sizeFactor
            imgThumb.Width = Int((OrgThumbWidth + TpP2) / TpP) * TpP
            imgLeft.Width = Int((OrgLeftWidth + TpP2) / TpP) * TpP
            imgRight.Width = Int((OrgRightWidth + TpP2) / TpP) * TpP
        End If
        imgBody.Width = ScaleWidth - imgLeft.Width * Sgn(imgLeft) - imgRight.Width * Sgn(imgRight)
        imgBody.Left = imgLeft.Width * Sgn(imgLeft)
        imgBody.Height = ScaleHeight
        imgThumb.Height = ScaleHeight
        imgLeft.Height = ScaleHeight
        imgRight.Height = ScaleHeight
        XMin = IIf(imgLeft.Picture = 0, 0, imgLeft.Width)
        XMax = ScaleWidth - imgRight.Width + IIf(imgRight.Picture = 0, 15, 0)
        imgRight.Left = XMax
        XMax = XMax - imgThumb.Width
        Value = myValue
        Alignment = myAlignment
        Notches = myIntervals + 1
        CurrentHeight = ScaleHeight
    End If

End Sub

Private Sub UserControl_Show()

    Alignment = myAlignment

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "PictureBody", imgBody, Nothing
        .WriteProperty "PictureThumb", imgThumb, Nothing
        .WriteProperty "PictureLeft", imgLeft, Nothing
        .WriteProperty "PictureRite", imgRight, Nothing
        .WriteProperty "Notches", Notches, 0
        .WriteProperty "VMin", myMin, 0
        .WriteProperty "VMax", myMax, 100
        .WriteProperty "Value", myValue, 0
        .WriteProperty "BackColor", BackColor, &HE0E0E0
        .WriteProperty "ForeColor", lblCapt.ForeColor, &H80000008
        .WriteProperty "Alignment", myAlignment, [Align Center]
        .WriteProperty "Caption", lblCapt.Caption, ""
        .WriteProperty "Font", lblCapt.Font, Ambient.Font
        .WriteProperty "ShowFocus", myFocus, True
        .WriteProperty "BorderStyle", Borderstyle, BorderNone
    End With 'PROPBAG

End Sub

Public Property Get Value() As Single
Attribute Value.VB_Description = "Sets the value and moves the thumb / returns the value at current the humb position."
Attribute Value.VB_HelpID = 10022
Attribute Value.VB_UserMemId = 0

    Value = myValue

End Property

Public Property Let Value(nwValue As Single)

  Dim tmp As Single

    If Not Sliding Then
        If ValueIsOut(nwValue) Then
            Err.Raise 380, UserControl.Name
        End If
        If NotchDistance = 0 Then
            myValue = nwValue
          Else 'NOT NOTCHDISTANCE...
            tmp = (myMax - myMin) / myIntervals
            myValue = Int(nwValue / tmp + 0.5) * tmp
            myValue = Int(myValue / Resolution + 0.5) * Resolution
        End If
        If myMin <> myMax Then
            imgThumb.Left = (myValue - myMin) / (myMax - myMin) * (XMax - XMin) + XMin
        End If
        OldPosn = imgThumb.Left
        PropertyChanged "Value"
    End If

End Property

Private Function ValueIsOut(Val As Single) As Boolean

    Select Case True
      Case myMin < myMax
        ValueIsOut = (Val < myMin Or Val > myMax)
      Case myMin = myMax
        ValueIsOut = (Val <> myMin)
      Case myMin > myMax
        ValueIsOut = (Val < myMax Or Val > myMin)
    End Select

End Function

':) Ulli's VB Code Formatter V2.10.8 (12.03.2002 21:19:16) 73 + 714 = 787 Lines
