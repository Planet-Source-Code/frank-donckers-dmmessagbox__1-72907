VERSION 5.00
Begin VB.UserControl DMmsgButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ControlContainer=   -1  'True
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   5880
   ToolboxBitmap   =   "DMmsgButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3510
      Top             =   2460
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMCommandButton"
         Height          =   195
         Left            =   960
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
      Begin VB.Image imgRight 
         Height          =   375
         Left            =   2400
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgLeft 
         Height          =   495
         Left            =   1320
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgBackGround 
         Height          =   495
         Left            =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblshadow 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DMCommandButton"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Image imgBackGroundDown 
      Height          =   375
      Left            =   840
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image imgBackGroundEn 
      Height          =   495
      Left            =   840
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image imgLeftEn 
      Height          =   495
      Left            =   2160
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image imgRightEn 
      Height          =   375
      Left            =   3240
      Top             =   3960
      Width           =   495
   End
   Begin VB.Image imgBackGroundDi 
      Height          =   495
      Left            =   840
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image imgLeftDi 
      Height          =   495
      Left            =   2160
      Top             =   3120
      Width           =   495
   End
   Begin VB.Image imgRightDi 
      Height          =   375
      Left            =   3240
      Top             =   3240
      Width           =   495
   End
End
Attribute VB_Name = "DMmsgButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Enums:
'=====================================================
'POINTAPI
'=====================================================
Private Type POINTAPI
    X As Long
    Y As Long
End Type

'=====================================================
'RECTANGLE
'=====================================================
Private Type RECT
    rLeft    As Long
    rTop     As Long
    rRight   As Long
    rBottom  As Long
End Type


Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'=====================================================
'ALIGNEMENT
'=====================================================
Public Enum Alignment
    [Alignleft] = 0
    [AlignCenter] = 1
    [AlignRight] = 2
    [AlignTop] = 3
    [AlignBottom] = 4
End Enum
'Events
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Default Property Values:
Const m_def_Alignment = 0
Const m_def_HooverColor = 0
Const m_def_DisabledColor = &H8000000F
Const m_def_Default = False
Const m_def_Enabled = True
'Property Variables:
Dim m_Alignment As Alignment
Dim m_HooverColor As OLE_COLOR
Dim m_DisabledColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_ShadowColor As OLE_COLOR
Dim m_Caption As String
Dim m_Default As Boolean
Dim m_Enabled As Boolean
Dim m_PictureBackGround As Picture
Dim m_PictureBackGroundDown As Picture
Dim m_PictureLeft As Picture
Dim m_PictureRight As Picture
Dim m_PictureBackGroundDisabled As Picture
Dim m_PictureLeftDisabled As Picture
Dim m_PictureRightDisabled As Picture
'other variables
Private OldScaleMode As Byte
Private cControl As Control
Private i, ii, iii, iiii As Integer
Private HighLighted As Boolean

'DRAW CAPTION AND PICTURE
'==========================================================
Public Sub DrawFront()
    Dim lLeft, lTop As Double
        Select Case Alignment
            Case Alignleft
                lblCaption.Left = 90
                lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
            Case AlignCenter
                lblCaption.Left = UserControl.Width / 2 - (lblCaption.Width / 2)
                lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
            Case AlignRight
                lblCaption.Left = UserControl.Width - lblCaption.Width - 15
                lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
            Case AlignTop
                lblCaption.Left = UserControl.Width / 2 - (lblCaption.Width / 2)
                lblCaption.Top = 90
            Case AlignBottom
                lblCaption.Left = UserControl.Width / 2 - (lblCaption.Width / 2)
                lblCaption.Top = UserControl.Height - lblCaption.Height - 15
        End Select
        If Enabled = True Then
            lblCaption.ForeColor = m_ForeColor
            lblShadow.Visible = True
            lblShadow.Left = lblCaption.Left + 15
            lblShadow.Top = lblCaption.Top + 15
        Else
            lblCaption.ForeColor = m_DisabledColor
            lblShadow.Visible = False
        End If
End Sub

'DRAW THE BACKGROUND
'==========================================================
Private Sub DrawBack()
    On Error Resume Next
    picBack.Left = 0
    picBack.Top = 0
    picBack.Width = UserControl.Width
    picBack.Height = UserControl.Height
    DoEvents
    imgLeft.Top = 0
    imgLeft.Left = 0
    imgBackGround.Top = 0
    If imgLeft.Picture <> 0 Then imgBackGround.Left = imgLeft.Width
    imgRight.Top = 0
    imgRight.Left = UserControl.Width - imgRight.Width
    setEnabled
End Sub
Private Sub cmdCommand_Click()
    picBack_Click
End Sub

Private Sub lblCaption_Click()
    picBack_Click
End Sub
Private Sub lblShadow_Click()
    picBack_Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseUp Button, Shift, X, Y
End Sub
Private Sub lblShadown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblShadown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseMove Button, Shift, X, Y
End Sub

Private Sub lblShadown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack_MouseUp Button, Shift, X, Y
End Sub
Private Sub picBack_Click()
    RaiseEvent Click
End Sub


Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack.Cls
    On Error GoTo ErrorHandle
    picBack.PaintPicture imgBackGroundDown.Picture, imgBackGround.Left, imgBackGround.Top
    GoTo OK
ErrorHandle:
    picBack.PaintPicture imgBackGround.Picture, imgBackGround.Left, imgBackGround.Top
OK:
    picBack.PaintPicture imgLeft.Picture, imgLeft.Left, imgLeft.Top
    picBack.PaintPicture imgRight.Picture, imgRight.Left, imgRight.Top
    Timer1.Enabled = True
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If HighLighted Then Exit Sub
        
    HighLighted = True
    ' Top Line Highlighted
    picBack.Line (0, 0)-(picBack.Width, 0), m_HooverColor         ' Outside Border
    ' Left Line Highlighted
    picBack.Line (0, 15)-(0, picBack.Height), m_HooverColor       ' Outside Border
    ' Right Line Highlighted
    picBack.Line (picBack.Width - 15, 0)-(picBack.Width - 15, picBack.Height), m_HooverColor   ' Outside Border
    ' Bottom Line Highlighted
    picBack.Line (0, picBack.Height - 15)-(picBack.Width - 15, picBack.Height - 15), m_HooverColor   ' Outside Border
    Timer1.Enabled = True

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBack.Cls
    picBack.PaintPicture imgBackGround.Picture, imgBackGround.Left, imgBackGround.Top
    picBack.PaintPicture imgLeft.Picture, imgLeft.Left, imgLeft.Top
    picBack.PaintPicture imgRight.Picture, imgRight.Left, imgRight.Top
    If HighLighted = False Then
        DrawBack
        Exit Sub
    End If
    HighLighted = False
    Timer1.Enabled = True
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI
    ' See where the cursor is.
    GetCursorPos pt
    ' Translate into window coordinates.
    If WindowFromPointXY(pt.X, pt.Y) <> picBack.hwnd _
        Then
        HighLighted = False
        DrawBack
        Timer1.Enabled = False
    End If
End Sub
Private Sub setEnabled()
    On Error Resume Next
    If Enabled = False Then
        imgBackGround.Picture = imgBackGroundDi.Picture
        imgLeft.Picture = imgLeftDi.Picture
        imgRight.Picture = imgRightDi.Picture
    Else
        imgBackGround.Picture = imgBackGroundEn.Picture
        imgLeft.Picture = imgLeftEn.Picture
        imgRight.Picture = imgRightEn.Picture
    End If
    picBack.Cls
    picBack.PaintPicture imgBackGround.Picture, imgBackGround.Left, imgBackGround.Top
    picBack.PaintPicture imgLeft.Picture, imgLeft.Left, imgLeft.Top
    picBack.PaintPicture imgRight.Picture, imgRight.Left, imgRight.Top
    'picBack.Refresh
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    HighLighted = False
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    HighLighted = False
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If imgBackGround.Picture <> 0 Then
        If UserControl.Height > imgBackGround.Height Then UserControl.Height = imgBackGround.Height
    End If
    DrawFront
    DrawBack
End Sub

'=======================================================================================================
' USERCONTROL PROPERTIES
'=======================================================================================================

' USERCONTROL INIT PROPERTIES
'==========================================================
Private Sub UserControl_InitProperties()
    m_HooverColor = m_def_HooverColor
    m_DisabledColor = m_def_DisabledColor
    m_Alignment = m_def_Alignment
    UserControl_Resize
    m_Enabled = m_def_Enabled
    Set m_PictureBackGround = LoadPicture("")
    Set m_PictureLeft = LoadPicture("")
    Set m_PictureRight = LoadPicture("")
    Set m_PictureBackGroundDisabled = LoadPicture("")
    Set m_PictureLeftDisabled = LoadPicture("")
    Set m_PictureRightDisabled = LoadPicture("")
    setEnabled
End Sub

' USERCONTROL READ PROPERTIES
'==========================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblShadow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "DMCommandButton")
    lblShadow.Caption = lblCaption.Caption
    m_Caption = PropBag.ReadProperty("Caption", "DMCommandButton")
    m_Default = PropBag.ReadProperty("Default", m_def_Default)
    m_HooverColor = PropBag.ReadProperty("HooverColor", m_def_HooverColor)
    m_DisabledColor = PropBag.ReadProperty("DisabledColor", m_def_DisabledColor)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    Set DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Set DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_PictureBackGround = PropBag.ReadProperty("PictureBackGround", Nothing)
    Set m_PictureBackGroundDown = PropBag.ReadProperty("PictureBackGroundDown", Nothing)
    Set imgBackGroundDown.Picture = m_PictureBackGroundDown
    'Set imgBackGround.Picture = m_PictureBackGround
    Set imgBackGroundEn.Picture = m_PictureBackGround
    Set m_PictureLeft = PropBag.ReadProperty("PictureLeft", Nothing)
    'Set imgLeft.Picture = m_PictureLeft
    Set imgLeftEn.Picture = m_PictureLeft
    Set m_PictureRight = PropBag.ReadProperty("PictureRight", Nothing)
    'Set imgRight.Picture = m_PictureRight
    Set imgRightEn.Picture = m_PictureRight
    
    Set m_PictureBackGroundDisabled = PropBag.ReadProperty("PictureBackGroundDisabled", Nothing)
    Set imgBackGroundDi.Picture = m_PictureBackGroundDisabled
    Set m_PictureLeftDisabled = PropBag.ReadProperty("PictureLeftDisabled", Nothing)
    Set imgLeftDi.Picture = m_PictureLeftDisabled
    Set m_PictureRightDisabled = PropBag.ReadProperty("PictureRightDisabled", Nothing)
    Set imgRightDi.Picture = m_PictureRightDisabled
    setEnabled
    UserControl_Resize
    UserControl.Enabled = Enabled
End Sub

' USERCONTROL WRITE PROPERTIES
'==========================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "DMCommandButton")
    Call PropBag.WriteProperty("Default", m_Default, m_def_Default)
    Call PropBag.WriteProperty("HooverColor", m_HooverColor, m_def_HooverColor)
    Call PropBag.WriteProperty("DisabledColor", m_DisabledColor, m_def_DisabledColor)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("DisabledPicture", DisabledPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", DownPicture, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("PictureBackGround", m_PictureBackGround, Nothing)
    Call PropBag.WriteProperty("PictureBackGroundDown", m_PictureBackGroundDown, Nothing)
    Call PropBag.WriteProperty("PictureLeft", m_PictureLeft, Nothing)
    Call PropBag.WriteProperty("PictureRight", m_PictureRight, Nothing)
    Call PropBag.WriteProperty("PictureBackGroundDisabled", m_PictureBackGroundDisabled, Nothing)
    Call PropBag.WriteProperty("PictureLeftDisabled", m_PictureLeftDisabled, Nothing)
    Call PropBag.WriteProperty("PictureRightDisabled", m_PictureRightDisabled, Nothing)
End Sub

'
'
'
''=======================================================================================================
''USER CONTROL GET AND LET
''=======================================================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    m_ForeColor = New_ForeColor
    UserControl_Resize
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = lblShadow.ForeColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    lblShadow.ForeColor() = New_ShadowColor
    PropertyChanged "ShadowColor"
    m_ShadowColor = New_ShadowColor
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Set lblShadow.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    lblShadow.Caption() = New_Caption
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdCommand,cmdCommand,-1,Default
Public Property Get Default() As Boolean
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
    m_Default = New_Default
    PropertyChanged "Default"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get HooverColor() As OLE_COLOR
    HooverColor = m_HooverColor
End Property

Public Property Let HooverColor(ByVal New_HooverColor As OLE_COLOR)
    m_HooverColor = New_HooverColor
    PropertyChanged "HooverColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DisabledColor() As OLE_COLOR
    DisabledColor = m_DisabledColor
End Property

Public Property Let DisabledColor(ByVal New_DisabledColor As OLE_COLOR)
    m_DisabledColor = New_DisabledColor
    PropertyChanged "DisabledColor"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Alignment() As Alignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Alignment)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdCommand,cmdCommand,-1,DisabledPicture
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is disabled, if Style is set to 1."
    Set DisabledPicture = cmdCommand.DisabledPicture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set cmdCommand.DisabledPicture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdCommand,cmdCommand,-1,DownPicture
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is in the down position, if Style is set to 1."
    Set DownPicture = cmdCommand.DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
    Set cmdCommand.DownPicture = New_DownPicture
    PropertyChanged "DownPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdCommand,cmdCommand,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a CommandButton, OptionButton or CheckBox control, if Style is set to 1."
    Set Picture = cmdCommand.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set cmdCommand.Picture = New_Picture
    PropertyChanged "Picture"
    UserControl_Resize
End Property
'
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl.Enabled = Enabled
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBackGround() As Picture
    Set PictureBackGround = m_PictureBackGround
End Property
Public Property Set PictureBackGround(ByVal New_PictureBackGround As Picture)
    Set m_PictureBackGround = New_PictureBackGround
    PropertyChanged "PictureBackGround"
    On Error Resume Next
    Set imgBackGroundEn.Picture = m_PictureBackGround
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBackGroundDown() As Picture
    Set PictureBackGroundDown = m_PictureBackGroundDown
End Property
Public Property Set PictureBackGroundDown(ByVal New_PictureBackGroundDown As Picture)
    Set m_PictureBackGroundDown = New_PictureBackGroundDown
    PropertyChanged "PictureBackGroundDown"
    On Error Resume Next
    Set imgBackGroundDown.Picture = m_PictureBackGroundDown
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureLeft() As Picture
    Set PictureLeft = m_PictureLeft
End Property
Public Property Set PictureLeft(ByVal New_PictureLeft As Picture)
    Set m_PictureLeft = New_PictureLeft
    PropertyChanged "PictureLeft"
    On Error Resume Next
    Set imgLeftEn.Picture = m_PictureLeft
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureRight() As Picture
    Set PictureRight = m_PictureRight
End Property
Public Property Set PictureRight(ByVal New_PictureRight As Picture)
    Set m_PictureRight = New_PictureRight
    PropertyChanged "PictureRight"
    On Error Resume Next
    Set imgRightEn.Picture = m_PictureRight
    UserControl_Resize
End Property




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBackGroundDisabled() As Picture
    Set PictureBackGroundDisabled = m_PictureBackGroundDisabled
End Property
Public Property Set PictureBackGroundDisabled(ByVal New_PictureBackGroundDisabled As Picture)
    Set m_PictureBackGroundDisabled = New_PictureBackGroundDisabled
    PropertyChanged "PictureBackGroundDisabled"
    On Error Resume Next
    Set imgBackGroundDi.Picture = m_PictureBackGroundDisabled
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureLeftDisabled() As Picture
    Set PictureLeftDisabled = m_PictureLeftDisabled
End Property
Public Property Set PictureLeftDisabled(ByVal New_PictureLeftDisabled As Picture)
    Set m_PictureLeftDisabled = New_PictureLeftDisabled
    PropertyChanged "PictureLeftDisabled"
    On Error Resume Next
    Set imgLeftDi.Picture = m_PictureLeftDisabled
    UserControl_Resize
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureRightDisabled() As Picture
    Set PictureRightDisabled = m_PictureRightDisabled
End Property
Public Property Set PictureRightDisabled(ByVal New_PictureRightDisabled As Picture)
    Set m_PictureRightDisabled = New_PictureRightDisabled
    PropertyChanged "PictureRightDisabled"
    On Error Resume Next
    Set imgRightDi.Picture = m_PictureRightDisabled
    UserControl_Resize
End Property

