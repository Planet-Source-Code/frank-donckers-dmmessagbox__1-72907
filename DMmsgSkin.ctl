VERSION 5.00
Begin VB.UserControl DMmsgSkin 
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   10260
   ToolboxBitmap   =   "DMmsgSkin.ctx":0000
   Begin VB.PictureBox imgClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   2520
      Picture         =   "DMmsgSkin.ctx":0312
      ScaleHeight     =   420
      ScaleWidth      =   240
      TabIndex        =   4
      ToolTipText     =   "Close window"
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox imgMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1800
      Picture         =   "DMmsgSkin.ctx":0914
      ScaleHeight     =   420
      ScaleWidth      =   240
      TabIndex        =   3
      ToolTipText     =   "Maximize window"
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox imgMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1080
      Picture         =   "DMmsgSkin.ctx":0F16
      ScaleHeight     =   420
      ScaleWidth      =   240
      TabIndex        =   2
      ToolTipText     =   "Minimize window"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "DMmsgSkin.ctx":1518
      Top             =   0
      Width           =   420
   End
   Begin VB.Image StandardimgBackGround 
      Height          =   555
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   420
   End
   Begin VB.Image StandardpicExitHoover 
      Height          =   420
      Left            =   6600
      Picture         =   "DMmsgSkin.ctx":1D8E
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMaxHoover 
      Height          =   420
      Left            =   6000
      Picture         =   "DMmsgSkin.ctx":2390
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMinHoover 
      Height          =   420
      Left            =   5400
      Picture         =   "DMmsgSkin.ctx":2992
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicExitDown 
      Height          =   420
      Left            =   6600
      Picture         =   "DMmsgSkin.ctx":2F94
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardPicExitUp 
      Height          =   420
      Left            =   6600
      Picture         =   "DMmsgSkin.ctx":3596
      Top             =   4320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMaxDown 
      Height          =   420
      Left            =   6000
      Picture         =   "DMmsgSkin.ctx":3B98
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMaxUp 
      Height          =   420
      Left            =   6000
      Picture         =   "DMmsgSkin.ctx":419A
      Top             =   4320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMinDown 
      Height          =   420
      Left            =   5400
      Picture         =   "DMmsgSkin.ctx":479C
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardpicMinUp 
      Height          =   420
      Left            =   5400
      Picture         =   "DMmsgSkin.ctx":4D9E
      Top             =   4320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image StandardimgMidRightMask 
      Height          =   135
      Left            =   5040
      Picture         =   "DMmsgSkin.ctx":53A0
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image StandardimgMidButtomMask 
      Height          =   135
      Left            =   4440
      Picture         =   "DMmsgSkin.ctx":584E
      Top             =   5640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image StandardimgMidLeftMask 
      Height          =   135
      Left            =   4200
      Picture         =   "DMmsgSkin.ctx":5CFC
      Top             =   5160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image StandardimgMidTopMask 
      Height          =   555
      Left            =   3600
      Picture         =   "DMmsgSkin.ctx":61AA
      Top             =   3840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image StandardimgLeftBottomMask 
      Height          =   135
      Left            =   4200
      Picture         =   "DMmsgSkin.ctx":6964
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image StandardimgRightBottomMask 
      Height          =   135
      Left            =   5040
      Picture         =   "DMmsgSkin.ctx":6E12
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image StandardimgRightTopMask 
      Height          =   555
      Left            =   3960
      Picture         =   "DMmsgSkin.ctx":72C0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image StandardimgLeftTopMask 
      Height          =   1170
      Left            =   2160
      Picture         =   "DMmsgSkin.ctx":83BA
      Top             =   3840
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM_Skined_Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Label lblShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM_Skined_Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Image picExitHoover 
      Height          =   420
      Left            =   8280
      Picture         =   "DMmsgSkin.ctx":BE7C
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMaxHoover 
      Height          =   420
      Left            =   7560
      Picture         =   "DMmsgSkin.ctx":C47E
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMinHoover 
      Height          =   420
      Left            =   6840
      Picture         =   "DMmsgSkin.ctx":CA80
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picExitDown 
      Height          =   420
      Left            =   8160
      Picture         =   "DMmsgSkin.ctx":D082
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PicExitUp 
      Height          =   420
      Left            =   8160
      Picture         =   "DMmsgSkin.ctx":D684
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMaxDown 
      Height          =   420
      Left            =   7560
      Picture         =   "DMmsgSkin.ctx":DC86
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMaxUp 
      Height          =   420
      Left            =   7560
      Picture         =   "DMmsgSkin.ctx":E288
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMinDown 
      Height          =   420
      Left            =   6960
      Picture         =   "DMmsgSkin.ctx":E88A
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image picMinUp 
      Height          =   420
      Left            =   6960
      Picture         =   "DMmsgSkin.ctx":EE8C
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgMidRightMask 
      Height          =   135
      Left            =   5160
      Picture         =   "DMmsgSkin.ctx":F48E
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgMidButtomMask 
      Height          =   135
      Left            =   4560
      Picture         =   "DMmsgSkin.ctx":F93C
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgMidLeftMask 
      Height          =   135
      Left            =   4320
      Picture         =   "DMmsgSkin.ctx":FDEA
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgMidTopMask 
      Height          =   555
      Left            =   4680
      Picture         =   "DMmsgSkin.ctx":10298
      Top             =   720
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLeftBottomMask 
      Height          =   135
      Left            =   4320
      Picture         =   "DMmsgSkin.ctx":10A52
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgRightBottomMask 
      Height          =   135
      Left            =   5160
      Picture         =   "DMmsgSkin.ctx":10F00
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgRightTopMask 
      Height          =   555
      Left            =   5160
      Picture         =   "DMmsgSkin.ctx":113AE
      Top             =   720
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Image imgLeftTopMask 
      Height          =   1170
      Left            =   3240
      Picture         =   "DMmsgSkin.ctx":124A8
      Top             =   480
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Image imgBackGround 
      Height          =   555
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "DMmsgSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private MotherForm As Form
Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1
'=====================================================
'TYPE RECTANGLE
'=====================================================
Private Type RECT
    rLeft       As Long
    rTop        As Long
    rRight      As Long
    rBottom     As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
' Position of mouse
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Make a Semi Transparent Form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, pSize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pBlend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
' Move and resize a Titleless Window
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112

' region combine consts
Private Const RGN_AND = 1 'Combines an intersection
Private Const RGN_OR = 2 'Creates a union of two regions
Private Const RGN_XOR = 3 'Creations a union of two objects with the exception of overlapping
Private Const RGN_DIFF = 4 'Combines two regions
Private Const RGN_COPY = 5 'Copy a region

' region API declarations
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

' declarations for retrieving colors
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' Paint Circular Gradient
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

' Show a Form in the Taskbar
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_APPWINDOW = &H40000

Const SW_HIDE = 0
Const SW_NORMAL = 1

'Default Property Values:
'=====================================================
'POINTAPI
'=====================================================
Private Type POINTAPI
        X As Long
        Y As Long
End Type
'=====================================================
'CAPTIONSTYLE
'=====================================================
Public Enum CaptionStyles
    StyleNone = 0
    StyleSunken = 1
    StyleRaised = 2
End Enum
'=====================================================
'ACTIONONCLOSE
'=====================================================
Public Enum ActionOnClosing
    EndApp = 0
    HideForm = 1
    UnloadForm = 2
End Enum

Const m_def_CaptionForeColor = vbWhite '&HC0C0C0
Const m_def_CaptionShadowColor = &HFFC0C0   '&H404040
Const m_def_Caption = "DM_Skined_Form"
Const m_def_CaptionLeft = 120
Const m_def_CaptionTop = 190
Const m_def_CaptionStyle = 0
'Const m_def_BackColor = &H8000000F
Const m_def_Opacity = 0
Const m_def_MinButton = True
Const m_def_MaxButton = True
Const m_def_CloseButton = True
Const m_Borderwidth = 5 '5
Const m_def_Sizable = True
Const m_def_LeftFromRightClose = 780
Const m_def_LeftFromRightMaximize = 1350
Const m_def_LeftFromRightMiniMize = 1915
Const m_def_TopClose = 30
Const m_def_TopMaximize = 30
Const m_def_TopMiniMize = 30
Const m_def_ActionOnClose = 0
Const m_def_BackgroundStartColor = &H808080
Const m_def_BackgroundEndColor = 0
'Property Variables:
Dim m_PictureExitHoover As Picture
Dim m_PictureMaxHoover As Picture
Dim m_PictureMinHoover As Picture
Dim m_PictureTopLeft As Picture
Dim m_PictureBackGround As Picture
Dim m_PictureTopRight As Picture
Dim m_PictureTopMiddle As Picture
Dim m_PictureBottomLeft As Picture
Dim m_PictureBottomRight As Picture
Dim m_PictureBottomMiddle As Picture
Dim m_PictureMiddleLeft As Picture
Dim m_PictureMiddleRight As Picture
Dim m_PictureMinUp As Picture
Dim m_PictureMinDown As Picture
Dim m_PictureMaxUp As Picture
Dim m_PictureMaxDown As Picture
Dim m_PictureExitUp As Picture
Dim m_PictureExitDown As Picture
'Dim m_BackColor As OLE_COLOR
Dim m_Opacity As Byte
Dim m_Caption As String
Dim m_MinButton As Boolean
Dim m_MaxButton As Boolean
Dim m_CloseButton As Boolean
Dim m_CaptionStyle As CaptionStyles
Dim m_CaptionShadowColor As OLE_COLOR
Dim m_CaptionLeft As Long
Dim m_CaptionTop As Long
Dim m_CaptionForeColor As OLE_COLOR
Dim m_Sizable As Boolean
Dim m_LeftFromRightClose As Long
Dim m_LeftFromRightMaximize As Long
Dim m_LeftFromRightMiniMize As Long
Dim m_TopClose As Long
Dim m_TopMaximize As Long
Dim m_TopMiniMize As Long
Dim m_ActionOnClose As ActionOnClosing
Dim m_BackGroundStartColor As OLE_COLOR
Dim m_BackGroundEndColor As OLE_COLOR

' other
Dim OldScaleMode As Byte
Dim RT As Integer
Dim ButtonPos As Long
Dim PosFrom1, PosTo1, PosFrom2, PosTo2 As Long
Private rcGripperLT As RECT, rcGripperRT As RECT, rcGripperLB As RECT, rcGripperRB As RECT
Private XLresize As Boolean
Private XRresize As Boolean
Private YTresize As Boolean
Private YBresize As Boolean
Private OldFont As Font
Private OldOpacity As Byte



' HANDLING CONTROLBOX CONTROLS
'=====================================================
Public Sub imgClose_Click()
    Dim NewOpacity
    Select Case m_ActionOnClose
        Case EndApp
            Unload UserControl.Parent
        Case HideForm
            UserControl.Parent.Hide
        Case UnloadForm
            Unload UserControl.Parent
    End Select
End Sub

Public Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    UserControl.Parent.imgClose.Picture = picExitDown.Picture
    UserControl.Parent.imgClose.Refresh
End Sub

Public Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If X > 30 And X < imgClose.Width - 30 And Y > 30 And Y < imgClose.Height - 30 Then
        UserControl.Parent.imgClose.Picture = picExitHoover.Picture
    Else
        UserControl.Parent.imgClose.Picture = PicExitUp.Picture
    End If
    If m_MaxButton = True Then
        UserControl.Parent.imgMax.Picture = picMaxUp.Picture
        UserControl.Parent.imgMax.Refresh
    End If
    If m_MinButton = True Then
        UserControl.Parent.imgMin.Picture = picMinUp.Picture
        UserControl.Parent.imgMin.Refresh
    End If
    UserControl.Parent.imgClose.Refresh
End Sub

Public Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    UserControl.Parent.imgClose.Picture = PicExitUp.Picture
    UserControl.Parent.imgClose.Refresh
End Sub

Public Sub imgMax_Click()
    If UserControl.Parent.WindowState = vbMaximized Then
        UserControl.Parent.WindowState = vbNormal
    Else
        UserControl.Parent.WindowState = vbMaximized
    End If
        RepaintSkin UserControl.Parent
End Sub

Public Sub imgMax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Parent.imgMax.Picture = picMaxDown.Picture
    UserControl.Parent.imgMax.Refresh
End Sub

Public Sub imgMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 30 And X < imgMax.Width - 30 And Y > 30 And Y < imgMax.Height - 30 Then
        UserControl.Parent.imgMax.Picture = picMaxHoover.Picture
    Else
        UserControl.Parent.imgMax.Picture = picMaxUp.Picture
    End If
    If m_CloseButton = True Then
        UserControl.Parent.imgClose.Picture = PicExitUp.Picture
        UserControl.Parent.imgClose.Refresh
    End If
    If m_MinButton = True Then
        UserControl.Parent.imgMin.Picture = picMinUp.Picture
        UserControl.Parent.imgMin.Refresh
    End If
    UserControl.Parent.imgMax.Refresh
End Sub

Public Sub imgMax_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Parent.imgMax.Picture = picMaxUp.Picture
    UserControl.Parent.imgMax.Refresh
End Sub

Public Sub imgMin_Click()
    ShowInTheTaskbar UserControl.Parent.hwnd, True
    UserControl.Parent.WindowState = vbMinimized
End Sub

Public Sub imgMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Parent.imgMin.Picture = picMinDown.Picture
    UserControl.Parent.imgMin.Refresh
End Sub

Public Sub imgMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > 30 And X < imgMin.Width - 30 And Y > 30 And Y < imgMin.Height - 30 Then
        UserControl.Parent.imgMin.Picture = picMinHoover.Picture
    Else
        UserControl.Parent.imgMin.Picture = picMinUp.Picture
    End If
    If m_CloseButton = True Then
        UserControl.Parent.imgClose.Picture = PicExitUp.Picture
        UserControl.Parent.imgClose.Refresh
    End If
    If m_MaxButton = True Then
        UserControl.Parent.imgMax.Picture = picMaxUp.Picture
        UserControl.Parent.imgMax.Refresh
    End If
    UserControl.Parent.imgMin.Refresh
End Sub

Public Sub imgMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.Parent.imgMin.Picture = picMinUp.Picture
    UserControl.Parent.imgMin.Refresh
End Sub

Private Sub moveForm_Resize()
 RepaintSkin UserControl.Parent
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    'UserControl.Extender.Align = vbAlignTop
    'RepaintSkin
End Sub


' POSITIONING CONTROLBOXMEMBERS
'=====================================================
Private Sub SetControlBox(MotherForm As Form)
    On Error Resume Next
    MotherForm.imgMin.Top = 0
    MotherForm.imgMax.Top = 0
    MotherForm.imgClose.Top = 0
    If m_CloseButton = True Then
        MotherForm.imgClose.Left = MotherForm.Width - m_LeftFromRightClose
        MotherForm.imgClose.Top = m_TopClose
        MotherForm.imgClose.Visible = True
    Else
        MotherForm.imgClose.Visible = False
    End If
    If m_MaxButton = True And m_Sizable = True Then
        MotherForm.imgMax.Visible = True
        MotherForm.imgMax.Left = MotherForm.Width - m_LeftFromRightMaximize
        MotherForm.imgMax.Top = m_TopMaximize
    Else
        MotherForm.imgMax.Visible = False
    End If
    If m_MinButton = True And m_Sizable = True Then
            MotherForm.imgMin.Visible = True
            MotherForm.imgMin.Top = m_TopMiniMize
        If m_MaxButton = False Then
            MotherForm.imgMin.Left = MotherForm.Width - m_LeftFromRightMaximize
        Else
            MotherForm.imgMin.Left = MotherForm.Width - m_LeftFromRightMiniMize
        End If
    Else
        MotherForm.imgMin.Visible = False
   End If
End Sub

' REPAINT USERCONTROL
'=====================================================
Public Sub RepaintSkin(MotherForm)
    On Error Resume Next
    Set OldFont = UserControl.Parent.Font
    Set UserControl.Parent.Font = lblCaption.Font
    Dim ColorForm As Form
    Set ColorForm = MotherForm
    MotherForm.Cls
    MotherForm.Picture = LoadPicture("")
    MotherForm.AutoRedraw = True
    MotherForm.BorderStyle = 0
    SetControlBox UserControl.Parent
    MotherForm.BackColor = m_BackColor
    If imgBackGround.Picture = 0 Then BackgroundCircularGradient ColorForm, m_BackGroundEndColor, m_BackGroundStartColor
    MotherForm.PaintPicture imgLeftTopMask.Picture, 0, 0
    If imgBackGround.Picture <> 0 Then MotherForm.PaintPicture imgBackGround.Picture, imgMidLeftMask.Width, imgMidTopMask.Height, MotherForm.Width - imgMidLeftMask.Width - imgMidRightMask.Width, MotherForm.Height - imgMidTopMask.Height - imgMidButtomMask.Height
    MotherForm.PaintPicture imgMidTopMask.Picture, imgLeftTopMask.Width, 0
    MotherForm.PaintPicture imgRightTopMask.Picture, MotherForm.Width - imgRightTopMask.Width, 0
    MotherForm.PaintPicture imgMidlefttopMask.Picture, imgLeftTopMask.Width, 0, MotherForm.Width - imgLeftTopMask.Width - imgRightTopMask.Width, imgMidTopMask.Height
    MotherForm.PaintPicture imgMidLeftMask.Picture, 0, imgLeftTopMask.Height, imgMidLeftMask.Width
    MotherForm.PaintPicture imgMidButtomMask.Picture, imgLeftBottomMask.Width, MotherForm.Height - imgMidButtomMask.Height
    MotherForm.PaintPicture imgMidRightMask.Picture, MotherForm.Width - imgMidRightMask.Width, imgRightTopMask.Height
    MotherForm.PaintPicture imgLeftBottomMask.Picture, 0, MotherForm.Height - imgLeftBottomMask.Height
    MotherForm.PaintPicture imgRightBottomMask.Picture, MotherForm.Width - imgRightBottomMask.Width, MotherForm.Height - imgRightBottomMask.Height
    SetCaptionProps UserControl.Parent
    Set UserControl.Parent.Font = OldFont
    Sleep 100
    If m_Opacity > 0 Then MakeSemiTransparent UserControl.Parent.hwnd, m_Opacity
    MotherForm.Picture = MotherForm.Image
    MotherForm.Refresh
End Sub

' SHOW IN TASKBAR AT MINIMIZE
'=====================================================
Private Sub ShowInTheTaskbar(hwnd As Long, bShow As Boolean)
    Dim lStyle As Long
    ShowWindow hwnd, SW_HIDE
    lStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    SetWindowLong hwnd, GWL_EXSTYLE, lStyle
    App.TaskVisible = bShow
    ShowWindow hwnd, SW_NORMAL
End Sub


' SET CAPTION PROPERTIES
'=====================================================
Private Sub SetCaptionProps(MotherForm)
    On Error Resume Next
    If m_Caption <> "" Then
        lblCaption.Left = m_CaptionLeft
        lblCaption.Top = m_CaptionTop
        lblShadow.Top = lblCaption.Top + 15
        lblShadow.Left = lblCaption.Left + 15
        Select Case m_CaptionStyle
             Case 0
                MotherForm.ForeColor = m_CaptionForeColor
                MotherForm.CurrentX = lblCaption.Left
                MotherForm.CurrentY = lblCaption.Top
                MotherForm.picBack.Print m_Caption
             Case 1
                MotherForm.ForeColor = m_CaptionShadowColor
                MotherForm.CurrentX = lblCaption.Left
                MotherForm.CurrentY = lblCaption.Top
                MotherForm.Print m_Caption
                MotherForm.ForeColor = m_CaptionForeColor
                MotherForm.CurrentX = lblShadow.Left
                MotherForm.CurrentY = lblShadow.Top
                MotherForm.Print m_Caption
             Case 2
                MotherForm.ForeColor = m_CaptionShadowColor
                MotherForm.CurrentX = lblShadow.Left
                MotherForm.CurrentY = lblShadow.Top
                MotherForm.Print m_Caption
                MotherForm.ForeColor = m_CaptionForeColor
                MotherForm.CurrentX = lblCaption.Left
                MotherForm.CurrentY = lblCaption.Top
                MotherForm.Print m_Caption
         End Select
    End If
    lblShadow.Refresh
    lblCaption.Refresh
End Sub

' PARENTCONTROL HANDLING
'=====================================================
Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim lngReturnValue As Long
     moveForm.MousePointer = 0
     If Button = 1 Then
        XRresize = False
        XLresize = False
        YTresize = False
        YBresize = False
        If Y <= imgMidTopMask.Height And Y > (m_Borderwidth * Screen.TwipsPerPixelY) And X > imgLeftTopMask.Width And X < (moveForm.Width - imgRightTopMask.Width) Then
            Call ReleaseCapture
            lngReturnValue = SendMessage(moveForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        Else
        If m_Sizable = True Then
                 If m_Sizable = False Then Exit Sub
                 If X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth Then
                     XRresize = True
                 End If
                 If X < Screen.TwipsPerPixelX * m_Borderwidth Then
                     XLresize = True
                 End If
                 If Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth Then
                     YBresize = True
                 End If
                 If Y < Screen.TwipsPerPixelY * m_Borderwidth Then
                     YTresize = True
                 End If
                 If X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth Or X < Screen.TwipsPerPixelX * m_Borderwidth Then
                     moveForm.MousePointer = 9
                 End If
                 If Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth Or Y < Screen.TwipsPerPixelY * m_Borderwidth Then
                     moveForm.MousePointer = 7
                 End If
                 If (X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth And Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth) Or (X < Screen.TwipsPerPixelX * m_Borderwidth And Y < Screen.TwipsPerPixelY * m_Borderwidth) Then
                     moveForm.MousePointer = 8
                 End If
                 If (X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth And Y < Screen.TwipsPerPixelY * m_Borderwidth) Or (X < Screen.TwipsPerPixelX * m_Borderwidth And Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth) Then
                     moveForm.MousePointer = 6
                 End If
             End If
        End If
    End If

End Sub
Private Sub moveForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moveForm.MousePointer = 0
    If m_Sizable = False Then Exit Sub
    If X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth Then
        XRresize = True
    End If
    If X < Screen.TwipsPerPixelX * m_Borderwidth Then
        XLresize = True
    End If
    If Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth Then
        YBresize = True
    End If
    If Y < Screen.TwipsPerPixelY * m_Borderwidth Then
        YTresize = True
    End If
    If X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth Or X < Screen.TwipsPerPixelX * m_Borderwidth Then
        moveForm.MousePointer = 9
    End If
    If Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth Or Y < Screen.TwipsPerPixelY * m_Borderwidth Then
        moveForm.MousePointer = 7
    End If
    If (X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth And Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth) Or (X < Screen.TwipsPerPixelX * m_Borderwidth And Y < Screen.TwipsPerPixelY * m_Borderwidth) Then
        moveForm.MousePointer = 8
    End If
    If (X > moveForm.Width - Screen.TwipsPerPixelX * m_Borderwidth And Y < Screen.TwipsPerPixelY * m_Borderwidth) Or (X < Screen.TwipsPerPixelX * m_Borderwidth And Y > moveForm.Height - Screen.TwipsPerPixelY * m_Borderwidth) Then
        moveForm.MousePointer = 6
    End If
End Sub
Private Sub moveForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Sizable = True Then
        If Button = vbLeftButton Then
             If XRresize And X > Screen.TwipsPerPixelX * m_Borderwidth * 2 Then
                moveForm.Width = X
             ElseIf XLresize And moveForm.Width - X > Screen.TwipsPerPixelX * m_Borderwidth * 2 Then
                moveForm.Width = moveForm.Width - X
                moveForm.Left = moveForm.Left + X
             ElseIf YBresize And Y > Screen.TwipsPerPixelY * m_Borderwidth Then
                moveForm.Height = Y
                    RepaintSkin UserControl.Parent
             ElseIf YTresize And moveForm.Height - Y > Screen.TwipsPerPixelY * m_Borderwidth * 2 Then
                moveForm.Height = moveForm.Height - Y
                moveForm.Top = moveForm.Top + Y
                    RepaintSkin UserControl.Parent
            End If
        End If
        moveForm.MousePointer = 0
    End If
    ReleaseCapture
End Sub
Private Sub SizeByGripper(ByVal iHwnd As Long)
  ReleaseCapture
  SendMessage iHwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
End Sub
Private Sub DeleteObjectReference(ByVal iReference As Long)
    DeleteObject iReference
End Sub

' Paint a circular gradient
'
' StartColor is the starting color (applied to the corner)
' EndColor is the ending color (applied to the center point)
' NUMSTEPS is the optional number of stripes (default is 256)
' XPos, XY are the coordinates of the center (default is the center of the form)
'
' Example: a vertical gradient from blue to Black
'   BackgroundCircularGradient Me, &HFF0000, 0, , 500, 200
Sub BackgroundCircularGradient(MotherForm As Form, ByVal StartColor As Long, _
    ByVal EndColor As Long, Optional ByVal numSteps As Integer = 256, _
    Optional ByVal XPos As Single = -1, Optional ByVal YPos As Single = -1)
    ' Draws a circular gradiant on a form starting from the center of the form
    ' or optional from a given position
    If StartColor = EndColor Then Exit Sub
    Dim StartRed As Integer, StartGreen As Integer, StartBlue As Integer
    Dim EndRed As Integer, EndGreen As Integer, EndBlue As Integer
    Dim Rad As Single, DRad As Single
    Dim Stp As Long
    Dim OldFillColor As Long
    Dim OldFillStyle As Long
    If XPos = -1 And YPos = -1 Then
        XPos = MotherForm.ScaleWidth / 2
        YPos = MotherForm.ScaleHeight / 2
    End If
    If XPos < MotherForm.ScaleWidth / 2 Then
        If YPos < MotherForm.ScaleHeight / 2 Then
            Rad = Sqr((MotherForm.ScaleWidth - XPos) ^ 2 + (MotherForm.ScaleHeight - YPos) ^ 2)
        Else
            Rad = Sqr((MotherForm.ScaleWidth - XPos) ^ 2 + YPos ^ 2)
        End If
    Else
        If YPos < MotherForm.ScaleHeight / 2 Then
            Rad = Sqr(XPos ^ 2 + (MotherForm.ScaleHeight - YPos) ^ 2)
        Else
            Rad = Sqr(XPos ^ 2 + YPos ^ 2)
        End If
    End If
    StartRed = StartColor And &HFF
    StartGreen = (StartColor And &HFF00&) \ 256
    StartBlue = (StartColor And &HFF0000) \ 65536
    EndRed = (EndColor And &HFF&) - StartRed
    EndGreen = (EndColor And &HFF00&) \ 256 - StartGreen
    EndBlue = (EndColor And &HFF0000) \ 65536 - StartBlue
    RealizePalette MotherForm.hdc
    DRad = Rad / numSteps
    OldFillColor = MotherForm.FillColor
    OldFillStyle = MotherForm.FillStyle
    MotherForm.FillStyle = vbSolid
    For Stp = 0 To numSteps - 1
        MotherForm.FillColor = RGB(StartRed + (EndRed * Stp) \ numSteps, _
            StartGreen + (EndGreen * Stp) \ numSteps, _
            StartBlue + (EndBlue * Stp) \ numSteps)
        MotherForm.Circle (XPos, YPos), Rad, MotherForm.FillColor
        Rad = Rad - DRad
    Next
    MotherForm.FillColor = OldFillColor
    MotherForm.FillStyle = OldFillStyle
End Sub

' MAKE USERCONTROL SEMITRANSPARENT BY OPACITY
'=====================================================
Public Function MakeSemiTransparent(ByVal hwnd As Long, ByVal Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    Perc = ((100 - Perc) / 100) * 255
    If Perc < 0 Or Perc > 255 Then
        MakeSemiTransparent = 1
    Else
        Msg = GetWindowLong(hwnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hwnd, G, Msg
        SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
        MakeSemiTransparent = 0
    End If
    If Err Then
        MakeSemiTransparent = 2
    End If
End Function

'=======================================================================================================
' USERCONTROL PROPERTIES
'=======================================================================================================

' USERCONTROL INITPROPERTIES
'=====================================================
Private Sub UserControl_InitProperties()
    m_Opacity = m_def_Opacity
    Set m_PictureTopLeft = StandardimgLeftTopMask.Picture
    Set m_PictureTopRight = StandardimgRightTopMask.Picture
    Set m_PictureTopMiddle = StandardimgMidTopMask.Picture
    Set m_PictureBottomLeft = StandardimgLeftBottomMask.Picture
    Set m_PictureBottomRight = StandardimgRightBottomMask.Picture
    Set m_PictureBottomMiddle = StandardimgMidButtomMask.Picture
    Set m_PictureMiddleLeft = StandardimgMidLeftMask.Picture
    Set m_PictureMiddleRight = StandardimgMidRightMask.Picture
    Set m_PictureMinUp = StandardpicMinUp.Picture
    Set m_PictureMinDown = StandardpicMinDown.Picture
    Set m_PictureMaxUp = StandardpicMaxUp.Picture
    Set m_PictureMaxDown = StandardpicMaxDown.Picture
    Set m_PictureExitUp = StandardPicExitUp.Picture
    Set m_PictureExitDown = StandardpicExitDown.Picture
    Set m_PictureExitHoover = StandardpicExitHoover.Picture
    Set m_PictureMaxHoover = StandardpicMaxHoover.Picture
    Set m_PictureMinHoover = StandardpicMinHoover.Picture
    Set m_PictureBackGround = StandardimgBackGround.Picture
    Set Font = lblCaption.Font
    m_CaptionLeft = m_def_CaptionLeft
    m_CaptionTop = m_def_CaptionTop
    m_CaptionStyle = StyleRaised
    m_CaptionForeColor = m_def_CaptionForeColor
    m_CaptionShadowColor = m_def_CaptionShadowColor
    m_Caption = "DM_Skined_" & UserControl.Parent.Name
    m_MinButton = m_def_MinButton
    m_MaxButton = m_def_MaxButton
    m_CloseButton = m_def_CloseButton
    m_Opacity = m_def_Opacity
    m_LeftFromRightClose = m_def_LeftFromRightClose
    m_LeftFromRightMaximize = m_def_LeftFromRightMaximize
    m_LeftFromRightMiniMize = m_def_LeftFromRightMiniMize
    m_Sizable = m_def_Sizable
    m_TopClose = m_def_TopClose
    m_TopMiniMize = m_def_TopMiniMize
    m_TopMaximize = m_def_TopMaximize
    m_ActionOnClose = m_def_ActionOnClose
    m_BackGroundEndColor = m_def_BackgroundEndColor
    m_BackGroundStartColor = m_def_BackgroundStartColor
    RepaintSkin UserControl.Parent
End Sub




' USERCONTROL READPROPERTIES
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 On Error Resume Next
'     Set MotherForm = UserControl.Parent
    Set moveForm = UserControl.Parent
    ShowInTheTaskbar MotherForm.hwnd, True
    m_Opacity = PropBag.ReadProperty("Opacity", m_def_Opacity)
    Set m_PictureTopLeft = PropBag.ReadProperty("PictureTopLeft", Nothing)
    If m_PictureTopLeft <> StandardimgLeftTopMask.Picture Then
        Set imgLeftTopMask.Picture = m_PictureTopLeft
    End If
    Set m_PictureBackGround = PropBag.ReadProperty("PictureBackGround", Nothing)
    If m_PictureBackGround <> StandardimgBackGround.Picture Then
        Set imgBackGround.Picture = m_PictureBackGround
    End If
    Set m_PictureTopRight = PropBag.ReadProperty("PictureTopRight", Nothing)
    If m_PictureTopRight <> StandardimgRightTopMask.Picture Then
        Set imgRightTopMask.Picture = m_PictureTopRight
    End If
    Set m_PictureTopMiddle = PropBag.ReadProperty("PictureTopMiddle", Nothing)
    If m_PictureTopMiddle <> StandardimgMidTopMask.Picture Then
        Set imgMidTopMask.Picture = m_PictureTopMiddle
    End If
    Set m_PictureBottomLeft = PropBag.ReadProperty("PictureBottomLeft", Nothing)
    If m_PictureBottomLeft <> StandardimgLeftBottomMask.Picture Then
        Set imgLeftBottomMask.Picture = m_PictureBottomLeft
    End If
    Set m_PictureBottomRight = PropBag.ReadProperty("PictureBottomRight", Nothing)
    If m_PictureBottomRight <> StandardimgRightBottomMask.Picture Then
        Set imgRightBottomMask.Picture = m_PictureBottomRight
    End If
    Set m_PictureBottomMiddle = PropBag.ReadProperty("PictureBottomMiddle", Nothing)
    If m_PictureBottomMiddle <> StandardimgMidButtomMask.Picture Then
        Set imgMidButtomMask.Picture = m_PictureBottomMiddle
    End If
    Set m_PictureMiddleLeft = PropBag.ReadProperty("PictureMiddleLeft", Nothing)
    If m_PictureMiddleLeft <> StandardimgMidLeftMask.Picture Then
        Set imgMidLeftMask.Picture = m_PictureMiddleLeft
    End If
    Set m_PictureMiddleRight = PropBag.ReadProperty("PictureMiddleRight", Nothing)
    If m_PictureMiddleRight <> StandardimgMidRightMask.Picture Then
        Set imgMidRightMask.Picture = m_PictureMiddleRight
    End If
    Set m_PictureMinUp = PropBag.ReadProperty("PictureMinUp", Nothing)
    If m_PictureMinUp <> StandardpicMinUp.Picture Then
        Set picMinUp.Picture = m_PictureMinUp
        Set imgMin.Picture = m_PictureMinUp
  End If
    Set m_PictureMinDown = PropBag.ReadProperty("PictureMinDown", Nothing)
    If m_PictureMinDown <> StandardpicMinDown.Picture Then
        Set picMinDown.Picture = m_PictureMinDown
    End If
    
    Set m_PictureMaxUp = PropBag.ReadProperty("PictureMaxUp", Nothing)
    If m_PictureMaxUp <> StandardpicMaxUp.Picture Then
        Set picMaxUp.Picture = m_PictureMaxUp
        Set imgMax.Picture = m_PictureMaxUp
    End If
    Set m_PictureMaxDown = PropBag.ReadProperty("PictureMaxDown", Nothing)
    If m_PictureMaxDown <> StandardpicMaxDown.Picture Then
        Set picMaxDown.Picture = m_PictureMaxDown
    End If
    Set m_PictureExitUp = PropBag.ReadProperty("PictureExitUp", Nothing)
    If m_PictureExitUp <> StandardPicExitUp.Picture Then
        Set PicExitUp.Picture = m_PictureExitUp
        Set imgClose.Picture = m_PictureExitUp
    End If
    Set m_PictureExitDown = PropBag.ReadProperty("PictureExitDown", Nothing)
    If m_PictureExitDown <> StandardpicExitDown.Picture Then
        Set picExitDown.Picture = m_PictureExitDown
    End If
    Set m_PictureExitHoover = PropBag.ReadProperty("PictureExitHoover", Nothing)
    If m_PictureExitHoover <> StandardpicExitHoover.Picture Then
        Set picExitHoover.Picture = m_PictureExitHoover
    End If
    Set m_PictureMaxHoover = PropBag.ReadProperty("PictureMaxHoover", Nothing)
    If m_PictureMaxHoover <> StandardpicMaxHoover.Picture Then
        Set picMaxHoover.Picture = m_PictureMaxHoover
    End If
    Set m_PictureMinHoover = PropBag.ReadProperty("PictureMinHoover", Nothing)
    If m_PictureMinHoover <> StandardpicMinHoover.Picture Then
        Set picMinHoover.Picture = m_PictureMinHoover
    End If

    'm_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_MinButton = PropBag.ReadProperty("MinButton", m_def_MinButton)
    m_MaxButton = PropBag.ReadProperty("MaxButton", m_def_MaxButton)
    m_CloseButton = PropBag.ReadProperty("closeButton", m_def_CloseButton)
    m_CaptionFont = PropBag.ReadProperty("CaptionFont", lblCaption.Font)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set picBack.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblShadow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_CaptionStyle = PropBag.ReadProperty("Captionstyle", 0)
    m_CaptionForeColor = PropBag.ReadProperty("CaptionForeColor", &HC0C0C0)
    m_CaptionShadowColor = PropBag.ReadProperty("CaptionShadowColor", &H404040)
    m_Caption = PropBag.ReadProperty("Caption", "DM_Skined_" & UserControl.Parent.Name)
    m_Sizable = PropBag.ReadProperty("Sizable", m_def_Sizable)
    m_LeftFromRightClose = PropBag.ReadProperty("LeftFromRightClose", m_def_LeftFromRightClose)
    m_LeftFromRightMaximize = PropBag.ReadProperty("LeftFromRightMaximize", m_def_LeftFromRightMaximize)
    m_LeftFromRightMiniMize = PropBag.ReadProperty("LeftFromRightMinimize", m_def_LeftFromRightMaximize)
    m_TopClose = PropBag.ReadProperty("TopClose", m_def_TopClose)
    m_TopMiniMize = PropBag.ReadProperty("TopMinimize", m_def_TopMiniMize)
    m_TopMaximize = PropBag.ReadProperty("TopMaximize", m_def_TopMaximize)
    m_CaptionLeft = PropBag.ReadProperty("CaptionLeft", m_def_CaptionLeft)
    m_CaptionTop = PropBag.ReadProperty("Captiontop", m_def_CaptionTop)
    m_ActionOnClose = PropBag.ReadProperty("ActionOnClose", m_def_ActionOnClose)
    m_BackGroundStartColor = PropBag.ReadProperty("BackGroundStartColor", m_def_BackgroundStartColor)
    m_BackGroundEndColor = PropBag.ReadProperty("BackGroundEndColor", m_def_BackgroundEndColor)
    lblCaption.Caption = m_Caption
    lblShadow.Caption = m_Caption
    Opening = True
    RepaintSkin UserControl.Parent
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 375
End Sub

' USERCONTROL WRITEPROPERTIES
'=====================================================
Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Opacity", m_Opacity, m_def_Opacity)
    Call PropBag.WriteProperty("MinButton", m_MinButton, m_def_MinButton)
    Call PropBag.WriteProperty("MaxButton", m_MaxButton, m_def_MaxButton)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
    Call PropBag.WriteProperty("PictureTopLeft", m_PictureTopLeft, Nothing)
    Call PropBag.WriteProperty("PictureBackGround", m_PictureBackGround, Nothing)
    Call PropBag.WriteProperty("PictureTopRight", m_PictureTopRight, Nothing)
    Call PropBag.WriteProperty("PictureTopMiddle", m_PictureTopMiddle, Nothing)
    Call PropBag.WriteProperty("PictureBottomLeft", m_PictureBottomLeft, Nothing)
    Call PropBag.WriteProperty("PictureBottomRight", m_PictureBottomRight, Nothing)
    Call PropBag.WriteProperty("PictureBottomMiddle", m_PictureBottomMiddle, Nothing)
    Call PropBag.WriteProperty("PictureMiddleLeft", m_PictureMiddleLeft, Nothing)
    Call PropBag.WriteProperty("PictureMiddleRight", m_PictureMiddleRight, Nothing)
    Call PropBag.WriteProperty("PictureMinUp", m_PictureMinUp, Nothing)
    Call PropBag.WriteProperty("PictureMinDown", m_PictureMinDown, Nothing)
    Call PropBag.WriteProperty("PictureMaxUp", m_PictureMaxUp, Nothing)
    Call PropBag.WriteProperty("PictureMaxDown", m_PictureMaxDown, Nothing)
    Call PropBag.WriteProperty("PictureExitUp", m_PictureExitUp, Nothing)
    Call PropBag.WriteProperty("PictureExitDown", m_PictureExitDown, Nothing)
    'Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("PictureExitHoover", m_PictureExitHoover, Nothing)
    Call PropBag.WriteProperty("PictureMaxHoover", m_PictureMaxHoover, Nothing)
    Call PropBag.WriteProperty("PictureMinHoover", m_PictureMinHoover, Nothing)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, StyleNone)
    Call PropBag.WriteProperty("CaptionForeColor", m_CaptionForeColor, &HC0C0C0)
    Call PropBag.WriteProperty("CaptionShadowColor", m_CaptionShadowColor, &H404040)
    Call PropBag.WriteProperty("Caption", m_Caption, "DM_Skined_" & UserControl.Parent.Name)
    Call PropBag.WriteProperty("CaptionLeft", m_CaptionLeft, m_def_CaptionLeft)
    Call PropBag.WriteProperty("Captiontop", m_CaptionTop, m_def_CaptionTop)
    Call PropBag.WriteProperty("Sizable", m_Sizable, m_def_Sizable)
    Call PropBag.WriteProperty("LeftFromRightMinimize", m_LeftFromRightMiniMize, m_def_LeftFromRightMiniMize)
    Call PropBag.WriteProperty("LeftFromRightMaximize", m_LeftFromRightMaximize, m_def_LeftFromRightMaximize)
    Call PropBag.WriteProperty("LeftFromRightClose", m_LeftFromRightClose, m_def_LeftFromRightClose)
    Call PropBag.WriteProperty("TopMinimize", m_TopMiniMize, m_def_TopMiniMize)
    Call PropBag.WriteProperty("TopMaximize", m_TopMaximize, m_def_TopMaximize)
    Call PropBag.WriteProperty("TopClose", m_TopClose, m_def_TopClose)
    Call PropBag.WriteProperty("ActionOnClose", m_ActionOnClose, m_def_ActionOnClose)
    Call PropBag.WriteProperty("BackGroundStartColor", m_BackGroundStartColor, m_def_BackgroundStartColor)
    Call PropBag.WriteProperty("BackGroundEndColor", m_BackGroundEndColor, m_def_BackgroundEndColor)

End Sub

'
'=======================================================================================================
'USER CONTROL GET AND LET
'=======================================================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get Opacity() As Byte
    Opacity = m_Opacity
End Property
Public Property Let Opacity(ByVal New_Opacity As Byte)
    On Error Resume Next
    If New_Opacity > 100 Then New_Opacity = 100
    m_Opacity = New_Opacity
    PropertyChanged "Opacity"
    MakeSemiTransparent UserControl.Parent.hwnd, m_Opacity
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MinButton() As Boolean
    MinButton = m_MinButton
End Property
Public Property Let MinButton(ByVal New_MinButton As Boolean)
    m_MinButton = New_MinButton
    PropertyChanged "MinButton"
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MaxButton() As Boolean
    MaxButton = m_MaxButton
End Property
Public Property Let MaxButton(ByVal New_MaxButton As Boolean)
    m_MaxButton = New_MaxButton
    PropertyChanged "MaxButton"
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property
Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    m_CloseButton = New_CloseButton
    PropertyChanged "CloseButton"
        RepaintSkin UserControl.Parent
End Property




'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureTopLeft() As Picture
    Set PictureTopLeft = m_PictureTopLeft
End Property
Public Property Set PictureTopLeft(ByVal New_PictureTopLeft As Picture)
    Set m_PictureTopLeft = New_PictureTopLeft
    PropertyChanged "PictureTopLeft"
    On Error Resume Next
    Set imgLeftTopMask.Picture = m_PictureTopLeft
    If imgLeftTopMask.Picture <> 0 Then
    Else
        imgLeftTopMask.Picture = StandardimgLeftTopMask.Picture
        Set m_PictureTopLeft = imgLeftTopMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureTopRight() As Picture
    Set PictureTopRight = m_PictureTopRight
End Property
Public Property Set PictureTopRight(ByVal New_PictureTopRight As Picture)
    Set m_PictureTopRight = New_PictureTopRight
    PropertyChanged "PictureTopRight"
    On Error Resume Next
    Set imgRightTopMask.Picture = m_PictureTopRight
    If imgRightTopMask.Picture <> 0 Then
    Else
        imgRightTopMask.Picture = StandardimgRightTopMask.Picture
        Set m_PictureTopRight = imgRightTopMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureTopMiddle() As Picture
    Set PictureTopMiddle = m_PictureTopMiddle
End Property
Public Property Set PictureTopMiddle(ByVal New_PictureTopMiddle As Picture)
    Set m_PictureTopMiddle = New_PictureTopMiddle
    PropertyChanged "PictureTopMiddle"
    On Error Resume Next
    Set imgMidTopMask.Picture = m_PictureTopMiddle
    If imgMidTopMask.Picture <> 0 Then
    Else
        imgMidTopMask.Picture = StandardimgMidTopMask.Picture
        Set m_PictureTopMiddle = imgMidTopMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBottomLeft() As Picture
    Set PictureBottomLeft = m_PictureBottomLeft
End Property
Public Property Set PictureBottomLeft(ByVal New_PictureBottomLeft As Picture)
    Set m_PictureBottomLeft = New_PictureBottomLeft
    PropertyChanged "PictureBottomLeft"
    On Error Resume Next
    Set imgLeftBottomMask.Picture = m_PictureBottomLeft
    If imgLeftBottomMask.Picture <> 0 Then
    Else
        imgLeftBottomMask.Picture = StandardimgLeftBottomMask.Picture
        Set m_PictureBottomLeft = imgLeftBottomMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBottomRight() As Picture
    Set PictureBottomRight = m_PictureBottomRight
End Property
Public Property Set PictureBottomRight(ByVal New_PictureBottomRight As Picture)
    Set m_PictureBottomRight = New_PictureBottomRight
    PropertyChanged "PictureBottomRight"
    On Error Resume Next
    Set imgRightBottomMask.Picture = m_PictureBottomRight
    If imgRightBottomMask.Picture <> 0 Then
    Else
        imgRightBottomMask.Picture = StandardimgRightBottomMask.Picture
        Set m_PictureBottomRight = imgRightBottomMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureBottomMiddle() As Picture
    Set PictureBottomMiddle = m_PictureBottomMiddle
End Property
Public Property Set PictureBottomMiddle(ByVal New_PictureBottomMiddle As Picture)
    Set m_PictureBottomMiddle = New_PictureBottomMiddle
    PropertyChanged "PictureBottomMiddle"
    On Error Resume Next
    Set imgMidButtomMask.Picture = m_PictureBottomMiddle
    If imgMidButtomMask.Picture <> 0 Then
    Else
        imgMidButtomMask.Picture = StandardimgMidButtomMask.Picture
        Set m_PictureBottomMiddle = imgRightBottomMask.Picture
    End If
        RepaintSkin UserControl.Parent
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
    Set imgBackGround.Picture = m_PictureBackGround
    If imgBackGround.Picture <> 0 Then
    Else
        imgBackGround.Picture = StandardimgBackGround.Picture
        Set m_PictureBackGround = imgBackGround.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMiddleLeft() As Picture
    Set PictureMiddleLeft = m_PictureMiddleLeft
End Property
Public Property Set PictureMiddleLeft(ByVal New_PictureMiddleLeft As Picture)
    Set m_PictureMiddleLeft = New_PictureMiddleLeft
    PropertyChanged "PictureMiddleLeft"
    On Error Resume Next
    Set imgMidLeftMask.Picture = m_PictureMiddleLeft
    If imgMidLeftMask.Picture <> 0 Then
    Else
        imgMidLeftMask.Picture = StandardimgMidLeftMask.Picture
        Set m_PictureMiddleLeft = imgMidLeftMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMiddleRight() As Picture
    Set PictureMiddleRight = m_PictureMiddleRight
End Property
Public Property Set PictureMiddleRight(ByVal New_PictureMiddleRight As Picture)
    Set m_PictureMiddleRight = New_PictureMiddleRight
    PropertyChanged "PictureMiddleRight"
    On Error Resume Next
    Set imgMidRightMask.Picture = m_PictureMiddleRight
    If imgMidRightMask.Picture <> 0 Then
    Else
        imgMidRightMask.Picture = StandardimgMidRightMask.Picture
        Set m_PictureMiddleRight = imgMidRightMask.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMinUp() As Picture
    Set PictureMinUp = m_PictureMinUp
End Property
Public Property Set PictureMinUp(ByVal New_PictureMinUp As Picture)
    Set m_PictureMinUp = New_PictureMinUp
    PropertyChanged "PictureMinUp"
    On Error Resume Next
    Set picMinUp.Picture = m_PictureMinUp
    Set imgMin.Picture = m_PictureMinUp
    If picMinUp.Picture <> 0 Then
    Else
        picMinUp.Picture = StandardpicMinUp.Picture
        imgMin.Picture = StandardpicMinUp.Picture
        Set m_PictureMinUp = picMinUp.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMinDown() As Picture
    Set PictureMinDown = m_PictureMinDown
End Property
Public Property Set PictureMinDown(ByVal New_PictureMinDown As Picture)
    Set m_PictureMinDown = New_PictureMinDown
    PropertyChanged "PictureMinDown"
    On Error Resume Next
    Set picMinDown.Picture = m_PictureMinDown
    If picMinDown.Picture <> 0 Then
    Else
        picMinDown.Picture = StandardpicMinDown.Picture
        Set m_PictureMinDown = picMinDown.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMaxUp() As Picture
    Set PictureMaxUp = m_PictureMaxUp
End Property
Public Property Set PictureMaxUp(ByVal New_PictureMaxUp As Picture)
    Set m_PictureMaxUp = New_PictureMaxUp
    PropertyChanged "PictureMaxUp"
    On Error Resume Next
    Set picMaxUp.Picture = m_PictureMaxUp
    Set imgMax.Picture = m_PictureMaxUp
    If picMaxUp.Picture <> 0 Then
    Else
        picMaxUp.Picture = StandardpicMaxUp.Picture
        imgMax.Picture = StandardpicMaxUp.Picture
        Set m_PictureMaxUp = picMaxUp.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMaxDown() As Picture
    Set PictureMaxDown = m_PictureMaxDown
End Property
Public Property Set PictureMaxDown(ByVal New_PictureMaxDown As Picture)
    Set m_PictureMaxDown = New_PictureMaxDown
    PropertyChanged "PictureMaxDown"
    On Error Resume Next
    Set picMaxDown.Picture = m_PictureMaxDown
    If picMaxDown.Picture <> 0 Then
    Else
        picMaxDown.Picture = StandardpicMaxDown.Picture
        Set m_PictureMaxDown = picMaxDown.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureExitUp() As Picture
    Set PictureExitUp = m_PictureExitUp
End Property
Public Property Set PictureExitUp(ByVal New_PictureExitUp As Picture)
    Set m_PictureExitUp = New_PictureExitUp
    PropertyChanged "PictureExitUp"
    On Error Resume Next
    Set PicExitUp.Picture = m_PictureExitUp
    Set imgClose.Picture = m_PictureExitUp
    If PicExitUp.Picture <> 0 Then
    Else
        PicExitUp.Picture = StandardPicExitUp.Picture
        imgClose.Picture = StandardPicExitUp.Picture
        Set m_PictureExitUp = PicExitUp.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureExitDown() As Picture
    Set PictureExitDown = m_PictureExitDown
End Property
Public Property Set PictureExitDown(ByVal New_PictureExitDown As Picture)
    Set m_PictureExitDown = New_PictureExitDown
    PropertyChanged "PictureExitDown"
    On Error Resume Next
    Set picExitDown.Picture = m_PictureExitDown
    If picExitDown.Picture <> 0 Then
    Else
        picExitDown.Picture = StandardpicExitDown.Picture
        Set m_PictureExitDown = picExitDown.Picture
    End If
        RepaintSkin UserControl.Parent
End Property



''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=10,0,0,&H8000000F&
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = m_BackColor
'End Property
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    m_BackColor = New_BackColor
'    PropertyChanged "BackColor"
'    RepaintSkin UserControl.Parent
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get BackGroundStartColor() As OLE_COLOR
    BackGroundStartColor = m_BackGroundStartColor
End Property
Public Property Let BackGroundStartColor(ByVal New_BackGroundStartColor As OLE_COLOR)
    m_BackGroundStartColor = New_BackGroundStartColor
    PropertyChanged "BackGroundStartColor"
    RepaintSkin UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get BackGroundEndColorColor() As OLE_COLOR
    BackGroundEndColorColor = m_BackGroundEndColorColor
End Property
Public Property Let BackGroundEndColorColor(ByVal New_BackGroundEndColorColor As OLE_COLOR)
    m_BackGroundEndColorColor = New_BackGroundEndColorColor
    PropertyChanged "BackGroundEndColorColor"
    RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureExitHoover() As Picture
    Set PictureExitHoover = m_PictureExitHoover
End Property
Public Property Set PictureExitHoover(ByVal New_PictureExitHoover As Picture)
    Set m_PictureExitHoover = New_PictureExitHoover
    PropertyChanged "PictureExitHoover"

    On Error Resume Next
    Set picExitHoover.Picture = m_PictureExitHoover
    If picExitHoover.Picture <> 0 Then
    Else
        picExitHoover.Picture = StandardpicExitHoover.Picture
        Set m_PictureExitHoover = picExitHoover.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMaxHoover() As Picture
    Set PictureMaxHoover = m_PictureMaxHoover
End Property
Public Property Set PictureMaxHoover(ByVal New_PictureMaxHoover As Picture)
    Set m_PictureMaxHoover = New_PictureMaxHoover
    PropertyChanged "PictureMaxHoover"
    On Error Resume Next
    Set picMaxHoover.Picture = m_PictureMaxHoover
    If picMaxHoover.Picture <> 0 Then
    Else
        picMaxHoover.Picture = StandardpicMaxHoover.Picture
        Set m_PictureMaxHoover = picMaxHoover.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureMinHoover() As Picture
    Set PictureMinHoover = m_PictureMinHoover
End Property
Public Property Set PictureMinHoover(ByVal New_PictureMinHoover As Picture)
    Set m_PictureMinHoover = New_PictureMinHoover
    PropertyChanged "PictureMinHoover"
    On Error Resume Next
    Set picMinHoover.Picture = m_PictureMinHoover
    If picMinHoover.Picture <> 0 Then
    Else
        picMinHoover.Picture = StandardpicMinHoover.Picture
        Set m_PictureMinHoover = picMinHoover.Picture
    End If
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
    RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionLeft() As Long
    CaptionLeft = m_CaptionLeft
End Property
Public Property Let CaptionLeft(ByVal New_CaptionLeft As Long)
    m_CaptionLeft = New_CaptionLeft
    PropertyChanged "CaptionLeft"
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionTop() As Long
    CaptionTop = m_CaptionTop
End Property
Public Property Let CaptionTop(ByVal New_CaptionTop As Long)
    m_CaptionTop = New_CaptionTop
    PropertyChanged "CaptionTop"
        RepaintSkin UserControl.Parent
End Property


'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get CaptionShadowColor() As OLE_COLOR
    CaptionShadowColor = m_CaptionShadowColor 'lblShadow.ForeColor
End Property
Public Property Let CaptionShadowColor(ByVal New_CaptionShadowColor As OLE_COLOR)
    lblShadow.ForeColor() = New_CaptionShadowColor
    PropertyChanged "CaptionShadowColor"
    m_CaptionShadowColor = New_CaptionShadowColor
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get CaptionForeColor() As OLE_COLOR
    CaptionForeColor = m_CaptionForeColor 'lblCaption.ForeColor
End Property
Public Property Let CaptionForeColor(ByVal New_CaptionForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_CaptionForeColor
    PropertyChanged "CaptionForeColor"
    m_CaptionForeColor = New_CaptionForeColor
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionStyle() As CaptionStyles
    CaptionStyle = m_CaptionStyle
End Property
Public Property Let CaptionStyle(ByVal New_CaptionStyle As CaptionStyles)
    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"
        RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
     'UserControl.Parent.Caption = m_Caption
       RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Sizable() As Boolean
    Sizable = m_Sizable
End Property
Public Property Let Sizable(ByVal New_Sizable As Boolean)
    m_Sizable = New_Sizable
    PropertyChanged "Sizable"
        RepaintSkin UserControl.Parent
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LeftFromRightClose() As Long
    LeftFromRightClose = m_LeftFromRightClose
End Property
Public Property Let LeftFromRightClose(ByVal New_LeftFromRightClose As Long)
    m_LeftFromRightClose = New_LeftFromRightClose
    PropertyChanged "LeftFromRightClose"
       RepaintSkin UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LeftFromRightMaximize() As Long
    LeftFromRightMaximize = m_LeftFromRightMaximize
End Property
Public Property Let LeftFromRightMaximize(ByVal New_LeftFromRightMaximize As Long)
    m_LeftFromRightMaximize = New_LeftFromRightMaximize
    PropertyChanged "LeftFromRightMaximize"
       RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LeftFromRightMinimize() As Long
    LeftFromRightMinimize = m_LeftFromRightMiniMize
End Property
Public Property Let LeftFromRightMinimize(ByVal New_LeftFromRightMinimize As Long)
    m_LeftFromRightMiniMize = New_LeftFromRightMinimize
    PropertyChanged "LeftFromRightMinimize"
       RepaintSkin UserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TopClose() As Long
    TopClose = m_TopClose
End Property
Public Property Let TopClose(ByVal New_TopClose As Long)
    m_TopClose = New_TopClose
    PropertyChanged "TopClose"
        RepaintSkin UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TopMaximize() As Long
    TopMaximize = m_TopMaximize
End Property
Public Property Let TopMaximize(ByVal New_TopMaximize As Long)
    m_TopMaximize = New_TopMaximize
    PropertyChanged "TopMaximize"
        RepaintSkin UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get TopMinimize() As Long
    TopMinimize = m_TopMiniMize
End Property
Public Property Let TopMinimize(ByVal New_TopMinimize As Long)
    m_TopMiniMize = New_TopMinimize
    PropertyChanged "TopMinimize"
       RepaintSkin UserControl.Parent
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ActionOnClose() As ActionOnClosing
    ActionOnClose = m_ActionOnClose
End Property
Public Property Let ActionOnClose(ByVal New_ActionOnClose As ActionOnClosing)
    m_ActionOnClose = New_ActionOnClose
    PropertyChanged "ActionOnClose"
End Property




