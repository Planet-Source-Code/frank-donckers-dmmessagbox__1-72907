VERSION 5.00
Begin VB.Form frmMSGbox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Message"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13365
   ForeColor       =   &H0088DEEF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMSGbox.frx":0000
   ScaleHeight     =   5850
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjMSGbox.DMTextBox txtInput 
      Height          =   420
      Left            =   4560
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9609633
      Text            =   ""
      BorderColor     =   8421504
      BorderColorOver =   8454143
      BorderStyle     =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   5040
      Width           =   735
   End
   Begin prjMSGbox.DMmsgButton cmdYes 
      Height          =   450
      Left            =   11400
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Yes"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":FECFA
      PictureBackGroundDown=   "frmMSGbox.frx":107B54
      PictureLeft     =   "frmMSGbox.frx":1109AE
      PictureRight    =   "frmMSGbox.frx":110CD0
      PictureBackGroundDisabled=   "frmMSGbox.frx":110FF2
      PictureLeftDisabled=   "frmMSGbox.frx":119E4C
      PictureRightDisabled=   "frmMSGbox.frx":11A16E
   End
   Begin prjMSGbox.DMmsgSkin DMmsgSkin1 
      Left            =   12480
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   661
      MinButton       =   0   'False
      MaxButton       =   0   'False
      PictureTopLeft  =   "frmMSGbox.frx":11A490
      PictureTopRight =   "frmMSGbox.frx":11C26A
      PictureTopMiddle=   "frmMSGbox.frx":11E044
      PictureBottomLeft=   "frmMSGbox.frx":12194E
      PictureBottomRight=   "frmMSGbox.frx":121AE0
      PictureBottomMiddle=   "frmMSGbox.frx":121C72
      PictureMiddleLeft=   "frmMSGbox.frx":130724
      PictureMiddleRight=   "frmMSGbox.frx":13C2F6
      PictureMinUp    =   "frmMSGbox.frx":147EC8
      PictureMinDown  =   "frmMSGbox.frx":1484DA
      PictureMaxUp    =   "frmMSGbox.frx":148AEC
      PictureMaxDown  =   "frmMSGbox.frx":1490FE
      PictureExitUp   =   "frmMSGbox.frx":149710
      PictureExitDown =   "frmMSGbox.frx":149D22
      PictureExitHoover=   "frmMSGbox.frx":14A334
      PictureMaxHoover=   "frmMSGbox.frx":14A946
      PictureMinHoover=   "frmMSGbox.frx":14AF58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionStyle    =   2
      CaptionForeColor=   8969967
      CaptionShadowColor=   16448
      Caption         =   "MessageBox"
      CaptionLeft     =   1400
      Captiontop      =   110
      LeftFromRightMinimize=   1350
   End
   Begin prjMSGbox.DMmsgButton cmdIgnore 
      Height          =   450
      Left            =   11400
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ignore"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":14B56A
      PictureBackGroundDown=   "frmMSGbox.frx":1543C4
      PictureLeft     =   "frmMSGbox.frx":15D21E
      PictureRight    =   "frmMSGbox.frx":15D540
      PictureBackGroundDisabled=   "frmMSGbox.frx":15D862
      PictureLeftDisabled=   "frmMSGbox.frx":165534
      PictureRightDisabled=   "frmMSGbox.frx":165886
   End
   Begin prjMSGbox.DMmsgButton cmdNo 
      Height          =   450
      Left            =   11400
      TabIndex        =   2
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "No"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":165BD8
      PictureBackGroundDown=   "frmMSGbox.frx":16EA32
      PictureLeft     =   "frmMSGbox.frx":17788C
      PictureRight    =   "frmMSGbox.frx":177BAE
      PictureBackGroundDisabled=   "frmMSGbox.frx":177ED0
      PictureLeftDisabled=   "frmMSGbox.frx":180D2A
      PictureRightDisabled=   "frmMSGbox.frx":18104C
   End
   Begin prjMSGbox.DMmsgButton cmdCancel 
      Height          =   450
      Left            =   11400
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":18136E
      PictureBackGroundDown=   "frmMSGbox.frx":18A1C8
      PictureLeft     =   "frmMSGbox.frx":193022
      PictureRight    =   "frmMSGbox.frx":193344
      PictureBackGroundDisabled=   "frmMSGbox.frx":193666
      PictureLeftDisabled=   "frmMSGbox.frx":19C4C0
      PictureRightDisabled=   "frmMSGbox.frx":19C7E2
   End
   Begin prjMSGbox.DMmsgButton cmdAbort 
      Height          =   450
      Left            =   11400
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Abort"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":19CB04
      PictureBackGroundDown=   "frmMSGbox.frx":1A595E
      PictureLeft     =   "frmMSGbox.frx":1AE7B8
      PictureRight    =   "frmMSGbox.frx":1AEADA
      PictureBackGroundDisabled=   "frmMSGbox.frx":1AEDFC
      PictureLeftDisabled=   "frmMSGbox.frx":1B7C56
      PictureRightDisabled=   "frmMSGbox.frx":1B7F78
   End
   Begin prjMSGbox.DMmsgButton cmdOK 
      Height          =   450
      Left            =   11400
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":1B829A
      PictureBackGroundDown=   "frmMSGbox.frx":1C10F4
      PictureLeft     =   "frmMSGbox.frx":1C9F4E
      PictureRight    =   "frmMSGbox.frx":1CA270
      PictureBackGroundDisabled=   "frmMSGbox.frx":1CA592
      PictureLeftDisabled=   "frmMSGbox.frx":1D33EC
      PictureRightDisabled=   "frmMSGbox.frx":1D370E
   End
   Begin prjMSGbox.DMmsgButton cmdHelp 
      Height          =   450
      Left            =   11400
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Help"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":1D3A30
      PictureBackGroundDown=   "frmMSGbox.frx":1DC88A
      PictureLeft     =   "frmMSGbox.frx":1E56E4
      PictureRight    =   "frmMSGbox.frx":1E5A06
      PictureBackGroundDisabled=   "frmMSGbox.frx":1E5D28
      PictureLeftDisabled=   "frmMSGbox.frx":1EEB82
      PictureRightDisabled=   "frmMSGbox.frx":1EEEA4
   End
   Begin prjMSGbox.DMmsgButton cmdRetry 
      Height          =   450
      Left            =   11400
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   794
      ForeColor       =   8969967
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Retry"
      HooverColor     =   8969967
      DisabledColor   =   12648447
      Alignment       =   1
      PictureBackGround=   "frmMSGbox.frx":1EF1C6
      PictureBackGroundDown=   "frmMSGbox.frx":1F8020
      PictureLeft     =   "frmMSGbox.frx":200E7A
      PictureRight    =   "frmMSGbox.frx":20119C
      PictureBackGroundDisabled=   "frmMSGbox.frx":2014BE
      PictureLeftDisabled=   "frmMSGbox.frx":20A318
      PictureRightDisabled=   "frmMSGbox.frx":20A63A
   End
   Begin VB.Image imgNothing 
      Height          =   135
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0088DEEF&
      Height          =   345
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   930
   End
   Begin VB.Image imgCritical 
      Height          =   1080
      Left            =   200
      Picture         =   "frmMSGbox.frx":20A95C
      Top             =   500
      Width           =   1080
   End
   Begin VB.Image imgExclamation 
      Height          =   1080
      Left            =   200
      Picture         =   "frmMSGbox.frx":20B989
      Top             =   500
      Width           =   1080
   End
   Begin VB.Image imgQuestion 
      Height          =   1080
      Left            =   200
      Picture         =   "frmMSGbox.frx":20C991
      Top             =   500
      Width           =   1080
   End
   Begin VB.Image imgInfo 
      Height          =   1080
      Left            =   200
      Picture         =   "frmMSGbox.frx":20D999
      Top             =   500
      Width           =   1080
   End
End
Attribute VB_Name = "frmMSGbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbort_Click()
    WriteSettingsAndUnload "ABORT"
End Sub

Private Sub cmdCancel_Click()
    WriteSettingsAndUnload "CANCEL"
End Sub

Private Sub cmdHelp_Click()
    WriteSettingsAndUnload "HELP"
End Sub

Private Sub cmdIgnore_Click()
    WriteSettingsAndUnload "IGNORE"
End Sub

Private Sub cmdNo_Click()
    WriteSettingsAndUnload "NO"
End Sub

Private Sub cmdOK_Click()
    If txtInput.Visible = True Then
        If Trim$(txtInput.Text) = "" Then Exit Sub
        WriteSettingsAndUnload txtInput.Text
    Else
        WriteSettingsAndUnload "OK"
    End If
End Sub


Private Sub cmdRetry_Click()
    WriteSettingsAndUnload "RETRY"
End Sub

Private Sub cmdYes_Click()
    WriteSettingsAndUnload "YES"
End Sub

Private Sub Form_Resize()
    If Me.Height < 1785 Then Me.Height = 1785
    If Me.Width < (DMmsgSkin1.PictureTopLeft.Width / Screen.TwipsPerPixelX) + (DMmsgSkin1.PictureTopRight.Width / Screen.TwipsPerPixelX) Then _
    Me.Width = Me.Width < (DMmsgSkin1.PictureTopLeft.Width / Screen.TwipsPerPixelX) + (DMmsgSkin1.PictureTopRight.Width / Screen.TwipsPerPixelX)
End Sub

Private Sub WriteSettingsAndUnload(strButton)
    ' check if the button pressed is valid
    If txtInput.Visible = True Then
        If strButton <> txtInput.Text And strButton <> "CANCEL" Then strButton = ""
    Else
        strButton = Trim$(UCase(strButton))
        If strButton <> "OK" And strButton <> "CANCEL" And strButton <> "YES" And strButton <> "NO" And strButton <> "ABORT" _
         And strButton <> "RETRY" And strButton <> "IGNORE" And strButton <> "HELP" Then strButton = ""
        ' Save the registery settings to tell the ActiveX control what color is selected
        ' and unload the form
    End If
    SaveSetting "DMMSGBOX", "VALUES", "BUTTON", strButton
    SaveSetting "DMMSGBOX", "VALUES", "LEFT", Me.Left
    SaveSetting "DMMSGBOX", "VALUES", "TOP", Me.Top
    Unload Me
    DoEvents
End Sub
