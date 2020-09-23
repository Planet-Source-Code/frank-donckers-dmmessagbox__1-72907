VERSION 5.00
Begin VB.Form frmredraw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin prjMSGbox.DMmsgSkin DMmsgSkin1 
      Left            =   480
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   661
      PictureTopLeft  =   "Form1.frx":3B53C2
      PictureTopRight =   "Form1.frx":3B8E94
      PictureTopMiddle=   "Form1.frx":3B9F9E
      PictureBottomLeft=   "Form1.frx":3C62B8
      PictureBottomRight=   "Form1.frx":3C6776
      PictureBottomMiddle=   "Form1.frx":3C6C34
      PictureMiddleLeft=   "Form1.frx":3CA10A
      PictureMiddleRight=   "Form1.frx":3CC98C
      PictureMinUp    =   "Form1.frx":3CF20E
      PictureMinDown  =   "Form1.frx":3CF820
      PictureMaxUp    =   "Form1.frx":3CFE32
      PictureMaxDown  =   "Form1.frx":3D0444
      PictureExitUp   =   "Form1.frx":3D0A56
      PictureExitDown =   "Form1.frx":3D1068
      PictureExitHoover=   "Form1.frx":3D167A
      PictureMaxHoover=   "Form1.frx":3D1C8C
      PictureMinHoover=   "Form1.frx":3D229E
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
      CaptionForeColor=   16777215
      CaptionShadowColor=   8421504
      Caption         =   "Testing the DMmessagebox"
      CaptionLeft     =   1100
      Captiontop      =   120
      Sizable         =   0   'False
      LeftFromRightMinimize=   1350
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton1 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OkOnly"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3D28B0
      PictureLeft     =   "Form1.frx":3D9A92
      PictureRight    =   "Form1.frx":3D9FD4
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton2 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OkCancel"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3DA516
      PictureLeft     =   "Form1.frx":3E16F8
      PictureRight    =   "Form1.frx":3E1C3A
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton3 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "YesNo"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3E217C
      PictureLeft     =   "Form1.frx":3E935E
      PictureRight    =   "Form1.frx":3E98A0
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton4 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "YesNoCancel"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3E9DE2
      PictureLeft     =   "Form1.frx":3F0FC4
      PictureRight    =   "Form1.frx":3F1506
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton5 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "AbortRetryIgnore"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3F1A48
      PictureLeft     =   "Form1.frx":3F8C2A
      PictureRight    =   "Form1.frx":3F916C
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton6 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "Form1.frx":3F96AE
      PictureLeft     =   "Form1.frx":400890
      PictureRight    =   "Form1.frx":400DD2
   End
   Begin prjMSGbox.DMmsgBox DMmsgBox1 
      Left            =   840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   794
      TextBoxBackColor=   16777215
      TextBoxForeColor=   8421504
      BtnTextNo       =   "No"
      BtnTextYes      =   "Yes"
      BeginProperty MessageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Brussels"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextBoxFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicExclamation  =   "Form1.frx":401314
      PicQuestion     =   "Form1.frx":4048CA
      PicCritical     =   "Form1.frx":4058E2
      PicInformation  =   "Form1.frx":40691F
      ButtonPicDown   =   "Form1.frx":407952
      ButtonPicBackGround=   "Form1.frx":4107AC
      ButtonPicLeft   =   "Form1.frx":419606
      ButtonPicRight  =   "Form1.frx":419928
      ButtonPicDown   =   "Form1.frx":419C4A
      ButtonPicBackGroundDisabled=   "Form1.frx":422AA4
      ButtonPicLeftDisabled=   "Form1.frx":42B8FE
      ButtonPicRightDisabled=   "Form1.frx":42BC20
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SkinPicTopleft  =   "Form1.frx":42BF42
      SkinPicTopMiddle=   "Form1.frx":42DD1C
      SkinPicTopRight =   "Form1.frx":431626
      SkinPicMiddleLeft=   "Form1.frx":433400
      SkinPicMiddleRight=   "Form1.frx":43EF72
      SkinPicBottomLeft=   "Form1.frx":44AAE4
      SkinPicBottomMiddle=   "Form1.frx":44AC76
      SkinPicBottomRight=   "Form1.frx":459728
      SkinPicBackGround=   "Form1.frx":4598BA
      BeginProperty SkinCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SkinCaptionTop  =   60
      SkinCaptionLeft =   1400
   End
End
Attribute VB_Name = "frmredraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

