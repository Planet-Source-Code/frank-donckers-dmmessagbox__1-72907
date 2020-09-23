VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Testing the DMmessagebox"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboStartUpPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmTest.frx":3B53C2
      Left            =   1320
      List            =   "frmTest.frx":3B53EA
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CheckBox chkHelp 
      BackColor       =   &H00404040&
      Caption         =   "Add help button"
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin prjMSGbox.DMTextBox DMTextBox1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      Locked          =   -1  'True
      BorderColor     =   4210752
      BorderColorOver =   4210752
      BorderStyle     =   2
      BackColor       =   14737632
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton1 
      Height          =   450
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3B54C9
      PictureLeft     =   "frmTest.frx":3BC6AB
      PictureRight    =   "frmTest.frx":3BCBED
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton2 
      Height          =   450
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3BD12F
      PictureLeft     =   "frmTest.frx":3C4311
      PictureRight    =   "frmTest.frx":3C4853
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton3 
      Height          =   450
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3C4D95
      PictureLeft     =   "frmTest.frx":3CBF77
      PictureRight    =   "frmTest.frx":3CC4B9
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton4 
      Height          =   450
      Left            =   3360
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3CC9FB
      PictureLeft     =   "frmTest.frx":3D3BDD
      PictureRight    =   "frmTest.frx":3D411F
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton5 
      Height          =   450
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3D4661
      PictureLeft     =   "frmTest.frx":3DB843
      PictureRight    =   "frmTest.frx":3DBD85
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton6 
      Height          =   450
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      PictureBackGround=   "frmTest.frx":3DC2C7
      PictureLeft     =   "frmTest.frx":3E34A9
      PictureRight    =   "frmTest.frx":3E39EB
   End
   Begin prjMSGbox.DMmsgSkin DMmsgSkin1 
      Left            =   240
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   661
      PictureTopLeft  =   "frmTest.frx":3E3F2D
      PictureTopRight =   "frmTest.frx":3E79FF
      PictureTopMiddle=   "frmTest.frx":3E8B09
      PictureBottomLeft=   "frmTest.frx":3F4E23
      PictureBottomRight=   "frmTest.frx":3F52E1
      PictureBottomMiddle=   "frmTest.frx":3F579F
      PictureMiddleLeft=   "frmTest.frx":3F8C75
      PictureMiddleRight=   "frmTest.frx":3FB4F7
      PictureMinUp    =   "frmTest.frx":3FDD79
      PictureMinDown  =   "frmTest.frx":3FE38B
      PictureMaxUp    =   "frmTest.frx":3FE99D
      PictureMaxDown  =   "frmTest.frx":3FEFAF
      PictureExitUp   =   "frmTest.frx":3FF5C1
      PictureExitDown =   "frmTest.frx":3FFBD3
      PictureExitHoover=   "frmTest.frx":4001E5
      PictureMaxHoover=   "frmTest.frx":4007F7
      PictureMinHoover=   "frmTest.frx":400E09
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
      LeftFromRightClose=   100
   End
   Begin prjMSGbox.DMmsgBox DMmsgBox1 
      Left            =   240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   794
      TextBoxBackColor=   16777215
      TextBoxForeColor=   8421504
      BtnTextNo       =   "No"
      BtnTextYes      =   "Yes"
      BeginProperty MessageFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      PicExclamation  =   "frmTest.frx":40141B
      PicQuestion     =   "frmTest.frx":4049D1
      PicCritical     =   "frmTest.frx":4059E9
      PicInformation  =   "frmTest.frx":406A26
      ButtonPicDown   =   "frmTest.frx":407A59
      ButtonPicBackGround=   "frmTest.frx":4108B3
      ButtonPicLeft   =   "frmTest.frx":41970D
      ButtonPicRight  =   "frmTest.frx":419A2F
      ButtonPicDown   =   "frmTest.frx":419D51
      ButtonPicBackGroundDisabled=   "frmTest.frx":422BAB
      ButtonPicLeftDisabled=   "frmTest.frx":42BA05
      ButtonPicRightDisabled=   "frmTest.frx":42BD27
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SkinPicTopleft  =   "frmTest.frx":42C049
      SkinPicTopMiddle=   "frmTest.frx":42DE23
      SkinPicTopRight =   "frmTest.frx":43172D
      SkinPicMiddleLeft=   "frmTest.frx":433507
      SkinPicMiddleRight=   "frmTest.frx":43F079
      SkinPicBottomLeft=   "frmTest.frx":44ABEB
      SkinPicBottomMiddle=   "frmTest.frx":44AD7D
      SkinPicBottomRight=   "frmTest.frx":45982F
      SkinPicBackGround=   "frmTest.frx":4599C1
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
      ButtonWidthOkOnly=   945
      ButtonWidthOkCancel=   1215
      ButtonWidthYesNo=   990
      ButtonWidthYesNoCancel=   1215
      ButtonWidthAbortRetryIgnore=   1170
      ButtonWidthOkOnlyHelp=   1050
      ButtonWidthOkCancelHelp=   1215
      ButtonWidthYesNoHelp=   1050
      ButtonWidthYesNoCancelHelp=   1215
      ButtonWidthAbortRetryIgnoreHelp=   1170
      ButtonHeight    =   775
      FormWidthOkOnly =   1345
      FormWidthOkCancel=   2490
      FormWidthYesNo  =   2580
      FormWidthYesNoCancel=   4445
      FormWidthAbortRetryIgnore=   4310
      FormWidthOkOnlyHelp=   2700
      FormWidthOkCancelHelp=   3950
      FormWidthYesNoHelp=   3950
      FormWidthYesNoCancelHelp=   5860
      FormWidthAbortRetryIgnoreHelp=   5680
   End
   Begin prjMSGbox.DMmsgButton DMmsgButton7 
      Height          =   450
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      Caption         =   "InputBox"
      HooverColor     =   8421504
      Alignment       =   1
      PictureBackGround=   "frmTest.frx":794813
      PictureLeft     =   "frmTest.frx":79B9F5
      PictureRight    =   "frmTest.frx":79BF37
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "StartupPosition"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   1065
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1200
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3360
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMessage As String
Dim strInput As String
Dim blnAddHelp As Boolean
Dim test As Integer


Private Sub chkHelp_Click()
    blnAddHelp = False
    If chkHelp.Value = 1 Then blnAddHelp = True
End Sub


Private Sub DMmsgButton1_Click()
    strMessage = "This is a messagebox" & vbCrLf & " With OK button" & vbCrLf & "and Information icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    DMTextBox1.Text = "User clicked the button " & DMmsgBox1.ShowMsgBox(strMessage, OKOnly, Information, "Information", , MessageBox, blnAddHelp, cboStartUpPosition.ListIndex)
End Sub

Private Sub DMmsgButton2_Click()
    strMessage = "This is a messagebox" & vbCrLf & " With OK and Cancel buttons" & vbCrLf & "and Exclamation icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    DMTextBox1.Text = "User clicked the button " & DMmsgBox1.ShowMsgBox(strMessage, OKCancel, Exclamation, "Whatch out", MessageBox, , blnAddHelp, cboStartUpPosition.ListIndex)
End Sub

Private Sub DMmsgButton3_Click()
    strMessage = "This is a messagebox" & vbCrLf & " With Yes and No buttons" & vbCrLf & "and Critical icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    DMTextBox1.Text = "User clicked the button " & DMmsgBox1.ShowMsgBox(strMessage, YesNo, Critical, "Critical error", MessageBox, Information, blnAddHelp, cboStartUpPosition.ListIndex)
End Sub

Private Sub DMmsgButton4_Click()
    strMessage = "This is a messagebox" & vbCrLf & " With Yes, No and Cancel buttons" & vbCrLf & "and Question icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    DMTextBox1.Text = "User clicked the button " & DMmsgBox1.ShowMsgBox(strMessage, YesNoCancel, Questionmark, "Question", MessageBox, , blnAddHelp, cboStartUpPosition.ListIndex)
End Sub

Private Sub DMmsgButton5_Click()
    strMessage = "This is a Wiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiide messagebox" & vbCrLf & " With Abort, Retry and Ignore buttons" & vbCrLf & "and exclamation icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    DMTextBox1.Text = "User clicked the button " & DMmsgBox1.ShowMsgBox(strMessage, AbortRetryIgnore, Exclamation, "Error", MessageBox, , blnAddHelp, cboStartUpPosition.ListIndex)
End Sub

Private Sub DMmsgButton6_Click()
    End
End Sub

Private Sub DMmsgButton7_Click()
    strMessage = "This is an inputbox" & vbCrLf & " With OK and Cancel button" & vbCrLf & "and Questionmark icon"
    If blnAddHelp = True Then strMessage = strMessage & vbCrLf & "Help button is added"
    strInput = DMmsgBox1.ShowMsgBox(strMessage, OKCancel, Questionmark, "InputBox", InputBox, "Inputtext is added", blnAddHelp, cboStartUpPosition.ListIndex)
    If strInput = "CANCEL" Then
        DMTextBox1.Text = "User clicked the button " & strInput
    Else
        DMTextBox1.Text = "User input is: " & strInput
    End If
End Sub

Private Sub Form_Load()
    test = 0
    blnAddHelp = False
    cboStartUpPosition.ListIndex = 1
End Sub
