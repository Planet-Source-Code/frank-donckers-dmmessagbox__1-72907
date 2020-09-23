VERSION 5.00
Begin VB.UserControl DMmsgBox 
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   2730
   ScaleWidth      =   5025
   ToolboxBitmap   =   "DMmsgBox.ctx":0000
   Begin VB.Image imgEmpty 
      Height          =   375
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblLabel2 
      AutoSize        =   -1  'True
      Caption         =   "TestLabel2"
      Height          =   195
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      Caption         =   "TestLabel1"
      Height          =   195
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image imgInfo 
      Height          =   1080
      Left            =   1080
      Picture         =   "DMmsgBox.ctx":0312
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Image imgQuestion 
      Height          =   1080
      Left            =   0
      Picture         =   "DMmsgBox.ctx":1335
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Image imgExclamation 
      Height          =   1080
      Left            =   2160
      Picture         =   "DMmsgBox.ctx":233D
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Image imgCritical 
      Height          =   1080
      Left            =   3240
      Picture         =   "DMmsgBox.ctx":3345
      Top             =   1320
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "DMmsgBox.ctx":4372
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgNothing 
      Height          =   1095
      Left            =   600
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "DMmsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Programmer:        Donckers Frank
'                    DarkManSoft@Gmail.com
'
' Description:       Active X Control MessageBox
'
' Use:               ShowMsgBox(Optional ByVal Message As String,
'                               Optional ByVal MsgButtons As MessageButton,
'                               Optional ByVal msgIcon As MessageIcons,
'                               Optional Title As String,
'                               Optional ByVal MSGboxType As MSGboxtypes,
'                               Optional InputText As String,
'                               Optional HelpButton As Boolean,
'                               Optional WindowsStartUpposition As StartupPositions)
'                               As String
'
'                   Example:
'                   Text1.Text = DMmsgBox1.ShowMsgBox("Message", OKOnly, , "Title", Information, MessageBox, False, LastUsed)


'
'=====================================================
' API's
'=====================================================
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long

'=====================================================
' Type messagebox
'=====================================================
Public Enum MSGboxtypes
    MessageBox = 0
    InputBox = 1
End Enum

'=====================================================
' Type Message icons
'=====================================================
Public Enum MessageIcons
    Questionmark = 1
    Information = 2
    Exclamation = 3
    Critical = 4
End Enum

'=====================================================
' Type Message buttons
'=====================================================
Public Enum MessageButton
    OKOnly = 0
    OKCancel = 1
    YesNo = 2
    YesNoCancel = 3
    AbortRetryIgnore = 4
End Enum

'=====================================================
' Type startup position of the messagebox
'=====================================================
Public Enum StartupPositions
    Manual = 0
    CenterOwner = 1
    CenterScreen = 2
    LeftTopOwner = 3
    LeftTopScreen = 4
    LeftBottomOwner = 5
    LeftBottomScreen = 6
    RightTopOwner = 7
    RightTopScreen = 8
    RightBottomOwner = 9
    RightBottomScreen = 10
    LastUsed = 11
End Enum


'=====================================================
' Default Property Values:
'=====================================================
Const m_def_DisplayPosition = 0
Const m_def_MessageIcon = 0
Const m_def_MessageBackStartColor = &H808080
Const m_def_MessageBackEndColor = 0
Const m_def_MessageForeColor = &H88DEEF
Const m_def_TextBoxBackColor = &H8000000F
Const m_def_TextBoxForeColor = 0
Const m_def_TextBoxBorderColor = 0
Const m_def_ButtonForeColor = &H88DEEF
Const m_def_ButtonHooverColor = &H88DEEF
Const m_def_ButtonShadowColor = &H404040
Const m_def_ButtonDisabledColor = &HC0FFFF
Const m_def_BtnTextOK = "OK"
Const m_def_BtnTextNo = "No"
Const m_def_BtnTextYes = "Yes"
Const m_def_BtnTextCancel = "Cancel"
Const m_def_BtnTextAbort = "Abort"
Const m_def_BtnTextRetry = "Retry"
Const m_def_BtnTextIgnore = "Ignore"
Const m_def_BtnTextHelp = "Help"
Const m_def_MSGboxType = 0
Const m_def_SkinCaptionForeColor = &H88DEEF
Const m_def_SkinCaptionShadowColor = &H404040
Const m_def_SkinOpacity = 0
Const m_def_ButtonPicRightDisabled = 0
Const m_def_MessageButtons = 0
Const m_def_SkinCaptionTop = 400
Const m_def_SkinCaptionLeft = 1110
Const m_def_RedrawSkinAtStart = False

'=====================================================
' Property Variables:
'=====================================================
Dim m_ButtonPicRightDisabled As Picture
Dim m_DisplayPosition As StartupPositions
Dim m_MessageIcon As MessageIcons
Dim m_MessageBackStartColor As OLE_COLOR
Dim m_MessageBackEndColor As OLE_COLOR
Dim m_MessageForeColor As OLE_COLOR
Dim m_MessageButtons As MessageButton
Dim m_TextBoxBackColor As OLE_COLOR
Dim m_TextBoxForeColor As OLE_COLOR
Dim m_TextBoxBorderColor As OLE_COLOR
Dim m_ButtonForeColor As OLE_COLOR
Dim m_ButtonhooverColor As OLE_COLOR
Dim m_ButtonShadowColor As OLE_COLOR
Dim m_ButtonDisabledColor As OLE_COLOR
Dim m_ButtonFont As Font
Dim m_BtnTextNo As String
Dim m_BtnTextYes As String
Dim m_BtnTextCancel As String
Dim m_BtnTextAbort As String
Dim m_BtnTextRetry As String
Dim m_BtnTextOK As String
Dim m_BtnTextIgnore As String
Dim m_BtnTextHelp As String
Dim m_MessageFont As Font
Dim m_TextBoxFont As Font
Dim m_PicExclamation As Picture
Dim m_PicQuestion As Picture
Dim m_PicCritical As Picture
Dim m_PicInformation As Picture
Dim m_MSGboxType As MSGboxtypes
Dim m_ButtonPicBackGround As Picture
Dim m_ButtonPicBackGroundDisabled As Picture
Dim m_ButtonPicDown As Picture
Dim m_ButtonPicDownDisabled As Picture
Dim m_ButtonPicLeft As Picture
Dim m_ButtonPicLeftDisabled As Picture
Dim m_ButtonPicRight As Picture
Dim m_SkinCaptionTop As Long
Dim m_SkinCaptionLeft As Long
Dim m_SkinPicTopleft As Picture
Dim m_SkinPicTopMiddle As Picture
Dim m_SkinPicTopRight As Picture
Dim m_SkinPicMiddleLeft As Picture
Dim m_SkinPicMiddleRight As Picture
Dim m_SkinPicBottomLeft As Picture
Dim m_SkinPicBottomMiddle As Picture
Dim m_SkinPicBottomRight As Picture
Dim m_SkinPicBackGround As Picture
Dim m_SkinCaptionFont As Font
Dim m_SkinCaptionForeColor As OLE_COLOR
Dim m_SkinCaptionShadowColor As OLE_COLOR
Dim m_SkinOpacity As Byte
Dim m_RedrawSkinAtStart As Boolean

'=====================================================
'Type Dimensions
'=====================================================
' Used to save calculated buttonwidths and minimum form widths to the propertybag
' (avoids always recalculating)
Private Type Dimension
    ButtonTop                           As Long
    ButtonWidthOkOnly                   As Long
    ButtonWidthOkCancel                 As Long
    ButtonWidthYesNo                    As Long
    ButtonWidthYesNoCancel              As Long
    ButtonWidthAbortRetryIgnore         As Long
    ButtonWidthOkOnlyHelp               As Long
    ButtonWidthOkCancelHelp             As Long
    ButtonWidthYesNoHelp                As Long
    ButtonWidthYesNoCancelHelp          As Long
    ButtonWidthAbortRetryIgnoreHelp     As Long
    FormWidthOkOnly                     As Long
    FormWidthOkCancel                   As Long
    FormWidthYesNo                      As Long
    FormWidthYesNoCancel                As Long
    FormWidthAbortRetryIgnore           As Long
    FormWidthOkOnlyHelp                 As Long
    FormWidthOkCancelHelp               As Long
    FormWidthYesNoHelp                  As Long
    FormWidthYesNoCancelHelp            As Long
    FormWidthAbortRetryIgnoreHelp       As Long
End Type
Dim m_Dimensions As Dimension

'=====================================================
' Event Declarations:
'=====================================================
Event Hide()
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."

'=====================================================
' Program variables:
'=====================================================
Dim ctlControl As Control
Dim lngButtonWidth As Long
Dim lngButtonHeight As Long
Dim intLongest As Integer
Dim lngHightNeeded As Long
Dim lngMinHeight As Long
Dim i, ii As Long
Dim lngLong As Long
Dim lngToAdd As Long


' Displaying the messagebox
' =====================================================
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo = 13
Public Function ShowMsgBox(Optional ByVal Message As String, Optional ByVal MsgButtons As MessageButton _
, Optional ByVal msgIcon As MessageIcons, Optional Title As String, Optional ByVal MSGboxType As MSGboxtypes _
, Optional InputText As String, Optional HelpButton As Boolean, Optional WindowsStartUpposition As StartupPositions) As String
    frmMSGbox.Hide
    If MSGboxType = InputBox Then MsgButtons = OKCancel
    frmMSGbox.Caption = Title
    'SetButtonProps MsgButtons, HelpButton
    SetMessageProps msgIcon, Message, Title, InputText, MsgButtons, HelpButton, MSGboxType
    ' Set Caption on dialogform
    If Title <> "" Then
        frmMSGbox.DMmsgSkin1.Caption = Title
        frmMSGbox.Caption = Title
    Else
        If msgIcon = Critical Then frmMSGbox.DMmsgSkin1.Caption = "Stop"
        If msgIcon = Information Then frmMSGbox.DMmsgSkin1.Caption = "Information"
        If msgIcon = Questionmark Then frmMSGbox.DMmsgSkin1.Caption = "?????"
        If msgIcon = Exclamation Then frmMSGbox.DMmsgSkin1.Caption = "Whatch out"
        frmMSGbox.Caption = frmMSGbox.DMmsgSkin1.Caption
    End If
    If m_SkinPicBackGround <> 0 Then
    Else
        ' You could use this, but the routine BackgroundCircularGradient calculates the startposition
        ' for the gradient automaticly
        'Dim Xpos As Long, Ypos As Long
        'Xpos = ((frmMSGbox.Width / 2) * Screen.TwipsPerPixelX) + (frmMSGbox.DMmsgSkin1.PictureMiddleLeft.Width * Screen.TwipsPerPixelX)
        'Ypos = ((frmMSGbox.Height / 2) * Screen.TwipsPerPixelY) + (frmMSGbox.DMmsgSkin1.PictureTopMiddle.Height * Screen.TwipsPerPixelY)
         BackgroundCircularGradient frmMSGbox, m_MessageBackEndColor, m_MessageBackStartColor       ' , 256, Xpos, Ypos
   End If
    DoEvents
    ' Set the startupposition of the messagebox
    If WindowsStartUpposition < 12 Then SetStartUpPosition UserControl.Parent, frmMSGbox, WindowsStartUpposition
    ' Now show the messagebox as modal form
    frmMSGbox.Show vbModal
    ShowMsgBox = GetSetting("DMMSGBOX", "VALUES", "BUTTON", "")
End Function

'=====================================================
' Set properties of message and icon
'=====================================================
Private Sub SetMessageProps(ByVal msgIcon As MessageIcons, Message As String _
, Title As String, InputText As String, MsgButtons As MessageButton, HelpButton As Boolean, MSGboxType As MSGboxtypes)
    On Error Resume Next
    Dim lngButtonsTop As Long
    ' Automate height inputtext when type is inputbox
    If MSGboxType = InputBox Then frmMSGbox.txtInput.Height = TextHeight(frmMSGbox.txtInput.Text) + 100
    ' Set icon
    frmMSGbox.imgCritical.Visible = False
    frmMSGbox.imgInfo.Visible = False
    frmMSGbox.imgQuestion.Visible = False
    frmMSGbox.imgExclamation.Visible = False
    If msgIcon = 0 Then Set ctlControl = frmMSGbox.imgNothing
    If msgIcon = Critical Then Set ctlControl = frmMSGbox.imgCritical
    If msgIcon = Information Then Set ctlControl = frmMSGbox.imgInfo
    If msgIcon = Questionmark Then Set ctlControl = frmMSGbox.imgQuestion
    If msgIcon = Exclamation Then Set ctlControl = frmMSGbox.imgExclamation
    ctlControl.Visible = True
    ' Set message
    frmMSGbox.lblMessage.Left = ctlControl.Left + ctlControl.Width + 100
    frmMSGbox.lblMessage.Font.Size = m_MessageFont.Size
    frmMSGbox.lblMessage.Font.Bold = m_MessageFont.Bold
    frmMSGbox.lblMessage.Font.Italic = m_MessageFont.Italic
    frmMSGbox.lblMessage.Font.Underline = m_MessageFont.Underline
    frmMSGbox.lblMessage.Font.Strikethrough = m_MessageFont.Strikethrough
    frmMSGbox.lblMessage = Message
    DoEvents
    ' Make sure the messageform is hight enough
    lngHightNeeded = 0
    ' In case of inputbox add hight needed for the textbox
    If MSGboxType = InputBox Then lngHightNeeded = frmMSGbox.txtInput.Height + 210
    ' Add hight needed for the message
    lngHightNeeded = lngHightNeeded + frmMSGbox.lblMessage.Top + frmMSGbox.lblMessage.Height
    ' Add hight needed for buttons
    lngHightNeeded = lngHightNeeded + m_Dimensions.ButtonTop '   frmMSGbox.cmdOK.Height + 200
    ' Add hight needed for bottom form-edge
    lngHightNeeded = lngHightNeeded + frmMSGbox.DMmsgSkin1.PictureBottomMiddle.Height
    ' Minimum height = icon-image + bottom-edge of form + some space
    lngMinHeight = (ctlControl.Top + ctlControl.Height) + (frmMSGbox.DMmsgSkin1.PictureBottomMiddle.Height + 100)
    ' Now set the height for the messagebox
    If lngHightNeeded < lngMinHeight Then
        frmMSGbox.Height = lngMinHeight
    Else
        lngHightNeeded = 0
        If MSGboxType = InputBox Then lngHightNeeded = frmMSGbox.txtInput.Height + 210
        frmMSGbox.Height = (frmMSGbox.lblMessage.Top + frmMSGbox.lblMessage.Height) + (m_Dimensions.ButtonTop) + frmMSGbox.DMmsgSkin1.PictureBottomMiddle.Height + lngHightNeeded
    End If
    ' Width of messageform must be > width of message (even with multiline)
    Dim intPos As Integer
    lblLabel1.Caption = ""
    lblLabel2.Caption = ""
    Set lblLabel1.Font = m_MessageFont
    lblLabel1.Font.Size = m_MessageFont.Size
    lblLabel1.Font.Bold = m_MessageFont.Bold
    lblLabel1.Font.Italic = m_MessageFont.Italic
    lblLabel1.Caption = Message
    ' First the logest needed with of the form is set
    ' to the width of the borders + the witdh of the icon
    intLongest = (frmMSGbox.lblMessage.Width + ctlControl.Left + ctlControl.Width + frmMSGbox.DMmsgSkin1.PictureMiddleRight.Width) + 120
    ' set the font of lbllabel2 to the captionfont of the skin
    ' to check if the length of the caption isn't bigger than the witdh of the form
    Set lblLabel2.Font = m_SkinCaptionFont
    lblLabel2.Font.Size = m_SkinCaptionFont.Size
    lblLabel2.Font.Bold = m_SkinCaptionFont.Bold
    lblLabel2.Font.Italic = m_SkinCaptionFont.Italic
    lblLabel2.Font.Underline = m_SkinCaptionFont.Underline
    ' Width of messageform must be > width of caption on messageform
    lblLabel2.Caption = Trim$(Title)
    If intLongest < lblLabel2.Width + m_SkinCaptionLeft + frmMSGbox.DMmsgSkin1.PictureTopRight.Width Then
        intLongest = lblLabel2.Width + m_SkinCaptionLeft + frmMSGbox.DMmsgSkin1.PictureTopRight.Width
    End If
    frmMSGbox.Width = intLongest
    ' Width of messageform must be > width of al the needed buttons ( minimum width is calculated to m_dimensions...)
    ' top of buttons is already calculated to m_Dimensions.ButtonTop
    ' calculate the left of the buttons from the half of the frmMSGbox.width
    DoEvents
    Select Case MsgButtons
        Case OKOnly
            frmMSGbox.cmdOK.Top = frmMSGbox.Height - m_Dimensions.ButtonTop
            If (frmMSGbox.cmdOK.Top - 60) < (ctlControl.Top + ctlControl.Height) Then frmMSGbox.cmdOK.Top = ctlControl.Top + ctlControl.Height + 60
            lngButtonsTop = frmMSGbox.cmdOK.Top
            If HelpButton = False Then
                If intLongest < m_Dimensions.FormWidthOkOnly Then frmMSGbox.Width = m_Dimensions.FormWidthOkOnly
                frmMSGbox.cmdOK.Width = m_Dimensions.ButtonWidthOkOnly
                frmMSGbox.cmdOK.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthOkOnly / 2)
            Else
                If intLongest < m_Dimensions.FormWidthOkOnlyHelp Then frmMSGbox.Width = m_Dimensions.FormWidthOkOnlyHelp
                frmMSGbox.cmdOK.Width = m_Dimensions.ButtonWidthOkOnlyHelp
                frmMSGbox.cmdHelp.Width = m_Dimensions.ButtonWidthOkOnlyHelp
                frmMSGbox.cmdHelp.Left = ((frmMSGbox.Width / 2) - m_Dimensions.ButtonWidthOkOnlyHelp) - 100
                frmMSGbox.cmdOK.Left = (frmMSGbox.Width / 2) + 100
                frmMSGbox.cmdHelp.Top = lngButtonsTop
                frmMSGbox.cmdHelp.Visible = True
            End If
            frmMSGbox.cmdOK.Visible = True
        Case OKCancel
            frmMSGbox.cmdOK.Top = frmMSGbox.Height - m_Dimensions.ButtonTop
            If (frmMSGbox.cmdOK.Top - 60) < (ctlControl.Top + ctlControl.Height) Then frmMSGbox.cmdOK.Top = ctlControl.Top + ctlControl.Height + 60
            lngButtonsTop = frmMSGbox.cmdOK.Top
            frmMSGbox.cmdCancel.Top = lngButtonsTop
            If HelpButton = False Then
                If intLongest < m_Dimensions.FormWidthOkCancel Then frmMSGbox.Width = m_Dimensions.FormWidthOkCancel
                frmMSGbox.cmdOK.Width = m_Dimensions.ButtonWidthOkCancel
                frmMSGbox.cmdCancel.Width = m_Dimensions.ButtonWidthOkCancel
                frmMSGbox.cmdOK.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthOkCancel + 100)
                frmMSGbox.cmdCancel.Left = (frmMSGbox.Width / 2) + 100
            Else
                If m_Dimensions.FormWidthOkCancelHelp < lngLong Then frmMSGbox.Width = m_Dimensions.FormWidthOkCancelHelp
                frmMSGbox.cmdOK.Width = m_Dimensions.ButtonWidthOkCancelHelp
                frmMSGbox.cmdCancel.Width = m_Dimensions.ButtonWidthOkCancelHelp
                frmMSGbox.cmdHelp.Width = m_Dimensions.ButtonWidthOkCancelHelp
                frmMSGbox.cmdHelp.Top = lngButtonsTop
                frmMSGbox.cmdOK.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthOkCancelHelp / 2)
                frmMSGbox.cmdCancel.Left = (frmMSGbox.cmdOK.Left) + (m_Dimensions.ButtonWidthOkCancelHelp) + 100
                frmMSGbox.cmdHelp.Left = (frmMSGbox.cmdOK.Left) - (m_Dimensions.ButtonWidthOkCancelHelp) - 100
                frmMSGbox.cmdHelp.Visible = True
            End If
            frmMSGbox.cmdOK.Visible = True
            frmMSGbox.cmdCancel.Visible = True
        Case YesNo
            frmMSGbox.cmdYes.Top = frmMSGbox.Height - m_Dimensions.ButtonTop
            If (frmMSGbox.cmdYes.Top - 60) < (ctlControl.Top + ctlControl.Height) Then frmMSGbox.cmdYes.Top = ctlControl.Top + ctlControl.Height + 60
            lngButtonsTop = frmMSGbox.cmdYes.Top
            frmMSGbox.cmdNo.Top = lngButtonsTop
            If HelpButton = False Then
                If intLongest < m_Dimensions.FormWidthYesNo Then frmMSGbox.Width = m_Dimensions.FormWidthYesNo
                frmMSGbox.cmdYes.Width = m_Dimensions.ButtonWidthYesNo
                frmMSGbox.cmdNo.Width = m_Dimensions.ButtonWidthYesNo
                frmMSGbox.cmdYes.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthYesNo + 100)
                frmMSGbox.cmdNo.Left = (frmMSGbox.Width / 2) + 100
            Else
                frmMSGbox.cmdHelp.Top = lngButtonsTop
                If intLongest < m_Dimensions.FormWidthYesNoHelp Then frmMSGbox.Width = m_Dimensions.FormWidthYesNoHelp
                frmMSGbox.cmdYes.Width = m_Dimensions.ButtonWidthYesNoHelp
                frmMSGbox.cmdNo.Width = m_Dimensions.ButtonWidthYesNoHelp
                frmMSGbox.cmdHelp.Width = m_Dimensions.ButtonWidthYesNoHelp
                frmMSGbox.cmdYes.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthYesNoHelp / 2)
                frmMSGbox.cmdNo.Left = (frmMSGbox.cmdYes.Left) + (m_Dimensions.ButtonWidthYesNoHelp) + 100
                frmMSGbox.cmdHelp.Left = (frmMSGbox.cmdYes.Left) - (m_Dimensions.ButtonWidthYesNoHelp) - 100
                frmMSGbox.cmdHelp.Visible = True
            End If
            frmMSGbox.cmdYes.Visible = True
            frmMSGbox.cmdNo.Visible = True
       Case YesNoCancel
            frmMSGbox.cmdYes.Top = frmMSGbox.Height - m_Dimensions.ButtonTop
            If (frmMSGbox.cmdYes.Top - 60) < (ctlControl.Top + ctlControl.Height) Then frmMSGbox.cmdYes.Top = ctlControl.Top + ctlControl.Height + 60
            lngButtonsTop = frmMSGbox.cmdYes.Top
            frmMSGbox.cmdNo.Top = lngButtonsTop
            frmMSGbox.cmdCancel.Top = lngButtonsTop
            If HelpButton = False Then
                If intLongest < m_Dimensions.FormWidthYesNoCancel Then frmMSGbox.Width = m_Dimensions.FormWidthYesNoCancel
                frmMSGbox.cmdYes.Width = m_Dimensions.ButtonWidthYesNoCancel
                frmMSGbox.cmdNo.Width = m_Dimensions.ButtonWidthYesNoCancel
                frmMSGbox.cmdCancel.Width = m_Dimensions.ButtonWidthYesNoCancel
                frmMSGbox.cmdNo.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthYesNoCancel / 2)
                frmMSGbox.cmdYes.Left = (frmMSGbox.cmdNo.Left) - (m_Dimensions.ButtonWidthYesNoCancel) - 100
                frmMSGbox.cmdCancel.Left = (frmMSGbox.cmdNo.Left) + (m_Dimensions.ButtonWidthYesNoCancel) + 100
            Else
                frmMSGbox.cmdHelp.Top = lngButtonsTop
                If intLongest < m_Dimensions.FormWidthYesNoCancelHelp Then frmMSGbox.Width = m_Dimensions.FormWidthYesNoCancelHelp
                frmMSGbox.cmdYes.Width = m_Dimensions.ButtonWidthYesNoCancelHelp
                frmMSGbox.cmdNo.Width = m_Dimensions.ButtonWidthYesNoCancelHelp
                frmMSGbox.cmdCancel.Width = m_Dimensions.ButtonWidthYesNoCancelHelp
                frmMSGbox.cmdHelp.Width = m_Dimensions.ButtonWidthYesNoCancelHelp
                frmMSGbox.cmdYes.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthYesNoCancelHelp + 100)
                frmMSGbox.cmdNo.Left = (frmMSGbox.cmdYes.Left) + (m_Dimensions.ButtonWidthYesNoCancelHelp) + 100
                frmMSGbox.cmdCancel.Left = (frmMSGbox.cmdNo.Left) + (m_Dimensions.ButtonWidthYesNoCancelHelp) + 100
                frmMSGbox.cmdHelp.Left = (frmMSGbox.cmdYes.Left) - (m_Dimensions.ButtonWidthYesNoCancelHelp) - 100
                frmMSGbox.cmdHelp.Visible = True
            End If
            frmMSGbox.cmdYes.Visible = True
            frmMSGbox.cmdNo.Visible = True
            frmMSGbox.cmdCancel.Visible = True
       Case AbortRetryIgnore
            frmMSGbox.cmdAbort.Top = frmMSGbox.Height - m_Dimensions.ButtonTop
            If (frmMSGbox.cmdAbort.Top - 60) < (ctlControl.Top + ctlControl.Height) Then frmMSGbox.cmdAbort.Top = ctlControl.Top + ctlControl.Height + 60
            lngButtonsTop = frmMSGbox.cmdAbort.Top
            frmMSGbox.cmdRetry.Top = lngButtonsTop
            frmMSGbox.cmdIgnore.Top = lngButtonsTop
            If HelpButton = False Then
                If intLongest < m_Dimensions.FormWidthAbortRetryIgnore Then frmMSGbox.Width = m_Dimensions.FormWidthAbortRetryIgnore
                frmMSGbox.cmdAbort.Width = m_Dimensions.ButtonWidthAbortRetryIgnore
                frmMSGbox.cmdRetry.Width = m_Dimensions.ButtonWidthAbortRetryIgnore
                frmMSGbox.cmdIgnore.Width = m_Dimensions.ButtonWidthAbortRetryIgnore
                frmMSGbox.cmdRetry.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthAbortRetryIgnore / 2)
                frmMSGbox.cmdAbort.Left = (frmMSGbox.cmdRetry.Left) - (m_Dimensions.ButtonWidthAbortRetryIgnore) - 100
                frmMSGbox.cmdIgnore.Left = (frmMSGbox.cmdRetry.Left) + (m_Dimensions.ButtonWidthAbortRetryIgnore) + 100
            Else
                frmMSGbox.cmdHelp.Top = lngButtonsTop
                If intLongest < m_Dimensions.FormWidthAbortRetryIgnoreHelp Then frmMSGbox.Width = m_Dimensions.FormWidthAbortRetryIgnoreHelp
                frmMSGbox.cmdAbort.Width = m_Dimensions.ButtonWidthAbortRetryIgnoreHelp
                frmMSGbox.cmdRetry.Width = m_Dimensions.ButtonWidthAbortRetryIgnoreHelp
                frmMSGbox.cmdIgnore.Width = m_Dimensions.ButtonWidthAbortRetryIgnoreHelp
                frmMSGbox.cmdHelp.Width = m_Dimensions.ButtonWidthAbortRetryIgnoreHelp
                frmMSGbox.cmdAbort.Left = (frmMSGbox.Width / 2) - (m_Dimensions.ButtonWidthAbortRetryIgnoreHelp + 100)
                frmMSGbox.cmdRetry.Left = (frmMSGbox.cmdAbort.Left) + (m_Dimensions.ButtonWidthAbortRetryIgnoreHelp) + 100
                frmMSGbox.cmdIgnore.Left = (frmMSGbox.cmdRetry.Left) + (m_Dimensions.ButtonWidthAbortRetryIgnoreHelp) + 100
                frmMSGbox.cmdHelp.Left = (frmMSGbox.cmdAbort.Left) - (m_Dimensions.ButtonWidthAbortRetryIgnoreHelp) - 100
                frmMSGbox.cmdHelp.Visible = True
            End If
            frmMSGbox.cmdAbort.Visible = True
            frmMSGbox.cmdRetry.Visible = True
            frmMSGbox.cmdIgnore.Visible = True
     End Select
     ' Set input textbox when needed
     If MSGboxType = InputBox Then
        If (lngButtonsTop - frmMSGbox.txtInput.Height - 210) > (ctlControl.Top + ctlControl.Height) Then
             frmMSGbox.txtInput.Width = frmMSGbox.Width - frmMSGbox.DMmsgSkin1.PictureMiddleLeft.Width - frmMSGbox.DMmsgSkin1.PictureMiddleRight.Width - 210
             frmMSGbox.txtInput.Top = lngButtonsTop - frmMSGbox.txtInput.Height - 105
             frmMSGbox.txtInput.Left = frmMSGbox.DMmsgSkin1.PictureMiddleLeft.Width + 105
        Else
             frmMSGbox.txtInput.Width = frmMSGbox.Width - (ctlControl.Width + ctlControl.Left) - frmMSGbox.DMmsgSkin1.PictureMiddleRight.Width - 210
             frmMSGbox.txtInput.Left = ctlControl.Width + ctlControl.Left + 105
             frmMSGbox.txtInput.Top = frmMSGbox.lblMessage.Top + frmMSGbox.lblMessage.Height + 105
        End If
         frmMSGbox.txtInput.Text = InputText
         frmMSGbox.txtInput.Visible = True
     Else
         frmMSGbox.txtInput.Visible = False
     End If
    Set ctlControl = Nothing
    If msgIcon = 0 Then frmMSGbox.lblMessage.Left = (frmMSGbox.Width / 2) - (frmMSGbox.lblMessage.Width / 2)
End Sub

'=====================================================
' Set startupposition of the messagebox
' =====================================================
Private Sub SetStartUpPosition(ByRef ParentForm As Form, ByRef MessageForm As Form, ByRef WindowsStartUpposition As StartupPositions)
    On Error Resume Next
    Select Case WindowsStartUpposition
        Case Manual             'Nothing to do here
        Case CenterOwner        'center of parent form
            MessageForm.Left = ParentForm.Left + (ParentForm.ScaleWidth / 2) - (MessageForm.ScaleWidth / 2)
            MessageForm.Top = ParentForm.Top + (ParentForm.ScaleHeight / 2) - (MessageForm.ScaleHeight / 2)
        Case CenterScreen       'center of screen
            MessageForm.Left = (Screen.Width / 2) - ((MessageForm.ScaleWidth / 2))
            MessageForm.Top = (Screen.Height / 2) - (MessageForm.ScaleHeight / 2)
        Case LeftTopOwner       'left top corner of parent form
            MessageForm.Left = ParentForm.Left
            MessageForm.Top = ParentForm.Top
        Case LeftTopScreen      'left top corner of screen
            MessageForm.Left = 0
            MessageForm.Top = 0
        Case LeftBottomOwner    'left bottom corner of parent form
            MessageForm.Left = ParentForm.Left
            MessageForm.Top = (ParentForm.Top + ParentForm.ScaleHeight) - (MessageForm.ScaleHeight)
        Case LeftBottomScreen   'left bottom corner of screen
            MessageForm.Left = 0
            MessageForm.Top = (Screen.Height) - (MessageForm.ScaleHeight)
        Case RightTopOwner      'right top corner of parent form
            MessageForm.Left = (ParentForm.Left + ParentForm.ScaleWidth) - (MessageForm.ScaleWidth)
            MessageForm.Top = ParentForm.Top
        Case RightTopScreen     'right top corner of screen
            MessageForm.Left = (Screen.Width) - (MessageForm.ScaleWidth)
            MessageForm.Top = 0
        Case RightBottomOwner   'right bottom corner of parent form
            MessageForm.Left = (ParentForm.Left + ParentForm.ScaleWidth) - (MessageForm.ScaleWidth)
            MessageForm.Top = (ParentForm.Top + ParentForm.ScaleHeight) - (MessageForm.ScaleHeight)
        Case RightBottomScreen  'right bottom corner of screen
            MessageForm.Left = (Screen.Width) - (MessageForm.ScaleWidth)
            MessageForm.Top = (Screen.Height) - (MessageForm.ScaleHeight)
        Case LastUsed           'position when form whas last closed
            MessageForm.Left = Val(GetSetting("DMMSGBOX", "VALUES", "LEFT", ""))
            MessageForm.Top = Val(GetSetting("DMMSGBOX", "VALUES", "TOP", ""))
            MessageForm.Left = Val(GetSetting("DMMSGBOX", "VALUES", "LEFT", ""))

    End Select
End Sub

'=====================================================
' Redraw all controls on the messageform
'=====================================================
Private Sub SetControlsProps()
    If m_RedrawSkinAtStart = False Then Exit Sub
    On Error Resume Next
    ' Skin
    frmMSGbox.DMmsgSkin1.CaptionShadowColor = m_SkinCaptionShadowColor
    frmMSGbox.DMmsgSkin1.CaptionForeColor = m_SkinCaptionForeColor
    frmMSGbox.DMmsgSkin1.Opacity = m_SkinOpacity
    frmMSGbox.DMmsgSkin1.CaptionLeft = m_SkinCaptionLeft
    frmMSGbox.DMmsgSkin1.CaptionTop = m_SkinCaptionTop
    Set frmMSGbox.DMmsgSkin1.PictureBottomMiddle = m_SkinPicBottomMiddle
    Set frmMSGbox.DMmsgSkin1.PictureBottomRight = m_SkinPicBottomRight
    Set frmMSGbox.DMmsgSkin1.PictureBackGround = m_SkinPicBackGround
    Set frmMSGbox.DMmsgSkin1.Font = m_SkinCaptionFont
    Set frmMSGbox.DMmsgSkin1.PictureMiddleLeft = m_SkinPicMiddleLeft
    Set frmMSGbox.DMmsgSkin1.PictureMiddleRight = m_SkinPicMiddleRight
    Set frmMSGbox.DMmsgSkin1.PictureTopLeft = m_SkinPicTopleft
    Set frmMSGbox.DMmsgSkin1.PictureTopMiddle = m_SkinPicTopMiddle
    Set frmMSGbox.DMmsgSkin1.PictureTopRight = m_SkinPicTopRight
    ' Icons
    Set frmMSGbox.imgQuestion.Picture = m_PicQuestion
    Set frmMSGbox.imgCritical.Picture = m_PicCritical
    Set frmMSGbox.imgInfo.Picture = m_PicInformation
    Set frmMSGbox.imgExclamation.Picture = m_PicExclamation
    ' Messagetext
    frmMSGbox.lblMessage.ForeColor = m_MessageForeColor
    Set frmMSGbox.lblMessage.Font = m_MessageFont
    Set frmMSGbox.lblTest.Font = m_MessageFont
    ' Inputtext
    frmMSGbox.txtInput.BorderColor = m_TextBoxBorderColor
    frmMSGbox.txtInput.ForeColor = m_TextBoxForeColor
    frmMSGbox.txtInput.BackColor = m_TextBoxBackColor
    ' DMCommandbuttons
    frmMSGbox.cmdIgnore.Caption = m_BtnTextIgnore
    frmMSGbox.cmdCancel.Caption = m_BtnTextCancel
    frmMSGbox.cmdAbort.Caption = m_BtnTextAbort
    frmMSGbox.cmdRetry.Caption = m_BtnTextRetry
    frmMSGbox.cmdNo.Caption = m_BtnTextNo
    frmMSGbox.cmdYes.Caption = m_BtnTextYes
    frmMSGbox.cmdOK.Caption = m_BtnTextOK
    frmMSGbox.cmdHelp.Caption = m_BtnTextHelp
    ResetSkinButtonProps
    Set ctlControl = Nothing
End Sub

'=====================================================
' Redraw all buttons on the messageform
'=====================================================
Private Sub ResetSkinButtonProps()
    For Each ctlControl In frmMSGbox
        If TypeOf ctlControl Is DMmsgButton Then
            ctlControl.ForeColor = m_ButtonForeColor
            ctlControl.HooverColor = m_ButtonhooverColor
            ctlControl.DisabledColor = m_ButtonDisabledColor
            ctlControl.ShadowColor = m_ButtonShadowColor
            Set ctlControl.PictureBackGroundDown = m_ButtonPicDown
            Set ctlControl.PictureRight = m_ButtonPicRight
            Set ctlControl.PictureLeft = m_ButtonPicLeft
            Set ctlControl.PictureRightDisabled = m_ButtonPicRightDisabled
            Set ctlControl.PictureLeftDisabled = m_ButtonPicLeftDisabled
            Set ctlControl.PictureBackGround = m_ButtonPicBackGround
            Set ctlControl.PictureBackGroundDisabled = m_ButtonPicBackGroundDisabled
            Set ctlControl.PictureRightDisabled = m_ButtonPicRightDisabled
            Set ctlControl.Font = m_ButtonFont
        End If
    Next
    Set ctlControl = Nothing
End Sub

'=====================================================
' Draw circular gradient on background of form
'=====================================================
Public Sub BackgroundCircularGradient(MotherForm As Form, ByVal StartColor As Long, _
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
Public Sub ReCalButtonAndFormWitdh(Optional WhatButton As String)
    ' To avoid always recalculating the Width for the buttons
    ' and the width for the form needed for the buttons
    ' only recalculate them when the properties of a button or an image on the skin changes
    ' to save them to the propertybag
    ' when Whatbutton is empty al the dimensions get calculated
    lngToAdd = frmMSGbox.DMmsgSkin1.PictureMiddleLeft.Width + frmMSGbox.DMmsgSkin1.PictureMiddleRight.Width
    lngLong = 200    ' space between the buttons
    If WhatButton = "" Or WhatButton = "OK" Then
        m_Dimensions.ButtonWidthOkOnly = TextWidth(frmMSGbox.cmdOK.Caption) + 720
        m_Dimensions.ButtonWidthOkOnlyHelp = TextWidth(frmMSGbox.cmdOK.Caption) + 720
        If TextWidth(frmMSGbox.cmdHelp.Caption) + 720 > m_Dimensions.ButtonWidthOkOnlyHelp Then
            m_Dimensions.ButtonWidthOkOnlyHelp = TextWidth(frmMSGbox.cmdHelp.Caption) + 720
        End If
        m_Dimensions.ButtonWidthOkCancel = TextWidth(frmMSGbox.cmdOK.Caption) + 720
        m_Dimensions.ButtonWidthOkCancelHelp = TextWidth(frmMSGbox.cmdOK.Caption) + 720
        If TextWidth(frmMSGbox.cmdCancel.Caption) > TextWidth(frmMSGbox.cmdOK.Caption) Then
            m_Dimensions.ButtonWidthOkCancel = TextWidth(frmMSGbox.cmdCancel.Caption) + 720
            m_Dimensions.ButtonWidthOkCancelHelp = TextWidth(frmMSGbox.cmdCancel.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdHelp.Caption) + 720 > m_Dimensions.ButtonWidthOkCancelHelp Then
            m_Dimensions.ButtonWidthOkCancelHelp = TextWidth(frmMSGbox.cmdHelp.Caption) + 720
        End If
        m_Dimensions.FormWidthOkOnly = (lngLong * 2) + m_Dimensions.ButtonWidthOkOnly
        m_Dimensions.FormWidthOkOnlyHelp = (lngLong * 3) + (m_Dimensions.ButtonWidthOkOnlyHelp * 2)
        m_Dimensions.FormWidthOkCancel = (lngLong * 3) + (m_Dimensions.ButtonWidthOkOnly * 2)
        m_Dimensions.FormWidthOkCancelHelp = (lngLong * 4) + (m_Dimensions.ButtonWidthOkOnlyHelp * 3)
    End If
    If WhatButton = "" Or WhatButton = "YES" Then
        m_Dimensions.ButtonWidthYesNo = TextWidth(frmMSGbox.cmdYes.Caption) + 720
        m_Dimensions.ButtonWidthYesNoHelp = TextWidth(frmMSGbox.cmdYes.Caption) + 720
        m_Dimensions.ButtonWidthYesNoCancel = TextWidth(frmMSGbox.cmdYes.Caption) + 720
        m_Dimensions.ButtonWidthYesNoCancelHelp = TextWidth(frmMSGbox.cmdYes.Caption) + 720
        If TextWidth(frmMSGbox.cmdNo.Caption) + 720 > m_Dimensions.ButtonWidthYesNo Then
            m_Dimensions.ButtonWidthYesNo = TextWidth(frmMSGbox.cmdYes.Caption) + 720
            m_Dimensions.ButtonWidthYesNoHelp = TextWidth(frmMSGbox.cmdYes.Caption) + 720
            m_Dimensions.ButtonWidthYesNoCancel = TextWidth(frmMSGbox.cmdYes.Caption) + 720
            m_Dimensions.ButtonWidthYesNoCancelHelp = TextWidth(frmMSGbox.cmdYes.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdHelp.Caption) + 720 > m_Dimensions.ButtonWidthYesNoHelp Then
            m_Dimensions.ButtonWidthYesNoHelp = TextWidth(frmMSGbox.cmdHelp.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdCancel.Caption) + 720 > m_Dimensions.ButtonWidthYesNoCancel Then
            m_Dimensions.ButtonWidthYesNoCancel = TextWidth(frmMSGbox.cmdCancel.Caption) + 720
            m_Dimensions.ButtonWidthYesNoCancelHelp = TextWidth(frmMSGbox.cmdCancel.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdHelp.Caption) + 720 > m_Dimensions.ButtonWidthYesNoCancelHelp Then
            m_Dimensions.ButtonWidthYesNoCancelHelp = TextWidth(frmMSGbox.cmdHelp.Caption) + 720
        End If
        m_Dimensions.FormWidthYesNo = (lngLong * 3) + (m_Dimensions.ButtonWidthYesNo * 2)
        m_Dimensions.FormWidthYesNoHelp = (lngLong * 4) + (m_Dimensions.ButtonWidthYesNoHelp * 3)
        m_Dimensions.FormWidthYesNoCancel = (lngLong * 4) + (m_Dimensions.ButtonWidthYesNoCancel * 3)
        m_Dimensions.FormWidthYesNoCancelHelp = (lngLong * 5) + (m_Dimensions.ButtonWidthYesNoCancelHelp * 4)
    End If
    If WhatButton = "" Or WhatButton = "ABORT" Then
        m_Dimensions.ButtonWidthAbortRetryIgnore = TextWidth(frmMSGbox.cmdAbort.Caption) + 720
        m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = TextWidth(frmMSGbox.cmdAbort.Caption) + 720
        If TextWidth(frmMSGbox.cmdRetry.Caption) + 720 > m_Dimensions.ButtonWidthAbortRetryIgnore Then
            m_Dimensions.ButtonWidthAbortRetryIgnore = TextWidth(frmMSGbox.cmdAbort.Caption) + 720
            m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = TextWidth(frmMSGbox.cmdAbort.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdIgnore.Caption) + 720 > m_Dimensions.ButtonWidthAbortRetryIgnore Then
            m_Dimensions.ButtonWidthAbortRetryIgnore = TextWidth(frmMSGbox.cmdIgnore.Caption) + 720
            m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = TextWidth(frmMSGbox.cmdIgnore.Caption) + 720
        End If
        If TextWidth(frmMSGbox.cmdHelp.Caption) + 720 > m_Dimensions.ButtonWidthAbortRetryIgnoreHelp Then
            m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = TextWidth(frmMSGbox.cmdHelp.Caption) + 720
        End If
        m_Dimensions.FormWidthAbortRetryIgnore = (lngLong * 4) + (m_Dimensions.ButtonWidthAbortRetryIgnore * 3)
        m_Dimensions.FormWidthAbortRetryIgnoreHelp = (lngLong * 5) + (m_Dimensions.ButtonWidthAbortRetryIgnoreHelp * 4)
    End If
    Dim intPlusHeight As Integer
    intPlusHeight = frmMSGbox.DMmsgSkin1.PictureBottomMiddle.Height + 60
    m_Dimensions.ButtonTop = frmMSGbox.cmdOK.Height + intPlusHeight
    If (frmMSGbox.cmdCancel.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdCancel.Height + intPlusHeight
    If (frmMSGbox.cmdYes.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdYes.Height + intPlusHeight
    If (frmMSGbox.cmdNo.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdNo.Height + intPlusHeight
    If (frmMSGbox.cmdAbort.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdAbort.Height + intPlusHeight
    If (frmMSGbox.cmdRetry.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdRetry.Height + intPlusHeight
    If (frmMSGbox.cmdIgnore.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdIgnore.Height + intPlusHeight
    If (frmMSGbox.cmdHelp.Height + 200) > m_Dimensions.ButtonTop Then m_Dimensions.ButtonTop = frmMSGbox.cmdHelp.Height + intPlusHeight
    PropertyChanged "m_Dimensions"
End Sub


'=======================================================================================================
' USERCONTROL PROPERTIES
'=======================================================================================================

'=====================================================
' Resizing the control
'=====================================================
Private Sub UserControl_Resize()
    UserControl.Height = 450
    UserControl.Width = 480
End Sub

'=====================================================
' Initialize Properties for User Control
'=====================================================
Private Sub UserControl_InitProperties()
    m_SkinCaptionTop = m_def_SkinCaptionTop
    m_SkinCaptionLeft = m_def_SkinCaptionLeft
    m_DisplayPosition = m_def_DisplayPosition
    m_MessageIcon = m_def_MessageIcon
    m_MessageBackStartColor = m_def_MessageBackStartColor
    m_MessageBackEndColor = m_def_MessageBackEndColor
    m_MessageForeColor = m_def_MessageForeColor
    m_TextBoxBackColor = m_def_TextBoxBackColor
    m_TextBoxBorderColor = m_def_TextBoxBorderColor
    m_TextBoxForeColor = m_def_TextBoxForeColor
    m_ButtonForeColor = m_def_ButtonForeColor
    m_ButtonhooverColor = m_def_ButtonHooverColor
    m_ButtonShadowColor = m_def_ButtonShadowColor
    m_ButtonDisabledColor = m_def_ButtonDisabledColor
    m_RedrawSkinAtStart = m_def_RedrawSkinAtStart
    Set m_ButtonFont = Ambient.Font
    m_BtnTextNo = m_def_BtnTextNo
    m_BtnTextOK = m_def_BtnTextOK
    m_BtnTextYes = m_def_BtnTextYes
    m_BtnTextCancel = m_def_BtnTextCancel
    m_BtnTextAbort = m_def_BtnTextAbort
    m_BtnTextRetry = m_def_BtnTextRetry
    m_BtnTextIgnore = m_def_BtnTextIgnore
    m_BtnTextHelp = m_def_BtnTextHelp
    Set m_MessageFont = Ambient.Font
    m_MessageFont.Name = "Brussels"
    m_MessageFont.Size = 14
    Set m_TextBoxFont = Ambient.Font
    Set m_PicExclamation = LoadPicture("")
    Set m_PicQuestion = LoadPicture("")
    Set m_PicCritical = LoadPicture("")
    Set m_PicInformation = LoadPicture("")
    m_MSGboxType = m_def_MSGboxType
    Set m_ButtonPicBackGround = LoadPicture("")
    Set m_ButtonPicBackGroundDisabled = LoadPicture("")
    Set m_ButtonPicDown = LoadPicture("")
    Set m_ButtonPicLeft = LoadPicture("")
    Set m_ButtonPicLeftDisabled = LoadPicture("")
    Set m_ButtonPicRight = LoadPicture("")
    Set m_SkinPicTopleft = LoadPicture("")
    Set m_SkinPicTopMiddle = LoadPicture("")
    Set m_SkinPicTopRight = LoadPicture("")
    Set m_SkinPicMiddleLeft = LoadPicture("")
    Set m_SkinPicMiddleRight = LoadPicture("")
    Set m_SkinPicBottomLeft = LoadPicture("")
    Set m_SkinPicBottomMiddle = LoadPicture("")
    Set m_SkinPicBottomRight = LoadPicture("")
    Set m_SkinPicBackGround = LoadPicture("")
    Set m_SkinCaptionFont = Ambient.Font
    Set m_ButtonPicRightDisabled = LoadPicture("")
    m_SkinCaptionForeColor = m_def_SkinCaptionForeColor
    m_SkinCaptionShadowColor = m_def_SkinCaptionShadowColor
    m_SkinOpacity = m_def_SkinOpacity
    Set m_ButtonPicRightDisabled = LoadPicture("")
    m_Dimensions.ButtonWidthOkOnly = frmMSGbox.cmdOK.Width + 720
    m_Dimensions.ButtonWidthOkCancel = frmMSGbox.cmdCancel.Width + 720
    m_Dimensions.ButtonWidthYesNo = frmMSGbox.cmdYes.Width + 720
    m_Dimensions.ButtonWidthYesNoCancel = frmMSGbox.cmdCancel.Width + 720
    m_Dimensions.ButtonWidthAbortRetryIgnore = frmMSGbox.cmdIgnore.Width + 720
    m_Dimensions.ButtonWidthOkOnlyHelp = frmMSGbox.cmdHelp.Width + 720
    m_Dimensions.ButtonWidthOkCancelHelp = frmMSGbox.cmdCancel.Width + 720
    m_Dimensions.ButtonWidthYesNoHelp = frmMSGbox.cmdHelp.Width + 720
    m_Dimensions.ButtonWidthYesNoCancelHelp = frmMSGbox.cmdCancel.Width + 720
    m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = frmMSGbox.cmdIgnore.Width + 720
    m_Dimensions.ButtonTop = frmMSGbox.cmdCancel.Height
    m_Dimensions.FormWidthOkOnly = frmMSGbox.Width + (3 * frmMSGbox.cmdCancel.Width) + 800
    m_Dimensions.FormWidthOkCancel = FormWidthOkOnly
    m_Dimensions.FormWidthYesNo = FormWidthOkOnly
    m_Dimensions.FormWidthYesNoCancel = FormWidthOkOnly
    m_Dimensions.FormWidthAbortRetryIgnore = FormWidthOkOnly
    m_Dimensions.FormWidthOkOnlyHelp = frmMSGbox.Width + (4 * frmMSGbox.cmdCancel.Width) + 1000
    m_Dimensions.FormWidthOkCancelHelp = FormWidthOkOnlyHelp
    m_Dimensions.FormWidthYesNoHelp = FormWidthOkOnlyHelp
    m_Dimensions.FormWidthYesNoCancelHelp = FormWidthOkOnlyHelp
    m_Dimensions.FormWidthAbortRetryIgnoreHelp = FormWidthOkOnlyHelp
End Sub

'=====================================================
' Load property values from storage
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DisplayPosition = PropBag.ReadProperty("DisplayPosition", m_def_DisplayPosition)
    m_MessageIcon = PropBag.ReadProperty("MessageIcon", m_def_MessageIcon)
    m_MessageBackStartColor = PropBag.ReadProperty("MessageBackStartColor", m_def_MessageBackStartColor)
    m_MessageBackEndColor = PropBag.ReadProperty("MessageBackEndColor", m_def_MessageBackEndColor)
    m_MessageForeColor = PropBag.ReadProperty("MessageForeColor", m_def_MessageForeColor)
    m_TextBoxBackColor = PropBag.ReadProperty("TextBoxBackColor", m_def_TextBoxBackColor)
    m_TextBoxForeColor = PropBag.ReadProperty("TextBoxForeColor", m_def_TextBoxForeColor)
    m_TextBoxBorderColor = PropBag.ReadProperty("TextBoxForeBorderColor", m_def_TextBoxBorderColor)
    m_ButtonForeColor = PropBag.ReadProperty("ButtonForeColor", m_def_ButtonForeColor)
    m_ButtonhooverColor = PropBag.ReadProperty("ButtonhooverColor", m_def_ButtonHooverColor)
    m_ButtonShadowColor = PropBag.ReadProperty("ButtonShadowColor", m_def_ButtonShadowColor)
    m_ButtonDisabledColor = PropBag.ReadProperty("ButtonDisabledColor", m_def_ButtonDisabledColor)
    Set m_ButtonFont = PropBag.ReadProperty("ButtonFont", Ambient.Font)
    m_BtnTextNo = PropBag.ReadProperty("BtnTextNo", m_def_BtnTextNo)
    m_BtnTextYes = PropBag.ReadProperty("BtnTextYes", m_def_BtnTextYes)
    m_BtnTextCancel = PropBag.ReadProperty("BtnTextCancel", m_def_BtnTextCancel)
    m_BtnTextAbort = PropBag.ReadProperty("BtnTextAbort", m_def_BtnTextAbort)
    m_BtnTextRetry = PropBag.ReadProperty("BtnTextRetry", m_def_BtnTextRetry)
    m_BtnTextIgnore = PropBag.ReadProperty("BtnTextIgnore", m_def_BtnTextIgnore)
    m_BtnTextHelp = PropBag.ReadProperty("BtnTextHelp", m_def_BtnTextHelp)
    m_BtnTextOK = PropBag.ReadProperty("BtnTextok", m_def_BtnTextOK)
    Set m_MessageFont = PropBag.ReadProperty("MessageFont", Ambient.Font)
    Set m_TextBoxFont = PropBag.ReadProperty("TextBoxFont", Ambient.Font)
    Set m_PicExclamation = PropBag.ReadProperty("PicExclamation", Nothing)
    Set m_PicQuestion = PropBag.ReadProperty("PicQuestion", Nothing)
    Set m_PicCritical = PropBag.ReadProperty("PicCritical", Nothing)
    Set m_PicInformation = PropBag.ReadProperty("PicInformation", Nothing)
    m_MSGboxType = PropBag.ReadProperty("MSGboxType", m_def_MSGboxType)
    Set m_ButtonPicBackGround = PropBag.ReadProperty("ButtonPicBackGround", Nothing)
    Set m_ButtonPicBackGroundDisabled = PropBag.ReadProperty("ButtonPicBackGroundDisabled", Nothing)
    Set m_ButtonPicLeft = PropBag.ReadProperty("ButtonPicLeft", Nothing)
    Set m_ButtonPicLeftDisabled = PropBag.ReadProperty("ButtonPicLeftDisabled", Nothing)
    Set m_ButtonPicRight = PropBag.ReadProperty("ButtonPicRight", Nothing)
    Set m_SkinPicTopleft = PropBag.ReadProperty("SkinPicTopleft", Nothing)
    Set m_SkinPicTopMiddle = PropBag.ReadProperty("SkinPicTopMiddle", Nothing)
    Set m_SkinPicTopRight = PropBag.ReadProperty("SkinPicTopRight", Nothing)
    Set m_SkinPicMiddleLeft = PropBag.ReadProperty("SkinPicMiddleLeft", Nothing)
    Set m_SkinPicMiddleRight = PropBag.ReadProperty("SkinPicMiddleRight", Nothing)
    Set m_SkinPicBottomLeft = PropBag.ReadProperty("SkinPicBottomLeft", Nothing)
    Set m_SkinPicBottomMiddle = PropBag.ReadProperty("SkinPicBottomMiddle", Nothing)
    Set m_SkinPicBottomRight = PropBag.ReadProperty("SkinPicBottomRight", Nothing)
    Set m_SkinPicBackGround = PropBag.ReadProperty("SkinPicBackGround", Nothing)
    Set m_SkinCaptionFont = PropBag.ReadProperty("SkinCaptionFont", Ambient.Font)
    m_SkinCaptionTop = PropBag.ReadProperty("SkinCaptionTop", m_def_SkinCaptionTop)
    m_SkinCaptionLeft = PropBag.ReadProperty("SkinCaptionLeft", m_def_SkinCaptionLeft)
    m_SkinCaptionForeColor = PropBag.ReadProperty("SkinCaptionForeColor", m_def_SkinCaptionForeColor)
    m_SkinCaptionShadowColor = PropBag.ReadProperty("SkinCaptionShadowColor", m_def_SkinCaptionShadowColor)
    m_SkinOpacity = PropBag.ReadProperty("SkinOpacity", m_def_SkinOpacity)
    Set m_ButtonPicDown = PropBag.ReadProperty("ButtonPicDown", Nothing)
    Set m_ButtonPicRightDisabled = PropBag.ReadProperty("ButtonPicRightDisabled", Nothing)
    Set m_ButtonPicRightDisabled = PropBag.ReadProperty("ButtonPicRightDisabled", Nothing)
    m_RedrawSkinAtStart = PropBag.ReadProperty("RedrawSkinAtStart", m_def_RedrawSkinAtStart)
    On Error Resume Next
    m_Dimensions.ButtonWidthOkOnly = PropBag.ReadProperty("ButtonWidthOkOnly")
    m_Dimensions.ButtonWidthOkCancel = PropBag.ReadProperty("ButtonWidthOkCancel")
    m_Dimensions.ButtonWidthYesNo = PropBag.ReadProperty("ButtonWidthYesNo")
    m_Dimensions.ButtonWidthYesNoCancel = PropBag.ReadProperty("ButtonWidthYesNoCancel")
    m_Dimensions.ButtonWidthAbortRetryIgnore = PropBag.ReadProperty("ButtonWidthAbortRetryIgnore")
    m_Dimensions.ButtonWidthOkOnlyHelp = PropBag.ReadProperty("ButtonWidthOkOnlyHelp")
    m_Dimensions.ButtonWidthOkCancelHelp = PropBag.ReadProperty("ButtonWidthOkCancelHelp")
    m_Dimensions.ButtonWidthYesNoHelp = PropBag.ReadProperty("ButtonWidthYesNoHelp")
    m_Dimensions.ButtonWidthYesNoCancelHelp = PropBag.ReadProperty("ButtonWidthYesNoCancelHelp")
    m_Dimensions.ButtonWidthAbortRetryIgnoreHelp = PropBag.ReadProperty("ButtonWidthAbortRetryIgnoreHelp")
    m_Dimensions.ButtonTop = PropBag.ReadProperty("ButtonHeight")
    m_Dimensions.FormWidthOkOnly = PropBag.ReadProperty("FormWidthOkOnly")
    m_Dimensions.FormWidthOkCancel = PropBag.ReadProperty("FormWidthOkCancel")
    m_Dimensions.FormWidthYesNo = PropBag.ReadProperty("FormWidthYesNo")
    m_Dimensions.FormWidthYesNoCancel = PropBag.ReadProperty("FormWidthYesNoCancel")
    m_Dimensions.FormWidthAbortRetryIgnore = PropBag.ReadProperty("FormWidthAbortRetryIgnore")
    m_Dimensions.FormWidthOkOnlyHelp = PropBag.ReadProperty("FormWidthOkOnlyHelp")
    m_Dimensions.FormWidthOkCancelHelp = PropBag.ReadProperty("FormWidthOkCancelHelp")
    m_Dimensions.FormWidthYesNoHelp = PropBag.ReadProperty("FormWidthYesNoHelp")
    m_Dimensions.FormWidthYesNoCancelHelp = PropBag.ReadProperty("FormWidthYesNoCancelHelp")
    m_Dimensions.FormWidthAbortRetryIgnoreHelp = PropBag.ReadProperty("FormWidthAbortRetryIgnoreHelp")
    SetControlsProps
End Sub


'=====================================================
' Write property values to storage
'=====================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DisplayPosition", m_DisplayPosition, m_def_DisplayPosition)
    Call PropBag.WriteProperty("MSGboxType", m_MSGboxType, m_def_MSGboxType)
    Call PropBag.WriteProperty("MessageIcon", m_MessageIcon, m_def_MessageIcon)
    Call PropBag.WriteProperty("MessageBackStartColor", m_MessageBackStartColor, m_def_MessageBackStartColor)
    Call PropBag.WriteProperty("MessageBackEndColor", m_MessageBackEndColor, m_def_MessageBackEndColor)
    Call PropBag.WriteProperty("MessageForeColor", m_MessageForeColor, m_def_MessageForeColor)
    Call PropBag.WriteProperty("TextBoxBackColor", m_TextBoxBackColor, m_def_TextBoxBackColor)
    Call PropBag.WriteProperty("TextBoxForeColor", m_TextBoxForeColor, m_def_TextBoxForeColor)
    Call PropBag.WriteProperty("TextBoxBorderColor", m_TextBoxBorderColor, m_def_TextBoxBorderColor)
    Call PropBag.WriteProperty("BtnTextNo", m_BtnTextNo, m_def_BtnTextYes)
    Call PropBag.WriteProperty("BtnTextYes", m_BtnTextYes, m_def_BtnTextNo)
    Call PropBag.WriteProperty("BtnTextCancel", m_BtnTextCancel, m_def_BtnTextCancel)
    Call PropBag.WriteProperty("BtnTextOK", m_BtnTextOK, m_def_BtnTextOK)
    Call PropBag.WriteProperty("BtnTextAbort", m_BtnTextAbort, m_def_BtnTextAbort)
    Call PropBag.WriteProperty("BtnTextRetry", m_BtnTextRetry, m_def_BtnTextRetry)
    Call PropBag.WriteProperty("BtnTextIgnore", m_BtnTextIgnore, m_def_BtnTextIgnore)
    Call PropBag.WriteProperty("BtnTextHelp", m_BtnTextHelp, m_def_BtnTextHelp)
    Call PropBag.WriteProperty("MessageFont", m_MessageFont, Ambient.Font)
    Call PropBag.WriteProperty("TextBoxFont", m_TextBoxFont, Ambient.Font)
    Call PropBag.WriteProperty("PicExclamation", m_PicExclamation, imgExclamation.Picture)
    Call PropBag.WriteProperty("PicQuestion", m_PicQuestion, imgQuestion.Picture)
    Call PropBag.WriteProperty("PicCritical", m_PicCritical, imgCritical.Picture)
    Call PropBag.WriteProperty("PicInformation", m_PicInformation, imgInfo.Picture)
    Call PropBag.WriteProperty("ButtonPicDown", m_ButtonPicDown, Nothing)
    Call PropBag.WriteProperty("ButtonPicBackGround", m_ButtonPicBackGround, Nothing)
    Call PropBag.WriteProperty("ButtonPicLeft", m_ButtonPicLeft, Nothing)
    Call PropBag.WriteProperty("ButtonPicRight", m_ButtonPicRight, Nothing)
    Call PropBag.WriteProperty("ButtonPicDown", m_ButtonPicDown, Nothing)
    Call PropBag.WriteProperty("ButtonPicBackGroundDisabled", m_ButtonPicBackGroundDisabled, Nothing)
    Call PropBag.WriteProperty("ButtonPicLeftDisabled", m_ButtonPicLeftDisabled, Nothing)
    Call PropBag.WriteProperty("ButtonPicRightDisabled", m_ButtonPicRightDisabled, Nothing)
    Call PropBag.WriteProperty("ButtonPicDownDisabled", m_ButtonPicDownDisabled, Nothing)
    Call PropBag.WriteProperty("ButtonForeColor", m_ButtonForeColor, m_def_ButtonForeColor)
    Call PropBag.WriteProperty("ButtonHooverColor", m_ButtonhooverColor, m_def_ButtonHooverColor)
    Call PropBag.WriteProperty("ButtonShadowColor", m_ButtonShadowColor, m_def_ButtonShadowColor)
    Call PropBag.WriteProperty("ButtonDisabledColor", m_ButtonDisabledColor, m_def_ButtonDisabledColor)
    Call PropBag.WriteProperty("ButtonFont", m_ButtonFont, Ambient.Font)
    Call PropBag.WriteProperty("SkinPicTopleft", m_SkinPicTopleft, Nothing)
    Call PropBag.WriteProperty("SkinPicTopMiddle", m_SkinPicTopMiddle, Nothing)
    Call PropBag.WriteProperty("SkinPicTopRight", m_SkinPicTopRight, Nothing)
    Call PropBag.WriteProperty("SkinPicMiddleLeft", m_SkinPicMiddleLeft, Nothing)
    Call PropBag.WriteProperty("SkinPicMiddleRight", m_SkinPicMiddleRight, Nothing)
    Call PropBag.WriteProperty("SkinPicBottomLeft", m_SkinPicBottomLeft, Nothing)
    Call PropBag.WriteProperty("SkinPicBottomMiddle", m_SkinPicBottomMiddle, Nothing)
    Call PropBag.WriteProperty("SkinPicBottomRight", m_SkinPicBottomRight, Nothing)
    Call PropBag.WriteProperty("SkinPicBackGround", m_SkinPicBackGround, Nothing)
    Call PropBag.WriteProperty("SkinCaptionFont", m_SkinCaptionFont, Ambient.Font)
    Call PropBag.WriteProperty("SkinCaptionForeColor", m_SkinCaptionForeColor, m_def_SkinCaptionForeColor)
    Call PropBag.WriteProperty("SkinCaptionShadowColor", m_SkinCaptionShadowColor, m_def_SkinCaptionShadowColor)
    Call PropBag.WriteProperty("SkinCaptionTop", m_SkinCaptionTop, m_def_SkinCaptionTop)
    Call PropBag.WriteProperty("SkinCaptionLeft", m_SkinCaptionLeft, m_def_SkinCaptionLeft)
    Call PropBag.WriteProperty("SkinOpacity", m_SkinOpacity, m_def_SkinOpacity)
    Call PropBag.WriteProperty("RedrawSkinAtStart", m_RedrawSkinAtStart, m_def_RedrawSkinAtStart)
    'ReCalButtonAndFormWitdh '!!!!! to test
    PropBag.WriteProperty "ButtonWidthOkOnly", m_Dimensions.ButtonWidthOkOnly
    PropBag.WriteProperty "ButtonWidthOkCancel", m_Dimensions.ButtonWidthOkCancel
    PropBag.WriteProperty "ButtonWidthYesNo", m_Dimensions.ButtonWidthYesNo
    PropBag.WriteProperty "ButtonWidthYesNoCancel", m_Dimensions.ButtonWidthYesNoCancel
    PropBag.WriteProperty "ButtonWidthAbortRetryIgnore", m_Dimensions.ButtonWidthAbortRetryIgnore
    PropBag.WriteProperty "ButtonWidthOkOnlyHelp", m_Dimensions.ButtonWidthOkOnlyHelp
    PropBag.WriteProperty "ButtonWidthOkCancelHelp", m_Dimensions.ButtonWidthOkCancelHelp
    PropBag.WriteProperty "ButtonWidthYesNoHelp", m_Dimensions.ButtonWidthYesNoHelp
    PropBag.WriteProperty "ButtonWidthYesNoCancelHelp", m_Dimensions.ButtonWidthYesNoCancelHelp
    PropBag.WriteProperty "ButtonWidthAbortRetryIgnoreHelp", m_Dimensions.ButtonWidthAbortRetryIgnoreHelp
    PropBag.WriteProperty "ButtonHeight", m_Dimensions.ButtonTop
    PropBag.WriteProperty "FormWidthOkOnly", m_Dimensions.FormWidthOkOnly
    PropBag.WriteProperty "FormWidthOkCancel", m_Dimensions.FormWidthOkCancel
    PropBag.WriteProperty "FormWidthYesNo", m_Dimensions.FormWidthYesNo
    PropBag.WriteProperty "FormWidthYesNoCancel", m_Dimensions.FormWidthYesNoCancel
    PropBag.WriteProperty "FormWidthAbortRetryIgnore", m_Dimensions.FormWidthAbortRetryIgnore
    PropBag.WriteProperty "FormWidthOkOnlyHelp", m_Dimensions.FormWidthOkOnlyHelp
    PropBag.WriteProperty "FormWidthOkCancelHelp", m_Dimensions.FormWidthOkCancelHelp
    PropBag.WriteProperty "FormWidthYesNoHelp", m_Dimensions.FormWidthYesNoHelp
    PropBag.WriteProperty "FormWidthYesNoCancelHelp", m_Dimensions.FormWidthYesNoCancelHelp
    PropBag.WriteProperty "FormWidthAbortRetryIgnoreHelp", m_Dimensions.FormWidthAbortRetryIgnoreHelp
End Sub


'=======================================================================================================
'USER CONTROL GET AND LET
'=======================================================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get MessageBackStartColor() As OLE_COLOR
    MessageBackStartColor = m_MessageBackStartColor
End Property
Public Property Let MessageBackStartColor(ByVal New_MessageBackStartColor As OLE_COLOR)
    m_MessageBackStartColor = New_MessageBackStartColor
    PropertyChanged "MessageBackStartColor"
    RepaintMessage UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000F&
Public Property Get MessageBackEndColor() As OLE_COLOR
    MessageBackEndColor = m_MessageBackEndColor
End Property
Public Property Let MessageBackEndColor(ByVal New_MessageBackEndColor As OLE_COLOR)
    m_MessageBackEndColor = New_MessageBackEndColor
    PropertyChanged "MessageBackEndColor"
    RepaintMessage UserControl.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SkinCaptionTop() As Long
    SkinCaptionTop = m_SkinCaptionTop
End Property
Public Property Let SkinCaptionTop(ByVal New_SkinCaptionTop As Long)
    m_SkinCaptionTop = New_SkinCaptionTop
    PropertyChanged "SkinCaptionTop"
    frmMSGbox.DMmsgSkin1.CaptionTop = m_SkinCaptionTop
    frmMSGbox.DMmsgSkin1.RepaintSkin frmMSGbox
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get RedrawSkinAtStart() As Boolean
    RedrawSkinAtStart = m_RedrawSkinAtStart
End Property
Public Property Let RedrawSkinAtStart(ByVal New_RedrawSkinAtStart As Boolean)
    m_RedrawSkinAtStart = New_RedrawSkinAtStart
    PropertyChanged "RedrawSkinAtStart"
'    frmMSGbox.DMmsgSkin1.RepaintSkin = frmMSGbox
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SkinCaptionLeft() As Long
    SkinCaptionLeft = m_SkinCaptionLeft
End Property
Public Property Let SkinCaptionLeft(ByVal New_SkinCaptionLeft As Long)
    m_SkinCaptionLeft = New_SkinCaptionLeft
    PropertyChanged "SkinCaptionLeft"
    frmMSGbox.DMmsgSkin1.CaptionLeft = m_SkinCaptionLeft
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,FALSE
Public Property Get DisplayPosition() As StartupPositions
    DisplayPosition = m_DisplayPosition
End Property

Public Property Let DisplayPosition(ByVal New_DisplayPosition As StartupPositions)
    m_DisplayPosition = New_DisplayPosition
    PropertyChanged "DisplayPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,0
Public Property Get MessageIcon() As MessageIcons
    MessageIcon = m_MessageIcon
End Property

Public Property Let MessageIcon(ByVal New_MessageIcon As MessageIcons)
    m_MessageIcon = New_MessageIcon
    PropertyChanged "MessageIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MessageForeColor() As OLE_COLOR
    MessageForeColor = m_MessageForeColor
End Property

Public Property Let MessageForeColor(ByVal New_MessageForeColor As OLE_COLOR)
    m_MessageForeColor = New_MessageForeColor
    PropertyChanged "MessageForeColor"
    frmMSGbox.lblMessage.ForeColor = m_MessageForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextBoxBackColor() As OLE_COLOR
    TextBoxBackColor = m_TextBoxBackColor
End Property

Public Property Let TextBoxBackColor(ByVal New_TextBoxBackColor As OLE_COLOR)
    m_TextBoxBackColor = New_TextBoxBackColor
    PropertyChanged "TextBoxBackColor"
    frmMSGbox.txtInput.BackColor = m_TextBoxBackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextBoxForeColor() As OLE_COLOR
    TextBoxForeColor = m_TextBoxForeColor
End Property

Public Property Let TextBoxForeColor(ByVal New_TextBoxForeColor As OLE_COLOR)
    m_TextBoxForeColor = New_TextBoxForeColor
    PropertyChanged "TextBoxForeColor"
    frmMSGbox.txtInput.ForeColor = m_TextBoxForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get TextBoxBorderColor() As OLE_COLOR
    TextBoxBorderColor = m_TextBoxBorderColor
End Property

Public Property Let TextBoxBorderColor(ByVal New_TextBoxBorderColor As OLE_COLOR)
    m_TextBoxBorderColor = New_TextBoxBorderColor
    PropertyChanged "TextBoxBorderColor"
    frmMSGbox.txtInput.BorderColor = m_TextBoxBorderColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonForeColor() As OLE_COLOR
    ButtonForeColor = m_ButtonForeColor
End Property

Public Property Let ButtonForeColor(ByVal New_ButtonForeColor As OLE_COLOR)
    m_ButtonForeColor = New_ButtonForeColor
    PropertyChanged "ButtonForeColor"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonhooverColor() As OLE_COLOR
    ButtonhooverColor = m_ButtonhooverColor
End Property

Public Property Let ButtonhooverColor(ByVal New_ButtonhooverColor As OLE_COLOR)
    m_ButtonhooverColor = New_ButtonhooverColor
    PropertyChanged "ButtonhooverColor"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonShadowColor() As OLE_COLOR
    ButtonShadowColor = m_ButtonShadowColor
End Property

Public Property Let ButtonShadowColor(ByVal New_ButtonShadowColor As OLE_COLOR)
    m_ButtonShadowColor = New_ButtonShadowColor
    PropertyChanged "ButtonShadowColor"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonDisabledColor() As OLE_COLOR
    ButtonDisabledColor = m_ButtonDisabledColor
End Property

Public Property Let ButtonDisabledColor(ByVal New_ButtonDisabledColor As OLE_COLOR)
    m_ButtonDisabledColor = New_ButtonDisabledColor
    PropertyChanged "ButtonDisabledColor"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get ButtonFont() As Font
    Set ButtonFont = m_ButtonFont
End Property

Public Property Set ButtonFont(ByVal New_ButtonFont As Font)
    Set m_ButtonFont = New_ButtonFont
    PropertyChanged "ButtonFont"
    ReCalButtonAndFormWitdh
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextNo() As String
    BtnTextNo = m_BtnTextNo
End Property

Public Property Let BtnTextNo(ByVal New_BtnTextNo As String)
    m_BtnTextNo = New_BtnTextNo
    PropertyChanged "BtnTextNo"
    frmMSGbox.cmdNo.Caption = m_BtnTextNo
    ReCalButtonAndFormWitdh "YES"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextOK() As String
    BtnTextOK = m_BtnTextOK
End Property

Public Property Let BtnTextOK(ByVal New_BtnTextOK As String)
    m_BtnTextOK = New_BtnTextOK
    PropertyChanged "BtnTextOK"
    frmMSGbox.cmdOK.Caption = m_BtnTextOK
    ReCalButtonAndFormWitdh "OK"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextYes() As String
    BtnTextYes = m_BtnTextYes
End Property

Public Property Let BtnTextYes(ByVal New_BtnTextYes As String)
    m_BtnTextYes = New_BtnTextYes
    PropertyChanged "BtnTextYes"
    frmMSGbox.cmdYes.Caption = m_BtnTextYes
    ReCalButtonAndFormWitdh "YES"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextCancel() As String
    BtnTextCancel = m_BtnTextCancel
End Property

Public Property Let BtnTextCancel(ByVal New_BtnTextCancel As String)
    m_BtnTextCancel = New_BtnTextCancel
    PropertyChanged "BtnTextCancel"
    frmMSGbox.cmdCancel.Caption = m_BtnTextCancel
    ReCalButtonAndFormWitdh "OK"
    ReCalButtonAndFormWitdh "YES"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextAbort() As String
    BtnTextAbort = m_BtnTextAbort
End Property

Public Property Let BtnTextAbort(ByVal New_BtnTextAbort As String)
    m_BtnTextAbort = New_BtnTextAbort
    PropertyChanged "BtnTextAbort"
    frmMSGbox.cmdAbort.Caption = m_BtnTextAbort
    ReCalButtonAndFormWitdh "ABORT"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextRetry() As String
    BtnTextRetry = m_BtnTextRetry
End Property

Public Property Let BtnTextRetry(ByVal New_BtnTextRetry As String)
    m_BtnTextRetry = New_BtnTextRetry
    PropertyChanged "BtnTextRetry"
    frmMSGbox.cmdRetry.Caption = m_BtnTextRetry
    ReCalButtonAndFormWitdh "ABORT"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextIgnore() As String
    BtnTextIgnore = m_BtnTextIgnore
End Property

Public Property Let BtnTextIgnore(ByVal New_BtnTextIgnore As String)
    m_BtnTextIgnore = New_BtnTextIgnore
    PropertyChanged "BtnTextIgnore"
    frmMSGbox.cmdIgnore.Caption = m_BtnTextIgnore
    ReCalButtonAndFormWitdh "ABORT"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get BtnTextHelp() As String
    BtnTextHelp = m_BtnTextHelp
End Property

Public Property Let BtnTextHelp(ByVal New_BtnTextHelp As String)
    m_BtnTextHelp = New_BtnTextHelp
    PropertyChanged "BtnTextHelp"
    frmMSGbox.cmdHelp.Caption = m_BtnTextHelp
    ReCalButtonAndFormWitdh
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get MessageFont() As Font
    Set MessageFont = m_MessageFont
End Property

Public Property Set MessageFont(ByVal New_MessageFont As Font)
    Set m_MessageFont = New_MessageFont
    PropertyChanged "MessageFont"
    Set frmMSGbox.lblMessage.Font = m_MessageFont
    Set frmMSGbox.lblTest.Font = m_MessageFont
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get TextBoxFont() As Font
    Set TextBoxFont = m_TextBoxFont
End Property

Public Property Set TextBoxFont(ByVal New_TextBoxFont As Font)
    Set m_TextBoxFont = New_TextBoxFont
    PropertyChanged "TextBoxFont"
    Set frmMSGbox.txtInput.Font = m_MessageFont
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicExclamation() As Picture
    Set PicExclamation = m_PicExclamation
End Property

Public Property Set PicExclamation(ByVal New_PicExclamation As Picture)
    Set m_PicExclamation = New_PicExclamation
    PropertyChanged "PicExclamation"
    Set frmMSGbox.imgExclamation.Picture = m_PicExclamation
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicQuestion() As Picture
    Set PicQuestion = m_PicQuestion
End Property

Public Property Set PicQuestion(ByVal New_PicQuestion As Picture)
    Set m_PicQuestion = New_PicQuestion
    PropertyChanged "PicQuestion"
    Set frmMSGbox.imgQuestion.Picture = m_PicQuestion
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicCritical() As Picture
    Set PicCritical = m_PicCritical
End Property

Public Property Set PicCritical(ByVal New_PicCritical As Picture)
    Set m_PicCritical = New_PicCritical
    PropertyChanged "PicCritical"
    Set frmMSGbox.imgCritical.Picture = m_PicCritical
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PicInformation() As Picture
    Set PicInformation = m_PicInformation
End Property

Public Property Set PicInformation(ByVal New_PicInformation As Picture)
    Set m_PicInformation = New_PicInformation
    PropertyChanged "PicInformation"
    Set frmMSGbox.imgInfo.Picture = m_PicInformation
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,0,0,0
Public Property Get MSGboxType() As MSGboxtypes
    MSGboxType = m_MSGboxType
End Property

Public Property Let MSGboxType(ByVal New_MSGboxType As MSGboxtypes)
    m_MSGboxType = New_MSGboxType
    PropertyChanged "MSGboxType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicBackGround() As Picture
    Set ButtonPicBackGround = m_ButtonPicBackGround
End Property

Public Property Set ButtonPicBackGround(ByVal New_ButtonPicBackGround As Picture)
    Set m_ButtonPicBackGround = New_ButtonPicBackGround
    PropertyChanged "ButtonPicBackGround"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicBackGroundDisabled() As Picture
    Set ButtonPicBackGroundDisabled = m_ButtonPicBackGroundDisabled
End Property

Public Property Set ButtonPicBackGroundDisabled(ByVal New_ButtonPicBackGroundDisabled As Picture)
    Set m_ButtonPicBackGroundDisabled = New_ButtonPicBackGroundDisabled
    PropertyChanged "ButtonPicBackGroundDisabled"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicLeft() As Picture
    Set ButtonPicLeft = m_ButtonPicLeft
End Property

Public Property Set ButtonPicLeft(ByVal New_ButtonPicLeft As Picture)
    Set m_ButtonPicLeft = New_ButtonPicLeft
    PropertyChanged "ButtonPicLeft"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicLeftDisabled() As Picture
    Set ButtonPicLeftDisabled = m_ButtonPicLeftDisabled
End Property

Public Property Set ButtonPicLeftDisabled(ByVal New_ButtonPicLeftDisabled As Picture)
    Set m_ButtonPicLeftDisabled = New_ButtonPicLeftDisabled
    PropertyChanged "ButtonPicLeftDisabled"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicRight() As Picture
    Set ButtonPicRight = m_ButtonPicRight
End Property

Public Property Set ButtonPicRight(ByVal New_ButtonPicRight As Picture)
    Set m_ButtonPicRight = New_ButtonPicRight
    PropertyChanged "ButtonPicRight"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicTopleft() As Picture
    Set SkinPicTopleft = m_SkinPicTopleft
End Property

Public Property Set SkinPicTopleft(ByVal New_SkinPicTopleft As Picture)
    Set m_SkinPicTopleft = New_SkinPicTopleft
    PropertyChanged "SkinPicTopleft"
    Set frmMSGbox.DMmsgSkin1.PictureTopLeft = m_SkinPicTopleft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicTopMiddle() As Picture
    Set SkinPicTopMiddle = m_SkinPicTopMiddle
End Property

Public Property Set SkinPicTopMiddle(ByVal New_SkinPicTopMiddle As Picture)
    Set m_SkinPicTopMiddle = New_SkinPicTopMiddle
    PropertyChanged "SkinPicTopMiddle"
    Set frmMSGbox.DMmsgSkin1.PictureTopMiddle = m_SkinPicTopMiddle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicTopRight() As Picture
    Set SkinPicTopRight = m_SkinPicTopRight
End Property

Public Property Set SkinPicTopRight(ByVal New_SkinPicTopRight As Picture)
    Set m_SkinPicTopRight = New_SkinPicTopRight
    PropertyChanged "SkinPicTopRight"
    Set frmMSGbox.DMmsgSkin1.PictureTopRight = m_SkinPicTopRight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicMiddleLeft() As Picture
    Set SkinPicMiddleLeft = m_SkinPicMiddleLeft
End Property

Public Property Set SkinPicMiddleLeft(ByVal New_SkinPicMiddleLeft As Picture)
    Set m_SkinPicMiddleLeft = New_SkinPicMiddleLeft
    PropertyChanged "SkinPicMiddleLeft"
    Set frmMSGbox.DMmsgSkin1.PictureMiddleLeft = m_SkinPicMiddleLeft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicMiddleRight() As Picture
    Set SkinPicMiddleRight = m_SkinPicMiddleRight
End Property

Public Property Set SkinPicMiddleRight(ByVal New_SkinPicMiddleRight As Picture)
    Set m_SkinPicMiddleRight = New_SkinPicMiddleRight
    PropertyChanged "SkinPicMiddleRight"
    Set frmMSGbox.DMmsgSkin1.PictureMiddleRight = m_SkinPicMiddleRight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicBottomLeft() As Picture
    Set SkinPicBottomLeft = m_SkinPicBottomLeft
End Property

Public Property Set SkinPicBottomLeft(ByVal New_SkinPicBottomLeft As Picture)
    Set m_SkinPicBottomLeft = New_SkinPicBottomLeft
    PropertyChanged "SkinPicBottomLeft"
    Set frmMSGbox.DMmsgSkin1.PictureBottomLeft = m_SkinPicBottomLeft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicBottomMiddle() As Picture
    Set SkinPicBottomMiddle = m_SkinPicBottomMiddle
End Property

Public Property Set SkinPicBottomMiddle(ByVal New_SkinPicBottomMiddle As Picture)
    Set m_SkinPicBottomMiddle = New_SkinPicBottomMiddle
    PropertyChanged "SkinPicBottomMiddle"
    Set frmMSGbox.DMmsgSkin1.PictureBottomMiddle = m_SkinPicBottomMiddle
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicBottomRight() As Picture
    Set SkinPicBottomRight = m_SkinPicBottomRight
End Property

Public Property Set SkinPicBottomRight(ByVal New_SkinPicBottomRight As Picture)
    Set m_SkinPicBottomRight = New_SkinPicBottomRight
    PropertyChanged "SkinPicBottomRight"
    Set frmMSGbox.DMmsgSkin1.PictureBottomRight = m_SkinPicBottomRight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get SkinPicBackGround() As Picture
    Set SkinPicBackGround = m_SkinPicBackGround
End Property

Public Property Set SkinPicBackGround(ByVal New_SkinPicBackGround As Picture)
    Set m_SkinPicBackGround = New_SkinPicBackGround
    PropertyChanged "SkinPicBackGround"
    Set frmMSGbox.DMmsgSkin1.PictureBackGround = m_SkinPicBackGround
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get SkinCaptionFont() As Font
    Set SkinCaptionFont = m_SkinCaptionFont
End Property

Public Property Set SkinCaptionFont(ByVal New_SkinCaptionFont As Font)
    Set m_SkinCaptionFont = New_SkinCaptionFont
    PropertyChanged "SkinCaptionFont"
    Set frmMSGbox.DMmsgSkin1.Font = m_SkinCaptionFont
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SkinCaptionForeColor() As OLE_COLOR
    SkinCaptionForeColor = m_SkinCaptionForeColor
End Property

Public Property Let SkinCaptionForeColor(ByVal New_SkinCaptionForeColor As OLE_COLOR)
    m_SkinCaptionForeColor = New_SkinCaptionForeColor
    PropertyChanged "SkinCaptionForeColor"
    frmMSGbox.DMmsgSkin1.CaptionForeColor = m_SkinCaptionForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SkinCaptionShadowColor() As OLE_COLOR
    SkinCaptionShadowColor = m_SkinCaptionShadowColor
End Property

Public Property Let SkinCaptionShadowColor(ByVal New_SkinCaptionShadowColor As OLE_COLOR)
    m_SkinCaptionShadowColor = New_SkinCaptionShadowColor
    PropertyChanged "SkinCaptionShadowColor"
    frmMSGbox.DMmsgSkin1.CaptionShadowColor = m_SkinCaptionShadowColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,0
Public Property Get SkinOpacity() As Byte
    SkinOpacity = m_SkinOpacity
End Property

Public Property Let SkinOpacity(ByVal New_SkinOpacity As Byte)
    m_SkinOpacity = New_SkinOpacity
    PropertyChanged "SkinOpacity"
    frmMSGbox.DMmsgSkin1.Opacity = m_SkinOpacity
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicDown() As Picture
    Set ButtonPicDown = m_ButtonPicDown
End Property

Public Property Set ButtonPicDown(ByVal New_ButtonPicDown As Picture)
    Set m_ButtonPicDown = New_ButtonPicDown
    PropertyChanged "ButtonPicDown"
    ResetSkinButtonProps
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get ButtonPicRightDisabled() As Picture
    Set ButtonPicRightDisabled = m_ButtonPicRightDisabled
End Property

Public Property Set ButtonPicRightDisabled(ByVal New_ButtonPicRightDisabled As Picture)
    Set m_ButtonPicRightDisabled = New_ButtonPicRightDisabled
    PropertyChanged "ButtonPicRightDisabled"
    ResetSkinButtonProps
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Private Property Let Dimensions(New_Dimensions As Dimension)
    m_Dimensions = New_Dimensions
    PropertyChanged "Dimensions"
End Property


