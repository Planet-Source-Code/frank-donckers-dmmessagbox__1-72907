VERSION 5.00
Begin VB.UserControl DMTextBox 
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "DMTextBox.ctx":0000
   Begin VB.PictureBox DotBottomRight 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   3000
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   2640
      Width           =   15
   End
   Begin VB.PictureBox DotBottomLeft 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   960
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   2520
      Width           =   15
   End
   Begin VB.PictureBox DotTopRight 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   3240
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   1680
      Width           =   15
   End
   Begin VB.PictureBox DotTopLeft 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   1560
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   1680
      Width           =   15
   End
   Begin VB.PictureBox lnRight 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   600
      ScaleHeight     =   780
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   2040
      Width           =   15
   End
   Begin VB.PictureBox lnTop 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   3000
      ScaleHeight     =   15
      ScaleWidth      =   1575
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox lnLeft 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   840
      ScaleHeight     =   900
      ScaleWidth      =   15
      TabIndex        =   3
      Top             =   2040
      Width           =   15
   End
   Begin VB.PictureBox lnBottom 
      BackColor       =   &H00B99D7F&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   2880
      ScaleHeight     =   15
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox MyTxt 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtXText 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Text            =   "DMTextBox"
      Top             =   15
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "DMTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:

'=====================================================
'States
'=====================================================
Public Enum States
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum

'=====================================================
'InputType
'=====================================================
Public Enum InputType
    AlfaNum = 0
    Alfa = 1
    Num = 2
End Enum

'=====================================================
'BorderStyle
'=====================================================
Public Enum BorderStyle
    BorderNone = 0
    BorderNormal = 1
    BorderRounded = 2
End Enum

'=====================================================
'Alignment
'=====================================================
Public Enum Alignment
    Alignleft = 0
    AlignRight = 1
    AlignCenter = 2
End Enum

Const m_def_ForeColor = vbBlack
Const m_def_BorderColor = &HB99D7F
Const m_def_BorderStyle = 1
Const m_def_BorderColorOver = &H96E7&
Const m_def_DataFields = ""
Const m_def_InputType = 0
Const m_def_BackColor = vbWhite
Const m_def_Enabled = True
Const m_def_Locked = False

'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_AutoSelect As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOver As OLE_COLOR
Dim m_DataFields As String
Dim m_InputType As InputType
Dim m_BorderStyle As BorderStyle
Dim m_BackColor As OLE_COLOR
Dim m_TransBorder As Byte
Dim m_Enabled As Boolean
Dim m_Locked As Boolean
Dim MyState As States

'=====================================================
' Event Declarations
'=====================================================
Event Change()
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=MyTxt,MyTxt,-1,MouseMove

' Draw borders
'=====================================================
Private Sub RedrawBorder()
On Error Resume Next
    With UserControl
        If TransBorder > 0 Then
            MyTxt.Width = .Width - 45 - (TransBorder * 2)
            MyTxt.Height = .Height - 45 - (TransBorder * 2)
            MyTxt.Left = TransBorder + 15
            MyTxt.Top = TransBorder + 15
        Else
            MyTxt.Width = .Width - 30
            MyTxt.Height = .Height - 30
            MyTxt.Left = 15
            MyTxt.Top = 15
        End If
    End With
End Sub


' Draw the textbox
'=====================================================
Private Function DrawTextBox(txt As TextBox, tBackcolor As ColorConstants, State As States)
    UserControl.Cls
    txt.BackColor = tBackcolor
    MyTxt.BackColor = tBackcolor
    UserControl.ScaleMode = 1
    txt.Appearance = 0
    txt.BorderStyle = 0
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    txt.ForeColor = m_ForeColor
    If BorderStyle = BorderRounded Then
        lnTop.Move 30, 0, UserControl.Width - 60, 15
        lnBottom.Move 30, UserControl.Height - 15, UserControl.Width - 60, UserControl.Height
        lnLeft.Move 0, 30, 15, UserControl.Height - 60
        lnRight.Move UserControl.Width - 15, 30, UserControl.Width, UserControl.Height - 60
        DotTopLeft.Move 15, 15
        DotTopRight.Move UserControl.Width - 30, 15
        DotBottomLeft.Move 15, UserControl.Height - 30
        DotBottomRight.Move UserControl.Width - 30, UserControl.Height - 30
    ElseIf BorderStyle = BorderNormal Then
        lnTop.Move 0, 0, UserControl.Width, 15
        lnBottom.Move 0, UserControl.Height - 15, UserControl.Width, UserControl.Height
        lnLeft.Move 0, 0, 15, UserControl.Height
        lnRight.Move UserControl.Width - 15, 0, UserControl.Width, UserControl.Height
    End If
    If State = Normal Then
        txt.BackColor = tBackcolor
        txt.Enabled = True
        txt.Locked = False
    ElseIf State = Disable Then
        txt.Enabled = False
        txt.BackColor = RGB(235, 235, 228)
        txt.ForeColor = RGB(161, 161, 146)
        lnTop.BackColor = RGB(161, 161, 146)
        lnBottom.BackColor = RGB(161, 161, 146)
        lnLeft.BackColor = RGB(161, 161, 146)
        lnRight.BackColor = RGB(161, 161, 146)
        DotTopLeft.BackColor = RGB(161, 161, 146)
        DotTopRight.BackColor = RGB(161, 161, 146)
        DotBottomLeft.BackColor = RGB(161, 161, 146)
        DotBottomRight.BackColor = RGB(161, 161, 146)
        txt.ForeColor = RGB(161, 161, 146)
    ElseIf State = ReadOnly Then
        txt.BackColor = tBackcolor
        txt.Enabled = True
        txt.Locked = True
    End If
    
End Function

'=====================================================
' Events
'=====================================================
Private Sub MyTxt_Change()
    RaiseEvent Change
End Sub

Private Sub MyTxt_Click()
    RaiseEvent Click
End Sub

Private Sub MyTxt_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub MyTxt_GotFocus()
    SetMyFocus m_BorderColorOver
    If m_AutoSelect = True Then
        MyTxt.SelStart = 0
        MyTxt.SelLength = Len(MyTxt.Text)
    End If
End Sub

Private Sub MyTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub MyTxt_KeyPress(KeyAscii As Integer)
    If InputType = Num Then
        Select Case KeyAscii
           Case 48 To 58 '0 - 9
           Case 8, 13, 27, 44, 46 'backspace, enter, esc,.,,
           Case 45
                If Len(Trim$(MyTxt.Text)) > 0 Then
                    KeyAscii = 0
                    Exit Sub
                End If
           
           Case 24, 3 'cut, copy
           Case 22 'paste (ctrl + v)
               If Not IsNumeric(Clipboard.GetText) Then Clipboard.Clear 'if not numeric
           Case Else
               KeyAscii = 0
        End Select
   ElseIf InputType = Alfa Then
        Select Case KeyAscii
           Case 48 To 58 '0 - 9
                KeyAscii = 0
           Case 8, 13, 27, 44, 46 'backspace, enter, esc,.,,
        End Select
   End If
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub SetMyFocus(LineColor As ColorConstants)
    UserControl.AutoRedraw = True
    UserControl.DrawWidth = 1
    lnTop.BackColor = LineColor
    lnBottom.BackColor = LineColor
    lnLeft.BackColor = LineColor
    lnRight.BackColor = LineColor
    DotTopLeft.BackColor = LineColor
    DotTopRight.BackColor = LineColor
    DotBottomLeft.BackColor = LineColor
    DotBottomRight.BackColor = LineColor
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    MyTxt.SetFocus
End Sub

Private Sub UserControl_ExitFocus()
    SetMyFocus m_BorderColor
End Sub

'=======================================================================================================
' USERCONTROL PROPERTIES
'=======================================================================================================

'=====================================================
' Resizing the control
'=====================================================
Private Sub UserControl_Resize()
    If UserControl.Height < (TextHeight(txtXText.Text) + 100) Then UserControl.Height = (TextHeight(txtXText.Text) + 100)
    If BorderStyle = BorderNone Then
        lnTop.Visible = False
        lnBottom.Visible = False
        lnLeft.Visible = False
        lnRight.Visible = False
        DotTopLeft.Visible = False
        DotTopRight.Visible = False
        DotBottomLeft.Visible = False
        DotBottomRight.Visible = False
    ElseIf BorderStyle = BorderNormal Then
        lnTop.Visible = True
        lnBottom.Visible = True
        lnLeft.Visible = True
        lnRight.Visible = True
        DotTopLeft.Visible = False
        DotTopRight.Visible = False
        DotBottomLeft.Visible = False
        DotBottomRight.Visible = False
    ElseIf BorderStyle = BorderRounded Then
        lnTop.Visible = True
        lnBottom.Visible = True
        lnLeft.Visible = True
        lnRight.Visible = True
        DotTopLeft.Visible = True
        DotTopRight.Visible = True
        DotBottomLeft.Visible = True
        DotBottomRight.Visible = True
    End If
    RedrawBorder
    Call DrawTextBox(MyTxt, BackColor, MyState)
End Sub


'=====================================================
' Initialize Properties for User Control
'=====================================================
Private Sub UserControl_InitProperties()
    m_DataFields = m_def_DataFields
    MyTxt.Text = UserControl.Extender.Name
    UserControl.Height = 330
    MyTxt.FontName = "Verdana"
    m_BorderColor = m_def_BorderColor
    m_BorderColorOver = m_def_BorderColorOver
    m_BorderStyle = BorderRounded
    m_BackColor = vbWhite
    UserControl_Resize
End Sub

'=====================================================
' Load property values from storage
'=====================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyState = Normal
    MyTxt.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_AutoSelect = PropBag.ReadProperty("AutoSelect", False)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    MyTxt.Enabled = m_Enabled
    Set MyTxt.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    MyTxt.ForeColor = m_ForeColor
    m_Locked = PropBag.ReadProperty("Locked", False)
    MyTxt.Locked = m_Locked
    MyTxt.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    MyTxt.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    MyTxt.SelStart = PropBag.ReadProperty("SelStart", 0)
    MyTxt.SelText = PropBag.ReadProperty("SelText", "")
    MyTxt.SelLength = PropBag.ReadProperty("SelLength", 0)
    MyTxt.Text = PropBag.ReadProperty("Text", "Text1")
    MyTxt.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOver = PropBag.ReadProperty("BorderColorOver", m_def_BorderColorOver)
    m_InputType = PropBag.ReadProperty("InputType", m_def_InputType)
    m_BorderStyle = PropBag.ReadProperty("Borderstyle", m_def_BorderStyle)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_TransBorder = PropBag.ReadProperty("TransBorder", 0)
    MyTxt.BackColor = m_BackColor
    lnTop.BackColor = m_BorderColor
    lnBottom.BackColor = m_BorderColor
    lnLeft.BackColor = m_BorderColor
    lnRight.BackColor = m_BorderColor
    DotTopLeft.BackColor = m_BorderColor
    DotTopRight.BackColor = m_BorderColor
    DotBottomLeft.BackColor = m_BorderColor
    DotBottomRight.BackColor = m_BorderColor
    If m_Locked = True Then MyState = ReadOnly
    If m_Enabled = False Then MyState = Disable
    If m_AutoSelect = True Then
        MyTxt.SelStart = 0
        MyTxt.SelLength = Len(MyTxt.Text)
    End If
End Sub

'=====================================================
' Write property values to storage
'=====================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Alignment", MyTxt.Alignment, 0)
    Call PropBag.WriteProperty("AutoSelect", m_AutoSelect, False)
    Call PropBag.WriteProperty("Enabled", MyTxt.Enabled, True)
    Call PropBag.WriteProperty("Font", MyTxt.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", MyTxt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("MaxLength", MyTxt.MaxLength, 0)
    Call PropBag.WriteProperty("PasswordChar", MyTxt.PasswordChar, "")
    Call PropBag.WriteProperty("SelStart", MyTxt.SelStart, 0)
    Call PropBag.WriteProperty("SelText", MyTxt.SelText, "")
    Call PropBag.WriteProperty("SelLength", MyTxt.SelLength, 0)
    Call PropBag.WriteProperty("Text", MyTxt.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipText", MyTxt.ToolTipText, "")
    Call PropBag.WriteProperty("Value", Val(MyTxt.Text), 0)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOver", m_BorderColorOver, m_def_BorderColorOver)
    Call PropBag.WriteProperty("InputType", m_InputType, m_def_InputType)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub

'=======================================================================================================
'USER CONTROL GET AND LET
'=======================================================================================================
Public Property Get Value() As Double
    Value = Val(MyTxt.Text)
End Property
Public Property Let Value(ByVal New_Value As Double)
    MyTxt.Text() = New_Value
    PropertyChanged "Value"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    DrawTextBox MyTxt, BackColor, MyState
    lnTop.BackColor = m_BorderColor
    lnBottom.BackColor = m_BorderColor
    lnLeft.BackColor = m_BorderColor
    lnRight.BackColor = m_BorderColor
    DotTopLeft.BackColor = m_BorderColor
    DotTopRight.BackColor = m_BorderColor
    DotBottomLeft.BackColor = m_BorderColor
    DotBottomRight.BackColor = m_BorderColor
End Property

Public Property Get BorderColorFocus() As OLE_COLOR
    BorderColorFocus = m_BorderColorOver
End Property
Public Property Let BorderColorFocus(ByVal New_BorderColorOver As OLE_COLOR)
    m_BorderColorOver = New_BorderColorOver
    PropertyChanged "BorderColorOver"
End Property

Public Property Get InputType() As InputType
    InputType = m_InputType
End Property
Public Property Let InputType(ByVal New_InputType As InputType)
    m_InputType = New_InputType
    PropertyChanged "InputType"
End Property

Public Property Get BorderStyle() As BorderStyle
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl_Resize
End Property

Public Property Get TransBorder() As Byte
    TransBorder = m_TransBorder
    If TransBorder > 256 Then TransBorder = 256
End Property
Public Property Let TransBorder(ByVal New_TransBorder As Byte)
    m_TransBorder = New_TransBorder
    If m_TransBorder > 256 Then m_TransBorder = 256
    PropertyChanged "TransBorder"
    UserControl_Resize
End Property

Public Property Get Alignment() As Alignment
    Alignment = MyTxt.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Alignment)
    If New_Alignment > 2 Then New_Alignment = 0
    MyTxt.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get AutoSelect() As Boolean
    AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(ByVal New_AutoSelect As Boolean)
    m_AutoSelect = New_AutoSelect
    PropertyChanged "AutoSelect"
End Property

Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
    If Locked = True Then MyState = ReadOnly
    UserControl_Resize
End Property

Public Property Get MaxLength() As Long
    MaxLength = MyTxt.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    MyTxt.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property


Public Property Get PasswordChar() As String
    PasswordChar = MyTxt.PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    MyTxt.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Public Property Get SelStart() As Long
    SelStart = MyTxt.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    MyTxt.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
    SelText = MyTxt.SelText
End Property
Public Property Let SelText(ByVal New_SelText As String)
    MyTxt.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get SelLength() As Long
    SelLength = MyTxt.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    MyTxt.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
Public Property Get Text() As String
    Text = MyTxt.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    MyTxt.Text() = New_Text
    PropertyChanged "Text"
End Property
Public Property Get ToolTipText() As String
    ToolTipText = MyTxt.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    MyTxt.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    MyTxt.Enabled = m_Enabled
    If New_Enabled Then
        MyState = Normal
        SetMyFocus RGB(127, 157, 185)
    Else
        MyState = Disable
        SetMyFocus RGB(191, 167, 128)
    End If
    UserControl_Resize
End Property

Public Property Get Font() As Font
    Set Font = MyTxt.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set MyTxt.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    MyTxt.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

