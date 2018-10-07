VERSION 5.00
Begin VB.UserControl ActiveDate 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ScaleHeight     =   1095
   ScaleWidth      =   1965
   ToolboxBitmap   =   "ActiveDate.ctx":0000
   Begin VB.PictureBox picBack 
      BackColor       =   &H80000005&
      Height          =   315
      Left            =   210
      ScaleHeight     =   255
      ScaleWidth      =   1365
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   300
      Width           =   1425
      Begin VB.TextBox txtDate 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         MaxLength       =   10
         TabIndex        =   0
         Top             =   15
         Width           =   1020
      End
      Begin VB.Image imgButton 
         Height          =   255
         Left            =   1125
         Picture         =   "ActiveDate.ctx":00FA
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblButton 
         Height          =   255
         Left            =   1125
         TabIndex        =   2
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Image imgDis 
      Height          =   255
      Left            =   825
      Picture         =   "ActiveDate.ctx":018F
      Top             =   735
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   255
      Left            =   510
      Picture         =   "ActiveDate.ctx":0226
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   255
      Left            =   225
      Picture         =   "ActiveDate.ctx":02B1
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ActiveDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Dim frm As Form

'Properties Default Constants
Const m_def_TodayCaption = "&Hoje"

'Event Declarations:
Event Change() 'MappingInfo=txtDate,txtDate,-1,Change
Event Click() 'MappingInfo=imgButton,imgButton,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtDate,txtDate,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."

Private Sub imgButton_Click()
    RaiseEvent Click
End Sub

Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = vbLeftButton And Not txtDate.Locked Then
        If frm Is Nothing Then
            imgButton.Picture = imgDown.Picture
            DoEvents
            DatePopUp
            Exit Sub
        Else
            If Not frm.Visible Then DestroyPopUp
            imgButton.Picture = imgDown.Picture
            DoEvents
            DatePopUp
        End If
    End If
End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    If Button = vbLeftButton Then
        imgButton.Picture = imgUp.Picture
    End If
End Sub

Private Sub txtDate_Change()
    RaiseEvent Change
    PropertyChanged "Text"
End Sub

Private Sub txtDate_GotFocus()
    If Not frm Is Nothing Then DestroyPopUp
    If Not txtDate.Locked Then
        txtDate.SelStart = 0
        txtDate.SelLength = Len(txtDate.Text)
    End If
End Sub

Private Sub txtDate_LostFocus()
    If IsDate(txtDate) Then
        txtDate.Text = Format$(txtDate.Text, "Short Date")
    Else
        txtDate.Text = ""
    End If
End Sub
Private Sub DatePopUp()
On Error Resume Next
   Dim lpRect As RECT
   Dim Style As Long
   Set frm = New frmActiveDate
   Call GetWindowRect(UserControl.hWnd, lpRect)
   frm.Top = lpRect.Bottom * Screen.TwipsPerPixelY
   If (frm.Top + frm.Height) > Screen.Height Then
      frm.Top = (lpRect.Top * Screen.TwipsPerPixelY) - frm.Height
   End If
   frm.Left = lpRect.Left * Screen.TwipsPerPixelX
   If (frm.Left + frm.Width) > Screen.Width Then
      frm.Left = (lpRect.Right * Screen.TwipsPerPixelX) - frm.Width
   End If
   Set frm.oComboDate = UserControl.txtDate
   If IsDate(txtDate.Text) Then
      frm.FillCalendar CDate(txtDate.Text)
   Else
      frm.FillCalendar Now
   End If
   'Atualiza Botão Hoje
   frm.btToday.Caption = TodayCaption
   frm.btToday.ToolTipText = Format$(Now, "Long Date")
    
   frm.ZOrder
   frm.Show
   If Err = 401 Then 'Controle em Janela Modal
      Style = GetWindowLong(frm.hWnd, GWL_STYLE)
      SetWindowLong frm.hWnd, GWL_STYLE, (Style And Not WS_POPUP) Or WS_CHILD 'Or WS_CLIPSIBLINGS
      imgButton.Picture = imgUp.Picture
      frm.Show vbModal
    Else
      HookWindow frm
   End If
End Sub

Private Sub UserControl_EnterFocus()
    txtDate.SetFocus
End Sub

Private Sub UserControl_Hide()
    On Error Resume Next
    DestroyPopUp
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    picBack.Top = 0
    picBack.Left = 0
    Set txtDate.Font = Ambient.Font
    txtDate.ToolTipText = Extender.ToolTipText
    txtDate.Tag = m_def_TodayCaption
End Sub

Private Sub UserControl_Resize()
    If Width < 1305 Then
        Width = 1305
        Exit Sub
    End If
    picBack.Width = UserControl.Width
    imgButton.Left = picBack.Width - imgButton.Width - 60
    lblButton.Left = picBack.Width - lblButton.Width - 60
    txtDate.Width = picBack.Width - 405
    If Height <> picBack.Height Then Height = picBack.Height
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtDate.Enabled = New_Enabled
    If New_Enabled Then
        imgButton.Picture = imgUp.Picture
    Else
        imgButton.Picture = imgDis.Picture
    End If
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And (frm Is Nothing) Then
        KeyCode = 0
        DatePopUp
    ElseIf KeyCode = vbKeyF4 Then
        KeyCode = 0
        DestroyPopUp
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If Not (frm Is Nothing) Then
        If frm.Visible Then
            DestroyPopUp
            KeyAscii = 0
            Exit Sub
        Else
            DestroyPopUp
        End If
    End If
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        If Len(txtDate.Text) = 2 Or Len(txtDate.Text) = 5 Then
            txtDate.Text = txtDate.Text + "/"
            txtDate.SelStart = Len(txtDate.Text)
        End If
    ElseIf KeyAscii < 32 Then
        'Go On
    Else
        KeyAscii = 0
        Exit Sub
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtDate_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get TodayCaption() As String
    TodayCaption = txtDate.Tag
End Property

Public Property Let TodayCaption(ByVal New_TodayCaption As String)
    txtDate.Tag = New_TodayCaption
    PropertyChanged "TodayCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "22c"
    Text = txtDate.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    If CanPropertyChange("Text") Then
        If IsDate(New_Text) Or Trim(New_Text) = "" Then
            txtDate.Text() = Format$(New_Text, "Short Date")
            PropertyChanged "Text"
        End If
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtDate.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtDate.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtDate.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtDate.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtDate.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtDate.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    picBack.Top = 0
    picBack.Left = 0
    txtDate.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtDate.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    txtDate.Enabled = UserControl.Enabled
    If Not UserControl.Enabled Then imgButton.Picture = imgDis.Picture
    Set txtDate.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtDate.Text = PropBag.ReadProperty("Text", "")
    txtDate.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtDate.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtDate.SelText = PropBag.ReadProperty("SelText", "")
    txtDate.Tag = PropBag.ReadProperty("TodayCaption", m_def_TodayCaption)
    txtDate.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picBack.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtDate.Locked = PropBag.ReadProperty("Locked", False)
    txtDate.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtDate.FontBold = PropBag.ReadProperty("FontBold", 0)
    txtDate.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    txtDate.FontName = PropBag.ReadProperty("FontName", "")
    txtDate.FontSize = PropBag.ReadProperty("FontSize", 0)
    txtDate.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    txtDate.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    txtDate.ToolTipText = Extender.ToolTipText
    imgButton.ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    DestroyPopUp
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtDate.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtDate.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("Text", txtDate.Text, "")
    Call PropBag.WriteProperty("SelStart", txtDate.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtDate.SelLength, 0)
    Call PropBag.WriteProperty("SelText", txtDate.SelText, "")
    Call PropBag.WriteProperty("Font", txtDate.Font, Ambient.Font)
    Call PropBag.WriteProperty("TodayCaption", txtDate.Tag, m_def_TodayCaption)
    Call PropBag.WriteProperty("BackColor", txtDate.BackColor, &H80000005)
    Call PropBag.WriteProperty("Locked", txtDate.Locked, False)
    Call PropBag.WriteProperty("ForeColor", txtDate.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FontBold", txtDate.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", txtDate.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", txtDate.FontName, "")
    Call PropBag.WriteProperty("FontSize", txtDate.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", txtDate.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", txtDate.FontUnderline, 0)
End Sub
Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub DestroyPopUp()
    UnHookWindow
    Unload frm
    Set frm = Nothing
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtDate.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtDate.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtDate.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtDate.BackColor = New_BackColor
    picBack.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtDate.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtDate.Locked = New_Locked
    imgButton.Enabled = Not New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtDate.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtDate.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = txtDate.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtDate.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = txtDate.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtDate.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_MemberFlags = "400"
    FontName = txtDate.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtDate.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = txtDate.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtDate.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = txtDate.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtDate.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = txtDate.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtDate.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

