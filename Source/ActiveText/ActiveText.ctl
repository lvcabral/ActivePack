VERSION 5.00
Begin VB.UserControl ActiveText 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "ActiveText.ctx":0000
   ScaleHeight     =   1035
   ScaleWidth      =   3000
   ToolboxBitmap   =   "ActiveText.ctx":0033
   Begin VB.TextBox txtActive 
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   345
      TabIndex        =   0
      Top             =   330
      Width           =   2220
   End
End
Attribute VB_Name = "ActiveText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'API declaration
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Constants
Enum TextMaskOptions
    [No Mask]
    [Date Mask]
    [Time Mask]
    [Integer Mask]
    [Float Mask]
    [Phone Mask]
    [CEP Mask]
    [CPF Mask]
    [CGC Mask]
    [Custom Mask]
End Enum
Enum TextCaseOptions
    Normal
    UpperCase
    LowerCase
    ProperCase
End Enum
Enum DateFormatOptions
    [dd/mm/yyyy]
    [dd/mm/yy]
    [mm/dd/yyyy]
    [mm/dd/yy]
End Enum
Enum TimeFormatOptions
    [hh:mm:ss]
    [hh:mm]
End Enum
Enum FloatFormatOptions
    [Decimal Point]
    [Windows Default]
    [Currency Format]
    [Percent Format]
End Enum
Enum AlignmentOptions
    [Left Justify]
    [Right Justify]
    [Center]
End Enum
Enum AppearanceOptions
    Flat
    [3D]
End Enum
Enum BorderStyleOptions
    None
    [Fixed Single]
End Enum
'Internal Variables
Private bolShow As Boolean

'Default Property Values:
Const m_def_About = 0
Const m_def_DateFormat = 0
Const m_def_TimeFormat = 0
Const m_def_FloatFormat = 0
Const m_def_MaxLength = 0
Const m_def_Decimals = 2
Const m_def_TextMask = 0
Const m_def_TextCase = 0
Const m_def_FocusSelect = True
Dim m_def_DecimalPoint As String

'Property Variables:
Dim m_About As Variant
Dim m_DateFormat As DateFormatOptions
Dim m_TimeFormat As TimeFormatOptions
Dim m_FloatFormat As FloatFormatOptions
Dim m_DecimalPoint As String
Dim m_Decimals As Byte
Dim m_MaxLength As Long
Dim m_TextMask As TextMaskOptions
Dim m_TextCase As TextCaseOptions
Dim m_FocusSelect As Boolean
Dim m_DataField As String
Dim m_RawText As String
Dim m_Mask As String

'Event Declarations:
Event Change() 'MappingInfo=txtActive,txtActive,-1,Change
Event Click() 'MappingInfo=txtActive,txtActive,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=txtActive,txtActive,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtActive,txtActive,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtActive,txtActive,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtActive,txtActive,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtActive,txtActive,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtActive,txtActive,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtActive,txtActive,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Public Property Get PasswordChar() As String
    PasswordChar = txtActive.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    If Len(New_PasswordChar) = 0 Then
        txtActive.PasswordChar = ""
    ElseIf Len(New_PasswordChar) > 1 Then
        txtActive.PasswordChar = Left(New_PasswordChar, 1)
    Else
        txtActive.PasswordChar = New_PasswordChar
    End If
    PropertyChanged "PasswordChar"
End Property

Public Property Get Alignment() As AlignmentOptions
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = txtActive.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentOptions)
    txtActive.Alignment = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get Appearance() As AppearanceOptions
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceOptions)
    UserControl.Appearance = New_Appearance
    UserControl.BackColor = txtActive.BackColor
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = txtActive.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtActive.BackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleOptions
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleOptions)
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Let DataField(ByVal New_DataField As String)
    m_DataField = New_DataField
End Property

Public Property Get DataField() As String
    DataField = m_DataField
End Property

Public Property Let FocusSelect(ByVal New_FocusSelect As Boolean)
    If Ambient.UserMode Then Err.Raise 393
    m_FocusSelect = New_FocusSelect
    PropertyChanged "FocusSelect"
End Property

Public Property Get FocusSelect() As Boolean
Attribute FocusSelect.VB_ProcData.VB_Invoke_Property = "ControlStyles;Behavior"
    FocusSelect = m_FocusSelect
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = txtActive.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtActive.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ControlStyles;Behavior"
    Enabled = txtActive.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtActive.Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = txtActive.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtActive.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Let RawText(ByVal New_RawText As String)
Dim i%, s$
If CanPropertyChange("RawText") Then
    If New_RawText = "" Then
        txtActive.Text = New_RawText
    ElseIf m_TextMask = [CEP Mask] Then
        New_RawText = Format$(New_RawText, "00000000")
        txtActive.Text = Format$(New_RawText, "@@@@@-@@@")
    ElseIf m_TextMask = [CPF Mask] Then
        New_RawText = Format$(New_RawText, "00000000000")
        txtActive.Text = Format$(New_RawText, "@@@.@@@.@@@-@@")
    ElseIf m_TextMask = [CGC Mask] Then
        New_RawText = Format$(New_RawText, "00000000000000")
        txtActive.Text = Format$(New_RawText, "@@.@@@.@@@/@@@@-@@")
    ElseIf m_TextMask = [Custom Mask] Or _
           m_TextMask = [Date Mask] Or _
           m_TextMask = [Time Mask] Or _
           m_TextMask = [Phone Mask] Then
        For i% = 1 To Len(m_Mask)
            If InStr("&#?", Mid$(m_Mask, i%, 1)) = 0 Then
                s$ = s$ & Mid$(m_Mask, i%, 1)
            ElseIf IsValidChar(Mid$(m_Mask, i%, 1), Asc(Left(New_RawText, 1))) Then
                s$ = s$ & Left(New_RawText, 1)
                If Len(New_RawText) > 1 Then
                    New_RawText = Mid$(New_RawText, 2)
                Else
                    Exit For
                End If
                
            Else
                If Len(New_RawText) > 1 Then
                    New_RawText = Mid$(New_RawText, 2)
                Else
                    Exit For
                End If
            End If
        Next
        txtActive.Text = s$
    Else
        txtActive.Text = New_RawText
    End If
    m_RawText = New_RawText
    PropertyChanged "RawText"
End If
End Property

Public Property Get RawText() As String
Attribute RawText.VB_MemberFlags = "1c"
Dim i%, s$
    If m_TextMask = [Custom Mask] Or _
       m_TextMask = [Date Mask] Or _
       m_TextMask = [Time Mask] Or _
       m_TextMask = [Phone Mask] Then
        For i% = 1 To Len(txtActive)
            If InStr("&#?", Mid$(m_Mask, i%, 1)) = 0 Then
                'Ignore Literals
            ElseIf IsValidChar(Mid$(m_Mask, i%, 1), Asc(Mid$(txtActive, i%, 1))) Then
                s$ = s$ & Mid$(txtActive, i%, 1)
            End If
        Next
        RawText = s$
    ElseIf m_TextMask >= [Integer Mask] Then
        If m_TextMask = [Float Mask] Then
            On Error Resume Next
            s$ = CStr(CDbl(txtActive))
            If Err = 0 Then
                RawText = s$
                Exit Property
            End If
        End If
        s$ = ""
        For i% = 1 To Len(txtActive)
            If IsNumeric(Mid$(txtActive, i%, 1)) Or _
               (Mid$(txtActive, i%, 1) = m_DecimalPoint And _
                m_TextMask = [Float Mask]) Then
                s$ = s$ + Mid$(txtActive, i%, 1)
            ElseIf i% = 1 And Left$(txtActive, 1) = "-" Then
                s$ = s$ + Mid$(txtActive, i%, 1)
            End If
        Next
        RawText = s$
    Else
        RawText = txtActive.Text
    End If
End Property

Private Sub txtActive_Change()
    RaiseEvent Change
    If bolShow Then RaiseChange Extender
    PropertyChanged "RawText"
    PropertyChanged "Text"
End Sub

Private Sub txtActive_Click()
    RaiseEvent Click
    RaiseClick Extender
End Sub

Private Sub txtActive_DblClick()
    RaiseDblClick Extender
    RaiseEvent DblClick
End Sub

Private Sub txtActive_GotFocus()
    If m_TextMask = [Float Mask] Then
        txtActive.Text = Me.RawText
    End If
    If m_FocusSelect Then
        SelectEditBox txtActive
    End If
    RaiseGotFocus Extender
End Sub

Private Sub txtActive_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseKeyDown Extender, KeyCode, Shift
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtActive_KeyPress(KeyAscii As Integer)
Dim mBuffer$, mPos%
    If txtActive.Locked Then Exit Sub
    If m_TextMask = [Float Mask] Then
        KeyAscii = FloatTextBox(m_MaxLength, m_Decimals, KeyAscii)
    ElseIf m_TextMask = [Integer Mask] Then
        KeyAscii = IntegerTextBox(KeyAscii)
    ElseIf m_TextMask = [Date Mask] Or m_TextMask = [Time Mask] Then
        KeyAscii = CustomTextBox(KeyAscii)
    ElseIf m_TextMask >= [Phone Mask] Then
        If m_TextCase = UpperCase Then
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
        ElseIf m_TextCase = LowerCase Then
            KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
        End If
        KeyAscii = CustomTextBox(KeyAscii)
    ElseIf KeyAscii >= 65 Then
        If m_TextCase = UpperCase Then
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
        ElseIf m_TextCase = LowerCase Then
            KeyAscii = Asc(LCase$(Chr$(KeyAscii)))
        ElseIf m_TextCase = ProperCase Then
            If txtActive.SelStart = Len(txtActive.Text) Then
                mBuffer$ = ProperCaseText(txtActive.Text + Chr$(KeyAscii))
                txtActive.Text = Left$(mBuffer$, Len(mBuffer$) - 1)
                txtActive.SelStart = Len(txtActive.Text)
                KeyAscii = Asc(Right$(mBuffer$, 1))
            ElseIf txtActive.SelLength = 0 Then
                mPos% = txtActive.SelStart
                txtActive.Text = ProperCaseText(txtActive.Text)
                txtActive.SelStart = mPos%
            End If
        End If
    End If
    RaiseKeyPress Extender, KeyAscii
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtActive_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseKeyUp Extender, KeyCode, Shift
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtActive_LostFocus()
    If m_TextMask = [No Mask] And m_TextCase = ProperCase Then
        txtActive.Text = ProperCaseText(txtActive.Text)
    ElseIf m_TextMask = [Float Mask] Then
        Me.Text = txtActive.Text
    ElseIf m_TextMask = [Date Mask] Then
        UpdateDate
    ElseIf m_TextMask = [Integer Mask] Then
        Me.Text = Format(Val(Me.Text))
    ElseIf m_TextMask >= [Phone Mask] Then
        RawText = RawText
    End If
    RaiseLostFocus Extender
End Sub

Private Sub txtActive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseMouseDown Extender, Button, Shift, X, Y
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtActive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseMouseMove Extender, Button, Shift, X, Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtActive_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseMouseUp Extender, Button, Shift, X, Y
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = txtActive.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = "ControlStyles"
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Select Case m_TextMask
    Case [Date Mask], [Time Mask], [Phone Mask], [CEP Mask], [CPF Mask], [CGC Mask], [Custom Mask]
        Exit Property
    Case [Float Mask]
        If New_MaxLength < 4 Then New_MaxLength = 4
        m_MaxLength = New_MaxLength
        txtActive.MaxLength = 0
    Case Else
        m_MaxLength = New_MaxLength
        txtActive.MaxLength = New_MaxLength
    End Select
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_MemberFlags = "400"
    SelLength = txtActive.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtActive.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_MemberFlags = "400"
    SelStart = txtActive.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtActive.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_MemberFlags = "400"
    SelText = txtActive.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtActive.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "3c"
    Text = txtActive.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    If CanPropertyChange("Text") Then
        If m_TextMask = [Integer Mask] Then
            New_Text = Format(Val(New_Text))
        ElseIf m_TextMask = [Float Mask] Then
            New_Text = Replace(New_Text, m_def_DecimalPoint, ".")
            New_Text = Replace(New_Text, m_DecimalPoint, ".")
            Select Case m_FloatFormat
            Case [Decimal Point]
                New_Text = Format$(Val(New_Text), String(m_MaxLength - m_Decimals - 2, "#") & _
                                 "0." & String(m_Decimals, "00"))
                New_Text = Replace(New_Text, m_def_DecimalPoint, m_DecimalPoint)
            Case [Windows Default]
                New_Text = Format$(Val(New_Text), "Standard")
            Case [Currency Format]
                New_Text = Format$(Val(New_Text), "Currency")
            Case [Percent Format]
                New_Text = Format$(Val(New_Text), String(m_MaxLength - m_Decimals - 2, "#") & _
                                 "0." & String(m_Decimals, "00"))
                New_Text = Replace(New_Text, m_def_DecimalPoint, m_DecimalPoint) & "%"
            End Select
        End If
        txtActive.Text = New_Text
        PropertyChanged "Text"
    End If
End Property

Public Property Get TextMask() As TextMaskOptions
Attribute TextMask.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TextMask = m_TextMask
End Property

Public Property Let TextMask(ByVal New_TextMask As TextMaskOptions)
    'If Ambient.UserMode Then Err.Raise 393
    m_TextMask = New_TextMask
    m_Mask = ""
    txtActive.Alignment = vbLeftJustify
    If m_TextMask > [No Mask] Then
        m_TextCase = Normal
        txtActive.Text = ""
    End If
    If m_TextMask = [Integer Mask] Then
        Me.Text = Me.RawText
        txtActive.Alignment = vbRightJustify
    ElseIf m_TextMask = [Float Mask] Then
        m_MaxLength = 13
        Me.Text = Me.RawText
        txtActive.Alignment = vbRightJustify
    ElseIf m_TextMask = [Date Mask] Then
        Select Case m_DateFormat
        Case [dd/mm/yyyy], [mm/dd/yyyy]
            m_Mask = "##/##/####"
            m_MaxLength = 10
        Case [dd/mm/yy], [mm/dd/yy]
            m_Mask = "##/##/##"
            m_MaxLength = 8
        End Select
    ElseIf m_TextMask = [Time Mask] Then
        If m_TimeFormat = [hh:mm:ss] Then
            m_Mask = "##:##:##"
            m_MaxLength = 8
        Else
            m_Mask = "##:##"
            m_MaxLength = 5
        End If
    ElseIf m_TextMask = [Phone Mask] Then
        m_Mask = "(###)####-####"
        m_MaxLength = 14
    ElseIf m_TextMask = [CEP Mask] Then
        m_Mask = "#####-###"
        m_MaxLength = 9
    ElseIf m_TextMask = [CPF Mask] Then
        m_Mask = "###.###.###-##"
        m_MaxLength = 14
    ElseIf m_TextMask = [CGC Mask] Then
        m_Mask = "##.###.###/####-##"
        m_MaxLength = 18
    ElseIf m_TextMask = [Custom Mask] Then
        m_MaxLength = Len(m_Mask)
    End If
    If m_TextMask <> [Float Mask] Then
        txtActive.MaxLength = 0
        txtActive.MaxLength = m_MaxLength
    End If
    PropertyChanged "MaxLength"
    PropertyChanged "TextMask"
End Property

Public Property Get TextCase() As TextCaseOptions
Attribute TextCase.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TextCase = m_TextCase
End Property

Public Property Let TextCase(ByVal New_TextCase As TextCaseOptions)
    'If Ambient.UserMode Then Err.Raise 393
    If m_TextMask = [No Mask] Then
        m_TextCase = New_TextCase
    ElseIf m_TextMask = [Custom Mask] Then
        If New_TextCase = ProperCase Then New_TextCase = UpperCase
        m_TextCase = New_TextCase
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error Resume Next
    Set Font = Ambient.Font
    m_TextMask = m_def_TextMask
    m_TextCase = m_def_TextCase
    m_FocusSelect = m_def_FocusSelect
    m_DataField = ""
    txtActive.Text = Extender.Name
    m_MaxLength = m_def_MaxLength
    m_DateFormat = m_def_DateFormat
    m_TimeFormat = m_def_TimeFormat
    m_FloatFormat = m_def_FloatFormat
    m_Mask = ""
    m_def_DecimalPoint = Mid(Format(12.34, "00.00"), 3, 1)
    m_DecimalPoint = m_def_DecimalPoint
    m_Decimals = m_def_Decimals
    m_About = m_def_About
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_def_DecimalPoint = Mid(Format(12.34, "00.00"), 3, 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", [Fixed Single])
    UserControl.Appearance = PropBag.ReadProperty("Appearance", [3D])
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtActive.Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    txtActive.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtActive.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtActive.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.Enabled = txtActive.Enabled
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtActive.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtActive.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtActive.SelText = PropBag.ReadProperty("SelText", "")
    txtActive.Text = PropBag.ReadProperty("Text", "")
    txtActive.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    m_MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
    m_TextMask = PropBag.ReadProperty("TextMask", m_def_TextMask)
    m_TextCase = PropBag.ReadProperty("TextCase", m_def_TextCase)
    m_FocusSelect = PropBag.ReadProperty("FocusSelect", m_def_FocusSelect)
    m_DataField = PropBag.ReadProperty("DataField", "")
    m_RawText = PropBag.ReadProperty("RawText", "")
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_TimeFormat = PropBag.ReadProperty("TimeFormat", m_def_TimeFormat)
    m_FloatFormat = PropBag.ReadProperty("FloatFormat", m_def_FloatFormat)
    m_Mask = PropBag.ReadProperty("Mask", "")
    m_DecimalPoint = PropBag.ReadProperty("DecimalPoint", m_def_DecimalPoint)
    m_Decimals = PropBag.ReadProperty("Decimals", m_def_Decimals)
    m_About = PropBag.ReadProperty("About", m_def_About)
    If m_TextMask <> [Float Mask] Then txtActive.MaxLength = m_MaxLength
    txtActive.FontBold = PropBag.ReadProperty("FontBold", False)
    txtActive.FontItalic = PropBag.ReadProperty("FontItalic", False)
    txtActive.FontName = PropBag.ReadProperty("FontName", txtActive.FontName)
    txtActive.FontSize = PropBag.ReadProperty("FontSize", txtActive.FontSize)
    txtActive.FontStrikethru = PropBag.ReadProperty("FontStrikethru", False)
    txtActive.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    txtActive.Locked = PropBag.ReadProperty("Locked", False)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim xy As Long
    If UserControl.BorderStyle = 0 Then 'No Border
        xy = 30
    ElseIf UserControl.Appearance = 0 Then 'Flat
        xy = 60
    Else
        xy = 90
    End If
    txtActive.Move 15, 15, UserControl.Width - xy, UserControl.Height - xy
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    txtActive.ToolTipText = Extender.ToolTipText
    bolShow = True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", txtActive.Alignment, vbLeftJustify)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, [3D])
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, [Fixed Single])
    Call PropBag.WriteProperty("BackColor", txtActive.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtActive.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtActive.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
    Call PropBag.WriteProperty("SelLength", txtActive.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtActive.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtActive.SelText, "")
    Call PropBag.WriteProperty("Text", txtActive.Text, "")
    Call PropBag.WriteProperty("TextMask", m_TextMask, m_def_TextMask)
    Call PropBag.WriteProperty("TextCase", m_TextCase, m_def_TextCase)
    Call PropBag.WriteProperty("PasswordChar", txtActive.PasswordChar, "")
    Call PropBag.WriteProperty("FocusSelect", m_FocusSelect, m_def_FocusSelect)
    Call PropBag.WriteProperty("DataField", m_DataField, "")
    Call PropBag.WriteProperty("RawText", m_TextMask, "")

    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    Call PropBag.WriteProperty("TimeFormat", m_TimeFormat, m_def_TimeFormat)
    Call PropBag.WriteProperty("FloatFormat", m_FloatFormat, m_def_FloatFormat)
    Call PropBag.WriteProperty("Mask", m_Mask, "")
    Call PropBag.WriteProperty("DecimalPoint", m_DecimalPoint, m_def_DecimalPoint)
    Call PropBag.WriteProperty("Decimals", m_Decimals, m_def_Decimals)
    Call PropBag.WriteProperty("About", m_About, m_def_About)
    Call PropBag.WriteProperty("FontBold", txtActive.FontBold, False)
    Call PropBag.WriteProperty("FontItalic", txtActive.FontItalic, False)
    Call PropBag.WriteProperty("FontName", txtActive.FontName, "Tahoma")
    Call PropBag.WriteProperty("FontSize", txtActive.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", txtActive.FontStrikethru, False)
    Call PropBag.WriteProperty("FontUnderline", txtActive.FontUnderline, False)
    Call PropBag.WriteProperty("Locked", txtActive.Locked, False)
End Sub

Public Property Get DateFormat() As DateFormatOptions
Attribute DateFormat.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As DateFormatOptions)
    m_DateFormat = New_DateFormat
    If m_TextMask = [Date Mask] Then
        Select Case m_DateFormat
        Case [dd/mm/yyyy], [mm/dd/yyyy]
            m_Mask = "##/##/####"
            m_MaxLength = 10
        Case [dd/mm/yy], [mm/dd/yy]
            m_Mask = "##/##/##"
            m_MaxLength = 8
        End Select
        txtActive.MaxLength = m_MaxLength
        UpdateDate
    End If
    PropertyChanged "DateFormat"
End Property

Public Property Get FloatFormat() As FloatFormatOptions
    FloatFormat = m_FloatFormat
End Property

Public Property Let FloatFormat(ByVal New_FloatFormat As FloatFormatOptions)
    m_FloatFormat = New_FloatFormat
    If m_TextMask = [Float Mask] Then
        Me.Text = Me.RawText
    End If
    PropertyChanged "FloatFormat"
End Property

Public Property Get TimeFormat() As TimeFormatOptions
Attribute TimeFormat.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TimeFormat = m_TimeFormat
End Property

Public Property Let TimeFormat(ByVal New_TimeFormat As TimeFormatOptions)
    m_TimeFormat = New_TimeFormat
    If m_TextMask = [Time Mask] Then
        If m_TimeFormat = [hh:mm:ss] Then
            m_Mask = "##:##:##"
            m_MaxLength = 8
        Else
            m_Mask = "##:##"
            m_MaxLength = 5
            txtActive.Text = Left$(txtActive.Text, 5)
        End If
        txtActive.MaxLength = m_MaxLength
    End If
    PropertyChanged "TimeFormat"
End Property

Public Property Get Mask() As String
Attribute Mask.VB_Description = "# - Numbers (0-9)\r\n? - Letters (A-z)\r\n& - Any Character"
Attribute Mask.VB_ProcData.VB_Invoke_Property = "ControlStyles;Behavior"
    Mask = m_Mask
End Property

Public Property Let Mask(ByVal New_Mask As String)
    m_Mask = New_Mask
    If Trim$(m_Mask) = "" Then
        m_TextMask = [No Mask]
    Else
        m_TextMask = [Custom Mask]
        m_MaxLength = Len(Trim(m_Mask))
        txtActive.MaxLength = m_MaxLength
        txtActive.Text = ""
    End If
    PropertyChanged "Mask"
End Property

Public Property Get Decimals() As String
Attribute Decimals.VB_ProcData.VB_Invoke_Property = "ControlStyles;Behavior"
    Decimals = m_Decimals
End Property

Public Property Let Decimals(ByVal New_Decimals As String)
    If New_Decimals < 1 Then
        MsgBox "Decimals must be at least 1 !", vbInformation, "ActiveText"
        Exit Property
    ElseIf New_Decimals > 8 Then
        MsgBox "Decimals can't be greater than 8 !", vbInformation, "ActiveText"
        Exit Property
    End If
    m_Decimals = New_Decimals
    If m_TextMask = [Float Mask] Then
        Me.Text = Me.RawText
    End If
    PropertyChanged "Decimals"
End Property

Public Property Get DecimalPoint() As String
Attribute DecimalPoint.VB_ProcData.VB_Invoke_Property = "ControlStyles;Behavior"
    DecimalPoint = m_DecimalPoint
End Property

Public Property Let DecimalPoint(ByVal New_DecimalPoint As String)
    If Len(New_DecimalPoint) = 0 Or Len(New_DecimalPoint) > 1 Then
        MsgBox "Decimal Point must have exactly 1 character!", vbInformation, "ActiveText"
        Exit Property
    End If
    If m_TextMask = [Float Mask] Then
        txtActive.Text = Replace(txtActive.Text, m_DecimalPoint, New_DecimalPoint)
    End If
    m_DecimalPoint = New_DecimalPoint
    If m_TextMask = [Float Mask] Then
        Me.Text = Me.RawText
    End If
    PropertyChanged "DecimalPoint"
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = txtActive.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtActive.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = txtActive.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtActive.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontName.VB_MemberFlags = "400"
    FontName = txtActive.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    txtActive.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = txtActive.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    txtActive.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = txtActive.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    txtActive.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = txtActive.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtActive.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtActive,txtActive,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Attribute Locked.VB_ProcData.VB_Invoke_Property = "ControlStyles"
    Locked = txtActive.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtActive.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Private Function FloatTextBox(ByVal IntLen As Integer, ByVal DecLen As Integer, ByVal KeyAscii As Integer) As Integer
Dim mInt As Integer, mVrg As Integer
    mInt = IntLen - (DecLen + 1)
    mVrg = InStr(txtActive.Text, m_DecimalPoint)
    If KeyAscii < 32 Then
        FloatTextBox = KeyAscii
        Exit Function
    ElseIf (Len(txtActive) = 0 Or txtActive.SelLength = Len(txtActive)) And _
            Chr(KeyAscii) = "-" Then
        FloatTextBox = KeyAscii
        Exit Function
    End If
    If txtActive.SelLength > 0 And (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
       'Nada
    ElseIf Len(txtActive.Text) >= IntLen Then
        KeyAscii = 0
    End If
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        If mVrg = 0 Then
            If Len(txtActive.Text) = mInt Then
                SendKeys Chr$(KeyAscii)
                KeyAscii = Asc(m_DecimalPoint)
            End If
        Else
            If txtActive.SelStart <= mVrg And (mVrg - 1) < mInt Then
                'Nada
            ElseIf txtActive.SelLength > 0 Then
                'Nada
            ElseIf Len(txtActive.Text) - mVrg >= DecLen Then
                KeyAscii = 0
            End If
        End If
    Else
        If KeyAscii = Asc(m_def_DecimalPoint) Or _
           KeyAscii = Asc(m_DecimalPoint) Or _
           KeyAscii = Asc(".") Then
            If InStr(txtActive.Text, m_DecimalPoint) = 0 Then
                KeyAscii = Asc(m_DecimalPoint)
            Else
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
    End If
    FloatTextBox = KeyAscii
End Function

Private Function IntegerTextBox(ByVal KeyAscii As Integer) As Integer
    If KeyAscii < 32 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        'Go On
    ElseIf (Len(txtActive) = 0 Or txtActive.SelLength = Len(txtActive)) And _
           Chr(KeyAscii) = "-" Then
        'Negative Integer
    Else
        KeyAscii = 0
    End If
    IntegerTextBox = KeyAscii
End Function

Private Function CustomTextBox(ByVal KeyAscii As Integer) As Integer
Dim Pos As Long, p As Long
Dim Txt As String
    If KeyAscii < 32 Or txtActive.SelStart = txtActive.MaxLength Then
        'NonVisibleKey or Full, Go On
        CustomTextBox = KeyAscii
        Exit Function
    ElseIf txtActive.SelLength > 0 And _
           IsValidChar("&", KeyAscii) Then
        'Selected
        txtActive.Text = ""
        txtActive.SelStart = 1
    End If
    If IsValidChar("&", KeyAscii) Then
        Pos = txtActive.SelStart + 1
        Txt = txtActive.Text
        For p = Pos To txtActive.MaxLength
            If IsValidChar(Mid$(m_Mask, p, 1), KeyAscii) Then
                Exit For
            ElseIf InStr("&#?", Mid$(m_Mask, p, 1)) > 0 Then
                KeyAscii = 0
                Exit For
            Else 'Literal
                Txt = Left$(Txt, p - 1) & Mid$(m_Mask, p, 1) & Mid$(Txt, p)
            End If
        Next
        txtActive.Text = Txt
        txtActive.SelStart = p
    Else
        KeyAscii = 0
    End If
    CustomTextBox = KeyAscii
End Function

Private Function IsValidChar(ByVal strMask As String, ByVal KeyAscii As Integer) As Boolean
    Select Case strMask
    Case "&"
        IsValidChar = (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Or _
                      (KeyAscii >= 64 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or _
                      (KeyAscii >= 192)
    Case "#"
        IsValidChar = (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57)
    Case "?"
        IsValidChar = (KeyAscii = 32) Or (KeyAscii >= 64 And KeyAscii <= 90) Or _
                      (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 192)
    End Select
End Function

Private Function ProperCaseText(cString As String) As String
Dim nPos As Integer
Dim nPos2 As Integer
Dim cWord As String
Dim cProper As String
Dim cExcept As String
Dim nFirst As Boolean
    nFirst = True
    cProper = ""
    cExcept = "de da das do dos em na nas no nos e é ao à "
    cString = LCase(cString)
    Do Until Trim$(cString) = ""
        nPos = InStr(cString, " ")
        If nPos <> 0 Then
            cWord = Mid$(cString, 1, nPos)
            cString = Mid$(cString, nPos + 1)
        Else
            cWord = cString
            cString = ""
        End If
        If InStr(cWord, "'") <> 0 Then
            nPos2 = InStr(cWord, "'")
        ElseIf InStr(cWord, "-") <> 0 Then
            nPos2 = InStr(cWord, "-")
        ElseIf InStr(cWord, ".") <> 0 Then
            nPos2 = InStr(cWord, ".")
        Else
            nPos2 = 0
        End If
        If InStr(cExcept, cWord) <> 0 And Not nFirst Then
            cProper = cProper + cWord
        ElseIf nPos2 <> 0 Then
            cProper = cProper + UCase$(Mid$(cWord, 1, 1)) + Mid$(cWord, 2, nPos2 - 1) + UCase$(Mid$(cWord, nPos2 + 1, 1)) + Mid$(cWord, nPos2 + 2)
        Else
            cProper = cProper + UCase$(Mid$(cWord, 1, 1)) + Mid$(cWord, 2)
        End If
        nFirst = False
    Loop
    ProperCaseText = cProper
End Function

Private Sub UpdateDate()
    If m_TextMask = [Date Mask] And IsDate(txtActive.Text) Then
        Select Case m_DateFormat
        Case [dd/mm/yyyy]
            txtActive.Text = Format$(txtActive.Text, "dd/mm/yyyy")
        Case [dd/mm/yy]
            txtActive.Text = Format$(txtActive.Text, "dd/mm/yy")
        Case [mm/dd/yyyy]
            txtActive.Text = Format$(txtActive.Text, "mm/dd/yyyy")
        Case [mm/dd/yy]
            txtActive.Text = Format$(txtActive.Text, "mm/dd/yy")
        End Select
    End If
End Sub

Private Sub SelectEditBox(ObjTxt As Object)
Const EM_SETSEL = &HB1
    SendMessage ObjTxt.hWnd, EM_SETSEL, 0&, -1&
End Sub


'Private Function PhoneTextBox(ByVal KeyAscii As Integer) As Integer
'    If KeyAscii = 32 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
'        Select Case Len(txtActive.Text)
'        Case 0
'            txtActive.Text = "("
'            txtActive.SelStart = 2
'        Case 4
'            txtActive.Text = txtActive.Text + ")"
'            txtActive.SelStart = Len(txtActive.Text)
'        Case 9
'            txtActive.Text = txtActive.Text + "-"
'            txtActive.SelStart = Len(txtActive.Text)
'        End Select
'        If txtActive.SelLength = Len(txtActive.Text) Then
'            txtActive.Text = "("
'            txtActive.SelStart = 2
'        End If
'    ElseIf KeyAscii < 32 Then
'        'Go On
'    Else
'        KeyAscii = 0
'    End If
'    PhoneTextBox = KeyAscii
'End Function

