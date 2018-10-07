VERSION 5.00
Begin VB.UserControl ActivePanel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ControlContainer=   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   2520
   ToolboxBitmap   =   "ActivePanel.ctx":0000
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1620
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ActivePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum PanelBorderConstants
    [None]
    [Fixed Single]
    [Etched]
    [Bump]
    [Raised]
    [Raised Inner]
    [Raised Outer]
    [Sunken]
    [Sunken Inner]
    [Sunken Outer]
End Enum

Public Enum CaptionPosConstants
    [Left Top]
    [Left Center]
    [Left Bottom]
    [Center Top]
    [Center Center]
    [Center Bottom]
    [Right Top]
    [Right Center]
    [Right Bottom]
End Enum


'DrawEdge Constants
Const BDR_RAISEDOUTER = &H1
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const BDR_SUNKENINNER = &H8

Const BDR_OUTER = &H3
Const BDR_INNER = &HC
Const BDR_RAISED = &H5
Const BDR_SUNKEN = &HA

Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const EDGE_BUMP = (BDR_SUNKENINNER Or BDR_RAISEDOUTER)

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Const BF_RECT& = &HF
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

'Default Property Values:
Const m_def_CaptionPos = [Center Center]
Const m_def_BorderStyle = [Raised]

'Property Variables:
Dim m_CaptionPos As CaptionPosConstants
Dim m_BorderStyle As PanelBorderConstants

'Event Declarations:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    DrawBorder m_BorderStyle
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As PanelBorderConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As PanelBorderConstants)
    m_BorderStyle = New_BorderStyle
    DrawBorder m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
On Error Resume Next
    lblCaption.Caption = Extender.Name
    Set lblCaption.Font = Ambient.Font
    UserControl.BackColor = Ambient.BackColor
    m_BorderStyle = m_def_BorderStyle
    DrawBorder m_BorderStyle
    m_CaptionPos = m_def_CaptionPos
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    DrawBorder m_BorderStyle
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Caption")
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.FontBold = PropBag.ReadProperty("FontBold", Ambient.Font.Bold)
    lblCaption.FontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font.Italic)
    lblCaption.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    lblCaption.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
    lblCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
    lblCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font.Underline)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Extender.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    lblCaption.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_CaptionPos = PropBag.ReadProperty("CaptionPos", m_def_CaptionPos)
End Sub

Private Sub UserControl_Resize()
Dim lngBorder As Long
Dim strCaption As String
On Error Resume Next
    If UserControl.Width = 0 Then Exit Sub
    lngBorder = Screen.TwipsPerPixelX * 6
    lblCaption.AutoSize = True
    lblCaption.Left = lngBorder '(UserControl.Width - lblCaption.Width) / 2
    lblCaption.Width = UserControl.Width - (lngBorder * 2)
    lblCaption.AutoSize = False
    Select Case m_CaptionPos
    Case [Left Top]
        lblCaption.Alignment = 0 'Left
        lblCaption.Top = lngBorder
    Case [Left Center]
        lblCaption.Alignment = 0 'Left
        lblCaption.Top = (UserControl.Height - lblCaption.Height) / 2
    Case [Left Bottom]
        lblCaption.Alignment = 0 'Left
        lblCaption.Top = (UserControl.Height - lblCaption.Height) - lngBorder
    Case [Center Top]
        lblCaption.Alignment = 2 'Center
        lblCaption.Top = lngBorder
    Case [Center Center]
        lblCaption.Alignment = 2 'Center
        lblCaption.Top = (UserControl.Height - lblCaption.Height) / 2
    Case [Center Bottom]
        lblCaption.Alignment = 2 'Center
        lblCaption.Top = (UserControl.Height - lblCaption.Height) - lngBorder
    Case [Right Top]
        lblCaption.Alignment = 1 'Right
        lblCaption.Top = lngBorder
    Case [Right Center]
        lblCaption.Alignment = 1 'Right
        lblCaption.Top = (UserControl.Height - lblCaption.Height) / 2
    Case [Right Bottom]
        lblCaption.Alignment = 1 'Right
        lblCaption.Top = (UserControl.Height - lblCaption.Height) - lngBorder
    End Select
    DrawBorder m_BorderStyle
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", lblCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblCaption.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ToolTipText", Extender.ToolTipText, "")
    Call PropBag.WriteProperty("CaptionPos", m_CaptionPos, m_def_CaptionPos)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CaptionPos() As CaptionPosConstants
Attribute CaptionPos.VB_Description = "Returns/sets the Caption position."
Attribute CaptionPos.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionPos = m_CaptionPos
End Property

Public Property Let CaptionPos(ByVal New_CaptionPos As CaptionPosConstants)
    m_CaptionPos = New_CaptionPos
    UserControl_Resize
    PropertyChanged "CaptionPos"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = lblCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = lblCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_MemberFlags = "400"
    FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = lblCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = lblCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Extender.ToolTipText() = New_ToolTipText
    lblCaption.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub DrawBorder(ByVal eStyle As Long)
Dim lpRect As RECT
    UserControl.Cls
    Select Case eStyle
    Case [None]: Exit Sub
    Case [Fixed Single]: eStyle = -1
    Case [Etched]: eStyle = EDGE_ETCHED
    Case [Bump]: eStyle = EDGE_BUMP
    Case [Raised]: eStyle = EDGE_RAISED
    Case [Raised Inner]: eStyle = BDR_RAISEDINNER
    Case [Raised Outer]: eStyle = BDR_RAISEDOUTER
    Case [Sunken]: eStyle = EDGE_SUNKEN
    Case [Sunken Inner]: eStyle = BDR_SUNKENINNER
    Case [Sunken Outer]: eStyle = BDR_SUNKENOUTER
    End Select
    lpRect.Left = 0: lpRect.Top = 0
    lpRect.Bottom = UserControl.Height / Screen.TwipsPerPixelY
    lpRect.Right = UserControl.Width / Screen.TwipsPerPixelX
    If UserControl.HasDC Then
        If eStyle = -1 Then 'Single
            UserControl.Line (0, 0)-(UserControl.ScaleWidth - Screen.TwipsPerPixelX, UserControl.Height - Screen.TwipsPerPixelY), UserControl.ForeColor, B
        Else
            Call DrawEdge(UserControl.hDC, lpRect, eStyle, BF_RECT)
        End If
        UserControl.Refresh
    End If
End Sub

