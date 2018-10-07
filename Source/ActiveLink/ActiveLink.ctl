VERSION 5.00
Begin VB.UserControl ActiveLink 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   ScaleHeight     =   375
   ScaleWidth      =   3105
   ToolboxBitmap   =   "ActiveLink.ctx":0000
   Begin VB.Label lblLink 
      Caption         =   "http://www.activepack.cjb.net/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   60
      MouseIcon       =   "ActiveLink.ctx":0532
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   60
      Width           =   2895
   End
End
Attribute VB_Name = "ActiveLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_Link = "http://www.activepack.cjb.net/"
'Property Variables:
Dim m_Link As String
'Event Declarations:
Event Click() 'MappingInfo=lblLink,lblLink,-1,Click
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=lblLink,lblLink,-1,DblClick
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblLink,lblLink,-1,MouseDown
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblLink,lblLink,-1,MouseMove
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lblLink,lblLink,-1,MouseUp
Attribute MouseUp.VB_UserMemId = -607

'API Declarations
Private Const SW_SHOW = 5
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Sub AbouBox()
Attribute AbouBox.VB_UserMemId = -552
Attribute AbouBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = lblLink.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblLink.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
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
'MappingInfo=lblLink,lblLink,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblLink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblLink.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub lblLink_Click()
    RaiseEvent Click
    Screen.MousePointer = vbArrowHourglass
    Call ShellExecute(UserControl.hWnd, "open", m_Link, vbNullString, CurDir$, SW_SHOW)
    Screen.MousePointer = vbNormal
End Sub

Private Sub lblLink_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,http://www.activepack.cjb.net/
Public Property Get Link() As String
Attribute Link.VB_Description = "URL Link"
    Link = m_Link
End Property

Public Property Let Link(ByVal New_Link As String)
    m_Link = New_Link
    PropertyChanged "Link"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = lblLink.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    lblLink.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = lblLink.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lblLink.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"
    Caption = lblLink.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLink.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Link = m_def_Link
    lblLink.BackColor = Ambient.BackColor
    Set lblLink.Font = Ambient.Font
    lblLink.FontBold = True
    lblLink.FontUnderline = True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblLink.ForeColor = PropBag.ReadProperty("ForeColor", &HFF0000)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblLink.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Link = PropBag.ReadProperty("Link", m_def_Link)
    lblLink.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lblLink.Caption = PropBag.ReadProperty("Caption", "http://www.activepack.cjb.net/")
    lblLink.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    lblLink.Move 0, 0, Width, Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", lblLink.ForeColor, &HFF0000)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblLink.Font, Ambient.Font)
    Call PropBag.WriteProperty("Link", m_Link, m_def_Link)
    Call PropBag.WriteProperty("MousePointer", lblLink.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Caption", lblLink.Caption, "http://www.activepack.cjb.net/")
    Call PropBag.WriteProperty("BackColor", lblLink.BackColor, &H8000000F)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblLink,lblLink,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = lblLink.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblLink.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

