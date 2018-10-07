VERSION 5.00
Begin VB.UserControl ActiveCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   ScaleHeight     =   555
   ScaleWidth      =   2520
   ToolboxBitmap   =   "ActiveCombo.ctx":0000
End
Attribute VB_Name = "ActiveCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ComboEx As ComboBoxEx
Dim tpx As Long, tpy As Long, Hooked As Boolean
Dim EditHwnd As Long
Enum IconSizeConstants
    [16x16] = 16
    [32x32] = 32
End Enum
Enum CboStyleConstants
    [Dropdown Combo] = 0
    [Dropdown List] = 2
End Enum
Enum ColorDepthConstants
    [Default]
    [16 Colors]
    [256 Colors]
    [HighColor]
    [TrueColor]
End Enum
'Event Declarations:
Event Change()
Event Click()
Event DropDown()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)

'Property Variables:
Dim m_SelStart As Long
Dim m_SelEnd As Long
Dim m_Locked As Boolean
Dim m_ImageList As Object
Dim m_Style As CboStyleConstants
Dim m_ShowIcons As Boolean
Dim m_IconSize As Integer
Dim m_ItemData As Collection
Dim m_ColorDepth As ColorDepthConstants

'Default Property Values:
Const m_def_Locked = False
Const m_def_ShowIcons = False

'API Constants
Const EM_SETREADONLY As Long = &HCF
Const CB_FINDSTRING = &H14C
Const CB_ERR = (-1)
Const EM_GETSEL = &HB0
Const EM_SETSEL = &HB1
Const EM_REPLACESEL = &HC2
Const ILC_COLOR = &H0
Const ILC_COLOR4 = &H4
Const ILC_COLOR8 = &H8
Const ILC_COLOR16 = &H10
Const ILC_COLOR24 = &H18

'API Functions
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal lEnable As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Sub UserControl_EnterFocus()
    ComboEx.SetFocus
End Sub

Private Sub UserControl_Initialize()
    Set ComboEx = New ComboBoxEx
    tpx = Screen.TwipsPerPixelX
    tpy = Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_InitProperties()
    m_IconSize = 16
    m_Style = [Dropdown Combo]
    Set Font = Ambient.Font
    With Ambient.Font
        ComboEx.FontName = .Name
        ComboEx.FontHeight = UserControl.TextHeight("M") / tpy
        ComboEx.FontBold = .Bold
        ComboEx.FontItalic = .Italic
        ComboEx.FontUnderlined = .Underline
    End With
    CreateCombo
    ComboEx.SetEditString Extender.Name
    m_ShowIcons = m_def_ShowIcons
    Set m_ItemData = New Collection
    m_Locked = m_def_Locked
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_IconSize = PropBag.ReadProperty("IconSize", 16)
    m_ColorDepth = PropBag.ReadProperty("IconColorDepth", 0)
    m_ShowIcons = PropBag.ReadProperty("ShowIcons", m_def_ShowIcons)
    m_Style = PropBag.ReadProperty("Style", 0)
    m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
    With Font
        ComboEx.FontName = .Name
        ComboEx.FontHeight = UserControl.TextHeight("M") / tpy
        ComboEx.FontBold = .Bold
        ComboEx.FontItalic = .Italic
        ComboEx.FontUnderlined = .Underline
    End With
    If m_ColorDepth = [16 Colors] Then
        ComboEx.ColorDepth = ILC_COLOR4
    ElseIf m_ColorDepth = [256 Colors] Then
        ComboEx.ColorDepth = ILC_COLOR8
    ElseIf m_ColorDepth = HighColor Then
        ComboEx.ColorDepth = ILC_COLOR16
    ElseIf m_ColorDepth = TrueColor Then
        ComboEx.ColorDepth = ILC_COLOR24
    Else
        ComboEx.ColorDepth = ILC_COLOR
    End If
    CreateCombo
    If m_Style = [Dropdown Combo] Then
        ComboEx.SetEditString PropBag.ReadProperty("Text", Extender.Name)
        Call SendMessageLong(ComboEx.GetEditHwnd, EM_SETREADONLY, m_Locked, 0&)
    End If
    Extender.Height = (ComboEx.GetItemHeight(-1) + 6) * tpy
    Set m_ItemData = New Collection
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Call EnableWindow(ComboEx.GetComboHwnd, UserControl.Enabled)
End Sub

Private Sub UserControl_Resize()
    If ComboEx.GetComboHwnd > 0 Then
        With Extender
            ComboEx.ResizeCombo 0, 0, .Width / tpx
            .Height = (ComboEx.GetItemHeight(-1) + 6) * tpy
        End With
    End If
End Sub

Private Sub UserControl_Show()
    ComboEx.Refresh
End Sub

Private Sub UserControl_Terminate()
    If Hooked Then
        UnHookWindow ComboEx.GetComboHwnd
        UnHookWindow GetParent(ComboEx.GetComboHwnd)
        UnHookWindow EditHwnd
    End If
    ComboEx.Destroy
    Set ComboEx = Nothing
    Set m_ItemData = Nothing
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    EnableWindow ComboEx.GetComboHwnd, Not Enabled
    EnableWindow ComboEx.GetComboHwnd, Enabled
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call EnableWindow(ComboEx.GetComboHwnd, New_Enabled)
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Dim hImageList As Long
    If Not m_ImageList Is Nothing Then hImageList = m_ImageList.hImageList
    Set UserControl.Font = New_Font
    With UserControl.Font
        ComboEx.FontName = .Name
        ComboEx.FontHeight = UserControl.TextHeight("M") / tpy
        ComboEx.FontBold = .Bold
        ComboEx.FontItalic = .Italic
        ComboEx.FontUnderlined = .Underline
    End With
    ComboEx.SetComboFont
    ComboEx.ShowIcons m_ShowIcons, hImageList
    Extender.Height = (ComboEx.GetItemHeight(-1) + 6) * tpy
    PropertyChanged "Font"
End Property

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("IconSize", m_IconSize, 16)
    Call PropBag.WriteProperty("IconColorDepth", m_ColorDepth, 0)
    If m_Style = [Dropdown Combo] Then
        Call PropBag.WriteProperty("Text", ComboEx.GetEditString, Extender.Name)
    End If
    Call PropBag.WriteProperty("ShowIcons", m_ShowIcons, m_def_ShowIcons)
    Call PropBag.WriteProperty("Style", m_Style, 0)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
End Sub

Public Property Get IconSize() As IconSizeConstants
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_IconSize As IconSizeConstants)
Dim hImageList As Long
    If Not m_ImageList Is Nothing Then hImageList = m_ImageList.hImageList
    m_IconSize = New_IconSize
    ComboEx.SetIconSize m_IconSize
    ComboEx.ShowIcons m_ShowIcons, hImageList
    Extender.Height = (ComboEx.GetItemHeight(-1) + 6) * tpy
    PropertyChanged "IconSize"
End Property
Public Sub Clear()
    ComboEx.Clear
    Set m_ItemData = New Collection
End Sub

Public Sub Refresh()
    ComboEx.Refresh
End Sub

Public Sub AddItem(Item As String, Optional Index As Variant, Optional ImgIndex As Variant, Optional Indent As Integer)
    If Not IsMissing(Index) Then
        m_ItemData.Add 0, , Index
    Else
        m_ItemData.Add 0
        Index = -1
    End If
    If IsMissing(ImgIndex) Then ImgIndex = -1
    ComboEx.AddItem Item, CInt(Index), CInt(ImgIndex) - 1, Indent
End Sub

Public Function AddIcon(hIcon As Variant)
    AddIcon = ComboEx.AddIcon(hIcon)
End Function

Public Property Get ListCount() As Integer
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = ComboEx.ListCount
End Property

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = ComboEx.GetIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    ComboEx.SetIndex New_ListIndex
    PropertyChanged "ListIndex"
End Property

Public Property Get List(Index) As String
Attribute List.VB_MemberFlags = "400"
    List = ComboEx.GetItem(CInt(Index))
End Property

Public Property Let List(Index, ByVal New_List As String)
    ComboEx.SetItem New_List, CInt(Index), -1, -1
    PropertyChanged "List"
End Property

Public Property Get ItemData(Index As Variant) As Variant
Attribute ItemData.VB_MemberFlags = "400"
    ItemData = m_ItemData(Index + 1)
End Property

Public Property Let ItemData(Index As Variant, ByVal New_Item As Variant)
    m_ItemData.Add New_Item, , Index + 1
    m_ItemData.Remove (Index + 2)
    PropertyChanged "ItemData"
End Property

Public Property Get NewIndex() As Integer
    NewIndex = ComboEx.NewIndex
End Property
Public Property Get Style() As CboStyleConstants
    Style = m_Style
End Property
Public Property Let Style(Value As CboStyleConstants)
    If Ambient.UserMode Then Err.Raise 382
    If m_Style <> Value Then
        m_Style = Value
        ComboEx.Destroy
        CreateCombo
    End If
    PropertyChanged "Style"
End Property
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "2c"
    If Style = [Dropdown Combo] Then
        Text = ComboEx.GetEditString
    ElseIf Not Ambient.UserMode Then
        Text = ""
    Else
        Text = ComboEx.GetText
    End If
End Property

Public Property Let Text(ByVal New_Text As String)
    If CanPropertyChange("Text") Then
        If Style = [Dropdown Combo] Then
            ComboEx.SetEditString New_Text
        Else
            MsgBox "Text Property is Read-Only with Dropdown List Style", vbInformation
        End If
        PropertyChanged "Text"
    End If
End Property

Private Sub CreateCombo()
Dim hImageList As Long
    If Not m_ImageList Is Nothing Then hImageList = m_ImageList.hImageList
    With Extender
        ComboEx.ParentHwnd = UserControl.hWnd
        ComboEx.Create 0, 0, .Width / tpx, (ComboHeight() * 9) - (ComboHeight() - 16) + 11, hImageList, CInt(m_IconSize), (m_Style = [Dropdown Combo]), m_ShowIcons
    End With
    'If Ambient.UserMode Then
        EditHwnd = ComboEx.GetEditHwnd
        HookWindow ComboEx.GetComboHwnd, Me
        HookWindow GetParent(ComboEx.GetComboHwnd), Me
        HookWindow EditHwnd, Me
    'End If
    Hooked = Ambient.UserMode
End Sub

Private Function ComboHeight() As Integer
Dim th As Integer
th = (UserControl.TextHeight("M") / tpy) + 3
    If m_ShowIcons And m_IconSize >= th Then
        ComboHeight = m_IconSize
    ElseIf m_ShowIcons Or th > 16 Then
        ComboHeight = th
    Else
        ComboHeight = 16
    End If
End Function

Public Property Get ShowIcons() As Boolean
    ShowIcons = m_ShowIcons
End Property

Public Property Let ShowIcons(ByVal New_ShowIcons As Boolean)
Dim hImageList As Long
    If Not m_ImageList Is Nothing Then hImageList = m_ImageList.hImageList
    m_ShowIcons = New_ShowIcons
    ComboEx.ShowIcons m_ShowIcons, hImageList
    Extender.Height = (ComboEx.GetItemHeight(-1) + 6) * tpy
    PropertyChanged "ShowIcons"
End Property

Public Property Get IconColorDepth() As ColorDepthConstants
   ' Returns the ColourDepth:
    IconColorDepth = m_ColorDepth
End Property

Public Property Let IconColorDepth(ByVal eDepth As ColorDepthConstants)
    If Ambient.UserMode Then Err.Raise 382
   ' Sets the ColourDepth.  NB no change at runtime unless you
   ' call Create and rebuild the image list.
    m_ColorDepth = eDepth
    PropertyChanged "IconColorDepth"
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
Attribute ImageList.VB_MemberFlags = "400"
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    If TypeName(New_ImageList) = "ImageList" Then
        Set m_ImageList = New_ImageList
        ShowIcons = True
        PropertyChanged "ImageList"
    End If
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get hWnd() As Long
    hWnd = ComboEx.GetComboHwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    If m_Style = [Dropdown Combo] Then
        Call SendMessageLong(ComboEx.GetEditHwnd, EM_SETREADONLY, m_Locked, 0&)
    End If
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    Dim lEnd As Long
    SendMessage ComboEx.GetEditHwnd, EM_GETSEL, m_SelStart, lEnd
    SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    m_SelStart = New_SelStart
    SendMessageLong ComboEx.GetEditHwnd, EM_SETSEL, m_SelStart, m_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    SendMessage ComboEx.GetEditHwnd, EM_GETSEL, m_SelStart, m_SelEnd
    SelLength = m_SelEnd - m_SelStart
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    m_SelEnd = SelStart + New_SelLength
    SendMessageLong ComboEx.GetEditHwnd, EM_SETSEL, SelStart, m_SelEnd
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    SelText = Mid$(Text, SelStart + 1, SelLength)
End Property

Public Property Let SelText(ByVal New_SelText As String)
    SendMessage ComboEx.GetEditHwnd, EM_REPLACESEL, True, ByVal New_SelText
    PropertyChanged "SelText"
End Property


Public Function WindowProc(ByVal CtlHwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
Attribute WindowProc.VB_MemberFlags = "40"
Const WM_COMMAND = &H111
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_CHAR = &H102
Const WM_USER = &H400
Const CBN_SELCHANGE = 1
Const CBN_DROPDOWN = 7
Const CBN_EDITCHANGE = 5
Dim Shift As Integer
    Select Case CtlHwnd
    Case EditHwnd
        If uMsg = WM_CHAR Then
            RaiseEvent KeyPress(CInt(wParam))
        ElseIf uMsg = WM_KEYDOWN Then
            If wParam = 16 Then Shift = 1
            If wParam = 17 Then Shift = Shift + 2
            If wParam = 18 Then Shift = Shift + 4
            RaiseEvent KeyDown(CInt(wParam), Shift)
        ElseIf uMsg = WM_KEYUP Then
            If wParam = 16 Then Shift = 1
            If wParam = 17 Then Shift = Shift + 2
            If wParam = 18 Then Shift = Shift + 4
            RaiseEvent KeyUp(CInt(wParam), Shift)
        End If
    Case GetParent(ComboEx.GetComboHwnd)
        If uMsg = WM_COMMAND Then
            If HiWord(wParam) = CBN_SELCHANGE Then
                RaiseEvent Click
            ElseIf HiWord(wParam) = CBN_DROPDOWN Then
                RaiseEvent DropDown
            ElseIf HiWord(wParam) = CBN_EDITCHANGE Then
                RaiseEvent Change
                PropertyChanged "Text"
            End If
        End If
    End Select
    'Trata o BackColor
    If uMsg = WM_CTLCOLOREDIT Or _
       uMsg = WM_CTLCOLORLISTBOX Or _
       uMsg = WM_CTLCOLORSTATIC Then
        WindowProc = False
    Else
        WindowProc = True
    End If
End Function
