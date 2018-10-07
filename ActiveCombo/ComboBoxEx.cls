VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ImgList As New CImageList
Public NewIndex As Integer

Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_SETFONT = &H30

Private NewComboFont As Long
Private cmbFontBold As Boolean
Private cmbFontItalic As Boolean
Private cmbFontName As String
Private cmbFontHeight As Integer
Private cmbCustomize As Boolean
Private cmbFontUnderlined As Boolean
Private cmbColorDepth As Integer
 
Private Const SYSTEM_FONT& = 13
Private LF As LOGFONT
Public Enum cbIconState
  cbNormal = 0
  cbDisabled = 1
End Enum

Private Type FONTSTRUC
    lStructSize As Long
    hWnd As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lffacename As String * 32
   'lfFaceName(LF_FACESIZE) As Byte
End Type

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SendStringMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetComboFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long


Private Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Const CF_BITMAP = 2
Private Const SWP_NOACTIVATE = &H10
Private ObjHwnd As Long
 
Private Const CB_SETCURSEL = &H14E
Private Const CB_GETCOUNT = &H146
Private Const CB_GETCURSEL = &H147
Private Const CB_GETEDITSEL = &H140
Private Const CB_GETITEMHEIGHT = &H154
Private Const CB_GETLBTEXT = &H148
Private Const CB_GETLBTEXTLEN = &H149
Private Const CB_SELECTSTRING = &H14D
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_RESETCONTENT = &H14B
Private Const CB_ERR = (-1)

Private cbItems As COMBOBOXEXITEMW

Const WS_VISIBLE = &H10000000
Const WS_CHILD = &H40000000
Const WS_BORDER = &H800000

Private Const WM_USER = &H400
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
 
Private ComboExhWnd As Long

Const HWND_TOPMOST = -1
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
 
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOREDRAW = &H8
Const SWP_SHOWWINDOW = &H40

Private Const ICC_USEREX_CLASSES = &H200

Private Const WC_COMBOBOXEXW = "ComboBoxEx32"
Private Const WC_COMBOBOXEXA = "ComboBoxEx32"
 
#If UNICODE Then
Private Const WC_COMBOBOXEX = WC_COMBOBOXEXW
#Else
Private Const WC_COMBOBOXEX = WC_COMBOBOXEXA
#End If

Private Const CBS_AUTOHSCROLL = &H40&
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBS_HASSTRINGS = &H200&
Private Const CBS_DISABLENOSCROLL = &H800&
Private Const CBS_NOINTEGRALHEIGHT = &H400&
Private Const CBS_OWNERDRAWFIXED = &H10&
Private Const CBS_OWNERDRAWVARIABLE = &H20&
Private Const CBS_SIMPLE = &H1&
Private Const CBS_SORT = &H100&
Private Const CB_SETEDITSEL = &H142
 
Private Const CBEIF_TEXT = &H1
Private Const CBEIF_IMAGE = &H2
Private Const CBEIF_SELECTEDIMAGE = &H4
Private Const CBEIF_OVERLAY = &H8
Private Const CBEIF_INDENT = &H10
Private Const CBEIF_LPARAM = &H20
Private Const CBEIF_DI_SETITEM = &H10000000
Private Const H_MAX As Long = &HFFFF + 1
Private Const CBEN_FIRST = (H_MAX - 800&)
Private Const CBEN_LAST = (H_MAX - 830&)
Private Const CBEN_GETDISPINFO = (CBEN_FIRST - 0)
Private Const CBEN_INSERTITEM = (CBEN_FIRST - 1)
Private Const CBEN_DELETEITEM = (CBEN_FIRST - 2)
Private Const CBEN_BEGINEDIT = (CBEN_FIRST - 4)
Private Const CBEN_ENDEDITA = (CBEN_FIRST - 5)
Private Const CBEN_ENDEDITW = (CBEN_FIRST - 6)
Private Const CBN_EDITCHANGE = 5
Private Const CBN_EDITUPDATE = 6
Private Const CBN_SELCHANGE = 1
Private Const CB_DELETESTRING = &H144
Private Const CBEM_INSERTITEMA = (WM_USER + 1)
Private Const CBEM_SETIMAGELIST = (WM_USER + 2)
Private Const CBEM_GETIMAGELIST = (WM_USER + 3)
Private Const CBEM_GETITEMA = (WM_USER + 4)
Private Const CBEM_SETITEMA = (WM_USER + 5)
Private Const CBEM_DELETEITEM = CB_DELETESTRING
Private Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL = (WM_USER + 7)
Private Const CBEM_SETEXSTYLE = (WM_USER + 8)
Private Const CBEM_GETEXSTYLE = (WM_USER + 9)
Private Const CBEM_HASEDITCHANGED = (WM_USER + 10)
Private Const CBEM_INSERTITEMW = (WM_USER + 11)
Private Const CBEM_SETITEMW = (WM_USER + 12)
Private Const CBEM_GETITEMW = (WM_USER + 13)

Private Type COMBOBOXEXITEMW
    mask As Long
    iItem As Long
    pszText As String
    cchTextMax  As Long
    iImage As Long
    iSelectedImage As Long
    iOverlay As Long
    iIndent As Long
    lParam As Long
End Type


#If UNICODE Then
Private Const CBEM_INSERTITEM = CBEM_INSERTITEMW
Private Const CBEM_SETITEM = CBEM_SETITEMW
Private Const CBEM_GETITEM = CBEM_GETITEMW
#Else
Private Const CBEM_INSERTITEM = CBEM_INSERTITEMA
Private Const CBEM_SETITEM = CBEM_SETITEMA
Private Const CBEM_GETITEM = CBEM_GETITEMA
#End If

Private Const CBES_EX_NOEDITIMAGE = &H1
Private Const CBES_EX_NOEDITIMAGEINDENT = &H2
Private Const CBES_EX_PATHWORDBREAKPROC = &H4

Public Function AddIcon(hIcon As Variant)
    AddIcon = ImgList.AddIcon(hIcon)
End Function

Public Sub Clear()
  Dim ComboCount As Integer, Total
  Total = ListCount
 
  For ComboCount = 0 To Total - 1
    Call SendMessage(ComboExhWnd, CBEM_DELETEITEM, 0, 0)
  Next

  cbItems.mask = 0&
  cbItems.pszText = ""
  cbItems.cchTextMax = 0
  cbItems.iIndent = 0
  cbItems.iImage = -1
  cbItems.iSelectedImage = -1
  cbItems.iItem = -1
  cbItems.iOverlay = -1
  
  Call SendMessage(ComboExhWnd, CBEM_INSERTITEM, -1, cbItems)
  Call SendMessage(ComboExhWnd, CBEM_DELETEITEM, 0, 0)
  
  SetIndex 0
End Sub

Public Sub RemoveItem(ByVal Item As Integer)
    Call SendMessage(ComboExhWnd, CBEM_DELETEITEM, Item, 0)
    SetIndex 0
End Sub

Public Sub Destroy()
    ImgList.Destroy
    Call DestroyWindow(ComboExhWnd)
End Sub

Public Function GetComboHwnd() As Long
    GetComboHwnd = ComboExhWnd
End Function

Public Function GetEditHwnd() As Long
    GetEditHwnd = SendMessage(ComboExhWnd, CBEM_GETEDITCONTROL, 0, 0)
End Function

Public Function GetDropDownHwnd() As Long
    GetDropDownHwnd = SendMessage(ComboExhWnd, CBEM_GETCOMBOCONTROL, 0, 0)
End Function
 
Public Function GetEditString() As String
Dim ComboString As String * 255, lResult As Long
    'Gets the Text in the Edit portion of the ComboBox
    lResult = GetWindowText(GetEditHwnd, ComboString, Len(ComboString))
    GetEditString = Left(ComboString, lResult)
 
End Function

Public Function ListCount() As Integer
    ListCount = SendMessage(GetComboHwnd, CB_GETCOUNT, 0, 0)
End Function

Public Sub SetEditString(EditString As String)
Dim lResult As Long
    'Sets the Text in the Edit portion of the ComboBox
    'Only if it's editable
    lResult = SetWindowText(GetEditHwnd, ByVal EditString)
    Call UpdateWindow(GetEditHwnd)

End Sub

Public Sub ResizeCombo(cmbLeft As Integer, cmbTop As Integer, cmbWidth As Integer)
Dim ComboRect As RECT
    Call GetWindowRect(ComboExhWnd, ComboRect)
    Call MoveWindow(ComboExhWnd, CLng(cmbLeft), CLng(cmbTop), CLng(cmbWidth), GetItemHeight(-1) + (GetItemHeight(0) * 8), True)
End Sub

Public Sub SetDropWidth(NewWidth As Integer)
    Call SendMessage(GetDropDownHwnd, CB_SETDROPPEDWIDTH, NewWidth, 0&)
End Sub
Public Function GetIndex() As Integer
    GetIndex = SendMessage(ComboExhWnd, CB_GETCURSEL, 0, 0)
End Function
Public Sub SetIndex(Index As Integer)
Dim CurSel As Long
     CurSel = SendMessage(ComboExhWnd, CB_SETCURSEL, Index, 0)
End Sub

Public Function GetItemHeight(Item As Long) As Long
     GetItemHeight = SendMessageByLong(GetDropDownHwnd, CB_GETITEMHEIGHT, Item, 0)
End Function

Public Sub SetItemHeight(Item As Long, ItemHeight As Integer)
    'Set the New Item Height of drop down
    Call SendMessageByLong(GetDropDownHwnd, CB_SETITEMHEIGHT, 0, ByVal (ItemHeight And &HFFFF))
    'Update the Window
    Call UpdateWindow(GetDropDownHwnd)
End Sub

Private Sub Class_Initialize()
Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    Call InitCommonControlsEx(iccex)
    NewIndex = -1
    ComboExhWnd = 0
End Sub
 
Public Function Create( _
 Optional Left As Variant, _
 Optional Top As Variant, _
 Optional Width As Variant, _
 Optional Height As Variant, _
 Optional ImageListHwnd As Long, _
 Optional IconSize As Integer = 16, _
 Optional Editable As Boolean, _
 Optional ShowIcons As Boolean = True) As Boolean
    
    'if we didn't pass a previously created imagelist then create one
    If ImageListHwnd = 0 Then
        With ImgList
            If IconSize = 16 Then .Create Size16, cmbColorDepth
            If IconSize = 32 Then .Create Size32, cmbColorDepth
        End With
        ImageListHwnd = ImgList.hWnd
    End If
    
    'Create the ComboBoxEx Control
    'CBS_DROPDOWNLIST - Not Editable
    'CBS_DROPDOWN
    If Editable = True Then
        ComboExhWnd = CreateWindowEX(0, WC_COMBOBOXEX, "", _
                      WS_CHILD Or WS_VISIBLE Or WS_BORDER Or CBS_DROPDOWN Or CBS_AUTOHSCROLL, _
                      Left, Top, Width, Height, _
                      ParentHwnd, 0&, App.hInstance, 0&)
    Else
        ComboExhWnd = CreateWindowEX(0, WC_COMBOBOXEX, "", _
                      WS_CHILD Or WS_VISIBLE Or WS_BORDER Or CBS_DROPDOWNLIST, _
                      Left, Top, Width, Height, _
                      ParentHwnd, 0&, App.hInstance, 0&)
    End If
    'Set the parent to receive the messages
    Call SetParent(ComboExhWnd, ParentHwnd)
    Call MoveWindow(ComboExhWnd, CLng(Left), CLng(Top), CLng(Width), GetItemHeight(-1) + (GetItemHeight(0) * 8), True)
    If cmbCustomize = True Then SetComboFont
    If ShowIcons Then
        'Set the Imagelist for the ComboBox
        Call SendMessage(ComboExhWnd, CBEM_SETIMAGELIST, 0, ByVal ImageListHwnd)
    Else
        Call SendMessage(ComboExhWnd, CBEM_SETIMAGELIST, 0, 0)
    End If
End Function
Public Sub Show()
    Call ShowWindow(ComboExhWnd, SW_SHOWNORMAL)
End Sub
Public Property Get ParentHwnd() As Long
    ParentHwnd = ObjHwnd
End Property

Public Property Let ParentHwnd(hWnd As Long)
    ObjHwnd = hWnd
End Property

Public Sub AddItem(StringToAdd As String, Index As Integer, ImgIndex As Integer, Indent As Integer)
    
    cbItems.mask = CBEIF_TEXT Or CBEIF_INDENT Or CBEIF_IMAGE Or CBEIF_LPARAM Or CBEIF_SELECTEDIMAGE Or CBEIF_OVERLAY
    cbItems.pszText = StringToAdd
    cbItems.cchTextMax = Len(StringToAdd)
    cbItems.iIndent = Indent
    cbItems.iImage = ImgIndex
    cbItems.iSelectedImage = ImgIndex
    cbItems.iItem = Index
    cbItems.iOverlay = ImgIndex
    'cbItems.lParam
   
    NewIndex = SendMessage(ComboExhWnd, CBEM_INSERTITEM, Index, cbItems)
 
End Sub

Public Function GetText() As String
Dim CurSel As Integer
    On Error Resume Next
    CurSel = SendMessage(ComboExhWnd, CB_GETCURSEL, 0, 0)
    GetText = GetItem(CurSel)
End Function

Public Function GetItem(Index As Integer) As String
Dim Tam As Integer, Item As String
    Tam = SendMessage(ComboExhWnd, CB_GETLBTEXTLEN, CLng(Index), 0)
    If Tam = CB_ERR Then Exit Function
    Item = Space(Tam)
    If SendMessage(ComboExhWnd, CB_GETLBTEXT, CLng(Index), ByVal (Item)) > 0 Then
        GetItem = Left(Item, Tam)
    End If
End Function

Public Sub SetItem(StringToAdd As String, Index As Integer, ImgIndex As Integer, Indent As Integer)
    cbItems.mask = CBEIF_TEXT Or CBEIF_INDENT Or CBEIF_IMAGE Or CBEIF_LPARAM Or CBEIF_SELECTEDIMAGE Or CBEIF_OVERLAY
    cbItems.pszText = StringToAdd
    cbItems.cchTextMax = Len(StringToAdd)
    cbItems.iIndent = Indent
    cbItems.iImage = ImgIndex
    cbItems.iSelectedImage = ImgIndex
    cbItems.iItem = Index
    cbItems.iOverlay = ImgIndex
    'cbItems.lParam
    Call SendMessage(ComboExhWnd, CBEM_SETITEM, Index, cbItems)
End Sub
Private Sub Class_Terminate()
Dim dl As Long
    On Error Resume Next
    If ComboExhWnd <> 0 Then
        ImgList.Destroy
        Call DestroyWindow(ComboExhWnd)
    End If
    If NewComboFont <> 0 Then
        dl = DeleteObject(NewComboFont)
    End If
End Sub

Public Sub SetComboFont()
Dim cbDC As Long
Dim dl As Long
Dim CurrentComboFont As Long
Dim mFlags As Long
Const TRANSPARENT = 1

    cbDC = GetDC(GetDropDownHwnd)
    CurrentComboFont = SelectObject(cbDC, GetStockObject(SYSTEM_FONT))
    dl = SetBkMode(cbDC, TRANSPARENT)
    
    LF.lffacename = cmbFontName & Chr$(0)
    LF.lfHeight = cmbFontHeight
    LF.lfUnderline = cmbFontUnderlined
    LF.lfItalic = cmbFontItalic
    
    If cmbFontBold = True Then
        LF.lfWeight = 600
    Else
        LF.lfWeight = 300
    End If
    NewComboFont = CreateFontIndirect(LF)
    dl = SelectObject(cbDC, NewComboFont)
      
    'Set the New Font to drop down
    Call SendMessage(GetDropDownHwnd, WM_SETFONT, NewComboFont, 1)
    
    'Set the New Font to the Edit window
    Call SendMessage(GetEditHwnd, WM_SETFONT, NewComboFont, 1)
       
    NewComboFont = SelectObject(cbDC, CurrentComboFont)
    'Restore original font
    dl = SelectObject(cbDC, CurrentComboFont)
    
    'NewComboFont is deleted in terminate Event
     
End Sub

Public Sub ShowIcons(ByVal Value As Boolean, ByVal hImageList As Long)
Dim ComboRect As RECT
    If hImageList = 0 Then hImageList = ImgList.hWnd
    Call GetWindowRect(ComboExhWnd, ComboRect)
    If Value Then
        Call SendMessage(ComboExhWnd, CBEM_SETIMAGELIST, 0, ByVal hImageList)
    Else
        Call SendMessage(ComboExhWnd, CBEM_SETIMAGELIST, 0, 0)
    End If
End Sub
Public Sub SetIconSize(IconSize As Integer)
    With ImgList
        If .hWnd <> 0 Then .Destroy
        If IconSize = 16 Then .Create Size16, cmbColorDepth
        If IconSize = 32 Then .Create Size32, cmbColorDepth
    End With
End Sub

Public Sub SetFocus()
    SetComboFocus ComboExhWnd
End Sub
Public Sub SetEditFocus()
    SetComboFocus GetEditHwnd
End Sub
'Public Sub SetBackColor(BkColor As Long)
'Dim hdc As Long
'    hdc = GetDC(GetEditHwnd)
'    SetBkColor hdc, BkColor
'End Sub
'
'Public Sub SetForeColor(FrColor As Long)
'Dim hdc As Long
'    hdc = GetDC(GetEditHwnd)
'    SetTextColor hdc, FrColor
'End Sub
Public Sub Refresh()
    Call UpdateWindow(ComboExhWnd)
End Sub
Public Property Let FontBold(ByVal vNewValue As Boolean)
    cmbFontBold = vNewValue
    cmbCustomize = True
End Property
Public Property Let FontItalic(ByVal vNewValue As Boolean)
    cmbFontItalic = vNewValue
    cmbCustomize = True
End Property
Public Property Let FontName(ByVal vNewValue As String)
    cmbFontName = vNewValue
    cmbCustomize = True
End Property
Public Property Let FontHeight(ByVal vNewValue As Integer)
    cmbFontHeight = vNewValue
    cmbCustomize = True
End Property
Public Property Let FontUnderlined(ByVal vNewValue As Boolean)
    cmbFontUnderlined = vNewValue
    cmbCustomize = True
End Property
Public Property Get ColorDepth() As Integer
   ' Returns the ColourDepth:
    ColorDepth = cmbColorDepth
End Property
Public Property Let ColorDepth(ByVal eDepth As Integer)
   ' Sets the ColourDepth.  NB no change at runtime unless you
   ' call Create and rebuild the image list.
    cmbColorDepth = eDepth
End Property

