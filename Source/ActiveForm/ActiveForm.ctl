VERSION 5.00
Begin VB.UserControl ActiveForm 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   450
   ToolboxBitmap   =   "ActiveForm.ctx":0000
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   0
      Picture         =   "ActiveForm.ctx":0312
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "ActiveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private FormHook As Boolean

Private oElastic As cElastic
Private WithEvents oForm As Form
Attribute oForm.VB_VarHelpID = -1
Private mhWnd As Long
Private bTrayFlag As Boolean
Private lWindowState As Long
Private pPicture As StdPicture
Private tTrayStuff As NOTIFYICONDATA

'Constants
Const hWnd_TOPMOST = -1
Const hWnd_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const GW_CHILD = 5

'Enums
Public Enum rdGradientConstants
    bgVertical
    bgHorizontal
    bgCircle
    bgRectangle
    bgDiagLeftRight
    bgDiagRightLeft
End Enum

Public Enum rdBackgroundConstants
    bgNone
    bgGradient
    bgTransparent
    bgTiledPicture
    bgLeftPicture
    bgCenterPicture
End Enum

Public Enum rdRestoreModeConstants
    rmManual
    rmSingleClick
    rmDoubleClick
End Enum

'TypeDefs
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
'
' System Tray Messages and Structures
'
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Const NIM_ADD As Long = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE As Long = &H2
Const NIF_MESSAGE As Long = &H1
Const NIF_ICON As Long = &H2
Const NIF_TIP As Long = &H4
Const ICON_SMALL = 0
Const ICON_BIG = 1
'
' Mouse Messages Captured from the System Tray
'
Const WM_MOUSEMOVE      As Long = &H200
Const WM_LBUTTONDOWN    As Long = &H201
Const WM_LBUTTONUP      As Long = &H202
Const WM_LBUTTONDBLCLK  As Long = &H203
Const WM_RBUTTONDOWN    As Long = &H204
Const WM_RBUTTONUP      As Long = &H205
Const WM_RBUTTONDBLCLK  As Long = &H206
Const WM_MBUTTONDOWN    As Long = &H207
Const WM_MBUTTONUP      As Long = &H208
Const WM_MBUTTONDBLCLK  As Long = &H209
Const WM_MOUSELAST      As Long = &H209
Const WM_GETICON        As Long = &H7F


'Property Variables:
Private m_MinWidth As Long
Private m_MinHeight As Long
Private m_MaxWidth As Long
Private m_MaxHeight As Long
Private m_ResizeControls As Boolean
Private m_ResizeFonts As Boolean
Private m_CloseButton As Boolean
Private m_AllwaysOnTop As Boolean
Private m_Gradient As rdGradientConstants
Private m_BackColor As Long
Private m_BackColor2 As Long
Private m_Background As rdBackgroundConstants
Private m_MinimizeToTray As Boolean
Private m_RestoreMode As rdRestoreModeConstants

'Default Property Values:
Const m_def_MinWidth = 0
Const m_def_MinHeight = 0
Const m_def_MaxWidth = 0
Const m_def_MaxHeight = 0
Const m_def_ResizeControls = False
Const m_def_ResizeFonts = False
Const m_def_CloseButton = True
Const m_def_AllwaysOnTop = False
Const m_def_Gradient = 0
Const m_def_BackGround = 0
Const m_def_BackColor = vbBlue
Const m_def_BackColor2 = vbBlack
Const m_def_MinimizeToTray = False
Const m_def_RestoreMode = rmDoubleClick

' API
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Event TrayLeftClick()
Event TrayRightClick()
Event TrayDblClick()

Private Sub oForm_Load()
    Select Case m_Background
    Case bgGradient
        Shade Extender.Parent, m_Gradient, m_BackColor, m_BackColor2
    Case bgTiledPicture
        PaintWallPaper
    Case bgLeftPicture
        PaintLeftPicture
    Case bgCenterPicture
        PaintCenterPicture
    End Select
End Sub

Private Sub oForm_Unload(Cancel As Integer)
    If Not Cancel And bTrayFlag Then KillSysTrayIcon
End Sub

Private Sub UserControl_Initialize()
    FormHook = False
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error Resume Next
    m_MinWidth = m_def_MinWidth
    m_MinHeight = m_def_MinHeight
    m_MaxWidth = m_def_MaxWidth
    m_MaxHeight = m_def_MaxHeight
    m_ResizeControls = m_def_ResizeControls
    m_ResizeFonts = m_def_ResizeFonts
    m_CloseButton = m_def_CloseButton
    m_Gradient = m_def_Gradient
    m_BackColor = m_def_BackColor
    m_BackColor2 = m_def_BackColor2
    m_Background = m_def_BackGround
    m_AllwaysOnTop = m_def_AllwaysOnTop
    m_MinimizeToTray = m_def_MinimizeToTray
    m_RestoreMode = m_def_RestoreMode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
    m_MaxWidth = PropBag.ReadProperty("MaxWidth", m_def_MaxWidth)
    m_MaxHeight = PropBag.ReadProperty("MaxHeight", m_def_MaxHeight)
    m_ResizeControls = PropBag.ReadProperty("ResizeControls", m_def_ResizeControls)
    m_ResizeFonts = PropBag.ReadProperty("ResizeFonts", m_def_ResizeFonts)
    m_Gradient = PropBag.ReadProperty("BackGradient", m_def_Gradient)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackColor2 = PropBag.ReadProperty("BackColor2", m_def_BackColor2)
    m_Background = PropBag.ReadProperty("Background", m_def_BackGround)
    m_CloseButton = PropBag.ReadProperty("CloseButton", m_def_CloseButton)
    m_AllwaysOnTop = PropBag.ReadProperty("AllwaysOnTop", m_def_AllwaysOnTop)
    m_MinimizeToTray = PropBag.ReadProperty("MinimizeToTray", m_def_MinimizeToTray)
    m_RestoreMode = PropBag.ReadProperty("RestoreMode", m_def_RestoreMode)
    'Se Estiver em Runtime e o container for um
    'Form com a borda de tamanho variável então
    'realiza o SubClass e Configura o Tamanho
    If Ambient.UserMode Then
        If TypeOf Extender.Parent Is Form Then
            On Error Resume Next
            mhWnd = Extender.Parent.hwnd
            If TypeOf Extender.Parent Is MDIForm Then
                FormHook = HookWindow(mhWnd, Me)
            Else
                Set oForm = Extender.Parent
                If oForm.BorderStyle = vbSizable Or _
                   oForm.BorderStyle = vbSizableToolWindow Or _
                   m_MinimizeToTray Or m_Background > bgNone Then
                    Set oElastic = New cElastic
                    oElastic.Link oForm
                    FormHook = HookWindow(mhWnd, Me)
                End If
            End If
            If Not m_CloseButton Then
                DisableCloseButton Extender.Parent.hwnd
            End If
            If m_AllwaysOnTop Then
                Call SetWindowPos(Extender.Parent.hwnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
            End If
            If m_Background = bgTransparent Then
                Call SetWindowLong(Extender.Parent.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
            ElseIf m_Background = bgCenterPicture Then
                Set pPicture = Extender.Parent.Picture
                Extender.Parent.Picture = LoadPicture()
            End If
        End If
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
    Call PropBag.WriteProperty("MaxWidth", m_MaxWidth, m_def_MaxWidth)
    Call PropBag.WriteProperty("MaxHeight", m_MaxHeight, m_def_MaxHeight)
    Call PropBag.WriteProperty("ResizeControls", m_ResizeControls, m_def_ResizeControls)
    Call PropBag.WriteProperty("ResizeFonts", m_ResizeFonts, m_def_ResizeFonts)
    Call PropBag.WriteProperty("CloseButton", m_CloseButton, m_def_CloseButton)
    Call PropBag.WriteProperty("BackGradient", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColor2", m_BackColor2, m_def_BackColor2)
    Call PropBag.WriteProperty("Background", m_Background, m_def_BackGround)
    Call PropBag.WriteProperty("AllwaysOnTop", m_AllwaysOnTop, m_def_AllwaysOnTop)
    Call PropBag.WriteProperty("MinimizeToTray", m_MinimizeToTray, m_def_MinimizeToTray)
    Call PropBag.WriteProperty("RestoreMode", m_RestoreMode, m_def_RestoreMode)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
' Here's where we handle the Icon Tray Messages
'
    Dim lMsg As Long
    Static bInHere As Boolean
    
    On Error GoTo vbErrorHandler
    
    lMsg = x / Screen.TwipsPerPixelX
    
    If bInHere Then Exit Sub
    
    bInHere = True
    Select Case lMsg
    Case WM_LBUTTONDBLCLK
        '
        ' On Mouse DoubleClick - Restore the window or Raise Event
        '
        If m_RestoreMode = rmDoubleClick Then
           Restore
        Else
           SetForegroundWindow oForm.hwnd
           RaiseEvent TrayDblClick
        End If
    Case WM_LBUTTONUP
        '
        ' On Mouse LeftClick - Restore the window or Raise Event
        '
        If m_RestoreMode = rmSingleClick Then
           Restore
        Else
           SetForegroundWindow oForm.hwnd
           RaiseEvent TrayLeftClick
        End If
        
    Case WM_RBUTTONUP
        '
        ' On Mouse RightClick - Raise Event
        '
        SetForegroundWindow oForm.hwnd
        RaiseEvent TrayRightClick
        
    End Select
    
    bInHere = False
    
    Exit Sub
    
vbErrorHandler:
    MsgBox Err.Number & " " & Err.Description & "  " & Err.Source & " frmCodeLib::picSysBar_MouseMove", , App.ProductName
End Sub

Private Sub UserControl_Resize()
    Width = ScaleX(Image1.Picture.Width)
    Height = ScaleX(Image1.Picture.Height)
End Sub

Private Sub UserControl_Terminate()
    If FormHook Then
        UnHookWindow mhWnd
        Set oElastic = Nothing
    End If
    If Not oForm Is Nothing Then Set oForm = Nothing
    If bTrayFlag Then KillSysTrayIcon
End Sub

Public Property Get AllwaysOnTop() As Boolean
Attribute AllwaysOnTop.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllwaysOnTop = m_AllwaysOnTop
End Property

Public Property Let AllwaysOnTop(ByVal New_AllwaysOnTop As Boolean)
    m_AllwaysOnTop = New_AllwaysOnTop
    If Ambient.UserMode Then
        If m_AllwaysOnTop Then
            Call SetWindowPos(Extender.Parent.hwnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        Else
            Call SetWindowPos(Extender.Parent.hwnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        End If
    End If
    PropertyChanged "AllwaysOnTop"
End Property

Public Property Get CloseButton() As Boolean
    CloseButton = m_CloseButton
End Property

Public Property Let CloseButton(ByVal New_CloseButton As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    m_CloseButton = New_CloseButton
    PropertyChanged "CloseButton"
End Property

Public Property Get BackGradient() As rdGradientConstants
Attribute BackGradient.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackGradient = m_Gradient
End Property

Public Property Let BackGradient(ByVal New_Gradient As rdGradientConstants)
    m_Gradient = New_Gradient
    Extender.Parent.Refresh
    PropertyChanged "BackGradient"
End Property

Public Property Get Background() As rdBackgroundConstants
Attribute Background.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Background = m_Background
End Property

Public Property Let Background(ByVal New_Background As rdBackgroundConstants)
    If TypeOf Extender.Parent Is MDIForm Then
        MsgBox "This Property can't be used in MDI Forms.", vbInformation
    Else
        m_Background = New_Background
        Extender.Parent.Refresh
        PropertyChanged "Background"
    End If
End Property

Public Property Let BackColor(ByVal New_Color As OLE_COLOR)
    m_BackColor = New_Color
    Extender.Parent.Refresh
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
End Property

Public Property Let BackColor2(ByVal New_Color As OLE_COLOR)
    m_BackColor2 = New_Color
    Extender.Parent.Refresh
    PropertyChanged "BackColor2"
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor2 = m_BackColor2
End Property

Public Property Get ResizeControls() As Boolean
Attribute ResizeControls.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ResizeControls = m_ResizeControls
End Property

Public Property Let ResizeControls(ByVal New_ResizeControls As Boolean)
    If TypeOf Extender.Parent Is MDIForm Then
        MsgBox "This Property can't be used in MDI Forms.", vbInformation
    Else
        m_ResizeControls = New_ResizeControls
        PropertyChanged "ResizeControls"
    End If
End Property

Public Property Get ResizeFonts() As Boolean
Attribute ResizeFonts.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ResizeFonts = m_ResizeFonts
End Property

Public Property Let ResizeFonts(ByVal New_ResizeFonts As Boolean)
    If TypeOf Extender.Parent Is MDIForm Then
        MsgBox "This Property can't be used in MDI Forms.", vbInformation
    Else
        m_ResizeFonts = New_ResizeFonts
        PropertyChanged "ResizeFonts"
    End If
End Property

Public Property Get MinimizeToTray() As Boolean
Attribute MinimizeToTray.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MinimizeToTray = m_MinimizeToTray
End Property

Public Property Let MinimizeToTray(ByVal New_MinimizeToTray As Boolean)
    m_MinimizeToTray = New_MinimizeToTray
    PropertyChanged "MinimizeToTray"
End Property

Public Property Get MinWidth() As Long
Attribute MinWidth.VB_ProcData.VB_Invoke_Property = ";Position"
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

Public Property Get MinHeight() As Long
Attribute MinHeight.VB_ProcData.VB_Invoke_Property = ";Position"
    MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    m_MinHeight = New_MinHeight
    PropertyChanged "MinHeight"
End Property

Public Property Get MaxWidth() As Long
Attribute MaxWidth.VB_ProcData.VB_Invoke_Property = ";Position"
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Long)
    m_MaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
End Property

Public Property Get MaxHeight() As Long
Attribute MaxHeight.VB_ProcData.VB_Invoke_Property = ";Position"
    MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Long)
    m_MaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
End Property

Public Property Get RestoreMode() As rdRestoreModeConstants
    RestoreMode = m_RestoreMode
End Property

Public Property Let RestoreMode(ByVal New_RestoreMode As rdRestoreModeConstants)
    m_RestoreMode = New_RestoreMode
    PropertyChanged "RestoreMode"
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub DisableCloseButton(hwnd As Long)
Const MF_BYPOSITION = &H400&
Const MF_REMOVE = &H1000&
Dim hMenu As Long
Dim nCount As Long
    hMenu = GetSystemMenu(hwnd, 0)
    nCount = GetMenuItemCount(hMenu)
    'Get rid of the Close menu and its separator
    Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
    Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)
    'Make sure the screen updates     'our change
    DrawMenuBar hwnd
End Sub

'***************************************************************
' Name: Form Shade
' Description:This code creates 6 diffrent form shading effects.
'     You can shade the form verticaly, horizontaly, diagnaly, rectangn
'     aly and in a circular pattern with out using Windows API calls.
' By: Lightning Programmers
'
'
' Inputs:sForm = The form to be shaded.
' sType = The type of shade to be apply.
' sColor1 = The starting color.
' sColor2 = The finishing color.
'
' Returns:There are no returns.
'
'Assumes:None
'
'Side Effects:There should not be any side effects.
'
'Code provided by Planet Source Code(tm) (http://www.PlanetSource
'     Code.com) 'as is', without warranties as to performance, fitness,
'     merchantability,and any other warranty (whether expressed or impl
'     ied).
'***************************************************************

Private Sub Shade(sForm As Form, ByRef sType As rdGradientConstants, ByVal sColor1 As Long, ByVal sColor2 As Long)
    Dim x As Integer, y As Integer, r As Integer, xy As Single
    Dim sStart As Integer, sFinish As Integer
    Dim sCRedInc As Single, sCGreenInc As Single, sCBlueInc As Single
    Dim sC1Red As Integer, sC1Green As Integer, sC1Blue As Integer
    Dim sC2Red As Integer, sC2Green As Integer, sC2Blue As Integer
    Dim tRed As Single, tGreen As Single, tBlue As Single
    Dim sScaleMode As Integer
    If sForm.WindowState = 1 Then Exit Sub
    If sColor1 And &H80000000 Then
        sColor1 = GetSysColor(sColor1 - &H80000000)
    End If
    If sColor2 And &H80000000 Then
        sColor2 = GetSysColor(sColor2 - &H80000000)
    End If
    'Form Preparation
    sScaleMode = sForm.ScaleMode
    sForm.Cls
    sForm.AutoRedraw = True
    sForm.DrawStyle = 6
    sForm.ScaleWidth = 99
    sForm.ScaleHeight = 99
    'Color Separation Proccess
    'Separate Color1 Color
    sC1Red = sColor1 And 255
    sC1Green = sColor1 / 256 And 255
    sC1Blue = sColor1 / 65536 And 255
    'Separate Color2 Color
    sC2Red = sColor2 And 255
    sC2Green = sColor2 / 256 And 255
    sC2Blue = sColor2 / 65536 And 255
    'Shading Process
    'Set Starting Color
    tRed = sC1Red
    tGreen = sC1Green
    tBlue = sC1Blue
    GoSub sIncrementCalc
    'Shade
    Select Case sType
    Case bgCircle
    xy = sForm.Height / sForm.Width
    sForm.DrawWidth = Int((Int((sForm.Width / sForm.ScaleWidth)) + _
    Int((sForm.Height / sForm.ScaleWidth))) / 30)
    sFinish = 149
    GoSub sCIRCLE
    Case bgRectangle
    sForm.DrawWidth = Int((Int((sForm.Width / sForm.ScaleWidth)) + _
    Int((sForm.Height / sForm.ScaleWidth))) / 30)
    sFinish = 99
    GoSub sRECTANGLE
    Case bgDiagLeftRight
    sForm.DrawWidth = Int((Int((sForm.Width / sForm.ScaleWidth)) + _
    Int((sForm.Height / sForm.ScaleWidth))) / 15)
    sFinish = 49
    GoSub sDIAGNALLEFTRIGHT
    Case bgDiagRightLeft
    sForm.DrawWidth = Int((Int((sForm.Width / sForm.ScaleWidth)) + _
    Int((sForm.Height / sForm.ScaleWidth))) / 15)
    sFinish = 49
    GoSub sDIAGNALRIGHTLEFT
    Case bgHorizontal
    sFinish = 99
    GoSub sHORIZONTAL
    Case bgVertical
    sFinish = 99
    GoSub sVERTICAL
    End Select
    'Shading Finished
    sForm.ScaleMode = sScaleMode
    Exit Sub
    'Shading Loops
sCIRCLE:


    For r = sFinish To sStart Step -1
        sForm.Circle (50, 50), r / 2, RGB(tRed, tGreen, tBlue), , , xy
        GoSub sIncrement
    Next r

    Return
sRECTANGLE:


    For x = sStart To sFinish
        sForm.Line (x / 2, x / 2)-(99 - (x / 2), 99 - (x / 2)), _
        RGB(tRed, tGreen, tBlue), B
        GoSub sIncrement
    Next x

    Return
sDIAGNALLEFTRIGHT:


    For y = sStart To sFinish
        sForm.Line (0, y * 2)-(y * 2, 0), RGB(tRed, tGreen, tBlue)
        GoSub sIncrement
    Next y



    For y = sStart To sFinish
        sForm.Line (y * 2, 99)-(99, y * 2), RGB(tRed, tGreen, tBlue)
        GoSub sIncrement
    Next y

    Return
sDIAGNALRIGHTLEFT:


    For y = sStart To sFinish
        sForm.Line (99 - (y * 2), 0)-(99, y * 2), RGB(tRed, tGreen, tBlue)
        GoSub sIncrement
    Next y



    For y = sStart To sFinish
        sForm.Line (0, y * 2)-(99 - (y * 2), 99), RGB(tRed, tGreen, tBlue)
        GoSub sIncrement
    Next y

    Return
sHORIZONTAL:


    For x = sStart To sFinish
        sForm.Line (x, 0)-(x + 1, sForm.ScaleHeight), RGB(tRed, tGreen, tBlue), BF
        GoSub sIncrement
    Next x

    Return
sVERTICAL:


    For y = sStart To sFinish
        sForm.Line (0, y)-(sForm.ScaleWidth, y + 1), RGB(tRed, tGreen, tBlue), BF
        GoSub sIncrement
    Next y

    Return
sIncrement:
    'Increment Red Color
    tRed = tRed + sCRedInc
    If tRed > 255 Then tRed = 255
    If tRed < 0 Then tRed = 0
    'Increment Green Color
    tGreen = tGreen + sCGreenInc
    If tGreen > 255 Then tGreen = 255
    If tGreen < 0 Then tGreen = 0
    'Increment Blue Color
    tBlue = tBlue + sCBlueInc
    If tBlue > 255 Then tBlue = 255
    If tBlue < 0 Then tBlue = 0
    Return
sIncrementCalc:
    'Calculate increment values
    sCRedInc = (sC2Red - sC1Red) / 100
    sCGreenInc = (sC2Green - sC1Green) / 100
    sCBlueInc = (sC2Blue - sC1Blue) / 100
    Return
End Sub

Private Sub PaintWallPaper()
Dim x As Long, y As Long
Dim picHeight As Long
Dim picWidth As Long
Dim SM As Long
    If oForm.Picture.Handle Then
        SM = oForm.ScaleMode          'save current value
        oForm.ScaleMode = 3           'pixel
        picHeight = oForm.ScaleY(oForm.Picture.Height)
        picWidth = oForm.ScaleX(oForm.Picture.Width)
        For x = 0 To oForm.ScaleWidth Step picWidth
            For y = 0 To oForm.ScaleHeight Step picHeight
                oForm.PaintPicture oForm.Picture, x, y
            Next y
        Next x
        oForm.ScaleMode = SM          'reset to previous value
    End If
End Sub
Private Sub PaintLeftPicture()
Dim y As Long
Dim picHeight As Long
Dim SM As Long
    If oForm.Picture.Handle Then
        SM = oForm.ScaleMode          'save current value
        oForm.ScaleMode = 3           'pixel
        picHeight = oForm.ScaleY(oForm.Picture.Height)
        For y = 0 To oForm.ScaleHeight Step picHeight
            oForm.PaintPicture oForm.Picture, 0, y
        Next y
        oForm.ScaleMode = SM          'reset to previous value
    End If
End Sub

Private Sub PaintCenterPicture()
Dim x As Single, y As Single
Dim picHeight As Long
Dim picWidth As Long
Dim SM As Long
    If pPicture.Handle Then
        SM = oForm.ScaleMode          'save current value
        oForm.ScaleMode = 3           'pixel
        If pPicture.Handle <> oForm.Picture.Handle And oForm.Picture.Handle Then
            Set pPicture = oForm.Picture
            oForm.Picture = LoadPicture()
        End If
        picHeight = oForm.ScaleY(pPicture.Height)
        picWidth = oForm.ScaleX(pPicture.Width)
        x = CSng((oForm.ScaleWidth - picWidth) / 2)
        y = CSng((oForm.ScaleHeight - picHeight) / 2)
        oForm.Cls
        oForm.PaintPicture pPicture, x, y
        oForm.ScaleMode = SM          'reset to previous value
    End If
End Sub

Private Sub SetupSysTrayIcon()
Dim hIcon As Long
    On Error GoTo vbErrorHandler
'
' Setup the System Tray Icon
'
    hIcon = SendMessageLong(Extender.Parent.hwnd, WM_GETICON, ICON_SMALL, 0)
    
    With tTrayStuff
        .cbSize = Len(tTrayStuff)
        .hwnd = UserControl.hwnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = IIf(hIcon > 0, hIcon, Extender.Parent.Icon)
        .szTip = Extender.Parent.Caption & vbNullChar
        Shell_NotifyIcon NIM_ADD, tTrayStuff
    End With
    bTrayFlag = True
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & "  " & Err.Description & " " & Err.Source & "::frmBrowser_SetupSysTrayIcon", , App.ProductName
End Sub

Private Sub KillSysTrayIcon()
    Dim t As NOTIFYICONDATA
'
' Kill the icon in the system tray
'
    On Error Resume Next
    With t
        .cbSize = Len(t)
        .hwnd = UserControl.hwnd
        .uId = 1&
    End With
    
    Shell_NotifyIcon NIM_DELETE, t
    bTrayFlag = False

End Sub

Public Sub SetTrayToolTip(ByVal sTip As String)
    If Not bTrayFlag Then Exit Sub
    If (sTip & Chr$(0) <> tTrayStuff.szTip) Then
        tTrayStuff.szTip = sTip & Chr$(0)
        tTrayStuff.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, tTrayStuff
    End If
End Sub

Public Sub SetTrayIcon(ByVal hIcon As Long)
    If Not bTrayFlag Then Exit Sub
    If (hIcon <> tTrayStuff.hIcon) Then
        tTrayStuff.hIcon = hIcon
        tTrayStuff.uFlags = NIF_ICON
        Shell_NotifyIcon NIM_MODIFY, tTrayStuff
    End If
End Sub

Public Sub Restore(Optional WindowState)
    On Error Resume Next
    If Not IsMissing(WindowState) Then lWindowState = WindowState
    If Extender.Parent.WindowState = vbMinimized Then
        Extender.Parent.WindowState = lWindowState
    End If
    Extender.Parent.ZOrder
    Extender.Parent.Show
    SetActiveWindow Extender.Parent.hwnd
    SetForegroundWindow Extender.Parent.hwnd
    Extender.Parent.ZOrder
End Sub

Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
Attribute WindowProc.VB_MemberFlags = "40"
'Window Messages
Const WM_GETMINMAXINFO = &H24
Const WM_SIZE = &H5
Const WM_PAINT = &HF
Const WM_ACTIVATEAPP = &H1C
Const WM_CLOSE = &H10
Const WM_ENTERSIZEMOVE = &H231

Dim mmi As MINMAXINFO
    On Error Resume Next
    If uMsg = WM_GETMINMAXINFO Then
        CopyMemory mmi, ByVal lParam, LenB(mmi)
        If m_MinWidth > 0 Then mmi.ptMinTrackSize.x = m_MinWidth / Screen.TwipsPerPixelX
        If m_MinHeight > 0 Then mmi.ptMinTrackSize.y = m_MinHeight / Screen.TwipsPerPixelY
        If m_MaxWidth > 0 Then mmi.ptMaxTrackSize.x = m_MaxWidth / Screen.TwipsPerPixelX
        If m_MaxHeight > 0 Then mmi.ptMaxTrackSize.y = m_MaxHeight / Screen.TwipsPerPixelX
        CopyMemory ByVal lParam, mmi, LenB(mmi)
        WindowProc = False
        Exit Function
    ElseIf uMsg = WM_SIZE Then
        If m_ResizeControls Then oElastic.FormResize m_ResizeFonts
        If m_MinimizeToTray And Extender.Parent.WindowState = vbMinimized Then
            Extender.Parent.Hide
            SetupSysTrayIcon
        ElseIf bTrayFlag Then
            KillSysTrayIcon
        End If
        Debug.Print Extender.Parent.WindowState
        If Extender.Parent.WindowState <> vbMinimized Then
            lWindowState = Extender.Parent.WindowState
            If m_Background = bgCenterPicture Then
                PaintCenterPicture
            ElseIf m_Background = bgGradient Then
                Shade Extender.Parent, m_Gradient, m_BackColor, m_BackColor2
            End If
        End If
    ElseIf uMsg = WM_PAINT Then
        Select Case m_Background
        Case bgTiledPicture
            PaintWallPaper
        Case bgLeftPicture
            PaintLeftPicture
        Case bgCenterPicture
            PaintCenterPicture
        End Select
    End If
    WindowProc = True
End Function

