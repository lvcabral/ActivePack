VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************
' Copyright ©Karl E. Peterson, 1997
' All Rights Reserved.
' ********************************************
' Written for publication in January 1998,
' Visual Basic Programmer's Journal.
' May be freely used in your applications.
' Redistribution of source prohibited.
' ********************************************
Option Explicit

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Const ICC_ANIMATE_CLASS = &H80
Private Const ANIMATE_CLASS = "SysAnimate32"

Private Const ACS_CENTER = &H1&
Private Const ACS_TRANSPARENT = &H2&
Private Const ACS_AUTOPLAY = &H4&
Private Const ACS_TIMER = &H8&

Private Const WM_PAINT = &HF
Private Const WM_USER = &H400&
Private Const ACM_OPEN = WM_USER + 100
Private Const ACM_PLAY = WM_USER + 101
Private Const ACM_STOP = WM_USER + 102

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_BORDER = &H800000
Private Const WS_CLIPSIBLINGS = &H4000000

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private m_hWnd As Long
Private m_hWndParent As Long
Private m_AutoPlay As Boolean
Private m_Center As Boolean
Private m_Transparent As Boolean
Private m_Visible As Boolean
Private m_Playing As Boolean
Private m_AniResID As Long
Private m_AniFile As String
Private m_Left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long

' ***********************************************
'  Initialization and Termination
' ***********************************************
Private Sub Class_Initialize()
   Dim iccex As tagInitCommonControlsEx
   '
   ' Initialize common controls DLL
   '
   With iccex
      .lngSize = LenB(iccex)
      .lngICC = ICC_ANIMATE_CLASS
   End With
   Call InitCommonControlsEx(iccex)
   '
   ' Initialize member variables
   '
   m_AniResID = 0
   m_AniFile = ""
   m_AutoPlay = True
   m_Center = True
   m_Playing = False
   m_Transparent = True
   m_Visible = True
   m_Left = 0     'arbitrary
   m_Top = 0      'arbitrary
   m_Width = 0    'arbitrary
   m_Height = 0   'arbitrary
End Sub

Private Sub Class_Terminate()
   If m_hWnd Then
      Call AniDestroy
   End If
End Sub

' ***********************************************
'  Public Properties -- ReadOnly!
' ***********************************************
Public Property Get hWnd() As Long
   hWnd = m_hWnd
End Property

' ***********************************************
'  Public Properties -- General
' ***********************************************
Public Property Let AutoPlay(ByVal NewVal As Boolean)
   m_AutoPlay = NewVal
   If m_hWnd Then Call AniCreate
End Property

Public Property Get AutoPlay() As Boolean
   AutoPlay = m_AutoPlay
End Property

Public Property Let Center(ByVal NewVal As Boolean)
   m_Center = NewVal
   If m_hWnd Then Call AniCreate
End Property

Public Property Get Center() As Boolean
   Center = m_Center
End Property

Public Property Let FileName(ByVal NewVal As String)
   m_AniFile = NewVal
   m_AniResID = 0
   Call OpenAnimation
End Property

Public Property Get FileName() As String
   FileName = m_AniFile
End Property

Public Property Let ResourceID(ByVal NewVal As Long)
   m_AniResID = NewVal
   m_AniFile = ""
   Call OpenAnimation
End Property

Public Property Get ResourceID() As Long
   ResourceID = m_AniResID
End Property

Public Property Let Parent(ByVal NewVal As Long)
   m_hWndParent = NewVal
   If m_hWnd Then Call AniCreate
End Property

Public Property Get Parent() As Long
   Parent = m_hWndParent
End Property

Public Property Let Transparent(ByVal NewVal As Boolean)
   m_Transparent = NewVal
   If m_hWnd Then Call AniCreate
End Property

Public Property Get Transparent() As Boolean
   Transparent = m_Transparent
End Property

' ***********************************************
'  Public Properties -- Position
' ***********************************************
Public Property Let Height(ByVal NewVal As Long)
   Me.Move Height:=NewVal
End Property

Public Property Get Height() As Long
   Height = m_Height
End Property

Public Property Get Left() As Long
   Left = m_Left
End Property

Public Property Let Left(ByVal NewVal As Long)
   Me.Move Left:=NewVal
End Property

Public Property Let Top(ByVal NewVal As Long)
   Me.Move Top:=NewVal
End Property

Public Property Get Top() As Long
   Top = m_Top
End Property

Public Property Let Width(ByVal NewVal As Long)
   Me.Move Width:=NewVal
End Property

Public Property Get Width() As Long
   Width = m_Width
End Property

Public Property Let Visible(ByVal NewVal As Boolean)
   m_Visible = NewVal
   If m_hWnd Then
      If m_Visible Then
         Call ShowWindow(m_hWnd, SW_SHOW)
      Else
         Call ShowWindow(m_hWnd, SW_HIDE)
      End If
   End If
End Property

Public Property Get Visible() As Boolean
   Visible = m_Visible
End Property

' ***********************************************
'  Public Methods
' ***********************************************
Public Sub AniPlay()
   If m_hWnd = 0 Then Call AniCreate
   '
   ' wParam = (WPARAM) (UINT) cRepeat;        <-- (-1)=indefinitely
   ' lParam = (LPARAM) MAKELONG(wFrom, wTo);  <-- frame numbers
   m_Playing = CBool(SendMessage(m_hWnd, ACM_PLAY, -1&, ByVal &HFFFF0000))
End Sub

Public Sub AniStop()
   Dim nRet As Long
   If m_hWnd Then
      nRet = SendMessage(m_hWnd, ACM_STOP, 0&, ByVal 0&)
      If nRet Then m_Playing = False
   End If
End Sub

Public Sub Move(Optional Left, Optional Top, Optional Width, Optional Height)
   If Not IsMissing(Left) Then m_Left = CLng(Left)
   If Not IsMissing(Top) Then m_Top = CLng(Top)
   If Not IsMissing(Width) Then m_Width = CLng(Width)
   If Not IsMissing(Height) Then m_Height = CLng(Height)
   If m_hWnd Then
      Call MoveWindow(m_hWnd, m_Left, m_Top, m_Width, m_Height, True)
   End If
End Sub

' ***********************************************
'  Private Methods
' ***********************************************
Private Sub AniCreate()
   Dim AniStyle As Long
   Dim WasPlaying As Boolean
   '
   ' Make sure we don't already have one
   '
   If m_hWnd Then
      WasPlaying = m_Playing
      Call AniDestroy
   End If
   '
   ' Combine style bits
   '
   AniStyle = WS_CHILD Or WS_CLIPSIBLINGS
   If m_Visible Then AniStyle = AniStyle Or WS_VISIBLE
   If m_AutoPlay Then AniStyle = AniStyle Or ACS_AUTOPLAY
   If m_Center Then AniStyle = AniStyle Or ACS_CENTER
   If m_Transparent Then AniStyle = AniStyle Or ACS_TRANSPARENT
   '
   ' Create animation window
   '
   m_hWnd = CreateWindowEx( _
      0, _
      ANIMATE_CLASS, _
      vbNullString, _
      AniStyle, _
      m_Left, m_Top, m_Width, m_Height, _
      m_hWndParent, 0&, App.hInstance, ByVal 0&)
   '
   ' Restart animation if was playing before window recreation
   '
   If WasPlaying Then
      Call OpenAnimation
      Call AniPlay
   End If
   '
   ' Refresh window
   '
   If m_Visible Then
      Call SendMessage(m_hWndParent, WM_PAINT, 0&, ByVal 0&)
   End If
End Sub

Private Sub OpenAnimation()
   Dim nRet As Long
   '
   ' Make sure we have a window to work with, or
   ' Stop any existing animations.
   '
   If m_hWnd = 0 Then
      Call AniCreate
   Else
      Call AniStop
   End If
   '
   ' Load animation.
   '
   If m_hWnd Then
      If m_AniResID Then
         nRet = SendMessage(m_hWnd, ACM_OPEN, 0&, ByVal m_AniResID)
      ElseIf Len(m_AniFile) Then
         nRet = SendMessage(m_hWnd, ACM_OPEN, 0&, ByVal m_AniFile)
      End If
      If m_AutoPlay Then
         m_Playing = CBool(nRet)
      Else
         m_Playing = False
      End If
   End If
End Sub

Private Sub AniDestroy()
   If m_hWnd Then
      Call AniStop
      DestroyWindow m_hWnd
   End If
End Sub

