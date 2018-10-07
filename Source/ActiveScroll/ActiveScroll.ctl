VERSION 5.00
Begin VB.UserControl ActiveScroll 
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   540
   ToolboxBitmap   =   "ActiveScroll.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "ActiveScroll.ctx":00FA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ActiveScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oScroll As cScrollBars
Private bCreated As Boolean
Private WithEvents frmContainer As Form
Attribute frmContainer.VB_VarHelpID = -1
'Public Enums
Public Enum EFSStyleConstants
    efsRegular = 0
    efsEncarta = 1
    efsFlat = 2
End Enum
'Default Property Values:
Const m_def_Style = efsRegular
'Property Variables:
Private m_Style As EFSStyleConstants

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" Then
        UserControl.BackColor = Ambient.BackColor
    End If
End Sub

Private Sub UserControl_Initialize()
    'Configura as Barras de Rolagem
    Set oScroll = New cScrollBars
    bCreated = False
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Style = m_def_Style
    UserControl.BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If Not Ambient.UserMode Then
        imgIcon.Visible = True
        UserControl.Width = imgIcon.Width + 60
        UserControl.Height = imgIcon.Height + 60
        UserControl.BorderStyle = vbFixedSingle
    Else
        UserControl.BorderStyle = 0
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        Set frmContainer = Extender.Parent
    End If
    UserControl.BackColor = Ambient.BackColor
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
End Sub

Private Sub UserControl_Terminate()
    'Destroi as Barras de Rolagem
    Set oScroll = Nothing
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Style() As EFSStyleConstants
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As EFSStyleConstants)
    m_Style = New_Style
    oScroll.Style = m_Style
    PropertyChanged "Style"
End Property

Private Sub frmContainer_Resize()
    If Ambient.UserMode And Not bCreated Then
        oScroll.Create Extender.Parent, Extender
        oScroll.Style = m_Style
        bCreated = True
    Else
        oScroll.Resize
    End If
End Sub

