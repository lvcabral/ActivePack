VERSION 5.00
Begin VB.UserControl ActiveStatus 
   Alignable       =   -1  'True
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   ScaleHeight     =   900
   ScaleWidth      =   5505
   ToolboxBitmap   =   "ActiveStatus.ctx":0000
   Begin VB.PictureBox picStatus 
      Height          =   375
      Left            =   570
      ScaleHeight     =   315
      ScaleWidth      =   4275
      TabIndex        =   0
      Top             =   255
      Width           =   4335
   End
End
Attribute VB_Name = "ActiveStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum sbStyleConstants
    sbrNormal = 0
    sbrSimple = 1
End Enum

Public Enum sbImageSizeConstants
    [16x16] = 16
    [24x24] = 24
    [32x32] = 32
    [48x48] = 48
End Enum

'Event Declarations:
Event Click() 'MappingInfo=picStatus,picStatus,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=picStatus,picStatus,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picStatus,picStatus,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picStatus,picStatus,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picStatus,picStatus,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Default Property Values:
Const m_def_Style = sbrNormal
Const m_def_SimpleText = ""
Const m_def_ImageSize = [16x16]

'Property Variables:
Dim m_Font As Font
Dim m_SimpleText As String
Dim mvarcPanels As cPanels

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub UserControl_Initialize()
    Set mStatus = New cStatusBar
End Sub

Private Sub UserControl_InitProperties()
    With mStatus
        .Create picStatus
        .lIconSize = 16
        picStatus.BorderStyle = 0
        picStatus.ZOrder
        Panels.Add , , "RainDrops ActiveStatus"
        Panels(1).AutoSize = sbrSpring
        Panels.Add
        Panels(2).Bevel = sbrNoBevel
        .SimpleMode = False
        .SizeGrip = True
        Set .Font = Ambient.Font
        Extender.Align = vbAlignBottom
    End With
    Set m_Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With mStatus
        .Create picStatus
        picStatus.BorderStyle = 0
        picStatus.ZOrder
        If Not Ambient.UserMode Then
            Panels.Add , , "RainDrops ActiveStatus"
            Panels(1).AutoSize = sbrSpring
            Panels.Add
            Panels(2).Bevel = sbrNoBevel
        End If
        .lIconSize = PropBag.ReadProperty("ImageSize", m_def_ImageSize)
        .SimpleMode = -(PropBag.ReadProperty("Style", m_def_Style))
        .SimpleText = PropBag.ReadProperty("SimpleText", m_def_SimpleText)
        Set .Font = PropBag.ReadProperty("Font", Ambient.Font)
    End With
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    
End Sub

Private Sub UserControl_Resize()
    If picStatus.Visible Then
        picStatus.Move 0, 0, Width, Height
        If Extender.Align = vbAlignBottom Then
            On Error Resume Next
            mStatus.SizeGrip = (Extender.Container.WindowState = vbNormal)
        Else
            mStatus.SizeGrip = False
        End If
    End If
End Sub

Private Sub UserControl_Show()
    picStatus.Move 0, 0, Width, Height
    mStatus.SizeGrip = (Extender.Align = vbAlignBottom)
End Sub

Private Sub UserControl_Terminate()
    Set mvarcPanels = Nothing
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Private Sub picStatus_Click()
    RaiseEvent Click
End Sub

Private Sub picStatus_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub picStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub picStatus_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", picStatus.Font, Ambient.Font)
    Call PropBag.WriteProperty("Style", Abs(mStatus.SimpleMode), m_def_Style)
    Call PropBag.WriteProperty("ImageSize", mStatus.lIconSize, m_def_ImageSize)
    Call PropBag.WriteProperty("SimpleText", mStatus.SimpleText, m_def_SimpleText)
End Sub

Private Sub picStatus_Paint()
    mStatus.Draw
End Sub

Private Sub picStatus_Resize()
    mStatus.Draw
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = picStatus.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picStatus.Font = New_Font
    mStatus.Draw
    PropertyChanged "Font"
End Property

Public Property Let ImageList(vThis As Variant)
    mStatus.ImageList = vThis
End Property

Public Property Get ImageSize() As sbImageSizeConstants
    ImageSize = mStatus.lIconSize
End Property

Public Property Let ImageSize(ByVal mImageSize As sbImageSizeConstants)
    mStatus.lIconSize = mImageSize
    PropertyChanged "ImageSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Style() As sbStyleConstants
    Style = Abs(mStatus.SimpleMode)
End Property

Public Property Let Style(ByVal New_Style As sbStyleConstants)
    mStatus.SimpleMode = -(New_Style)
    PropertyChanged "Style"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SimpleText() As String
    SimpleText = mStatus.SimpleText
End Property

Public Property Let SimpleText(ByVal New_SimpleText As String)
    mStatus.SimpleText = New_SimpleText
    PropertyChanged "SimpleText"
End Property

Public Property Get Panels() As cPanels
    If mvarcPanels Is Nothing Then
        Set mvarcPanels = New cPanels
    End If
    Set Panels = mvarcPanels
End Property
