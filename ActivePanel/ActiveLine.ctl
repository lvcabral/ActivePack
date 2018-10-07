VERSION 5.00
Begin VB.UserControl ActiveLine 
   CanGetFocus     =   0   'False
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   2505
   ToolboxBitmap   =   "ActiveLine.ctx":0000
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "ActiveLine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpLine 
      BorderColor     =   &H80000014&
      Height          =   15
      Index           =   1
      Left            =   825
      Top             =   120
      Width           =   1635
   End
   Begin VB.Shape shpLine 
      BorderColor     =   &H80000010&
      Height          =   15
      Index           =   0
      Left            =   825
      Top             =   105
      Width           =   1635
   End
End
Attribute VB_Name = "ActiveLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Property Variables:
Dim m_Alignment As AlignmentConstants

Public Sub About()
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    shpLine(0).Height = Screen.TwipsPerPixelY
    shpLine(1).Height = Screen.TwipsPerPixelY
    UserControl.BackColor = Extender.Container.BackColor
    lblCaption.Caption = Extender.Name
    m_Alignment = vbLeftJustify
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    shpLine(0).Height = Screen.TwipsPerPixelY
    shpLine(1).Height = Screen.TwipsPerPixelY
    lblCaption.Caption = PropBag.ReadProperty("Caption", "")
    m_Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    lblCaption.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    UserControl.BackColor = lblCaption.BackColor
    
    Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next
    PropBag.WriteProperty "Caption", lblCaption.Caption, ""
    PropBag.WriteProperty "Alignment", m_Alignment, vbLeftJustify
    PropBag.WriteProperty "Font", lblCaption.Font, Ambient.Font
    PropBag.WriteProperty "ForeColor", lblCaption.ForeColor, Ambient.ForeColor
    PropBag.WriteProperty "BackColor", UserControl.BackColor, Ambient.BackColor
    
End Sub

Private Sub UserControl_Resize()
Dim PixelX As Long, PixelY As Long
On Error Resume Next
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    If lblCaption.Caption = "" Then
        If Height <> (PixelY * 8) Then 'Evita Stack Overflow
            Height = (PixelY * 8)
        End If
        shpLine(0).Top = PixelY * 3
        shpLine(1).Top = PixelY * 4
        shpLine(0).Left = 0
        shpLine(1).Left = 0
        shpLine(0).Width = Width
        shpLine(1).Width = Width
        lblCaption.Visible = False
    Else
        'Altura e Posição Y
        lblCaption.Visible = True
        If Height <> lblCaption.Height Then 'Evita Stack Overflow
            Height = lblCaption.Height
        End If
        shpLine(0).Top = (Height - (PixelY * 2)) / 2
        shpLine(1).Top = shpLine(0).Top + PixelY
        'Largura e Posição X
        Select Case m_Alignment
        Case vbLeftJustify
            shpLine(0).Left = lblCaption.Width + (PixelX * 2)
            shpLine(1).Left = lblCaption.Width + (PixelX * 2)
            'Largura
            lblCaption.Left = 0
            If Width < lblCaption.Width Then
                Width = lblCaption.Width
            ElseIf Width > lblCaption.Width + (PixelX * 2) Then
                shpLine(0).Width = Width - shpLine(0).Left
                shpLine(1).Width = Width - shpLine(1).Left
            Else
                shpLine(0).Width = 0
                shpLine(1).Width = 0
            End If
        Case vbRightJustify
            shpLine(0).Left = 0
            shpLine(1).Left = 0
            lblCaption.Left = Width - lblCaption.Width
            'Largura
            If Width < lblCaption.Width Then
                shpLine(0).Width = 0
                shpLine(1).Width = 0
                Width = lblCaption.Width
            ElseIf Width > lblCaption.Width + (PixelX * 2) Then
                shpLine(0).Width = Width - (lblCaption.Width + (PixelX * 2))
                shpLine(1).Width = Width - (lblCaption.Width + (PixelX * 2))
            Else
                shpLine(0).Width = 0
                shpLine(1).Width = 0
            End If
        Case vbCenter
            shpLine(0).Left = 0
            shpLine(1).Left = 0
            lblCaption.Left = (Width - lblCaption.Width) / 2
            'Largura
            If Width < lblCaption.Width Then
                shpLine(0).Width = 0
                shpLine(1).Width = 0
                Width = lblCaption.Width
            ElseIf Width > lblCaption.Width + (Screen.TwipsPerPixelX * 2) Then
                shpLine(0).Width = Width
                shpLine(1).Width = Width
            Else
                shpLine(0).Width = 0
                shpLine(1).Width = 0
            End If
        End Select
    End If
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"

    Caption = lblCaption.Caption
    
End Property

Public Property Let Caption(ByVal NewCaption As String)

    lblCaption.Caption = NewCaption
    UserControl_Resize
    PropertyChanged "Caption"
    
End Property

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    UserControl_Resize
    PropertyChanged "Alignment"
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512

    Set Font = lblCaption.Font
    
End Property

Public Property Set Font(ByVal NewFont As StdFont)

    Set lblCaption.Font = NewFont
    UserControl_Resize
    PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513

    ForeColor = lblCaption.ForeColor
    
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)

    lblCaption.ForeColor = NewColor
    PropertyChanged "ForeColor"
    
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501

    BackColor = UserControl.BackColor
    
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)

    UserControl.BackColor = NewColor
    lblCaption.BackColor = NewColor
    PropertyChanged "BackColor"
    
End Property
