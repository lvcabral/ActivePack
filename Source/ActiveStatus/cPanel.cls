VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum sbPanelBevelConstants
   sbrNoBevel
   sbrInset
   sbrRaised
End Enum

Public Enum sbPanelAlignmentConstants
    sbrLeft
    sbrCenter
    sbrRight
End Enum

Public Enum sbPanelStyleConstants
    sbrText
    sbrCaps
    sbrNum
    sbrIns
    sbrScrl
    sbrTime
    sbrDate
End Enum

Public Enum sbAutoSizeConstants
    sbrNoAutoSize
    sbrSpring
    sbrContents
End Enum

'local variable(s) to hold property value(s)
Private mvarTag  As String 'local copy
Private mvarBevel As sbPanelBevelConstants  'local copy
Private mvarStyle As sbPanelStyleConstants  'local copy
Private mvarIconIndex As Integer 'local copy
Private mvarAlignment As sbPanelAlignmentConstants 'local copy
Private mvarAutoSize As sbAutoSizeConstants 'local copy
Private mvarIndex As Long 'local copy
Private mvarMinWidth As Long 'local copy
Private mvarWidth As Long 'local copy

Public Property Let Alignment(ByVal vData As sbPanelAlignmentConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Alignment = 5
    mvarAlignment = vData
    mStatus.PanelAlignment(mvarIndex) = vData
End Property


Public Property Get Alignment() As sbPanelAlignmentConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property

Public Property Let AutoSize(ByVal vData As sbAutoSizeConstants)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoSize = 5
    mvarAutoSize = vData
    Select Case vData
    Case sbrNoAutoSize
         mStatus.PanelFitToContents(mvarIndex) = False
         mStatus.PanelSpring(mvarIndex) = False
    Case sbrSpring
         mStatus.PanelFitToContents(mvarIndex) = False
         mStatus.PanelSpring(mvarIndex) = True
    Case sbrContents
         mStatus.PanelFitToContents(mvarIndex) = True
         mStatus.PanelSpring(mvarIndex) = False
    End Select

End Property

Public Property Get AutoSize() As sbAutoSizeConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoSize
    AutoSize = mvarAutoSize
End Property

Public Property Let IconIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconIndex = 5
    mvarIconIndex = vData
    mStatus.PanelIcon(mvarIndex) = vData
End Property


Public Property Get IconIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconIndex
    IconIndex = mvarIconIndex
End Property


Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Text = 5
    mStatus.PanelText(mvarIndex) = vData
End Property


Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Text
    Text = mStatus.PanelText(mvarIndex)
End Property

Public Property Let ToolTipText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToolTipText = 5
    mStatus.PanelToolTipText(mvarIndex) = vData
End Property


Public Property Get ToolTipText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToolTipText
    Text = mStatus.PanelToolTipText(mvarIndex)
End Property

Public Property Let Bevel(ByVal vData As sbPanelBevelConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Bevel = Form1
    mvarBevel = vData
    Select Case vData
    Case sbrInset
        mStatus.PanelStyle(mvarIndex) = estbrStandard
    Case sbrNoBevel
        mStatus.PanelStyle(mvarIndex) = estbrNoBorders
    Case sbrRaised
        mStatus.PanelStyle(mvarIndex) = estbrRaisedBorder
    End Select
End Property


Public Property Get Bevel() As sbPanelBevelConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bevel
    Bevel = mvarBevel
End Property

Public Property Let Style(ByVal vData As sbPanelStyleConstants)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Style = Form1
    mvarStyle = vData
End Property


Public Property Get Style() As sbPanelStyleConstants
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Style
    Style = mvarStyle
End Property

Public Property Let Index(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
End Property

Public Property Get Index() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
End Property

Public Property Let Tag(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property

Public Property Let MinWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinWidth= 5
    mvarMinWidth = vData
    mStatus.PanelMinWidth(mvarIndex) = vData
End Property

Public Property Get MinWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinWidth
    MinWidth = mvarMinWidth
End Property

Public Property Let Picture(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Picture= 5
Dim pic As frmPicCache
    If TypeName(vData) = "Picture" Then
        Set pic = New frmPicCache
        pic.Picture = vData
        mStatus.PanelhIcon(mvarIndex) = pic.Picture.Handle
        Set pic = Nothing
    End If
End Property

Public Property Get Picture() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Width
    Picture.Handle = mStatus.PanelhIcon(mvarIndex)
End Property

Public Property Let Width(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Width= 5
    mvarWidth = vData
    mStatus.PanelIdealWidth(mvarIndex) = vData
End Property

Public Property Get Width() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Width
    Width = mvarWidth
End Property


