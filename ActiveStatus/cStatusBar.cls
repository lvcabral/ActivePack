VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cPanels"
Option Explicit

' =========================================================================
' Declares, constants and types required for fake status bar:
' =========================================================================
Private Type POINTAPI
   x As Long
   Y As Long
End Type
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function DrawStatusText Lib "COMCTL32" Alias "DrawStatusTextA" (ByVal hdc As Long, lprc As RECT, ByVal pszText As String, ByVal uFlags As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_BOTTOM = &H8
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Const SBT_NOBORDERS = &H100
Private Const SBT_POPOUT = &H200
Private Const SBT_RTLREADING = &H400
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNHIGHLIGHT = 20

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Const WM_USER = &H400
Private Const SB_SETTIPTEXT = (WM_USER + 16)
Private Const SB_GETTIPTEXT = (WM_USER + 18)

' =========================================================================
' Implementation of fake status bar:
' =========================================================================
Public Enum ENSBRPanelStyleConstants
   estbrStandard = &H0&
   estbrNoBorders = &H100
   estbrRaisedBorder = &H200
End Enum

Private Type tStatusPanel
   lID As Long
   sKey As String
   lItemData As Long
   iImgIndex As Long
   hIcon As Long
   sText As String
   sToolTipText As String
   lAlignment As Long
   lMinWidth As Long
   lIdealWidth As Long
   lSetWidth As Long
   bSpring As Boolean
   bFit As Boolean
   eType As sbPanelStyleConstants
   eStyle As ENSBRPanelStyleConstants
   bState As Boolean
   tR As RECT
End Type
Private m_tPanels() As tStatusPanel
Private m_iPanelCount As Long
Private m_bSizeGrip As Boolean
Private m_hIml As Long
Private m_pic As PictureBox
Private m_obj As Object
Private m_lLeft As Long, m_lTop As Long
Private m_lHeight As Long
Private m_bSimpleMode As Boolean
Private m_sSimpleText As String
Private m_lPanelAlign As Long

Private WithEvents m_Timer As CTimer
Attribute m_Timer.VB_VarHelpID = -1

Public lIconSize As Long


Public Property Get SimpleMode() As Boolean
   SimpleMode = m_bSimpleMode
End Property
Public Property Let SimpleMode(ByVal bSimple As Boolean)
   m_bSimpleMode = bSimple
   Draw
End Property
Public Property Get SimpleText() As String
   SimpleText = m_sSimpleText
End Property
Public Property Let SimpleText(ByVal sText As String)
   m_sSimpleText = sText
   If (m_bSimpleMode) Then
      Draw
   End If
End Property

Public Property Let SizeGrip(ByVal bSizeGrip As Boolean)
   m_bSizeGrip = bSizeGrip
   Draw
End Property
Public Property Get SizeGrip() As Boolean
   SizeGrip = m_bSizeGrip
End Property

Public Function AddPanel( _
      Optional ByVal eStyle As ENSBRPanelStyleConstants = estbrStandard, _
      Optional ByVal sText As String = "", _
      Optional ByVal iImgIndex As Long = -1, _
      Optional ByVal lMinWidth As Long = 64, _
      Optional ByVal bSpring As Boolean = False, _
      Optional ByVal bFitContents As Boolean = False, _
      Optional ByVal lItemData As Long = 0, _
      Optional ByVal sKey As String = "", _
      Optional ByVal vKeyBefore As Variant, _
      Optional ByVal eType As sbPanelStyleConstants = sbrText _
   ) As Long
Dim iIndex As Long
Dim i As Long
Dim bTimer As Boolean
Dim bEnabled As Boolean
Dim tR As RECT
   
   If (m_iPanelCount >= &HFF) Then
      Err.Raise vbObjectError + 1051, App.EXEName & ".cStatusBar", "Too many panels."
      Exit Function
   End If
      
   If Not IsMissing(vKeyBefore) Then
      ' Determine if vKeyBefore is valid:
      iIndex = PanelIndex(vKeyBefore)
      If (iIndex > 0) Then
         ' ok. Insert a space:
         m_iPanelCount = m_iPanelCount + 1
         ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
         For i = m_iPanelCount To iIndex + 1 Step -1
            LSet m_tPanels(i) = m_tPanels(i - 1)
         Next i
         m_tPanels(iIndex).hIcon = 0
      Else
         ' Failed
         Exit Function
      End If
   Else
      ' Insert a space at the end:
      m_iPanelCount = m_iPanelCount + 1
      ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
      iIndex = m_iPanelCount
   End If
   
   ' Set up the info:
   If (bSpring) Then
      For i = 1 To m_iPanelCount
         If (i <> iIndex) Then
            m_tPanels(i).bSpring = False
         End If
      Next i
   End If
   
   bTimer = (eType <> sbrText)
   
   With m_tPanels(iIndex)
      .bFit = True
      .bSpring = False
      .eStyle = eStyle
      .eType = eType
      .lItemData = lItemData
      .sKey = sKey
      Select Case eType
      Case sbrCaps
         .sText = "CAPS"
         .lAlignment = 1 'Center
      Case sbrDate
         .sText = Format$(Now, "Short Date")
         .lAlignment = 1 'Center
      Case sbrIns
         .sText = "INS"
         .lAlignment = 1 'Center
      Case sbrNum
         .sText = "NUM"
         .lAlignment = 1 'Center
      Case sbrScrl
         .sText = "SCRL"
         .lAlignment = 1 'Center
      Case sbrText
         .sText = sText
         .bFit = bFitContents
         .bSpring = bSpring
         .iImgIndex = iImgIndex
         .lMinWidth = lMinWidth
      Case sbrTime
         .sText = Format$(Now, "hh:nn")
         .lAlignment = 1 'Center
      End Select
   End With
   
   ' Add the information to the status bar:
   pEvaluateIdealSize iIndex
   pResizeStatus
   
   ' Now ensure the text, style, tooltip and icon are actually correct:
   PanelText(iIndex) = m_tPanels(iIndex).sText
   PanelIcon(iIndex) = m_tPanels(iIndex).iImgIndex
      
   Draw
   
   If bTimer Then m_Timer.Interval = 250
   
   AddPanel = iIndex
End Function

Public Sub Draw()
Dim i As Long, iEnd As Long
Dim lhDC As Long
Dim lX As Long, lY As Long
Dim hBr As Long, tR As RECT, tOR As RECT, tBR As RECT
Dim fntThis As StdFont
Dim bEnd As Boolean
   
   GetClientRect m_obj.hwnd, tR
   hBr = GetSysColorBrush(COLOR_BTNFACE)
   FillRect m_obj.hdc, tR, hBr
   DeleteObject hBr

   LSet tOR = tR

   pResizeStatus
   lhDC = m_obj.hdc
   If (m_bSimpleMode) Then
      DrawStatusText lhDC, tR, "", 0
      DrawText lhDC, " " & m_sSimpleText, Len(" " & m_sSimpleText), tR, DT_VCENTER Or DT_SINGLELINE
   Else
      For i = 1 To m_iPanelCount
         With m_tPanels(i)
            LSet tBR = .tR
            If (m_bSimpleMode) Then
               tBR.Right = tOR.Right
            Else
               If (tBR.Right > tOR.Right) Then
                  tBR.Right = tOR.Right - 1
                  bEnd = True
               End If
            End If
            If (.hIcon <> 0) Then
               DrawStatusText lhDC, tBR, "", .eStyle
               ' Draw the icon:
               lY = tBR.Top + 1 + (tBR.Bottom - tBR.Top - 2 - lIconSize) \ 2
               lX = tBR.Left + 2
               DrawIconEx lhDC, lX, lY, .hIcon, lIconSize, lIconSize, 0, 0, DI_NORMAL
               ' Draw the text:
               If (Len(.sText) > 0) Then
                  tBR.Left = tBR.Left + lIconSize + 4
                  tBR.Right = tBR.Right - 4
                  DrawText lhDC, .sText, Len(.sText), tBR, DT_VCENTER Or DT_SINGLELINE Or .lAlignment
               End If
            Else
               DrawStatusText lhDC, tBR, "", .eStyle
               tBR.Left = tBR.Left + 4
               tBR.Right = tBR.Right - 4
                Select Case .eType
                Case sbrDate
                    .sText = Format$(Now, "Short Date")
                Case sbrTime
                    .sText = Format$(Now, "hh:nn")
                Case sbrIns
                    .sText = "INS"
                    If GetInsertKey() Then
                       SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                    Else
                       SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                    End If
                Case sbrCaps
                    .sText = "CAPS"
                    If GetCapslock() Then
                       SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                    Else
                       SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                    End If
                Case sbrNum
                    .sText = "NUM"
                    If GetNumlock() Then
                       SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                    Else
                       SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                    End If
                Case sbrScrl
                    .sText = "SCRL"
                    If GetScrollLock() Then
                       SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                    Else
                       SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                    End If
                End Select
                DrawText lhDC, .sText, Len(.sText), tBR, DT_VCENTER Or DT_SINGLELINE Or .lAlignment
                SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
            End If
            If bEnd Then
               Exit For
            End If
         End With
      Next i
   
   End If
   
   If (m_bSizeGrip) Then
      Set fntThis = New StdFont
      With fntThis
         .Name = m_obj.Font.Name
         .Size = m_obj.Font.Size
         .Bold = m_obj.Font.Bold
         .Italic = m_obj.Font.Italic
         .Underline = m_obj.Font.Underline
      End With
      m_obj.Font.Name = "Marlett"
      m_obj.Font.Size = fntThis.Size * 4 / 3
      m_obj.ForeColor = vb3DHighlight
      OffsetRect tOR, -2, -1
      DrawText lhDC, "o", 1, tOR, DT_BOTTOM Or DT_RIGHT Or DT_SINGLELINE
      m_obj.ForeColor = vbButtonShadow
      'OffsetRect tOR, 1, 0
      DrawText lhDC, "p", 1, tOR, DT_BOTTOM Or DT_RIGHT Or DT_SINGLELINE
      Set m_obj.Font = fntThis
      m_obj.ForeColor = vbWindowText
   End If
End Sub

Public Function RemovePanel( _
      ByVal vKey As Variant _
   )
Dim iIndex As Long
Dim i As Long
   iIndex = PanelIndex(vKey)
   If (iIndex > 0) Then
      If (m_tPanels(iIndex).hIcon <> 0) Then
         DestroyIcon m_tPanels(iIndex).hIcon
      End If
      For i = iIndex To m_iPanelCount - 1
         LSet m_tPanels(i) = m_tPanels(i + 1)
      Next i
      m_iPanelCount = m_iPanelCount - 1
      If (m_iPanelCount > 0) Then
         ReDim Preserve m_tPanels(1 To m_iPanelCount) As tStatusPanel
      End If
      Draw
   End If
End Function


Public Sub SetLeftTopOffsets(ByVal lLeft As Long, ByVal lTop As Long)
   m_lLeft = lLeft
   m_lTop = lTop
End Sub

Public Property Let ImageList(vThis As Variant)
Dim cy As Long, lR As Long
    
    ' Set the ImageList handle property either from a VB
    ' image list or directly:
    m_hIml = 0
    If TypeName(vThis) = "ImageList" Then
        ' VB ImageList control.  Note that unless
        ' some call has been made to an object within a
        ' VB ImageList the image list itself is not
        ' created.  Therefore hImageList returns error. So
        ' ensure that the ImageList has been initialised by
        ' drawing into nowhere:
        On Error Resume Next
        ' Get the image list initialised..
        vThis.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vThis.hImageList
        If (Err.Number <> 0) Then
            ' No images.
            m_hIml = 0
        Else
            ' Get the icon size:
            lR = ImageList_GetIconSize(m_hIml, lIconSize, cy)
        End If
        On Error GoTo 0
    ElseIf VarType(vThis) = vbLong Then
        ' (Note that the default property of a vbAccelerator ImageList
        ' is the hIml property.)
        
        ' Assume ImageList handle:
        m_hIml = vThis
        ' Get the icon size:
        lR = ImageList_GetIconSize(m_hIml, lIconSize, cy)
    Else
        Err.Raise vbObjectError + 1049, App.EXEName & ".cStatusBar", "ImageList property expects ImageList object or long hImageList handle."
    End If
       
End Property

Public Sub Create(ByRef objThis As Object)
Dim lhDC As Long
Dim lWidth As Long
Dim lHeight As Long
Dim tR As RECT

   Set m_obj = objThis
   
   ' Check if required methods are supported:
   On Error Resume Next
   lhDC = m_obj.hdc
   lWidth = m_obj.ScaleWidth
   lHeight = m_obj.ScaleHeight
   If (Err.Number <> 0) Then
      Set m_obj = Nothing
      Err.Raise 9, App.EXEName & ".cNoStatusBar", "Invalid object passed to Create."
   Else
      ' Get the height of the font and store:
      DrawText lhDC, "Xy", 2, tR, DT_CALCRECT
      m_lHeight = tR.Bottom - tR.Top + 6
   End If
   
End Sub

Public Property Set Font(ByRef fntThis As StdFont)
Dim tR As RECT
   Set m_obj.Font = fntThis
   ' Get the height of the font and store:
   DrawText m_obj.hdc, "Xy", 2, tR, DT_CALCRECT
   m_lHeight = tR.Bottom - tR.Top + 6
End Property
Public Property Get Font() As StdFont
   Font = m_obj.Font
End Property

Public Property Get Height() As Long
   Height = m_lHeight * Screen.TwipsPerPixelY
End Property

Public Property Get PanelCount() As Long
   PanelCount = m_iPanelCount
End Property
Public Sub GetPanelRect( _
      ByVal vKey As Variant, _
      Optional ByRef iLeftPixels As Long, _
      Optional ByRef iTopPixels As Long, _
      Optional ByRef iRightPixels As Long, _
      Optional ByRef iBottomPixels As Long _
   )
Dim iPanel As Long
Dim tR As RECT
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      With m_tPanels(iPanel).tR
         iLeftPixels = .Left
         iTopPixels = .Top
         iRightPixels = .Right
         iBottomPixels = .Bottom
      End With
   End If
End Sub

Public Property Get PanelKey(ByVal lIndex As Long) As Variant
Dim iPanel As Long
   If (lIndex > 0) And (lIndex <= m_iPanelCount) Then
      PanelKey = m_tPanels(lIndex).sKey
   Else
      Err.Raise vbObjectError + 1050, App.EXEName & ".cStatusBar", "Invalid Panel Index: " & lIndex
   End If
   
End Property
Public Property Let PanelKey(ByVal lIndex As Long, ByVal vKey As Variant)
   If (lIndex > 0) And (lIndex <= m_iPanelCount) Then
      m_tPanels(lIndex).sKey = vKey
   Else
      Err.Raise vbObjectError + 1050, App.EXEName & ".cStatusBar", "Invalid Panel Index: " & lIndex
   End If
   
End Property
Public Property Get PanelIndex(ByVal vKey As Variant) As Long
Dim i As Long
Dim iFound As Long

   If (IsNumeric(vKey)) Then
      If (vKey > 0) And (vKey <= m_iPanelCount) Then
         PanelIndex = vKey
      Else
         Err.Raise vbObjectError + 1050, App.EXEName & ".cStatusBar", "Invalid Panel Index: " & vKey
      End If
   Else
      For i = 1 To m_iPanelCount
         If m_tPanels(i).sKey = vKey Then
            iFound = i
            Exit For
         End If
      Next i
      If (iFound > 0) Then
         PanelIndex = iFound
      Else
         Err.Raise vbObjectError + 1050, App.EXEName & ".cStatusBar", "Invalid Panel Index: " & vKey
      End If
   End If
   
End Property
Public Property Let PanelText(ByVal vKey As Variant, ByVal sText As String)
Dim iPanel As Long
Dim iPartuType As Long
Dim lR As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).sText = sText
      Draw
   End If
End Property
Public Property Get PanelText(ByVal vKey As Variant) As String
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelText = m_tPanels(iPanel).sText
   End If
End Property
Public Property Let PanelSpring(ByVal vKey As Variant, ByVal bState As Boolean)
Dim iPanel As Long
Dim i As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).bSpring <> bState) Then
         For i = 1 To m_iPanelCount
            If i = iPanel Then
               m_tPanels(i).bSpring = bState
            Else
               m_tPanels(i).bSpring = False
            End If
         Next i
         pEvaluateIdealSize iPanel
         pResizeStatus
      End If
   End If
End Property
Public Property Get PanelSpring(ByVal vKey As Variant) As Boolean
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelSpring = m_tPanels(iPanel).bSpring
   End If
End Property
Public Property Let PanelFitToContents(ByVal vKey As Variant, ByVal bState As Boolean)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).bFit <> bState) Then
         m_tPanels(iPanel).bFit = bState
         pEvaluateIdealSize iPanel
         pResizeStatus
      End If
   End If
End Property
Public Property Get PanelFitToContents(ByVal vKey As Variant) As Boolean
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelFitToContents = m_tPanels(iPanel).bFit
   End If
End Property
Public Property Get PanelIcon(ByVal vKey As Variant) As Long
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelIcon = m_tPanels(iPanel).iImgIndex
   End If
End Property
Public Property Get PanelhIcon(ByVal vKey As Variant) As Long
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      ' Returns a hIcon if any:
      PanelhIcon = m_tPanels(iPanel).hIcon
   End If
End Property
Public Property Let PanelIcon(ByVal vKey As Variant, ByVal iImgIndex As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      If (m_tPanels(iPanel).hIcon <> 0) Then
         DestroyIcon m_tPanels(iPanel).hIcon
      End If
      m_tPanels(iPanel).hIcon = 0
      m_tPanels(iPanel).iImgIndex = iImgIndex
      If (iImgIndex > -1) Then
         ' extract a copy of the icon and add to sbar:
         m_tPanels(iPanel).hIcon = ImageList_GetIcon(m_hIml, iImgIndex, 0)
      End If
      Draw
   End If
End Property
Public Property Let PanelhIcon(ByVal vKey As Variant, ByVal hIcon As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      ' Destroy existing hIcon:
      If (m_tPanels(iPanel).hIcon <> 0) Then
         DestroyIcon m_tPanels(iPanel).hIcon
      End If
      m_tPanels(iPanel).hIcon = hIcon
      Draw
   End If
End Property
Public Property Let PanelStyle(ByVal vKey As Variant, ByVal eStyle As ENSBRPanelStyleConstants)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      'iPanel = iPanel - 1
      m_tPanels(iPanel).eStyle = eStyle
      Draw
   End If
End Property
Public Property Get PanelStyle(ByVal vKey As Variant) As ENSBRPanelStyleConstants
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelStyle = m_tPanels(iPanel).eStyle
   End If
End Property
Public Property Get PanelMinWidth(ByVal vKey As Variant) As Long
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelMinWidth = m_tPanels(iPanel).lMinWidth
   End If
End Property
Public Property Let PanelMinWidth(ByVal vKey As Variant, ByVal lWidth As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).lMinWidth = lWidth
      pResizeStatus
   End If
End Property
Public Property Get PanelAlignment(ByVal vKey As Variant) As Long
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelAlignment = m_tPanels(iPanel).lAlignment
   End If
End Property
Public Property Let PanelAlignment(ByVal vKey As Variant, ByVal lAlignment As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).lAlignment = lAlignment
      pResizeStatus
   End If
End Property

Public Property Get PanelIdealWidth(ByVal vKey As Variant) As Long
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelIdealWidth = m_tPanels(iPanel).lIdealWidth
   End If
End Property
Public Property Let PanelIdealWidth(ByVal vKey As Variant, ByVal lWidth As Long)
Dim iPanel As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).lIdealWidth = lWidth
      pResizeStatus
   End If
End Property

Public Property Let PanelToolTipText(ByVal vKey As Variant, ByVal sText As String)
Dim iPanel As Long
Dim lR As Long
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      m_tPanels(iPanel).sToolTipText = sText
      lR = SendMessageString(m_obj.hwnd, SB_SETTIPTEXT, iPanel - 1, sText & Chr$(0))
   End If
End Property

Public Property Get PanelToolTipText(ByVal vKey As Variant) As String
Dim iPanel As Long
Dim sTest As String
   iPanel = PanelIndex(vKey)
   If (iPanel > 0) Then
      PanelToolTipText = m_tPanels(iPanel).sToolTipText
   End If
End Property

Private Sub pEvaluateIdealSize( _
      ByVal iStartPanel As Long, _
      Optional ByVal iEndPanel As Long = -1 _
   )
Dim i As Long
Dim tR As RECT
Dim lhDC As Long

   If (m_iPanelCount > 0) Then
      If (iEndPanel < iStartPanel) Then
         iEndPanel = iStartPanel
      End If
      lhDC = m_obj.hdc
      For i = iStartPanel To iEndPanel
         DrawText lhDC, m_tPanels(i).sText, Len(m_tPanels(i).sText), tR, DT_CALCRECT
         m_tPanels(i).lIdealWidth = tR.Right - tR.Left + 12
         If (m_tPanels(i).lIdealWidth < m_tPanels(i).lMinWidth) Then
            m_tPanels(i).lIdealWidth = m_tPanels(i).lMinWidth
         End If
      Next i
   End If
End Sub
Private Sub pResizeStatus()
Dim tR As RECT
Dim i As Long
Dim iSpringIndex As Long
Dim lpParts() As Long
   
   If (m_iPanelCount > 0) Then
      
      GetClientRect m_obj.hwnd, tR
      tR.Left = tR.Left + m_lLeft
      tR.Top = tR.Top + m_lTop
      
      ' Initiallly set to minimum widths:
      ReDim lpParts(0 To m_iPanelCount - 1) As Long
      If (m_tPanels(1).bFit) Then
         lpParts(0) = m_tPanels(1).lIdealWidth
      Else
         lpParts(0) = m_tPanels(1).lMinWidth
      End If
      If (m_tPanels(1).hIcon) Then
         lpParts(0) = lpParts(0) + lIconSize
      End If
      If (m_tPanels(1).bSpring) Then
         iSpringIndex = 1
      End If
      For i = 2 To m_iPanelCount
         If (m_tPanels(i).bFit) Then
            lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lIdealWidth
         Else
            lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lMinWidth
         End If
         If (m_tPanels(i).bSpring) Then
            iSpringIndex = i
         End If
         If (m_tPanels(i).hIcon <> 0) Then
            ' Add space for the icon:
            lpParts(i - 1) = lpParts(i - 1) + lIconSize
         End If
         If (i = m_iPanelCount) Then
            lpParts(i - 1) = lpParts(i - 1) + (tR.Bottom - tR.Top) \ 2
         End If
      Next i
      
      ' Will all bars fit in at maximum size?
      If (lpParts(m_iPanelCount - 1) > tR.Right) Then
         ' Draw all panels at min width
      Else
         ' Spring the spring panel to fit:
         If (iSpringIndex = 0) Then
            iSpringIndex = m_iPanelCount
         End If
         lpParts(iSpringIndex - 1) = lpParts(iSpringIndex - 1) + (tR.Right - lpParts(m_iPanelCount - 1))
         For i = iSpringIndex + 1 To m_iPanelCount
            If (m_tPanels(i).bFit) Then
               lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lIdealWidth
            Else
               lpParts(i - 1) = lpParts(i - 2) + m_tPanels(i).lMinWidth
            End If
            If (m_tPanels(i).hIcon <> 0) Then
               ' Add space for the icon:
               lpParts(i - 1) = lpParts(i - 1) + lIconSize
            End If
            If (i = m_iPanelCount) Then
               lpParts(i - 1) = lpParts(i - 1) + (tR.Bottom - tR.Top) \ 2
            End If
         Next i
      End If
      
      m_tPanels(1).lSetWidth = lpParts(0)
      For i = 2 To m_iPanelCount
         m_tPanels(i).lSetWidth = lpParts(i - 1) - lpParts(i - 2)
      Next i
      
      ' Set the sizes:
      For i = 1 To m_iPanelCount
         With m_tPanels(i).tR
            If (i = 1) Then
               .Left = tR.Left
            Else
               .Left = lpParts(i - 2)
            End If
            If (i = m_iPanelCount) Then
               .Right = lpParts(i - 1)
            Else
               .Right = lpParts(i - 1) - 1
            End If
            .Top = tR.Top
            .Bottom = tR.Bottom
         End With
      Next i
      
   End If
   
End Sub

Private Sub Class_Initialize()
    Set m_Timer = New CTimer
    m_Timer.Interval = -1
End Sub

Private Sub Class_Terminate()
Dim i As Long
Dim lR As Long
   ' Delete any icons owned by the sbar:
   For i = 1 To m_iPanelCount
      If (m_tPanels(i).hIcon <> 0) Then
         lR = DestroyIcon(m_tPanels(i).hIcon)
         m_tPanels(i).hIcon = 0
      End If
   Next i
   Set m_Timer = Nothing
End Sub

Private Sub m_Timer_ThatTime()
    DrawPanelText
End Sub

Private Sub DrawPanelText()
Dim i As Long, iEnd As Long
Dim lhDC As Long
Dim lX As Long, lY As Long
Dim hBr As Long, tR As RECT, tOR As RECT, tBR As RECT
Dim fntThis As StdFont
Dim bEnd As Boolean
Dim bDraw As Boolean

    GetClientRect m_obj.hwnd, tR

    LSet tOR = tR

    lhDC = m_obj.hdc
    If (m_bSimpleMode) Then Exit Sub
    For i = 1 To m_iPanelCount
       bDraw = False
       With m_tPanels(i)
          LSet tBR = .tR
          If (tBR.Right > tOR.Right) Then
             tBR.Right = tOR.Right - 1
             bEnd = True
          End If
          If (.hIcon = 0) And (.eType <> sbrText) Then
             Select Case .eType
             Case sbrDate
                If .sText <> Format$(Now, "Short Date") Then
                    DrawStatusText lhDC, tBR, "", .eStyle
                    .sText = Format$(Now, "Short Date")
                    bDraw = True
                End If
             Case sbrTime
                 If .sText <> Format$(Now, "hh:nn") Then
                    DrawStatusText lhDC, tBR, "", .eStyle
                    .sText = Format$(Now, "hh:nn")
                    bDraw = True
                End If
             Case sbrIns
                 .sText = "INS"
                 If GetInsertKey() Then
                    SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                 Else
                    SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                 End If
                 bDraw = True
             Case sbrCaps
                 .sText = "CAPS"
                 If GetCapslock() Then
                    SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                 Else
                    SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                 End If
                 bDraw = True
             Case sbrNum
                 .sText = "NUM"
                 If GetNumlock() Then
                    SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                 Else
                    SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                 End If
                 bDraw = True
             Case sbrScrl
                 .sText = "SCRL"
                 If GetScrollLock() Then
                    SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
                 Else
                    SetTextColor lhDC, GetSysColor(COLOR_GRAYTEXT)
                 End If
                 bDraw = True
             End Select
             tBR.Left = tBR.Left + 4
             tBR.Right = tBR.Right - 4
             If bDraw Then
                DrawText lhDC, .sText, Len(.sText), tBR, DT_VCENTER Or DT_SINGLELINE Or .lAlignment
                SetTextColor lhDC, GetSysColor(COLOR_BTNTEXT)
             End If
          End If
          If bEnd Then
             Exit For
          End If
       End With
    Next i
End Sub


