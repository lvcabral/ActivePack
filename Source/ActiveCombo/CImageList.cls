VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
           Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
           Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Enum IconSize
  Size16 = 0
  Size32 = 1
End Enum

Public Enum IconState
  Normal = 0
  Disabled = 1
End Enum


Private Type PictDesc
  cbSizeofStruct As Long
  picType As Long
  hImage As Long
  xExt As Long
  yExt As Long
End Type
 
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type ImageFileInfo
  FileName As String
  IconIndex As Integer
  SystemIndex As Integer
End Type

Private ImgListImgInfo() As ImageFileInfo
Private ImageListHwnd As Long

Private Type ImageData
    hbmImage As Long
    hbmMask  As Long
    Unused1  As Long
    Unused2  As Long
    xLeft    As Long
    yTop     As Long
    xRight   As Long
    yBottom  As Long
End Type

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &H0
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20
 
Private Const ILD_NORMAL = &H0
Private Const ILD_TRANSPARENT = &H1
Private Const ILD_MASK = &H10
Private Const ILD_IMAGE = &H20
Private Const ILD_BLEND25 = &H2
Private Const ILD_BLEND50 = &H4
Private Const ILD_OVERLAYMASK = &H0
 
Private Const DI_NORMAL = 3
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_NORMAL = &H0&
Private Const DSS_RIGHT = &H8000
Private Const DSS_UNION = &H10
Private Const DST_BITMAP = &H4
Private Const DST_COMPLEX = &H0
Private Const DST_ICON = &H3&
Private Const DST_PREFIXTEXT = &H2
Private Const DST_TEXT = &H1

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Integer, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Boolean
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function DrawStateByString Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
 
Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ImageList_SetBkColor Lib "COMCTL32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_GetBkColor Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal Y As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_Convert Lib "COMCTL32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal Y As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Replace Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Private Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, hbmMask As Long) As Long
Private Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hImageList As Long, ImgIndex As Long) As Long
Private Declare Function ImageList_GeImageData Lib "COMCTL32" (ByVal himl As Long, ByVal ImgIndex As Long, pImageInfo As ImageData) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal himl As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, hbmMask As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "COMCTL32" (ByVal hImageList As Long, uNewCount As Long)
Private Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_SetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long

 Function ConvertIcon(hIcon) As Picture
    If hIcon = hNull Then Exit Function
    
    Dim NewPic As Picture, PicConv As PictDesc, IGuid As Guid
    
    PicConv.cbSizeofStruct = Len(PicConv)
    PicConv.picType = vbPicTypeIcon
    PicConv.hImage = hIcon
    
    IGuid.Data1 = &H20400
    IGuid.Data4(0) = &HC0
    IGuid.Data4(7) = &H46
   
    Call OleCreatePictureIndirect(PicConv, IGuid, True, NewPic)
    
    Set ConvertIcon = NewPic
End Function

Public Function Create(ImgSize As IconSize, ColorDepth As Integer) As Boolean
 
Dim SizeofIcon As Integer
 
    If ImgSize = 0 Then
       SizeofIcon = 16
    Else
       SizeofIcon = 32
    End If
    'Create the Imagelist
    ImageListHwnd = ImageList_Create(SizeofIcon, SizeofIcon, ILC_MASK Or ColorDepth, 0, 0)
 
    ReDim ImgListImgInfo(0)
  
End Function
Public Sub Destroy()
    Call ImageList_Destroy(ImageListHwnd)
End Sub

Public Function DrawImage(ImgIndex As Integer, Pic As Object)
Dim hIcon As Long

    Call ImageList_Draw(ImageListHwnd, ImgIndex, Pic.hdc, 0, 0, ILD_TRANSPARENT)
    Pic.Picture = Pic.Image

End Function

Public Function GetIconSize() As Integer
 Dim ImgHeight As Long, ImgWidth As Long
 Call ImageList_GetIconSize(ImageListHwnd, ImgHeight, ImgWidth)
 GetIconSize = ImgHeight
End Function

Public Function GetImageCount() As Integer
GetImageCount = ImageList_GetImageCount(ImageListHwnd)
End Function

Public Property Get Parent() As Object
    Set Parent = ObjParent
End Property

Public Property Set Parent(frm As Object)
    Set ObjParent = frm
End Property

Public Sub RemoveImage(Optional Index As Integer = -1)
    'If you don't specify the Index to remove it clears them all
    Call ImageList_Remove(ImageListHwnd, ByVal Index)
End Sub

Public Function AddIcon(hIcon As Variant) As Integer
    On Error Resume Next
    Call ImageList_AddIcon(ImageListHwnd, ConvertIcon(hIcon))
    AddIcon = GetImageCount
End Function

Public Function hWnd() As Long
    hWnd = ImageListHwnd
End Function

Private Sub Class_Terminate()
    If ImageListHwnd <> 0 Then
        Call ImageList_Destroy(ImageListHwnd)
    End If
End Sub

