VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cActiveCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
' Name: Elasic
' Description:This class can change size and location of controls
' on your form if you
' 1. Resize form
' 2. Change screen resolution
' By: Mikhail Shmukler
'
' Inputs:None
' Returns:None
' Assumes:1. Add Elastic.cls
'         2. Add declaration 'Dim El as New Elastic'
'         3. Insert string like 'El.Link Me' (formload event)
'         4. Insert string like 'El.FormResize' (Resize event)
'         5. Press 'F5' and resize form ....
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************

Option Explicit
Private iFormHeight As Integer, iFormWidth As Integer, iNumOfControls As Integer
Private iTop() As Integer, iLeft() As Integer, iHeight() As Integer, iWidth() As Integer, iFontSize() As Integer, iRightMargin() As Integer
Private bFirstTime As Boolean
Private LinkedForm As Form

Sub Link()
    Set LinkedForm = FormName
    iFormHeight = LinkedForm.Height
    iFormWidth = LinkedForm.Width
    bFirstTime = True
End Sub

Sub FormResize(ByVal ResizeFonts As Boolean)
Dim I As Integer, Inc As Integer, CaptionSize As Integer
Dim RatioX As Double, RatioY As Double
Dim SaveRedraw%
    On Error Resume Next
    SaveRedraw% = LinkedForm.AutoRedraw
    
    LinkedForm.AutoRedraw = True
    
    If bFirstTime Then
        Init
        bFirstTime = False
        Exit Sub
    End If
    
    If LinkedForm.Height < iFormHeight / 2 Then LinkedForm.Height = iFormHeight / 2
    If LinkedForm.Width < iFormWidth / 2 Then LinkedForm.Width = iFormWidth / 2
    CaptionSize = 400
    RatioY = 1# * (iFormHeight - CaptionSize) / (LinkedForm.Height - CaptionSize)
    RatioX = 1# * iFormWidth / LinkedForm.Width
    On Error Resume Next ' for comboboxes, timeres and other nonsizible controls

    For I = 0 To iNumOfControls

        If TypeOf LinkedForm.Controls(I) Is Line Then
            LinkedForm.Controls(I).Y1 = Int(iTop(I) / RatioY)
            LinkedForm.Controls(I).X1 = Int(iLeft(I) / RatioX)
            LinkedForm.Controls(I).Y2 = Int(iHeight(I) / RatioY)
            LinkedForm.Controls(I).X2 = Int(iWidth(I) / RatioX)
        Else
            LinkedForm.Controls(I).Top = Int(iTop(I) / RatioY)
            LinkedForm.Controls(I).Left = Int(iLeft(I) / RatioX)
            LinkedForm.Controls(I).Height = Int(iHeight(I) / RatioY)
            LinkedForm.Controls(I).Width = Int(iWidth(I) / RatioX)
            If ResizeFonts Then
                LinkedForm.Controls(I).Font.Size = Int(iFontSize(I) / RatioX) + Int(iFontSize(I) / RatioX) Mod 2
                LinkedForm.Controls(I).RightMargin = Int(iRightMargin(I) / RatioY)
            End If
        End If
    
    Next

    LinkedForm.AutoRedraw = SaveRedraw%
    Exit Sub
End Sub

Private Sub Init()
    Dim I As Integer
    iNumOfControls = LinkedForm.Controls.Count - 1
    ReDim iTop(iNumOfControls)
    ReDim iLeft(iNumOfControls)
    ReDim iHeight(iNumOfControls)
    ReDim iWidth(iNumOfControls)
    ReDim iFontSize(iNumOfControls)
    ReDim iRightMargin(iNumOfControls)
    On Error Resume Next
    For I = 0 To iNumOfControls
        If TypeOf LinkedForm.Controls(I) Is Line Then
            iTop(I) = LinkedForm.Controls(I).Y1
            iLeft(I) = LinkedForm.Controls(I).X1
            iHeight(I) = LinkedForm.Controls(I).Y2
            iWidth(I) = LinkedForm.Controls(I).X2
        Else
            iTop(I) = LinkedForm.Controls(I).Top
            iLeft(I) = LinkedForm.Controls(I).Left
            iHeight(I) = LinkedForm.Controls(I).Height
            iWidth(I) = LinkedForm.Controls(I).Width
            iFontSize(I) = LinkedForm.Font.Size
            iRightMargin(I) = LinkedForm.Controls(I).RightMargin
        End If
    Next
End Sub
