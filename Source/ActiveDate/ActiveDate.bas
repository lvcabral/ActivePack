Attribute VB_Name = "mActiveDate"
Option Explicit
Private x As Long
Private mForm As Form

'Window Messages
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ACTIVATE = &H6
 
' SetWindowLong / GetWindowLong
Public Const GWL_STYLE = -16
Public Const GWL_WNDPROC = -4
Public Const GWL_HINSTANCE = -6
Public Const GWL_ID = -12
Public Const GWL_USERDATA = -21
Public Const GWL_NEWUSERDATA = -8

'Window Style
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WM_SETREDRAW = &HB

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPRevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Sub HookWindow(oFrm As Form)
Dim lpPRevWndFunc As Long
    DoEvents
    Set mForm = oFrm
    If Not mForm Is Nothing Then
        ' Redireciona a função de janela
        lpPRevWndFunc = SetWindowLong(mForm.hWnd, GWL_WNDPROC, AddressOf ActiveDate_WindowProc)
        ' Guarda endereco da funcao antiga
        x = SetWindowLong(mForm.hWnd, GWL_USERDATA, lpPRevWndFunc)
    End If
End Sub

Public Sub UnHookWindow()
    'Retorna o endereço original
    If Not mForm Is Nothing Then
        x = SetWindowLong(mForm.hWnd, GWL_WNDPROC, GetWindowLong(mForm.hWnd, GWL_USERDATA))
        Unload mForm
        Set mForm = Nothing
    End If
End Sub

Private Function ActiveDate_WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If uMsg = WM_ACTIVATEAPP And wParam = 0 Then
        'mForm.Hide
        UnHookWindow
        Exit Function
    ElseIf uMsg = WM_ACTIVATE And wParam = 0 Then
        'mForm.Hide
        UnHookWindow
        Exit Function
    End If
    ActiveDate_WindowProc = CallWindowProc(GetWindowLong(hWnd, GWL_USERDATA), hWnd, uMsg, wParam, lParam)
End Function

