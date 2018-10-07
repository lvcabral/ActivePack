Attribute VB_Name = "mSubClass"
Option Explicit
Private x As Long
Private ctlHooked As New Collection

' SetWindowLong / GetWindowLong
Public Const GWL_WNDPROC = -4
Public Const GWL_USERDATA = -21
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_EXSTYLE = (-20)

Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPRevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long


Public Function HookWindow(hWnd As Long, oCtl As Object) As Boolean
Dim lpPRevWndFunc As Long
On Error GoTo HookError
    If hWnd > 0 Then
        ' Guarda Endereço do Controle
        ctlHooked.Add oCtl, CStr(hWnd)
        ' Redireciona a função de janela
        lpPRevWndFunc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf HookedWindowProc)
        ' Guarda endereco da funcao antiga
        x = SetWindowLong(hWnd, GWL_USERDATA, lpPRevWndFunc)
        HookWindow = True
    End If
HookExit:
    Exit Function
HookError:
    HookWindow = False
    Resume HookExit
End Function

Public Sub UnHookWindow(hWnd As Long)
Dim lOldProc As Long
    'Retorna o endereço original
    If hWnd > 0 Then
        lOldProc = GetWindowLong(hWnd, GWL_USERDATA)
        x = SetWindowLong(hWnd, GWL_WNDPROC, lOldProc)
        ctlHooked.Remove CStr(hWnd)
    End If
End Sub

Private Function HookedWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lDefaultProc As Long
Dim hBrush As Long
    On Error Resume Next
    If ctlHooked(CStr(hWnd)).WindowProc(hWnd, uMsg, wParam, lParam) Then
        lDefaultProc = GetWindowLong(hWnd, GWL_USERDATA)
        HookedWindowProc = CallWindowProc(lDefaultProc, hWnd, uMsg, wParam, lParam)
    Else
        If uMsg = WM_CTLCOLOREDIT Or _
           uMsg = WM_CTLCOLORLISTBOX Or _
           uMsg = WM_CTLCOLORSTATIC Then
            If ctlHooked(CStr(hWnd)).BackColor And &H80000000 Then
                hBrush = GetSysColorBrush(ctlHooked(CStr(hWnd)).BackColor - &H80000000)
            Else
                hBrush = CreateSolidBrush(ctlHooked(CStr(hWnd)).BackColor)
            End If
            HookedWindowProc = hBrush
        Else
            HookedWindowProc = 0
        End If
    End If
End Function

Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Function HiWord(ByVal dw As Long) As Integer
    HiWord = (dw And &HFFFF0000) \ 65536
End Function
