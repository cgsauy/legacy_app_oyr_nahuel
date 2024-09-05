Attribute VB_Name = "modSubClass"
Option Explicit

Private Const WM_ACTIVATE = &H6

Public lHookHwnd As Long
Public lPrevWndProc As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)

Public Sub Hook(ByVal hwnd As Long)
'Inicializo la subClass
    lHookHwnd = hwnd
    lPrevWndProc = SetWindowLong(lHookHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
Dim temp As Long
    'Cease subclassing.
    SetWindowLong lHookHwnd, GWL_WNDPROC, lPrevWndProc
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
         ByVal wParam As Long, ByVal lParam As Long) As Long

          'Check for request for min/max window sizes.
    If uMsg = WM_ACTIVATE Then
        'Deactivo
        If wParam = 0 Then Unload frmLista
    End If
    WindowProc = CallWindowProc(lPrevWndProc, hw, uMsg, wParam, lParam)
    
End Function



