Attribute VB_Name = "modAATrayIcon"
Option Explicit
'This module allows you to easily place a form's icon on the system tray.

Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uID                 As Long
    uFlags              As Long
    uCallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type


'Constantes de Botones del Mousse
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205
'-------------------------------------------------------

'Constantes para pasar al shell
Const NIM_ADD = &H0
Const NIM_DELETE = &H2
Const NIM_MODIFY = &H1
Const NIM_SETFOCUS = &H4
Const NIM_SETVERSION = &H8

Const NIF_ICON = 2
Const NIF_MESSAGE = 1
Const NIF_TIP = 4
'-------------------------------------------------------

Const PK_TRAYICON = WM_MOUSEMOVE

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'Call this sub to place a form's icon on the system tray.
'
' hwnd    : The Hwnd property of the form the icon of which you
'           want to place in the system tray.
' Icon    : The form object.
' TipText : Self expanatory,the tool tip of the icon in the tray.
'
' example :  TrayNotify Me.Hwnd,Me,"This icon is on the tray"

Sub TrayNotify(hwnd As stdole.OLE_HANDLE, Icon As Form, tipText As String)

   Dim nd As NOTIFYICONDATA
   Dim nRet As Long
                    
   nd.hwnd = hwnd
   nd.uID = 1
   nd.uCallbackMessage = PK_TRAYICON
   nd.hIcon = Icon.Icon
   nd.szTip = tipText & Chr$(0)
   nd.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
   nd.cbSize = Len(nd)
   nRet = Shell_NotifyIcon(NIM_ADD, nd)
   
   If nRet = 0 Then Exit Sub

End Sub

'This sub gets executed when the user clicks on the tray icon.
'In order to handle the mouse events of the tray icon,you need
'to place this code:
'
'Call TrayClick(Button,Shift,X,Y)
'
'in the MouseMove event of the form the icon of which is in the tray.

'Retorna 1- Para accion       2- Para Pop Up Menu
Public Function TrayClick(Button As Integer, Shift As Integer, X As Single, Y As Single) As Integer

    On Error GoTo errTrayClick
    TrayClick = 0
    
    If (Button + Shift + Y) = 0 Or Button > 0 Then
        
        X = X / Screen.TwipsPerPixelX
        
        If Button = 0 Then
            Select Case X
                Case WM_LBUTTONDBLCLK: TrayClick = 1
                Case WM_RBUTTONUP: TrayClick = 2
            End Select
        Else
            'Agregado problema Win2000 (aparentemente retorna el nro. del boton).
            If Button = vbLeftButton And X = WM_LBUTTONDBLCLK Then
                TrayClick = 1
            ElseIf Button = vbRightButton And X = WM_RBUTTONUP Then
                TrayClick = 2
            End If
        End If
    End If
    
    Exit Function
    
errTrayClick:
    MsgBox "Ocurrió un error al activar el menú." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error"
End Function


'Call this sub to modify an existing icon in the tray.
'
'Hwnd    : The Hwnd property of the form the tray icon of which you
'          with to modify.
'Icon    : The form object with the new tray icon.
'TipText : The new tool tip of the tray icon.

Public Sub TrayModify(hwnd As stdole.OLE_HANDLE, Icon As Form, tipText As String)
On Error Resume Next

   Dim nd As NOTIFYICONDATA
   Dim nRet As Long
                    
   nd.hwnd = hwnd
   nd.uID = 1
   nd.uCallbackMessage = PK_TRAYICON
   nd.hIcon = Icon.Icon
   nd.szTip = tipText & vbNullChar
   nd.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
   nd.cbSize = Len(nd)
   
   nRet = Shell_NotifyIcon(NIM_MODIFY, nd)
   
   If nRet = 0 Then Exit Sub

End Sub


'Call this sub to remove a form's icon from the tray.
'
'Hwnd  : The Hwnd property of the form the icon of which you want to
'        remove from the tray.
Public Sub TrayRemove(hwnd As stdole.OLE_HANDLE)

On Error Resume Next

   Dim nd As NOTIFYICONDATA
   Dim nRet As Long
                    
   nd.hwnd = hwnd
   nd.uID = 1
   nd.uCallbackMessage = PK_TRAYICON
   nd.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
   nd.cbSize = Len(nd)
   
   nRet = Shell_NotifyIcon(NIM_DELETE, nd)

End Sub

