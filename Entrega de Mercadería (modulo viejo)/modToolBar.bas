Attribute VB_Name = "modToolBar"
' Module      : modToolBar
' Description : This module implements routines for manipulating toolbars
' Source      : Total VB SourceBook 6
  
Private Declare Function GetWindowLong _
  Lib "user32" _
  Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
    ByVal nIndex As Long) _
  As Long
  
Private Declare Function SetWindowLong _
  Lib "user32" _
  Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) _
  As Long
  
Private Declare Function EnumChildWindows _
  Lib "user32" _
  (ByVal hWndParent As Long, _
    ByVal lpEnumFunc As Long, _
    lParam As Long) _
  As Long

Private Const GWL_STYLE = (-16)

' This is the magic style
Private Const TBSTYLE_FLAT As Integer = &H800&

Private Function FlatToolBarEnumChildProc( _
  ByVal lnghWnd As Long, _
  lnghWndOut As Long) _
  As Long
  ' Comments  : This function should never be called directly by user code. It
  '             is used as a callback function for the Windows API
  '             EnumChildWindows function. See the Windows SDK for more
  '             documentation on this procedure
  ' Parameters: lnghWnd - the handle to a child window
  '             lnghWndOut - the variable used to store the handle to the
  '             child window. This is a user defined parameter
  ' Returns   : 1 - A nonzero value indicating that we should continue
  '             enumeration
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  'Store the last child window enumerated
  lnghWndOut = lnghWnd
  
  ' Return success code to the window API
  FlatToolBarEnumChildProc = 1

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FlatToolBarEnumChildProc"
  Resume PROC_EXIT

End Function

Public Sub SetToolBarFlat(tlbToolBar As Object, fFlat As Boolean)
  ' Comments  : This procedure sets a toolbar to the flat style. This requires
  '             Internet Explorer 3.0 or higher. This procedure has no effect
  '             if it is not installed.
  ' Parameters: tlbToolBar - The toolbar to set the style in
  '             fFlat - Flag indicating if we should set or clear the flat
  '             style
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim lngStyle As Long
  Dim lngNewStyle As Long
  Dim lngRetVal As Long
  Dim lnghWndToolBar As Long
  
  On Error GoTo PROC_ERR
  
  ' The last child window of the toolbar ocx is the actual toolbar control
  ' so we enumerate the child windows, and return the last child window in
  ' lnghWndToolBar. This allows you to place controls, such as combo boxes
  ' on the toolbar
  EnumChildWindows tlbToolBar.hwnd, AddressOf FlatToolBarEnumChildProc, lnghWndToolBar

  'Get the current window style
  lngStyle = GetWindowLong(lnghWndToolBar, GWL_STYLE)
  
  ' Set or clear the style based on the fFlat parameter
  If fFlat Then
    'Set the flat style
    lngNewStyle = lngStyle Or TBSTYLE_FLAT
  Else
    ' clear the flat style
    lngNewStyle = lngStyle And Not TBSTYLE_FLAT
  End If
  
  'Set the window style to the new style
  lngRetVal = SetWindowLong(lnghWndToolBar, GWL_STYLE, lngNewStyle)
  
  ' Refresh the toolbar, forcing it to repaint with the new style
  tlbToolBar.Refresh

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SetToolBarFlat"
  Resume PROC_EXIT

End Sub




