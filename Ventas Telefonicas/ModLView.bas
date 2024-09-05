Attribute VB_Name = "ModLView"
Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const WM_SETREDRAW = &HB

Private Const LVIS_STATEIMAGEMASK = &HF000
Private Const LVM_GETITEM = LVM_FIRST + 5 '75 for unicode?

Private Const LVM_SETITEMSTATE = LVM_FIRST + 43
Private Const LVM_GETITEMSTATE = LVM_FIRST + 44

Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Enum lvValores
    Grilla = &H1
    SubItemImage = &H2
    CheckBoxes = &H4
    TrackSelect = &H8
    HeaderDragDrop = &H10
    FullRow = &H20
    UnClickIcono = &H40
    DosClickIcono = &H80
    FlatSB = &H100
    Regional = &H200
    InfoTip = &H400
End Enum

Private Type LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As String
  cchTextMax As Long
  iImage As Long
  lParam As Long
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
  (ByVal hWnd1 As Long, _
   ByVal hWnd2 As Long, _
   ByVal lpsz1 As String, _
   ByVal lpsz2 As String) As Long
   
Public Sub SetearLView(Valor As Long, lv As MSComCtlLib.ListView)
        
    On Error Resume Next
   Valor = SendMessageLong(lv.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Valor)

End Sub

Public Function GetCheckState(lngIndex As Long, lv As MSComCtlLib.ListView) As Boolean
  Dim lngRet As Long
  Dim lngMask As Long
  Dim lvi As LV_ITEM
  Dim plvi As Long
  
  On Error GoTo PROC_ERR
  
  ' fill ListView Info structure to determine values to return
  lvi.iItem = lngIndex
  lvi.mask = lvValores.TrackSelect 'LVIF_STATE
  lvi.stateMask = LVIS_STATEIMAGEMASK
  
  ' get pointer to this structure
  plvi = VarPtr(lvi)
  ' retrieve current settings
  lngRet = SendMessageLong(lv.hwnd, LVM_GETITEM, 0&, plvi)
  
  ' get current state
  lngMask = lvi.state
  If lngMask And &H2000 Then
    GetCheckState = True
  Else
    GetCheckState = False
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetCheckState"
  Resume PROC_EXIT

End Function

Public Sub SetCheckState(lngIndex As Long, fValue As Boolean, lv As MSComCtlLib.ListView)

  Dim lngRet As Long
  Dim lngNewMask As Long
  Dim lvi As LV_ITEM
  Dim plvi As Long
  
  On Error GoTo PROC_ERR
  
  ' fill ListView Info structure to determine values to return
  lvi.iItem = lngIndex
  lvi.mask = lvValores.TrackSelect 'LVIF_STATE
  lvi.stateMask = LVIS_STATEIMAGEMASK
  
  ' get pointer to this structure
  plvi = VarPtr(lvi)
  ' retrieve current settings
  lngRet = SendMessageLong(lv.hwnd, LVM_GETITEM, 0&, plvi)

  ' get current state
  lngNewMask = lvi.state
    
  ' set appropriate mask bit on or off
  If fValue Then
    lngNewMask = (lngNewMask And (Not &H1000)) Or &H2000
  Else
    lngNewMask = (lngNewMask And (Not &H2000)) Or &H1000
  End If
 
  ' assign the new value
  lvi.state = lngNewMask
  
  ' send message to apply the new value
  lngRet = SendMessageLong(lv.hwnd, LVM_SETITEMSTATE, lngIndex, plvi)

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SetCheckState"
  Resume PROC_EXIT

End Sub

'Con cero inicio redraw, y con 1 finalizo.
Public Sub MarcoRedraw(lnValor As Long, lv As MSComCtlLib.ListView)
    SendMessageLong lv.hwnd, WM_SETREDRAW, lnValor, ByVal 0&
End Sub



Public Sub AutoSizeColumns(lv As MSComCtlLib.ListView)
  ' Comments  : Sizes each column in the listview control to fit
  '             the widest data in each column
  ' Parameters: None
  ' Returns   : Nothing
  ' Source    : Total VB SourceBook 6
  '
  Dim intColumn As Integer
  
  On Error GoTo PROC_ERR
  
  For intColumn = 0 To lv.ColumnHeaders.Count - 1
    SendMessageLong _
      lv.hwnd, _
      LVM_SETCOLUMNWIDTH, _
      intColumn, _
      LVSCW_AUTOSIZE_USEHEADER
  Next intColumn

PROC_EXIT:
  Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AutoSizeColumns"
  Resume PROC_EXIT

End Sub


