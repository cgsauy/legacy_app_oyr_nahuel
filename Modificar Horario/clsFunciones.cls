VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetBDConnect(ByRef rdoConn As rdoConnection, ByVal ConnectKey As String, Optional iQTO As Byte = 15)
On Error GoTo errC
Dim oCnn As rdoConnection
    
    Dim oAC As New clsConexion
    Dim sConnect As String
    sConnect = oAC.TextoConexion(ConnectKey)
    Set oAC = Nothing
    
    rdoEnvironments(0).CursorDriver = rdUseServer
    Set oCnn = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, , sConnect)
    oCnn.QueryTimeout = iQTO
    Set rdoConn = oCnn
    Set oCnn = Nothing
    GetBDConnect = True
    Exit Function
errC:
    Screen.MousePointer = 0
    MsgBox "Error al iniciar la conexión, error: " & Err.Description, vbCritical, "Conexión"
End Function


Public Sub GetPositionForm(Frm As Form)
    
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    
    'Busco si tengo seteada la última posición y tamaño del formulario
    'Sino le marco yo los tamaños iniciales. ------------------------------------------
    With Frm
        .Left = GetSetting(App.Title, "Settings", "AA" & .Name & "Left", 1000)
        .Top = GetSetting(App.Title, "Settings", "AA" & .Name & "Top", 1000)
        .Width = GetSetting(App.Title, "Settings", "AA" & .Name & "Width", .Width)
        .Height = GetSetting(App.Title, "Settings", "AA" & .Name & "Height", .Height)
    End With
    
End Sub

Public Sub SetPositionForm(ByRef Form As Form)
    'Guarda la posicion y tamaño del form, si su estado es normal.
    If Form.WindowState <> vbMinimized And Form.WindowState <> vbMaximized Then
        SaveSetting App.Title, "Settings", "AA" & Form.Name & "Left", Form.Left
        SaveSetting App.Title, "Settings", "AA" & Form.Name & "Top", Form.Top
        SaveSetting App.Title, "Settings", "AA" & Form.Name & "Width", Form.Width
        SaveSetting App.Title, "Settings", "AA" & Form.Name & "Height", Form.Height
    End If
End Sub

Public Sub CargoCombo(ByVal Consulta As String, ByRef Combo As Control)

Dim RsAuxiliar As rdoResultset
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
    
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

