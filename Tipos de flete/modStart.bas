Attribute VB_Name = "modStart"
Option Explicit

Public cBase As rdoConnection       'Conexion a la Base de Datos
Public clsGeneral As New clsorCGSA

Public Sub Main()
On Error GoTo ErrMain
Dim miConexion As New clsConexion
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        
        If InicioConexionBD(miConexion.TextoConexion("comercio")) Then
            Set miConexion = Nothing
            frmAgenda.Show
        Else
            Set miConexion = Nothing
            Set clsGeneral = Nothing
        End If
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Set miConexion = Nothing
        End
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrMain:
    clsGeneral.OcurrioError "Ocurrií un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub


Private Function InicioConexionBD(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
On Error GoTo ErrICBD
    
    InicioConexionBD = False
    rdoEnvironments(0).CursorDriver = rdUseServer
    'Conexion a la base de datos----------------------------------------
    Set cBase = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, , strConexion)
    cBase.QueryTimeout = sqlTimeOut
    InicioConexionBD = True
    Exit Function
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al intentar comunicarse con la Base de Datos, se cancelará la ejecución.", vbExclamation, "ATENCIÓN"
End Function

Public Sub CargoCombo(Consulta As String, Combo As Control, Optional Seleccionado As String = "")

Dim RsAuxiliar As rdoResultset
Dim iSel As Integer: iSel = -1     'Guardo el indice del seleccionado
    
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then iSel = Combo.ListCount
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then Combo.ListIndex = iSel Else Combo.ListIndex = iSel - 1
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Public Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)
Dim I As Integer
    
    If cCombo.ListCount > 0 Then
        For I = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(I) = lngCodigo Then
                cCombo.ListIndex = I
                Exit Sub
            End If
        Next I
        cCombo.ListIndex = -1
    Else
        cCombo.ListIndex = -1
    End If

End Sub

