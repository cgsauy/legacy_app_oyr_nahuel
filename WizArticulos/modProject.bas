Attribute VB_Name = "modProject"
Option Explicit

'Definición del entorno RDO
Public cBase As rdoConnection            'Conexion a la Base de Datos
Public eBase As rdoEnvironment          'Definicion de entorno
Public RsAux As rdoResultset               'Resultset Auxiliar
Public Cons As String

Public clsGeneral As New clsorCGSA
Public prmUID As Long

Public Sub Main()
On Error GoTo errMain
Dim sPrm As String
    Screen.MousePointer = 11
    If db_Acceso(App.Title, prmUID) Then
        If ConnectBD Then
            sPrm = Trim(Command())
            If IsNumeric(sPrm) Then frmWizArticulo.prmID = Val(sPrm)
            frmWizArticulo.Show
        Else
            End
        End If
    Else
        MsgBox "Sin permiso", vbExclamation, "Acceso"
    End If
    Screen.MousePointer = 0
    Exit Sub
errMain:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la aplicación.", Err.Description
End Sub

Private Function ConnectBD() As Boolean
On Error GoTo errCBD
    Dim objAcceso As New clsConexion
    ConnectBD = False
    Set eBase = rdoCreateEnvironment("", "", "")
    eBase.CursorDriver = rdUseServer
    'Conexion a la base de datos----------------------------------------
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , objAcceso.TextoConexion("Comercio"))
    cBase.QueryTimeout = 15
    ConnectBD = True
    Exit Function
    
errCBD:
    On Error Resume Next
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar la conexión con el servidor de datos, se cancelará la ejecución.", Err.Description
End Function

Public Function db_Acceso(ByVal sName As String, lUID As Long) As Boolean
On Error GoTo errA
Dim objAcceso As New clsConexion
    db_Acceso = objAcceso.AccesoAlMenu(sName)
    If db_Acceso Then lUID = objAcceso.UsuarioLogueado(True)
    Set objAcceso = Nothing
    Exit Function
errA:
    clsGeneral.OcurrioError "Error al validar el acceso, la aplicación se cancelará.", Err.Description
    db_Acceso = False
    End
End Function

Public Sub db_CloseConnect()
    On Error GoTo ErrCC
    cBase.Close
    eBase.Close
    Exit Sub
ErrCC:
    On Error Resume Next
End Sub

Public Sub prj_SetFocus(ByVal ctrl As Control)
On Error Resume Next
    With ctrl
        If .Enabled Then
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End If
    End With
End Sub

Public Function prj_SetFilterFind(ByVal sTexto As String) As String
    sTexto = RTrim(sTexto)
    sTexto = Replace(sTexto, " ", "%")
    sTexto = Replace(sTexto, "*", "%")
    prj_SetFilterFind = "'" & sTexto & "%'"
End Function

