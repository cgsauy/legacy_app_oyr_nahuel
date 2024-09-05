Attribute VB_Name = "ModStart"
Option Explicit

Public paClienteEmpresa As Long
Public paEstadoArticuloEntrega As Integer, paEstadoARecuperar As Integer

Public clsGeneral As New clsLibGeneral
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miconexion.TextoConexion("comercio")
            CargoParametrosComercio
            paCodigoDeUsuario = miconexion.UsuarioLogueado(True)
            frmListado.Status.Panels("terminal").Text = "Terminal: " & miconexion.NombreTerminal
            frmListado.Status.Panels("usuario").Text = "Usuario: " & miconexion.UsuarioLogueado(False, True)
            frmListado.Show vbModeless
        Else
            MsgBox "No tiene permisos de ingreso.", vbExclamation, "ATENCIÓN"
            End
        End If
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Public Sub CargoParametrosComercio()

    'Parametros a cero--------------------------
    paEstadoArticuloEntrega = 0
    
    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

