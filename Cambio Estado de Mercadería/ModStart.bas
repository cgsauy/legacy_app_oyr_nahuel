Attribute VB_Name = "ModStart"
Option Explicit

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public prmEsp As Boolean
Public clsGeneral As New clsorCGSA
Public UsuLogueado As Long
Public miConexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
    prmEsp = False
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        prmEsp = miConexion.AccesoAlMenu("Ingreso de MercaderiaE")
    
        InicioConexionBD miConexion.TextoConexion("comercio")
        CargoParametrosSucursal
        CamEstadoMercaderia.Show vbModeless
        Screen.MousePointer = 0
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

