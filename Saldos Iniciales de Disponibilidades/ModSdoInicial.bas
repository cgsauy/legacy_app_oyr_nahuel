Attribute VB_Name = "ModSdoInicial"
Option Explicit

Public clsGeneral As New clsLibGeneral
Public miConexion As New clsConexion

Public paEstadoArticuloEntrega As Long
Public paMonedaDolar As Long, paMonedaPesos As Long
Public paDepartamento  As Long, paLocalidad As Long
Public paCategoriaCliente  As Long
Public paTipoTelefonoE  As Long

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        SdoDisponibilidad.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        SdoDisponibilidad.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        SdoDisponibilidad.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        SdoDisponibilidad.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title
    End
End Sub

