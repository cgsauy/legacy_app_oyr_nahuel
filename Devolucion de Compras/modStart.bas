Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paLocalZF As Long, paLocalPuerto As Long
Public txtConexion As String

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion(logComercio)
        InicioConexionBD txtConexion
        
        CargoParametrosComercio
        aSucursal = CargoParametrosSucursal
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmInFactura.Status.Panels("sucursal") = "Sucursal: " & aSucursal
        frmInFactura.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmInFactura.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmInFactura.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        frmInFactura.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub
