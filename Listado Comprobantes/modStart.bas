Attribute VB_Name = "modStart"

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA ' clsLibGeneral

Public paLocalZF As Long, paLocalPuerto As Long


Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu("Listado Compra") Then
        InicioConexionBD miConexion.TextoConexion(logComercio)
        
        CargoParametrosComercio
        aSucursal = CargoParametrosSucursal
        
        frmCompra.Status.Panels("sucursal") = "Sucursal: " & aSucursal
        frmCompra.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmCompra.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmCompra.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmCompra.Show vbModeless
    
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
