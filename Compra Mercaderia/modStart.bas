Attribute VB_Name = "modStart"
Option Explicit
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paLocalPuerto As Long
Public paLocalZF As Long


Public Sub Main()

Dim aSucursal As String
Dim aIdCompra As Long

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logComercio)
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        CargoParametrosComercio
        aSucursal = CargoParametrosSucursal
        
        aIdCompra = Val(Command())
        If aIdCompra <> 0 Then frmInFactura.prmIDCompra = aIdCompra
        frmInFactura.Show vbModeless
        
        frmInFactura.Status.Panels("sucursal") = "Sucursal: " & aSucursal
        frmInFactura.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmInFactura.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmInFactura.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub
