Attribute VB_Name = "modStart"

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral


Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logComercio)
        
        'CargoParametrosImportaciones
        CargoParametrosComercio
        aSucursal = CargoParametrosSucursal
'        CargoPathAppGeneral
        
        frmCatDto.Status.Panels("sucursal") = "Sucursal: " & aSucursal
        frmCatDto.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmCatDto.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmCatDto.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        frmCatDto.Show vbModeless
    
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
