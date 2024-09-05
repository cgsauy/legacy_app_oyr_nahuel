Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathApp As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logComercio)
        
        CargoParametrosComercio
        CargoParametrosCaja
        CargoParametrosImportaciones
        
        prmPathApp = App.Path & "\"
        
        frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        frmListado.Show vbModeless
    
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
