Attribute VB_Name = "ModVtasEstimadas"
Option Explicit
Public Const TipoEnCodTexto = 56
Public miConexion As New clsConexion

Public clsGeneral As New clsorCGSA

Public Sub MensajeError(ByVal msg As String, ByVal errdesc As String)
    clsGeneral.OcurrioError msg, errdesc, "Ventas estimadas"
End Sub

Sub Main()
    
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        VtasEstimadas.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        VtasEstimadas.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        VtasEstimadas.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        VtasEstimadas.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub
Public Sub RelojA()
    Screen.MousePointer = 11
End Sub

