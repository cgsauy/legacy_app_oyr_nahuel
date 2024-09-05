Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    
    Screen.MousePointer = 11
    If miconexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miconexion.TextoConexion("comercio")
        FechaDelServidor
        
        'prm id artículo.
        Dim sPRM As String
        Dim idArtPrm As Long
        sPRM = Command()
        
        If sPRM <> "" And IsNumeric(sPRM) Then frmMovFisico.prmArticulo = Val(sPRM)
        frmMovFisico.Show vbModeless
    Else
        If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
