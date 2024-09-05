Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion
Public txtConexion As String

Public Sub Main()

On Error GoTo ErrMain
Dim aTexto As String, aPos As Integer

    
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu("Quotation Items") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    txtConexion = miConexion.TextoConexion("comercio")
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    If Not InicioConexionBD(txtConexion) Then End
    
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        'id Item
        
        frmItems.prmIDItem = Val(aTexto)
    End If
    
    frmItems.Show
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar la aplicación.", Trim(Err.Description)
    Screen.MousePointer = 0
    End
End Sub


