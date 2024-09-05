Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion
Public txtConexion As String

Public prmPathApp As String

Public Sub Main()

On Error GoTo ErrMain

    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu(App.Title) Then
        MsgBox "Ud. no tiene permiso de ejecución para la aplicación " & App.Title, vbExclamation, "Falta Acceso"
        End
    End If
    
    txtConexion = miConexion.TextoConexion("comercio")
    InicioConexionBD txtConexion
    
    Cons = "Select * from Parametro Where ParNombre = 'pathapp'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        prmPathApp = Trim(RsAux!ParTexto) & "\"
    End If
    RsAux.Close
    
    frmLista.Show
    
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
    End
End Sub


