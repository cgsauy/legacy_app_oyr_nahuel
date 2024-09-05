Attribute VB_Name = "modStart"
Option Explicit
Public gFechaServidor As String
Public txtConexion As String
Public miConexion As New clsConexion
    
Public Sub Main()
On Error GoTo ErrMain
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu("Accesos") Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        Form1.Show
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrMain:
    MsgBox "Ocurrió un error al activar el ejecutable." & vbCrLf & Trim(Err.Description), vbExclamation, "ATENCIÓN"
    Screen.MousePointer = 0
End Sub

Private Sub start_FechaDelServidor()

    Dim RsF As rdoResultset
    
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    Time = gFechaServidor
    Date = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Date
End Sub


