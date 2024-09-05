Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion
Public txtConexion As String

Public paLocalZF As Long
Public paLocalPuerto As Long
Public paSubRubroAcreedores As Long
Public paTComentarioCtaCorr As Long

Public prmPathApp As String

Private Const prmKeyApp = "Cuentas Corrientes"

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        
        If miConexion.AccesoAlMenu(prmKeyApp) Then
            txtConexion = miConexion.TextoConexion("comercio")
            If Not InicioConexionBD(txtConexion) Then End
            
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            CargoParametrosComercio
            CargoParametrosLocal
            
            frmCuentas.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
            frmCuentas.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
            frmCuentas.Status.Panels("bd") = "BD: " & PropiedadesConnect(txtConexion, Database:=True) & " "
            
            frmCuentas.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miConexion.AccesoAlMenu prmKeyApp
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub CargoParametrosLocal()

    cons = "Select * from Parametro Where ParNombre like '%acreedore%' or ParNombre like '%comentario%' or ParNombre like '%pathapp%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "subrubroacreedoresvarios": paSubRubroAcreedores = rsAux!ParValor
            Case "tcomentarioctascorr": paTComentarioCtaCorr = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
End Sub
