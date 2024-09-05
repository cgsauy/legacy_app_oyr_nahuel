Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public prmKeyConnect As String
Public Const prmKeyApp = "Recibos de Pago"
Public Const prmKeyAppADM = "GastosADM"

Public prmPathApp As String

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu(prmKeyApp) Then
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Acceso Denegado"
        End If
        Screen.MousePointer = 0: End
    End If
    
    prmKeyConnect = miConexion.TextoConexion("comercio")
    If Not InicioConexionBD(prmKeyConnect) Then End
        
    CargoParametrosImportaciones
    CargoParametrosComercio
    CargoParametrosSucursal
    
    CargoParametrosLocales
    frmRecibos.Show
    
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & prmKeyApp & vbCrLf & _
                Err.Number & " - " & Err.Description
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP
    
    prmFCierreIVA = CDate("1/1/2002")
    
    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathApp', 'FechaCierreIVA')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
            Case "fechacierreiva": prmFCierreIVA = CDate(rsAux!ParTexto)
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros (locales).", Err.Description
End Sub


