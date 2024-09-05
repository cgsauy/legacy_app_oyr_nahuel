Attribute VB_Name = "modStart"
Option Explicit
Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

Public paEstadoArticuloARecuperar As Long
Public paEstadoArticuloRoto As Long
Public paGrupoRepuesto As Long
Public paTipoArticuloServicio As Long

Public paPlBalance As String

Public paLocalCompania As Long
Public paLocalEduardo As Long

'Datos de la base de datos II (para tambien generar movimientos)
Public cBaseMov As rdoConnection

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu("Diferencias Balance") Then End
        
    If Not InicioConexionBD(miConexion.TextoConexion("balance")) Then End
            
    If Not InicioConexionBDMov(miConexion.TextoConexion("comercio")) Then
        MsgBox "No se pudo conectar a la base de datos del comercio." & vbCrLf & "Se aconseja no corregir el lifo", vbExclamation, "Base de Datos"
    End If
    
    CargoParametrosLocal
    
    frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmListado.Status.Panels("bd") = "Balance: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
    
    frmListado.Status.Panels("bdcomercio") = "Comercio: " & PropiedadesConnect(cBaseMov.Connect, Database:=True) & " "
                    
    frmListado.Show vbModeless
    
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()

    cons = "Select * from cgsa.dbo.Parametro"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "estadoarecuperar": paEstadoArticuloARecuperar = rsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = rsAux!ParValor
            Case "estadoroto": paEstadoArticuloRoto = rsAux!ParValor
            
            Case "repuesto": paGrupoRepuesto = rsAux!ParValor
            
            Case "tipoarticuloservicio": paTipoArticuloServicio = rsAux!ParValor
            
            Case "plbalance": paPlBalance = Trim(rsAux!ParTexto)
            
            Case "localcompañia": paLocalCompania = rsAux!ParValor
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select * from cgsa.dbo.Local Where LocNombre = 'Eduardo'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then paLocalEduardo = rsAux!LocCodigo
    rsAux.Close
    
End Sub

'****     FUNCIONES DOBLES PARA LA BASE DE DATOS SECUNDARIA (comercio)            cBaseMov    ---------------------****************
Public Function InicioConexionBDMov(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
    
    On Error GoTo ErrICBD
    InicioConexionBDMov = False
    
    'Conexion a la base de datos----------------------------------------
    Set cBaseMov = eBase.OpenConnection("", , , strConexion)
    cBaseMov.QueryTimeout = sqlTimeOut
    
    InicioConexionBDMov = True
    Exit Function
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al intentar comunicarse con la Base de Datos." & vbCrLf & _
                "Error: " & Err.Description, vbExclamation, "Error de Conexión"
End Function

Public Function CierroConexionBDMov()
    On Error Resume Next
    cBaseMov.Close
End Function



