Attribute VB_Name = "modStart"
Option Explicit

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDeTerminal As Long

Public Enum Asuntos
    Solicitudes = 1
    Servicios = 2
    GastosAAutorizar = 3
    SucesosAAutorizar = 4
    SolicitudesResueltas = 5
End Enum

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmLocalPort As Long
Public prmServerIP As String
Public prmQueryInterval As Long

Public Const CTE_KeyConnect As String = "comercio"


Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    prmQueryInterval = 15000
    
    If App.PrevInstance Then
        MsgBox "Esta aplicación está activa. " & vbCrLf & _
        "No se puede abrir una nueva instanica.", vbExclamation, "Servidor de Asuntos Pendientes está activo..."
        End
    End If
    
    If Not ObtenerConexionBD(cBase, CTE_KeyConnect) Then Screen.MousePointer = 0: Exit Sub
    
    CargoParametrosLocales
    arrInicializoVariables
    FechaDelServidor
    frmServer.Show
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Public Sub EndMain()
On Error Resume Next
    cBase.Close
    End
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParametro
    
    Cons = "Select * from Parametro " & _
               " Where ParNombre IN ( 'ServerAsuntos_Port_IP', 'ServerAsuntos_QueryTM' )"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, CTE_KeyConnect, rdOpenForwardOnly, rdConcurReadOnly, True) <> RAQ_SinError Then Exit Sub
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "serverasuntos_port_ip"
                    prmLocalPort = RsAux!ParValor
                    prmServerIP = Trim(RsAux!ParTexto)
            
            Case LCase("ServerAsuntos_QueryTM")
                    prmQueryInterval = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close

    
    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros locales.", Err.Description
End Sub

Public Function z_BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
    z_BuscoUsuario = BuscoUsuario(Codigo, Identificacion, Digito, Iniciales)
End Function

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function

