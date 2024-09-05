Attribute VB_Name = "modConnect"
Option Explicit
'MODULO Conección
'Contiene rutinas y variables del entorno RDO.

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

Public Cons As String

'Usuario y Terminal
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDGI As Long
Public paCodigoDeTerminal As Long
Public paNombreSucursal As String
Public paDisponibilidad As Long

'Fecha del Servidor
Public gFechaServidor As Date

Public Function InicioConexionBD(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
On Error GoTo ErrICBD
    
    InicioConexionBD = False
    Set eBase = rdoCreateEnvironment("", "", "")
    eBase.CursorDriver = rdUseServer
    'Conexion a la base de datos----------------------------------------
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , strConexion)
    cBase.QueryTimeout = sqlTimeOut
    InicioConexionBD = True
    Exit Function
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al intentar comunicarse con la Base de Datos, se cancelará la ejecución.", vbExclamation, "ATENCIÓN"
End Function

Public Sub CierroConexion()
    On Error GoTo ErrCC
    cBase.Close
    eBase.Close
    Exit Sub
ErrCC:
    On Error Resume Next
End Sub

Public Function CargoDatosSucursal(ByVal sNombreTerminal As String, _
                                        Optional ByRef sNameCtdo As String = "", Optional ByRef sNameCred As String = "", _
                                        Optional ByRef sNameNCtdo As String = "", Optional ByRef sNameNCred As String = "", _
                                        Optional ByRef sNameRecibo As String = "", Optional ByRef sNameNEsp As String = "", Optional ByRef sNameNRemito As String = "") As Boolean
'................................................................................................................................................................
'Dado el nombre de la terminal
'   Cargo el código de la misma, el código de la sucursal y el nombre de los documentos.
'................................................................................................................................................................
On Error GoTo errCDS

    CargoDatosSucursal = False
    
    paCodigoDeSucursal = 0: paCodigoDeTerminal = 0
    sNameCtdo = "": sNameCred = ""
    sNameNCtdo = "": sNameNCred = ""
    sNameRecibo = "": sNameNEsp = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & sNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        paCodigoDGI = RsAux("SucCodDGI")
        paNombreSucursal = Trim(RsAux!SucAbreviacion)
        
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        
        If Not IsNull(RsAux!SucDContado) Then sNameCtdo = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then sNameCred = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then sNameNCtdo = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then sNameNCred = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDNEspecial) Then sNameNEsp = Trim(RsAux!SucDNEspecial)
        If Not IsNull(RsAux!SucDRecibo) Then sNameRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux("SucDRemito")) Then sNameNRemito = Trim(RsAux("SucDRemito"))
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(sNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    CargoDatosSucursal = (paCodigoDeSucursal > 0)
    Exit Function

errCDS:
    MsgBox "Error al leer la información de la sucursal." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Datos de Sucursal"
End Function

'----------------------------------------------------------------------------------------------------
'   Consulta por la fecha del servidor y la carga en la variable global gFechaServidor
'----------------------------------------------------------------------------------------------------
Public Sub FechaDelServidor()

    Dim RsF As rdoResultset
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    On Error Resume Next
    Time = gFechaServidor
    Date = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Now
End Sub
'---------------------------------------------------------------------------------------

