Attribute VB_Name = "modProyecto"
Option Explicit
Public objG As New clsorCGSA
Public oUsers As New clsConexion

Public paEstadoArticuloEntrega As Integer
Public gFechaServidor As Date
Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoDocumento
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    ReciboDePago = 5
    Remito = 6
    ContadoDomicilio = 7
    CreditoDomicilio = 8
    ServicioDomicilio = 9
    NotaEspecial = 10
    
    'Documentos de Compras
    Compracontado = 11
    CompraCredito = 12
    CompraNotaDevolucion = 13
    CompraNotaCredito = 14
    CompraRemito = 15
    CompraCarta = 16
    CompraCarpeta = 17
    
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
End Enum

Public paCodigoDeSucursal As Long, paCodigoDeTerminal As Long

Private Sub loc_GetSucursal()
Dim Cons As String, rsAux As rdoResultset
    Cons = "Select * From Terminal, Local" _
            & " Where TerNombre = '" & oUsers.NombreTerminal & "'" _
            & " And TerSucursal = LocCodigo"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeSucursal = rsAux!TerSucursal
        paCodigoDeTerminal = rsAux!TerCodigo
    End If
    rsAux.Close
End Sub

Public Sub Main()
On Error GoTo errC
    Screen.MousePointer = 11
    
    If oUsers.AccesoAlMenu("Seccion Entrega") Then
        If InicioConexionBD(oUsers.TextoConexion("Comercio")) Then
            loc_GetSucursal
            If paCodigoDeSucursal > 0 Then
                frmEntRapida.Show
            Else
                MsgBox "No se obtuvo el código de sucursal, no se abrira la aplicación.", vbCritical, "Atención"
            End If
        End If
    Else
        MsgBox "Ud. no tiene acceso a la aplicación.", vbExclamation, "Atención"
    End If
    Screen.MousePointer = 0
Exit Sub
errC:
    Screen.MousePointer = 0
    objG.OcurrioError "Error al iniciar la aplicación.", Err.Description, "Entrega rápida"
End Sub

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
