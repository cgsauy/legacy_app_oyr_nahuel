Attribute VB_Name = "modStart"
Option Explicit

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
    NotaDebito = 40
    
    'Documentos de Compras
    Compracontado = 11
    CompraCredito = 12
    CompraNotaDevolucion = 13
    CompraNotaCredito = 14
    CompraRemito = 15
    CompraCarta = 16
    CompraCarpeta = 17
    CompraRecibo = 18
    CompraReciboDePago = 19
    CompraSalidaCaja = 30       'Pedidos el 11/12 por carlos y juliana
    CompraEntradaCaja = 31
    
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
    Servicio = 26
    ServicioCambioEstado = 27
    Devolucion = 28
    
    VentasWebAConfirmar = 32
    VentasWebConfirmada = 33
    
    RemitoRecepcion = 34
End Enum


Public clsGeneral As New clsorCGSA

Public Sub Main()
Dim miConexion As clsConexion
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Set miConexion = New clsConexion
    'Si da error la conexión la misma despliega el msg de error
    If Not miConexion.AccesoAlMenu("Detalle de Factura") Then 'App.Title) Then
        Screen.MousePointer = 0
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    Else
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        Set miConexion = Nothing
        
        Screen.MousePointer = 0
        frmEntProgramada.Show
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr, vbCritical, "ATENCIÓN"
    End
End Sub

