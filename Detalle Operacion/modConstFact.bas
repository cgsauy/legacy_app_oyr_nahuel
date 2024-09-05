Attribute VB_Name = "modConstantesFacturacion"
'Contiene Constantes propias del Sistema de Facturacion.
Option Explicit

'CONSTANTES..........................................................................
Public Const logFacturacion = "comercio" 'Constante de Logueo a la base de datos.
'..................................................................................................

'ENUMERALES..........................................................................
'Tipo en CodigoTexto.------------------------------
Public Enum TablasCT
    Departamento = 1
    RubroEmpresa = 2
    EstadoCivil = 3
    VigenciaEmpleo = 4
    Ocupacion = 5
    TipoIngreso = 6
    InquilinoPropietario = 7
    Telefono = 8
    TipoPlazo = 9
    CategoriaCliente = 10
    CategoriaArticulo = 11
    TipoOcupacion = 12
    Garantia = 13
    Piso = 14
    RangoHoraEnvio = 15
    ComentarioSolicitud = 16    'Comentarios de Solicitudes de Crédito
    TipoAgencia = 17
    EstadoFisicoMercaderia = 18
    
    ListaDePrecios = 26
    TextoVisita = 31
End Enum
'---------------------------------------------------------

'Formularios------------------------------------------
'Nro. de Llamado a los formularios de Manejo de Precios.
Public Enum frmManejoPrecios
    Moneda = 0
    PlanFinanciacion = 1
    TipoCuota = 2
    Coeficientes = 3
    PrecioArticulo = 4
    ListaDePrecios = 5
    CategoriaArticulo = 6
    CategoriaCliente = 7
End Enum
'---------------------------------------------------------

'Constantes de clientes------------------------------------------
Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum
'---------------------------------------------------------

Public Enum TipoCredito     '(campo) Tipo - tabla Credito
    Normal = 0
    Gestor = 1
    Incobrable = 2
    Clearing = 3
End Enum

'Definicion de Tipos de Documentos----------------------
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
End Enum

Public Enum TipoSolicitud
    AlMostrador = 1
    Reserva = 2
    Servicio = 3
End Enum

Public Enum TipoResolucionSolicitud
    Automatica = 1
    Manual = 2
    Facturada = 3
    Facturando = 4
    LlamarA = 5
End Enum

Public Enum TipoPagoSolicitud
    Efectivo = 1
    ChequeDiferido = 2
End Enum

Public Enum EstadoSolicitud
    Pendiente = 0
    Aprovada = 1
    Rechazada = 2
    Condicional = 3
    ParaRetomar = 4
    SinEfecto = 5       'Nuevo, se agregó el 14/07/2006
End Enum

'Trato de Stock.-------------------------------------
Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum
'---------------------------------------------------------

Public Enum TipoFormaDePago
    Efectivo = 1
    Cheque = 2
    Tarjeta = 3
End Enum

'ENVIO.-------------------------------------------------------------------------------------------------------------------
Public Enum TipoPagoEnvio     'Definicion de Tipos de Pagos de Envio
    PagaAhora = 1
    PagaDomicilio = 2
    FacturaCamión = 3
End Enum

Public Enum TipoEnvio       'Definicion de Tipos de Envios
    Entrega = 1
    Service = 2
    Cobranza = 3
End Enum

Public Enum EstadoEnvio     'Estados posibles que puede tener un envío.
    AImprimir = 0
    AConfirmar = 1
    Rebotado = 2
    Impreso = 3
    Entregado = 4
    Anulado = 5
End Enum
'----------------------------------------------------------------------------------------------------------------------------------------------------------

Public Enum TipoControlMercaderia       'New Desde Adrian
    CambioEstado = 1
    EntregaMercaderia = 2
End Enum

'Definicion de Acciones de Comentario----------------------------------------------------------------------------
Public Enum AccionComentario
    Informacion = 1     'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Parametros ------------------------------------------------------------------------------------------------------------
Public paDepartamento As Long
Public paLocalidad As Long
Public paTipoCuotaContado As Long       'Indica cual es el tipo de cuota Contado
Public paMonedaFacturacion As Long      'Indica la moneda fija para facturar.
Public paArticuloPisoAgencia As Long
Public paArticuloDiferenciaEnvio As Long
Public paTipoArticuloServicio As Long       'New Desde Adrian

Public paVaToleranciaDiasExh As Integer
Public paVaToleranciaMonedaPorc As Integer
Public paVaToleranciaDiasExhTit As Integer
Public paCategoriaCliente As Long
Public paPlanPorDefecto As Long
Public paEstadoArticuloEntrega As Integer

'.......................................................................................................
Public Function RetornoFormaDePago(IdForma As Integer)
    Select Case IdForma
        Case 1: RetornoFormaDePago = "Efectivo"
        Case 2: RetornoFormaDePago = "Con Cheque"
        Case 3: RetornoFormaDePago = "Tarjeta de Crédito"
    End Select
End Function
Public Sub CargoComboConFormaDePago(Combo As Control)
    Combo.Clear
    Combo.AddItem RetornoFormaDePago(1)
    Combo.ItemData(Combo.NewIndex) = 1
    Combo.AddItem RetornoFormaDePago(2)
    Combo.ItemData(Combo.NewIndex) = 2
    Combo.AddItem RetornoFormaDePago(3)
    Combo.ItemData(Combo.NewIndex) = 3
End Sub

Public Function RetornoEstadoVirtual(Codigo As Integer) As String
    RetornoEstadoVirtual = ""
    Select Case Codigo
        Case TipoMovimientoEstado.AEntregar: RetornoEstadoVirtual = "A Entregar"
        Case TipoMovimientoEstado.ARetirar: RetornoEstadoVirtual = "A Retirar"
        Case TipoMovimientoEstado.Reserva: RetornoEstadoVirtual = "Reserva"
    End Select
End Function

Public Function RetornoNombreDocumento(Codigo As Integer, Optional Abreviacion As Boolean = False) As String
    Dim aRet As String
    
    aRet = ""
    Select Case Codigo
        Case TipoDocumento.CompraCarta: aRet = "Carta"
        Case TipoDocumento.Compracontado: If Abreviacion Then aRet = "CON" Else aRet = "Contado"
        Case TipoDocumento.CompraCredito: If Abreviacion Then aRet = "CRE" Else aRet = "Crédito"
        Case TipoDocumento.CompraNotaCredito: If Abreviacion Then aRet = "NCR" Else aRet = "Nota de Crédito"
        Case TipoDocumento.CompraNotaDevolucion: If Abreviacion Then aRet = "NCO" Else aRet = "Nota de Devolución"
        Case TipoDocumento.CompraRemito:  If Abreviacion Then aRet = "REM" Else aRet = "Remito"
        Case TipoDocumento.CompraCarpeta: If Abreviacion Then aRet = "IMP" Else aRet = "Carpeta Importación"
        Case TipoDocumento.CompraRecibo: If Abreviacion Then aRet = "RPR" Else aRet = "Recibo Provisorio"
        Case TipoDocumento.CompraReciboDePago: If Abreviacion Then aRet = "RPA" Else aRet = "Recibo de Pago"
        
        Case TipoDocumento.CompraSalidaCaja: If Abreviacion Then aRet = "SAL" Else aRet = "Salida de Caja"
        Case TipoDocumento.CompraEntradaCaja: If Abreviacion Then aRet = "ENT" Else aRet = "Entrada de Caja"
        
        Case TipoDocumento.Contado: If Abreviacion Then aRet = "CON" Else aRet = "Contado"
        Case TipoDocumento.Credito: If Abreviacion Then aRet = "CRE" Else aRet = "Crédito"
        Case TipoDocumento.NotaCredito: If Abreviacion Then aRet = "NCR" Else aRet = "Nota de Crédito"
        Case TipoDocumento.NotaDevolucion: If Abreviacion Then aRet = "NDE" Else aRet = "Nota de Devolución"
        Case TipoDocumento.NotaEspecial: If Abreviacion Then aRet = "NES" Else aRet = "Nota Especial"
        Case TipoDocumento.ReciboDePago: If Abreviacion Then aRet = "REC" Else aRet = "Recibo"
        
        Case TipoDocumento.Envios: If Abreviacion Then aRet = "REP" Else aRet = "Reparto"
        Case TipoDocumento.Traslados: If Abreviacion Then aRet = "TRA" Else aRet = "Traslado"
        Case TipoDocumento.CambioEstadoMercaderia: If Abreviacion Then aRet = "CEM" Else aRet = "Cambio Estado Mercadería"
        
        Case TipoDocumento.IngresoMercaderiaEspecial: If Abreviacion Then aRet = "IME" Else aRet = "Ingreso Especial"
        Case TipoDocumento.ArregloStock: If Abreviacion Then aRet = "AST" Else aRet = "Arreglo Stock"
        Case TipoDocumento.Servicio: If Abreviacion Then aRet = "SER" Else aRet = "Servicio"
        Case TipoDocumento.ServicioCambioEstado: If Abreviacion Then aRet = "CEM" Else aRet = "Cambio Estado Mercadería"
        Case TipoDocumento.Devolucion: If Abreviacion Then aRet = "DEV" Else aRet = "Devolución"
    End Select
    
    RetornoNombreDocumento = aRet
    
End Function

Public Function RetornoNombreVistaCodigoTexto(IdTipo As Integer) As String

    Select Case IdTipo
        
        Case 1: RetornoNombreVistaCodigoTexto = "Departamento"
        Case 2: RetornoNombreVistaCodigoTexto = "Ramo"
        Case 3: RetornoNombreVistaCodigoTexto = "EstadoCivil"
        Case 4: RetornoNombreVistaCodigoTexto = "VigenciaEmpleo"
        Case 5: RetornoNombreVistaCodigoTexto = "Ocupacion"
        Case 6: RetornoNombreVistaCodigoTexto = "TipoIngreso"
        Case 7: RetornoNombreVistaCodigoTexto = "InquilinoPropietario"
        Case 8: RetornoNombreVistaCodigoTexto = "TipoTelefono"
        Case 9: RetornoNombreVistaCodigoTexto = "TipoPlazo"
        
        Case 10: RetornoNombreVistaCodigoTexto = "CategoriaCliente"
        Case 11: RetornoNombreVistaCodigoTexto = "CategoriaArticulo"
        Case 12: RetornoNombreVistaCodigoTexto = "TipoOcupacion"
        Case 13: RetornoNombreVistaCodigoTexto = "Garantia"
        Case 14: RetornoNombreVistaCodigoTexto = "Piso"
        Case 15: RetornoNombreVistaCodigoTexto = "HorarioEnvio"
        Case 16: RetornoNombreVistaCodigoTexto = "ComentarioSolicitud"
        Case 17: RetornoNombreVistaCodigoTexto = "TipoAgencia"
        Case 18: RetornoNombreVistaCodigoTexto = "EstadoMercaderia"
        
        Case 26: RetornoNombreVistaCodigoTexto = "ListaDePrecios"
    End Select
    
End Function

Public Function RetornoEstadoEnvio(Estado As Integer) As String

    Select Case Estado
        Case EstadoEnvio.AImprimir: RetornoEstadoEnvio = "Confirmado"
        Case EstadoEnvio.AConfirmar: RetornoEstadoEnvio = "A Confirmar"
        Case EstadoEnvio.Impreso: RetornoEstadoEnvio = "Impreso"
        Case EstadoEnvio.Anulado: RetornoEstadoEnvio = "Anulado"
        Case EstadoEnvio.Entregado: RetornoEstadoEnvio = "Entregado"
        Case EstadoEnvio.Rebotado: RetornoEstadoEnvio = "Rebotado"
    End Select
    
End Function

Public Function RetornoPagoEnvio(Pago As Integer) As String

    Select Case Pago
        Case TipoPagoEnvio.PagaAhora: RetornoPagoEnvio = "Caja"
        Case TipoPagoEnvio.FacturaCamión: RetornoPagoEnvio = "Fact. Camión"
        Case TipoPagoEnvio.PagaDomicilio: RetornoPagoEnvio = "Domicilio"
    End Select
    
End Function



