Attribute VB_Name = "Constantes"
Option Explicit

Public txtConexion As String

'Definicion de Tipos de Campos (Tabla CodigoTexto)-------------------------------------------------------------
Public Enum TipoCampo
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
    'ComentarioMercaderiaPendiente = 19
    'Terminal = 20
    'Sucesos = 21
    'TipoMovCaja = 22
    'TipoExhibido = 23
    'Cargo = 24
    'CamposDireccion = 25
    'ListadePrecio = 26
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'ENUM DE SOLICITUDES------------------------------------------------------------------------------------------
'Definicion de Tipos de Solicitud
Public Enum TipoSolicitud
    AlMostrador = 1
    Reserva = 2
    Servicio = 3
End Enum

'Definicion de Tipos Resolucion de Solicitud
Public Enum TipoResolucionSolicitud
    Automatica = 1
    Manual = 2
    Facturada = 3
    Facturando = 4
    LlamarA = 5
End Enum

'Definicion de Tipos Resolucion de Solicitud
Public Enum EstadoSolicitud
    Pendiente = 0
    Aprovada = 1
    Rechazada = 2
    Condicional = 3
    ParaRetomar = 4
End Enum

'Constantes para las Formas de Pago de Solicitud
Global Const cFPSEfectivo = "Efectivo"
Global Const cFPSChequeD = "Cheque Diferido"

Public Enum TipoPagoSolicitud
    Efectivo = 1
    ChequeDiferido = 2
End Enum

Public Enum TipoCredito     '(campo) Tipo - tabla Credito
    Normal = 0
    Gestor = 1
    Incobrable = 2
    Clearing = 3
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Definiciones de Tipos de Locales
Public Enum TipoLocal
    Camion = 1
    Deposito = 2
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

'Definicion de Tipos de Clientes------------------------------------------------------------------------------------
Public Enum TipoCliente
    Persona = 1
    Empresa = 2
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Definicion de Acciones de Comentario----------------------------------------------------------------------------
Public Enum Accion
    Informacion = 1     'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Definicion para giros de empresas----------------------
Public Enum TipoEmpresa
    Proveedor = 1
    Banco = 2
    Agencia = 3
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Definicion para Llamados a los formularios----------------------
Public Enum TipoLlamado
    Normal = 0                         'Desde el menú
    IngresoNuevo = 3                'Para ingresar nuevos datos
    Modificacion = 7                  'Para modificar datos
    Visualizacion = 5                  'Llamado a clietnes

    'AL FORMULARIO DE CLIENTES
    CreditoAClientes = 2                           '(NuevoCompleto) Para dar nuevos clientes desde Credito
    ContadoAClientes = 1                         '(Nuevo) Para dar nuevos clientes desde Contado
    ClienteParaIngresoConyuge = 4           '(Conyuge) Al formulario de Clientes para ingresar el conyuege
    ClienteParaIngresoFNacimiento = 3      '(Modificar) Al formulario de Clientes para ingresar el la fecha de Nacimiento
    
    'Al Formulario de Cheques diferidos
    CuotasACheques = 8
    
End Enum
'-----------------------------------------------------------------------------------------------------------------------

Public Enum TipoImpresion
    EtiquetaAgencia = 1
    EnviosRepartoDirecto = 2
    EnviosAgencia = 3
    EnviosServicio = 4
    EntregaMercaderiaCamion = 5
    DocumentosContado = 6
    LiquidacionCamionero = 7
End Enum

'ENVIO.-------------------------------------------------------------------------------------------------------------------
'Definicion de Tipos de Pagos de Envio------------------------------------------------------------------------------------
Public Enum TipoPagoEnvio
    PagaAhora = 1
    PagaDomicilio = 2
    FacturaCamión = 3
End Enum

'Definicion de Tipos de Envios------------------------------------------------------------------------------------
Public Enum TipoEnvio
    Entrega = 1
    Service = 2
    Cobranza = 3
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Estados posibles que puede tener un envío.-----------------------------------------------------------------------------
Public Enum EstadoEnvio
    AImprimir = 0
    AConfirmar = 1
    Rebotado = 2
    Impreso = 3
    Entregado = 4
    Anulado = 5
End Enum

'----------------------------------------------------------------
Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public Enum TipoControlMercaderia
    CambioEstado = 1
    EntregaMercaderia = 2
End Enum

'REGISTRO DE SUCESOS--------------------------------------
Public Enum TipoSuceso
    ModificacionDeMora = 1
    AnulacionDeDocumentos = 2
    ModificacionDePrecios = 3
    RecepcionDeTraslados = 4
    AnulacionDeEnvios = 5
    CambioCostoDeFlete = 6
    Direcciones = 7
    ChequesDiferidos = 8
    CambioCategoriaCliente = 9
    Reimpresiones = 10
    DiferenciaDeArticulos = 11
    CederProductoServicio = 12
    FacturaArticuloInhabilitado = 13
    Notas = 14
    FacturaPlanInhabilitado = 15
    VariosStock = 98
    Varios = 99
End Enum
'----------------------------------------------------------------

'Constantes Para los Colores ------------------------------
Global Const Blanco = &HFFFFFF
Global Const Obligatorio = &HC0FFFF
Global Const Inactivo = &HE0E0E0
Global Const Rojo = &HFF&
Global Const Gris = &HE0E0E0
Global Const Cyan = &HC0C000
Global Const Busqueda = &HFFFFC0

Global Const FormatoCedula = "_.___.___-_"
Global Const FormatoFH = "mm/dd/yyyy hh:mm:ss"
Global Const FormatoMonedaP = "#,##0.00"

'ENVIOS
'Constantes para los Pagos de Envíos.---------
Global Const cPagaAhora = "Caja"
Global Const cPagaDomicilio = "Domicilio"
Global Const cFacturaCamion = "Fact. Camión"

'Constantes para los Estados de envios
Global Const cEnvAImprimir = "Confirmado"
Global Const cEnvAConfirmar = "A Confirmar"
Global Const cEnvRebotado = "Rebotado"
Global Const cEnvImpreso = "Impreso"
Global Const cEnvEntregado = "Entregado"
Global Const cEnvAnulado = "Anulado"


'VALORES DE TABLA PARAMETROS
Public paECivilConyuge As Long        'Sexo que requiere ingreso del conyuge
Public paDepartamento As Long       'Departamento por defecto
Public paLocalidad As Long              'Localidad por defecto
Public paMonedaEmpleo As Long      'Valor por defecto de la moneda en empleos
Public paMonedaFacturacion As Long
Public paTipoIngreso As Long           'Valor por defecto del tipo de ingreso
Public paTipoTelefonoP As Long              'Valor por defecto del tipo de telefono para las personas
Public paTipoTelefonoE As Long              'Valor por defecto del tipo de telefono para las empresas
Public paCategoriaCliente As Long     'Valor por defecto de la categoria del cliente
Public paCatCliPersonal As Long
Public paVigenciaEmpleo As Long
Public paTipoCuotaContado As Long
Public paEnvioFechaPrometida As Long

Public paArticuloPisoAgencia As Long
Public paArticuloDiferenciaEnvio As Long
Public paTipoArticuloServicio As Long
Public paArticuloCobroServicio As Long

Public paPrimeraHoraEnvio As Long
Public paUltimaHoraEnvio As Long
Public paMañanaEnvio As Long
Public paTardeEnvio As Long

Public paMonedaFija As Long
Public paMonedaFijaTexto As String

Public paVaToleranciaDiasExh As Integer
Public paVaToleranciaMonedaPorc As Integer
Public paVaToleranciaDiasExhTit As Integer
Public paDiasCobranzaCuota As Integer

Public paToleranciaMora As Integer       'Dias de Tolerancia de la mora
Public paCoeficienteMora As Currency    'Coeficiente para el cálculo de la mora
Public paIvaMora As Currency                'Porcentaje del IVA de la Mora
Public paCofis As Currency                      'Porcentaje del Cofis

Public paEstadoArticuloEntrega As Integer
Public paNoPagaPiso As Integer

'Valores en dias para los niveles de iconos en cuotas
Public paIconoPendienteN2Dias As Integer
Public paIconoVencimientoN2Dias As Integer

Public paMonedaDeuda As Long
Public paCartaATitular As Long
Public paCartaAGarantia As Long

'Parametros de Comentarios-------------
Public paLlamadaAMoroso As Long
Public paPasajeAGestor As Long

'Parametros de Movimientos de Salida de Caja-----------------------------------------
Public paMCNotaCredito As Long
Public paMCAnulacion As Long
Public paMCLiquidacionCamionero As Long
Public paMCChequeDiferido As Long
Public paMCVtaTelefonica As Long

Public paMayorDeEdad As Integer
Public paResolucionEstandar As Integer
Public paCantidadMaxCheques As Integer  'Maximo de Cheques Diferidos

Public paTipoFleteVentaTelefonica As Long
Public paLocalCompañia As Long

Public paPlanPorDefecto As Long

'OJO ESTOS SON PARAMETROS DE IMPORTACIONES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Public paLocalPuerto As Integer
Public paLocalZF As Long
Public paMonedaPesos As Integer
Public paMonedaDolar As Integer
Public paDisponibilidad As Long

Public Enum Folder      'Tipos de Folder
    cFCarpeta = 1
    cFEmbarque = 2
    cFSubCarpeta = 3
    cFPedido = 4
End Enum
'----------------------------------------------------------------------------------------------


Public Function NombreDocumentoCompra(Codigo As Integer) As String

    Select Case Codigo
        Case TipoDocumento.CompraCarta: NombreDocumentoCompra = "Carta"
        Case TipoDocumento.Compracontado: NombreDocumentoCompra = "Contado"
        Case TipoDocumento.CompraCredito: NombreDocumentoCompra = "Crédito"
        Case TipoDocumento.CompraNotaCredito: NombreDocumentoCompra = "Nota de Crédito"
        Case TipoDocumento.CompraNotaDevolucion: NombreDocumentoCompra = "Nota de Devolución"
        Case TipoDocumento.CompraRemito: NombreDocumentoCompra = "Remito"
         Case TipoDocumento.CompraCarpeta: NombreDocumentoCompra = "Carpeta Importación"
    End Select
    
End Function
