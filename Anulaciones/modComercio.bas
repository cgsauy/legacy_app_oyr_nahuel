Attribute VB_Name = "modComercio"
Option Explicit

Public Const logComercio = "Comercio"

Public Enum DocPendiente
    Servicio = 2
End Enum

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
End Enum

'Definicion de Tipos Resolucion de Solicitud
Public Enum EstadoSolicitud
    Pendiente = 0
    Aprovada = 1
    Rechazada = 2
    Condicional = 3
    ParaRetomar = 4
    SinEfecto = 5       'Nuevo, se agregó el 14/07/2006
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
'-----------------------------------------------------------------------------------------------------------------------

'Definicion de Acciones de Comentario----------------------------------------------------------------------------
Public Enum Accion
    Informacion = 1     'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum
'-----------------------------------------------------------------------------------------------------------------------

'Definicion para Llamados a los formularios----------------------
Public Enum TipoLlamado
    Normal = 0                         'Desde el menú
    IngresoNuevo = 3                'Para ingresar nuevos datos
    Modificacion = 7                  'Para modificar datos
    Visualizacion = 5                  'Llamado a clietnes

    'AL FORMULARIO DE CLIENTES
    CreditoAClientes = 1                           'Para dar nuevos clientes desde Credito
    ContadoAClientes = 2                         'Para dar nuevos clientes desde Contado
    ClienteParaIngresoConyuge = 4           'Al formulario de Clientes para ingresar el conyuege
    ClienteParaIngresoFNacimiento = 6      'Al formulario de Clientes para ingresar el la fecha de Nacimiento
    
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
    FacturaCambioNombre = 16
    CambioTipoArticuloServicio = 17
    ConfiguracionSistema = 18
    EliminarInstalacion = 21
    NotaArticuloSinDocumento = 22
    ClienteNoVender = 23
    VariosStock = 98
    Varios = 99
End Enum

Public Enum Cuenta
    Personal = 1
    Colectivo = 2
End Enum

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
Public paMonedaEmpleo As Long      'Valor por defecto de la moneda en empleos
Public paMonedaFacturacion As Long
Public paTipoIngreso As Long           'Valor por defecto del tipo de ingreso
Public paTipoTelefonoP As Long              'Valor por defecto del tipo de telefono para las personas
Public paCofis As Currency

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
Public paMCIngresosOperativos As Long
Public paMCTransferencias As Long
Public paMCSenias As Long

Public paMayorDeEdad As Integer
Public paResolucionEstandar As Integer
Public paCantidadMaxCheques As Integer  'Maximo de Cheques Diferidos

Public paTipoFleteVentaTelefonica As Long
Public paLocalCompañia As Long

Public paPlanPorDefecto As Long

Public paSubrubroCompraMercaderia As Long
Public paSubrubroAcreedoresVarios As Long

'Public paLocalZonaFranca As Long
'Public paLocalPuerto As Long

Public Sub CargoParametrosComercio()

    'Parametros a cero-----------------
    paMonedaFijaTexto = ""
    paCoeficienteMora = 1
    paIvaMora = 1
    paEstadoArticuloEntrega = 1
    paTipoFleteVentaTelefonica = 0
    
    'Dias para niveles de iconos
    paIconoPendienteN2Dias = 60
    paIconoVencimientoN2Dias = 60

    cons = "Select * from Parametro"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case Trim(rsAux!ParNombre)
            Case "ECivilConyuge": paECivilConyuge = rsAux!ParValor
            
            Case "Departamento": paDepartamento = rsAux!ParValor
            Case "Localidad": paLocalidad = rsAux!ParValor
            
            Case "MonedaEmpleo": paMonedaEmpleo = rsAux!ParValor
            Case "MonedaFacturacion": paMonedaFacturacion = rsAux!ParValor
                
            Case "TipoIngreso": paTipoIngreso = rsAux!ParValor
            
            Case "TipoTelefonoP": paTipoTelefonoP = rsAux!ParValor
            Case "TipoTelefonoE": paTipoTelefonoE = rsAux!ParValor
                
            Case "CategoriaCliente": paCategoriaCliente = rsAux!ParValor
            
            Case "VigenciaEmpleo": paVigenciaEmpleo = rsAux!ParValor
            
            Case "TipoCuotaContado": paTipoCuotaContado = rsAux!ParValor
                
            Case "ArticuloPisoAgencia": paArticuloPisoAgencia = rsAux!ParValor
            Case "ArticuloDiferenciaEnvio": paArticuloDiferenciaEnvio = rsAux!ParValor
            Case "TipoArticuloServicio": paTipoArticuloServicio = rsAux!ParValor
            Case "TipoFleteVentaTelefonica": paTipoFleteVentaTelefonica = rsAux!ParValor
            Case "ArticuloCobroServicio": paArticuloCobroServicio = rsAux!ParValor
                
            Case "PrimeraHoraEnvio": paPrimeraHoraEnvio = rsAux!ParValor
            Case "UltimaHoraEnvio": paUltimaHoraEnvio = rsAux!ParValor
                
            Case "EnvioFechaPrometida": paEnvioFechaPrometida = rsAux!ParValor
                
            Case "MonedaFija": paMonedaFija = rsAux!ParValor
            
            Case "VaToleranciaMonedaPorc": paVaToleranciaMonedaPorc = rsAux!ParValor
            Case "VaToleranciaDiasExh": paVaToleranciaDiasExh = rsAux!ParValor
            Case "VaToleranciaDiasExhTit": paVaToleranciaDiasExhTit = rsAux!ParValor
                
            Case "ToleranciaMora": paToleranciaMora = rsAux!ParValor
                
            Case "CoeficienteMora": paCoeficienteMora = ((rsAux!ParValor / 100) + 1) ^ (1 / 30)         'Como es mensual calculo el diario
            
            Case "IvaMora": paIvaMora = rsAux!ParValor
                
            Case "DiasCobranzaCuota": paDiasCobranzaCuota = rsAux!ParValor
            
            Case "EstadoArticuloEntrega": paEstadoArticuloEntrega = rsAux!ParValor
                
            Case "IconoPendienteN2Dias": paIconoPendienteN2Dias = rsAux!ParValor
            Case "IconoVencimientoN2Dias": paIconoVencimientoN2Dias = rsAux!ParValor
            
            Case "LlamadaAMoroso": paLlamadaAMoroso = rsAux!ParValor
            Case "PasajeAGestor": paPasajeAGestor = rsAux!ParValor
            
            Case "MonedaDeuda": paMonedaDeuda = rsAux!ParValor
                
            Case "CartaAGarantia": paCartaAGarantia = rsAux!ParValor
            Case "CartaATitular": paCartaATitular = rsAux!ParValor
                
            Case "MCNotaCredito": paMCNotaCredito = rsAux!ParValor
            Case "MCLiquidacionCamionero": paMCLiquidacionCamionero = rsAux!ParValor
            Case "MCChequeDiferido": paMCChequeDiferido = rsAux!ParValor
            Case "MCAnulacion": paMCAnulacion = rsAux!ParValor
            Case "MCVtaTelefonica": paMCVtaTelefonica = rsAux!ParValor
            Case "MCIngresosOperativos": paMCIngresosOperativos = rsAux!ParValor
            Case "MCTransferencias": paMCTransferencias = rsAux!ParValor
            Case "MCSenias": paMCSenias = rsAux!ParValor
                        
            Case "LocalCompañia": paLocalCompañia = rsAux!ParValor
            
            Case "SubrubroCompraMercaderia": paSubrubroCompraMercaderia = rsAux!ParValor
            Case "SubrubroAcreedoresVarios": paSubrubroAcreedoresVarios = rsAux!ParValor
                        
            'OJO ESTOS SON LOS PARAMETROS DE IMPORTACIONES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'Case "LocalZF": paLocalZF = rsAux!ParValor
            'Case "LocalPuerto": paLocalPuerto = rsAux!ParValor
            
            'OJO ESTOS SON LOS PARAMETROS COMUNES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            Case "MonedaDolar": paMonedaDolar = rsAux!ParValor
            Case "MonedaPesos": paMonedaPesos = rsAux!ParValor
            
            Case "MayorDeEdad": paMayorDeEdad = rsAux!ParValor
            Case "CantidadMaxCheques": paCantidadMaxCheques = rsAux!ParValor
            Case "PlanPorDefecto": paPlanPorDefecto = rsAux!ParValor
            
            Case "Cofis": paCofis = rsAux!ParValor
            
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If paMonedaFija <> 0 Then
        cons = "Select * from Moneda Where MonCodigo = " & paMonedaFija
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            paMonedaFijaTexto = Trim(rsAux!MonSigno)
        Else
            MsgBox "El código de moneda fija (parámetro) no existe en la base de datos.", vbCritical, "ERROR"
            paMonedaFija = 0
        End If
        rsAux.Close
    End If
    
End Sub

Public Function IVAArticulo(lnCodigo As Long)

    Dim RsIva As rdoResultset
    On Error GoTo ErrIA
    
    cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
            & " Where AFaArticulo = " & lnCodigo _
            & " And AFaIVA = IVACodigo"
    Set RsIva = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    IVAArticulo = 0
    If Not RsIva.EOF Then IVAArticulo = Format(RsIva(0), "#0.00")
    RsIva.Close
    Exit Function
    
ErrIA:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al buscar el tipo de iva del artículo.", Err.Description
End Function

Public Function DireccionATexto(Codigo As Long, Optional Departamento As Boolean = False, Optional Localidad As Boolean = False, Optional Zona As Boolean = False, _
                                               Optional EntreCalles As Boolean = False, Optional Ampliacion As Boolean = False, Optional ConfYVD As Boolean = False, Optional ConEnter As Boolean = False)

Dim aTexto As String
Dim RsAux2 As rdoResultset
Dim RsADir As rdoResultset

    aTexto = ""
    cons = "Select Direccion.*, LocNombre, DepNombre, CalNombre From Direccion, Calle, Localidad, Departamento" _
            & " Where DirCodigo = " & Codigo _
            & " And DirCalle = CalCodigo And CalLocalidad = LocCodigo" _
            & " And LocDepartamento = DepCodigo"
    
    Set RsADir = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If Departamento Then aTexto = Trim(RsADir!DepNombre)
    
    If Localidad Then
        If Departamento Then
            If Trim(RsADir!DepNombre) <> Trim(RsADir!LocNombre) Then aTexto = aTexto & ", " & Trim(RsADir!LocNombre)
        Else
            aTexto = aTexto & Trim(RsADir!LocNombre)
        End If
    End If
    
    If Zona Then        'Saco La ZONA------------------------------------------------------------------------------------------
        cons = "Select ZonNombre from CalleZona, Zona" _
               & " Where CZoCalle = " & RsADir!DirCalle _
               & " And CZoDesde <= " & RsADir!DirPuerta _
               & " And CZoHasta >= " & RsADir!DirPuerta _
               & " And CZoZona = ZonCodigo"
        Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " (" & Trim(RsAux2!ZonNombre) & ")"
        End If
        RsAux2.Close
    End If
    '------------------------------------------------------------------------------------------------------------
    If ConEnter Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) Else aTexto = aTexto & " "
    
    aTexto = aTexto & Trim(RsADir!CalNombre) & " "
    
    If Trim(RsADir!DirPuerta) = 0 Then
        aTexto = aTexto & "S/N"
    Else
        aTexto = aTexto & Trim(RsADir!DirPuerta)
    End If
    
    If Not IsNull(RsADir!DirLetra) Then aTexto = aTexto & Trim(RsADir!DirLetra)
    If Not IsNull(RsADir!DirApartamento) Then aTexto = aTexto & "/" & Trim(RsADir!DirApartamento)
    If RsADir!DirBis Then aTexto = aTexto & " Bis"
    
    
    'Campo 1 de la Direccion---------------------------------------------------------------------------------------
    If Not IsNull(RsADir!DirCampo1) Then
        cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsADir!DirCampo1
        Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsADir!DirSenda) Then aTexto = aTexto & " " & Trim(RsADir!DirSenda)
        End If
        RsAux2.Close
    End If
    'Campo 2 de la Direccion---------------------------------------------------------------------------------------
    If Not IsNull(RsADir!DirCampo2) Then
        cons = "Select CDiAbreviacion from CamposDireccion Where CDiCodigo = " & RsADir!DirCampo2
        Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux2.EOF Then
            aTexto = aTexto & " " & Trim(RsAux2!CDiAbreviacion)
            If Not IsNull(RsADir!DirBloque) Then aTexto = aTexto & " " & Trim(RsADir!DirBloque)
        End If
        RsAux2.Close
    End If
    '---------------------------------------------------------------------------------------------------------------------
    
    'Entre calles--------------------------------------------------------------------------------------------------------
    If ConEnter And EntreCalles Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10) Else aTexto = aTexto & " "
    If EntreCalles Then
        If Not IsNull(RsADir!DirEntre1) Then
            cons = "Select * from Calle Where CalCodigo = " & RsADir!DirEntre1
            Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux2.EOF Then
                If Not IsNull(RsADir!DirEntre2) Then aTexto = aTexto & "Entre " Else aTexto = aTexto & "Esq. "
                aTexto = aTexto & Trim(RsAux2!CalNombre) & " "
            End If
            RsAux2.Close
        End If
        If Not IsNull(RsADir!DirEntre2) Then
            cons = "Select * from Calle Where CalCodigo = " & RsADir!DirEntre2
            Set RsAux2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux2.EOF Then aTexto = aTexto & "y " & Trim(RsAux2!CalNombre)
            RsAux2.Close
        End If
    End If
    '---------------------------------------------------------------------------------------------------------------------
    
    'Ampliacion de Direccion------------------------------------------------------------------------------------------
    If ConEnter And Ampliacion Then
        If Asc(Mid(aTexto, Len(aTexto) - 1, 1)) <> vbKeyReturn Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10)
    Else
        If Mid(aTexto, Len(aTexto) - 1, 1) <> " " Then aTexto = aTexto & " "
    End If
    If Ampliacion Then If Not IsNull(RsADir!DirAmpliacion) Then aTexto = aTexto & Trim(RsADir!DirAmpliacion)
    '---------------------------------------------------------------------------------------------------------------------
    
    'Confirmada y vive Desde----------------------------------------------------------------------------------------------------
    If ConEnter And ConfYVD Then
        If Asc(Mid(aTexto, Len(aTexto) - 1, 1)) <> vbKeyReturn Then aTexto = aTexto & Chr(vbKeyReturn) & Chr(10)
    Else
        If Mid(aTexto, Len(aTexto) - 1, 1) <> " " Then aTexto = aTexto & " "
    End If
    
    If ConfYVD Then
        If Not IsNull(RsADir!DirVive) Then aTexto = aTexto & "VD: " & Format(RsADir!DirVive, "Mmm/YYYY")
        If RsADir!DirConfirmada Then aTexto = aTexto & " (Cf.)"
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    RsADir.Close
    DireccionATexto = Trim(aTexto)
    
End Function


