Attribute VB_Name = "modComun"
'Function Redondeo  la saque el 16/01/2002  esta en el modCredito

Option Explicit

'Parametros Comunes a los sistemas
Public paDisponibilidad As Long

Public paDepartamento As Long       'Departamento por defecto
Public paLocalidad As Long              'Localidad por defecto
Public paCategoriaCliente As Long
Public paTipoTelefonoE As Long              'Valor por defecto del tipo de telefono para las empresas
Public paEstadoArticuloEntrega As Integer

Public paMonedaDolar As Integer
Public paMonedaPesos As Integer

'-----------------------------------------------------------------------------------------------------------------

'Definicion de Tipos de Documentos----------------------
Public Enum TipoDocumento
    'Documentos Facturacion
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
    
    'Otros
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
    Servicio = 26
    ServicioCambioEstado = 27
    Devolucion = 28
    
End Enum

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

Public Enum Pendiente
    Comercio = 1
    Servicio = 2
End Enum

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
        Case TipoDocumento.ContadoDomicilio, TipoDocumento.CreditoDomicilio: If Abreviacion Then aRet = "VTE" Else aRet = "Vta. Telefónica"
        
        Case TipoDocumento.CompraSalidaCaja: If Abreviacion Then aRet = "SAL" Else aRet = "Salida de Caja"
        Case TipoDocumento.CompraEntradaCaja: If Abreviacion Then aRet = "ENT" Else aRet = "Entrada de Caja"
        
        Case TipoDocumento.Contado: If Abreviacion Then aRet = "CON" Else aRet = "Contado"
        Case TipoDocumento.Credito: If Abreviacion Then aRet = "CRE" Else aRet = "Crédito"
        Case TipoDocumento.NotaCredito: If Abreviacion Then aRet = "NCR" Else aRet = "Nota de Crédito"
        Case TipoDocumento.NotaDevolucion: If Abreviacion Then aRet = "NDE" Else aRet = "Nota de Devolución"
        Case TipoDocumento.NotaEspecial: If Abreviacion Then aRet = "NES" Else aRet = "Nota Especial"
        Case TipoDocumento.ReciboDePago:  If Abreviacion Then aRet = "REC" Else aRet = "Recibo"
        Case TipoDocumento.NotaDebito:  If Abreviacion Then aRet = "DEB" Else aRet = "Nota de Débito"

        
        Case TipoDocumento.Envios: If Abreviacion Then aRet = "REP" Else aRet = "Reparto"
        Case TipoDocumento.Traslados: If Abreviacion Then aRet = "TRA" Else aRet = "Traslado"
        Case TipoDocumento.CambioEstadoMercaderia: If Abreviacion Then aRet = "CEM" Else aRet = "Cambio Estado Mercadería"
        
        Case TipoDocumento.ArregloStock: If Abreviacion Then aRet = "AST" Else aRet = "Arreglo Stock"
        Case TipoDocumento.IngresoMercaderiaEspecial: If Abreviacion Then aRet = "IME" Else aRet = "Ingreso Mercadería Especial"
        Case TipoDocumento.Servicio: If Abreviacion Then aRet = "SER" Else aRet = "Servicio"
        Case TipoDocumento.ServicioCambioEstado: If Abreviacion Then aRet = "CEM" Else aRet = "Cambio Estado Mercadería"
        Case TipoDocumento.Devolucion: If Abreviacion Then aRet = "DEV" Else aRet = "Devolución"
        
    End Select
    
    RetornoNombreDocumento = aRet
    
End Function

'--------------------------------------------------------------------------------------------------------------
'   Inserta un Registro en la tabla Suceso
'--------------------------------------------------------------------------------------------------------------
Public Sub RegistroSuceso(Fecha As Date, Tipo As Integer, Terminal As Long, Usuario As Long, Documento As Long, _
                                     Optional Articulo As Long = 0, Optional Descripcion As String = "", Optional Defensa As String = "", _
                                     Optional Valor As Currency = 0)


    Cons = "Insert into Suceso" _
           & " (SucFecha, SucTipo, SucTerminal, SucUsuario, SucDocumento, SucArticulo, SucDescripcion, SucDefensa, SucValor)" _
           & " Values ( " _
           & "'" & Format(Fecha, "mm/dd/yyyy hh:mm:ss:nn") & "', " _
           & Tipo & "," _
           & Terminal & ", " _
           & Usuario & ", " _
           & Documento & ", "
           
        If Articulo <> 0 Then Cons = Cons & Articulo & ", " Else: Cons = Cons & "Null, "
        If Descripcion <> "" Then Cons = Cons & "'" & Trim(Descripcion) & "' , " Else: Cons = Cons & "Null, "
        If Defensa <> "" Then Cons = Cons & "'" & Trim(Defensa) & "', " Else: Cons = Cons & "Null, "
        If Valor <> 0 Then Cons = Cons & Valor & " )" Else: Cons = Cons & "Null)"
    
    cBase.Execute Cons

End Sub

Public Function CargoParametrosSucursal() As String

Dim aNombreTerminal As String

    CargoParametrosSucursal = ""
    aNombreTerminal = miConexion.NombreTerminal
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
        
        'lo borre pq la variable pdRecibo no esta definida en algunos forms pasar este proc a local
        'If Not IsNull(rsAux!SucDRecibo) Then paDRecibo = Trim(rsAux!SucDRecibo)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Sub MovimientoDeCaja(Tipo As Long, Fecha As Date, Disponibilidad As Long, Moneda As Integer, Importe As Currency, _
                                            Optional Comentario As String = "", Optional Salida As Boolean = False, Optional IdCompra As Long = 0)

Dim RsMov As rdoResultset, rs1 As rdoResultset
Dim aMovimiento As Long, aMonedaD As Integer
Dim TC As Currency, aImporteD As Currency
    
    Importe = Abs(Importe)
    'Saco la Moneda de la disponibilidad
    Cons = "Select * from Disponibilidad Where DisID = " & Disponibilidad
    Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aMonedaD = rs1!DisMoneda
    rs1.Close
    '------------------------------------------------------------------------------------------------------------

    'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidad Where MDiID = 0"
    Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsMov.AddNew
    RsMov!MDiFecha = Format(Fecha, "mm/dd/yyyy")
    RsMov!MDiHora = Format(Fecha, "hh:mm:ss")
    RsMov!MDiTipo = Tipo
    If IdCompra <> 0 Then RsMov!MDiIdCompra = IdCompra
    If Comentario <> "" Then RsMov!MDiComentario = Trim(Comentario)
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Saco el Id de movimiento-------------------------------------------------------------------------------
    Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
              " Where MDiFecha = " & Format(Fecha, "'mm/dd/yyyy'") & _
              " And MDiHora = " & Format(Fecha, "'hh:mm:ss'") & _
              " And MDiTipo = " & Tipo
    If IdCompra <> 0 Then Cons = Cons & " And MDiIdCompra = " & IdCompra
    
    Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aMovimiento = RsMov(0)
    RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
    Cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & aMovimiento
    Set RsMov = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsMov.AddNew
    RsMov!MDRIdMovimiento = aMovimiento
    RsMov!MDRIdDisponibilidad = Disponibilidad
    RsMov!MDRIdCheque = 0
    
    RsMov!MDRImporteCompra = Importe
    
    If aMonedaD = Moneda Then        'Disponibilidad = Mov
        If Salida Then RsMov!MDRHaber = Importe Else RsMov!MDRDebe = Importe
    Else                                            'Tasa de cambio (del Mov a Disp)
        TC = TasadeCambio(Moneda, aMonedaD, Fecha)
        aImporteD = Importe * TC
        If Salida Then RsMov!MDRHaber = aImporteD Else RsMov!MDRDebe = aImporteD
    End If
    
    If Moneda = paMonedaPesos Then  'Mov = Pesos
        RsMov!MDRImportePesos = Importe
    Else
        If aMonedaD = paMonedaPesos Then    'Disp = Pesos
            RsMov!MDRImportePesos = aImporteD
        Else
            'Tasa de cambio a pesos
            TC = TasadeCambio(Moneda, paMonedaPesos, Fecha)
            RsMov!MDRImportePesos = Importe * TC
        End If
    End If
    
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
                
End Sub

Public Function TasadeCambio(MOriginal As Integer, MDestino As Integer, Fecha As Date, Optional FechaTC As String = "", Optional TipoTC As Integer = -1) As Currency

On Error GoTo errTC

Dim retTasaC As Currency
Dim bOK As Boolean

    If TipoTC = -1 Then TipoTC = 1
    TasadeCambio = 1
        
    bOK = tc_Cotizacion(MOriginal, MDestino, Fecha, FechaTC, TipoTC, retTasaC)
    
    If Not bOK Then 'Consulto por el campo cotiza en hasta llegar a la moneda final.
        Dim rsMon As rdoResultset
        Dim mDatos As String: mDatos = ""
        
        Cons = "Select MonCodigo, MonCotizaEn From Moneda  Where MonCotizaEn Is Not Null"
        Set rsMon = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsMon.EOF
            mDatos = mDatos & rsMon!MonCodigo & ":" & rsMon!MonCotizaEn
            rsMon.MoveNext
            If Not rsMon.EOF Then mDatos = mDatos & "|"
        Loop
        rsMon.Close
        
        If Trim(mDatos) <> "" Then
            retTasaC = 1
            Dim arrData() As String
            Dim bSalir As Boolean: bSalir = False
            Dim mMonO As Integer, mCotizaEn As Integer
            Dim idx As Integer
            Dim mretTC As Currency
            
            arrData = Split(mDatos, "|")
            mMonO = MOriginal
            mDatos = ""
            
            Do While Not bSalir
                If Trim(mDatos) <> "" Then
                    If InStr(mDatos, "|" & mMonO & "|") <> 0 Then
                        bOK = False: Exit Do
                    End If
                End If
                
                mDatos = mDatos & "|" & mMonO & "|"
                mCotizaEn = 0
                'Busco en que cotiza la moneda original
                For idx = LBound(arrData) To UBound(arrData)
                    If Val(Mid(arrData(idx), 1, InStr(arrData(idx), ":") - 1)) = mMonO Then
                        mCotizaEn = Val(Mid(arrData(idx), InStr(arrData(idx), ":") + 1))
                        Exit For
                    End If
                Next
                
                If mCotizaEn = 0 Then
                    bOK = False: Exit Do
                End If
                
                bOK = tc_Cotizacion(mMonO, mCotizaEn, Fecha, FechaTC, TipoTC, mretTC)
                If bOK Then
                    retTasaC = retTasaC * mretTC
                    mMonO = mCotizaEn
                Else
                    bOK = False: Exit Do
                End If
                
                If mMonO = MDestino Then
                    bOK = True: bSalir = True
                End If
            Loop
            
        End If
    End If
    
    If Not bOK Then
        retTasaC = 1
        FechaTC = ""
    End If
    
    TasadeCambio = Format(retTasaC, "#.000")
    FechaTC = Format(FechaTC, "dd/mm/yyyy")
    
    Exit Function
    
errTC:
End Function

Private Function tc_Cotizacion(mDe As Integer, mA As Integer, mFecha As Date, mFechaTC As String, mTipoTC As Integer, mTC As Currency) As Boolean

On Error GoTo errTC
Dim rsTC As rdoResultset
Dim mDate As String

    tc_Cotizacion = False
        
    Cons = "Select Top 1 '1' as Tipo, TCaFecha, TCaComprador, TCaVendedor " & _
              " From TasaCambio " & _
              " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " & _
                                            " Where TCaFecha < '" & Format(mFecha, "mm/dd/yyyy 23:59") & "'" & _
                                            " And TCaOriginal = " & mDe & _
                                            " And TCaDestino = " & mA & _
                                            " And TCaTipo = " & mTipoTC & ")" & _
              " And TCaOriginal = " & mDe & _
              " And TCaDestino = " & mA & _
              " And TCaTipo = " & mTipoTC
    
    Cons = Cons & " UNION ALL "
    
    Cons = Cons & "Select Top 1 '2' as Tipo, TCaFecha, 1/TCaComprador, 1/TCaVendedor " & _
              " From TasaCambio " & _
              " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " & _
                                            " Where TCaFecha < '" & Format(mFecha, "mm/dd/yyyy 23:59") & "'" & _
                                            " And TCaOriginal = " & mA & _
                                            " And TCaDestino = " & mDe & _
                                            " And TCaTipo = " & mTipoTC & ")" & _
              " And TCaOriginal = " & mA & _
              " And TCaDestino = " & mDe & _
              " And TCaTipo = " & mTipoTC
            
    Set rsTC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsTC.EOF Then
        tc_Cotizacion = True
        mTC = rsTC!TCaComprador
        mDate = rsTC!TCaFecha
        
        rsTC.MoveNext
        
        If Not rsTC.EOF Then
            If Format(mDate, "yyyymmdd") < Format(rsTC!TCaFecha, "yyyymmdd") Then
                mTC = rsTC!TCaComprador
                mDate = rsTC!TCaFecha
            End If
        End If
    
    End If
    rsTC.Close
    
    mFechaTC = mDate 'Format(mDate, "dd/mm/yyyy")
    Exit Function
    
errTC:
End Function


