Attribute VB_Name = "modStock"
'----------------------------------------------------------------------------------------------------
'   ModStock: Modulo de funciones de manejo del stock.
'----------------------------------------------------------------------------------------------------

Option Explicit

Public Sub MarcoStockVenta(lnUsuario As Long, lnArticulo As Long, cCantARetirar As Currency, cCantAEntregar As Currency, cCantReservado As Currency, m_TipoDoc As Integer, lnDocumento As Long, Optional iLocal As Long = -1)
Dim RsStock As rdoResultset

    Cons = "Select * From StockTotal" _
        & " Where StTArticulo = " & lnArticulo _
        & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
        & " And StTEstado = " & paEstadoArticuloEntrega
        
    Set RsStock = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsStock.EOF Then
        RsStock.AddNew
        RsStock!StTArticulo = lnArticulo
        RsStock!StTTipoEstado = TipoEstadoMercaderia.Fisico
        RsStock!StTEstado = paEstadoArticuloEntrega
        'Como no hay lo marco negativo.
        RsStock!StTCantidad = (cCantAEntregar + cCantARetirar + cCantReservado) * -1
        RsStock.Update
    Else
        RsStock.Edit
        RsStock!StTCantidad = RsStock!StTCantidad + ((cCantAEntregar + cCantARetirar + cCantReservado) * -1)
        RsStock.Update
    End If
    RsStock.Close
    
    If cCantARetirar <> 0 Then
    
        'En Stock Total guardamos el stock virtual.
        Cons = "Select * From StockTotal" _
            & " Where StTArticulo = " & lnArticulo _
            & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
            & " And StTEstado = " & TipoMovimientoEstado.ARetirar
            
        Set RsStock = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsStock.EOF Then
            RsStock.AddNew
            RsStock!StTArticulo = lnArticulo
            RsStock!StTTipoEstado = TipoEstadoMercaderia.Virtual
            RsStock!StTEstado = TipoMovimientoEstado.ARetirar
            'Como no hay lo marco negativo.
            RsStock!StTCantidad = cCantARetirar
            RsStock.Update
        Else
            RsStock.Edit
            RsStock!StTCantidad = RsStock!StTCantidad + cCantARetirar
            RsStock.Update
        End If
        RsStock.Close
    
        If m_TipoDoc = -1 Then
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantARetirar, TipoMovimientoEstado.ARetirar, 1, -1, -1, iLocal
        Else
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantARetirar, TipoMovimientoEstado.ARetirar, 1, m_TipoDoc, lnDocumento, iLocal
        End If
        
    End If
    
    If cCantAEntregar <> 0 Then
    
        'En Stock Total guardamos el stock virtual.
        Cons = "Select * From StockTotal" _
            & " Where StTArticulo = " & lnArticulo _
            & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
            & " And StTEstado = " & TipoMovimientoEstado.AEntregar
            
        Set RsStock = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsStock.EOF Then
            RsStock.AddNew
            RsStock!StTArticulo = lnArticulo
            RsStock!StTTipoEstado = TipoEstadoMercaderia.Virtual
            RsStock!StTEstado = TipoMovimientoEstado.AEntregar
            'Como no hay lo marco negativo.
            RsStock!StTCantidad = cCantAEntregar
            RsStock.Update
        Else
            RsStock.Edit
            RsStock!StTCantidad = RsStock!StTCantidad + cCantAEntregar
            RsStock.Update
        End If
        RsStock.Close
    
        If m_TipoDoc = -1 Then
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantAEntregar, TipoMovimientoEstado.AEntregar, 1, -1, -1, iLocal
        Else
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantAEntregar, TipoMovimientoEstado.AEntregar, 1, m_TipoDoc, lnDocumento, iLocal
        End If
    End If
    
    If cCantReservado <> 0 Then
        
    'En Stock Total guardamos el stock virtual.
        Cons = "Select * From StockTotal" _
            & " Where StTArticulo = " & lnArticulo _
            & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
            & " And StTEstado = " & TipoMovimientoEstado.Reserva
            
        Set RsStock = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If RsStock.EOF Then
            RsStock.AddNew
            RsStock!StTArticulo = lnArticulo
            RsStock!StTTipoEstado = TipoEstadoMercaderia.Virtual
            RsStock!StTEstado = TipoMovimientoEstado.Reserva
            'Como no hay lo marco negativo.
            RsStock!StTCantidad = cCantReservado
            RsStock.Update
        Else
            RsStock.Edit
            RsStock!StTCantidad = RsStock!StTCantidad + cCantReservado
            RsStock.Update
        End If
        RsStock.Close
        
        If m_TipoDoc = -1 Then
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantReservado, TipoMovimientoEstado.Reserva, 1, -1, -1, iLocal
        Else
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantReservado, TipoMovimientoEstado.Reserva, 1, m_TipoDoc, lnDocumento, iLocal
        End If
    End If

End Sub

Public Sub MarcoMovimientoStockFisico(lnUsuario As Long, iTipoLocal As Integer, iLocal As Long, lnArticulo As Long, cCantidad As Currency, iEstadoMercaderia As Integer, iAltaoBaja As Integer, Optional iTipoDocumento As Integer = -1, Optional lnDocumento As Long = -1)
        
Dim rsFis As rdoResultset

    Cons = "Select * from MovimientoStockFisico Where MSFCodigo = 0"
    Set rsFis = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsFis.AddNew
    
    rsFis!MSFFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsFis!MSFTipoLocal = iTipoLocal
    rsFis!MSFLocal = iLocal
    
    rsFis!MSFArticulo = lnArticulo
    rsFis!MSFCantidad = cCantidad * iAltaoBaja
    rsFis!MSFEstado = iEstadoMercaderia
    
    If iTipoDocumento <> -1 Then
        rsFis!MSFTipoDocumento = iTipoDocumento
        If lnDocumento <> -1 Then rsFis!MSFDocumento = lnDocumento Else rsFis!MSFDocumento = Null
    Else
        rsFis!MSFTipoDocumento = Null
        rsFis!MSFDocumento = Null
    End If
    
    rsFis!MSFUsuario = lnUsuario
    
    If paCodigoDeTerminal > 0 Then rsFis!MSFTerminal = paCodigoDeTerminal Else rsFis!MSFTerminal = Null
    
    rsFis.Update
    rsFis.Close
    
End Sub

'Public Sub MarcoMovimientoStockFisico(lnUsuario As Long, iTipoLocal As Integer, iLocal As Long, lnArticulo As Long, cCantidad As Currency, iEstadoMercaderia As Integer, iAltaoBaja As Integer, Optional iTipoDocumento As Integer = -1, Optional lnDocumento As Long = -1)
        
 '   cons = "Insert Into MovimientoStockFisico (MSFFecha, MSFTipoLocal, MSFLocal, MSFArticulo, MSFCantidad, MSFEstado, MSFTipoDocumento, MSFDocumento, MSFUsuario, MSFTerminal)" _
            & " Values( '" & Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss") & "'" _
            & ", " & iTipoLocal & ", " & iLocal _
            & ", " & lnArticulo & ", " & cCantidad * iAltaoBaja _
            & ", " & iEstadoMercaderia
            
  '  If iTipoDocumento <> -1 Then
   '     cons = cons & ", " & iTipoDocumento & ", " & lnDocumento
    'Else
     '   cons = cons & ", Null, Null"
'    End If
 '   cons = cons & ", " & lnUsuario
  '  If paCodigoDeTerminal > 0 Then cons = cons & ", " & paCodigoDeTerminal & ")" Else cons = cons & ", Null)"
   ' cBase.Execute (cons)

'End Sub

Public Sub MarcoMovimientoStockEstado(lnUsuario As Long, lnArticulo As Long, cCantidad As Currency, iEstadoMercaderia As Integer, iAltaoBaja As Integer, Optional iTipoDocumento As Integer = -1, Optional lnDocumento As Long = -1, Optional iLocal As Long = -1)

    Cons = "Insert Into MovimientoStockEstado (MSEFecha, MSEArticulo, MSECantidad, MSEEstado, MSETipoDocumento, MSEDocumento, MSELocal, MSEUsuario)" _
            & " values( '" & Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss") & "'" _
            & ", " & lnArticulo & ", " & cCantidad * iAltaoBaja _
            & ", " & iEstadoMercaderia
            
    If iTipoDocumento <> -1 Then
        Cons = Cons & ", " & iTipoDocumento & ", " & lnDocumento & ""
    Else
        Cons = Cons & ", Null, Null"
    End If
    If iLocal > 0 Then
        Cons = Cons & ", " & iLocal
    Else
        Cons = Cons & ", Null"
    End If
    Cons = Cons & ", " & lnUsuario & ")"
    
    cBase.Execute (Cons)

End Sub

'------------------------------------------------------------------------------------------------------------------------------------
'   MarcoMovimientoStockFisicoEnLocal
'       Realiza movimiento del STOCK en la tabla STOCK LOCAL
'------------------------------------------------------------------------------------------------------------------------------------
Public Sub MarcoMovimientoStockFisicoEnLocal(TipoLocal As Integer, CodigoLocal As Long, Articulo As Long, Cantidad As Currency, Estado As Integer, AltaOBaja As Integer)

Dim RsSLo As rdoResultset

    Cons = "Select * From StockLocal " _
            & " Where StLTipoLocal = " & TipoLocal & " And StlLocal = " & CodigoLocal _
            & " And StLArticulo = " & Articulo & " And StLEstado = " & Estado
    Set RsSLo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsSLo.EOF Then
        RsSLo.AddNew
        RsSLo!StLTipoLocal = TipoLocal
        RsSLo!StlLocal = CodigoLocal
        RsSLo!StLArticulo = Articulo
        RsSLo!StLEstado = Estado
        RsSLo!StLCantidad = Cantidad * AltaOBaja
        RsSLo.Update
    Else
        RsSLo.Edit
        RsSLo!StLCantidad = RsSLo!StLCantidad + (Cantidad * AltaOBaja)
        RsSLo.Update
    End If
    RsSLo.Close
    
End Sub

Public Sub MarcoMovimientoStockTotal(Articulo As Long, TipoEstado As Integer, Estado As Integer, Cantidad As Currency, AltaOBaja As Integer)
Dim RsSTo As rdoResultset
 
    Cons = "Select * From StockTotal" _
            & " Where StTArticulo = " & Articulo _
            & " And StTTipoEstado = " & TipoEstado _
            & " And StTEstado = " & Estado
            
    Set RsSTo = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsSTo.EOF Then
        RsSTo.AddNew
        RsSTo!StTArticulo = Articulo
        RsSTo!StTTipoEstado = TipoEstado
        RsSTo!StTEstado = Estado
        RsSTo!StTCantidad = Cantidad * AltaOBaja
        RsSTo.Update
    Else
        RsSTo.Edit
        RsSTo!StTCantidad = RsSTo!StTCantidad + (Cantidad * AltaOBaja)
        RsSTo.Update
    End If
    RsSTo.Close
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'   Esta rutina trabaja de dos formas 1) Procesa el stock para la devolucion de los articulos (Alta = False - defecto).
'                                                    2) Procesa el stock para la anulacion de la devolucion (Alta = True).
'   El Valor ALTA debe venir en TRUE si es Anulacion de una NOTA.
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MarcoStockXDevolucion(lnArticulo As Long, cCantARetirar As Currency, cCantTotalDevueltos As Currency, iTipoLocal As Integer, lnLocal As Long, lnUsuario As Long, Optional iTipoDocumento As Integer = -1, Optional lnDocumento As Long = -1, Optional Alta As Boolean = False)

Dim cFisicosDev As Currency     'Cantidad de Articulos Físicos devueltos (Q' tenia el cliente)
    
    cFisicosDev = cCantTotalDevueltos - cCantARetirar   'Los Q' tenia el cliente
    
    If Not Alta Then    '------------------------------------------------------------------------------------------------------------
        'MUEVO EL STOCK POR UNA DEVOLUCION
        'Si hay a retirar entonces se los quito al stock total.
        If cCantARetirar > 0 Then
                MarcoMovimientoStockTotal lnArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, cCantARetirar, -1
                MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantARetirar, TipoMovimientoEstado.ARetirar, -1, iTipoDocumento, lnDocumento, lnLocal
        End If
        
        MarcoMovimientoStockTotal lnArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, cCantTotalDevueltos, 1
        
        If cFisicosDev > 0 Then
            MarcoMovimientoStockFisico lnUsuario, iTipoLocal, lnLocal, lnArticulo, cFisicosDev, paEstadoArticuloEntrega, 1, iTipoDocumento, lnDocumento
            MarcoMovimientoStockFisicoEnLocal iTipoLocal, lnLocal, lnArticulo, cFisicosDev, paEstadoArticuloEntrega, 1
        End If
    
    Else                    '------------------------------------------------------------------------------------------------------------
        'MUEVO EL STOCK POR UNA ANULACION DE DEVOLUCION
        If cCantARetirar > 0 Then
            MarcoMovimientoStockTotal lnArticulo, TipoEstadoMercaderia.Virtual, TipoMovimientoEstado.ARetirar, cCantARetirar, 1
            MarcoMovimientoStockEstado lnUsuario, lnArticulo, cCantARetirar, TipoMovimientoEstado.ARetirar, 1, iTipoDocumento, lnDocumento, lnLocal
        End If
        
        MarcoMovimientoStockTotal lnArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, cCantTotalDevueltos, -1
        
        If cFisicosDev > 0 Then
            MarcoMovimientoStockFisico lnUsuario, iTipoLocal, lnLocal, lnArticulo, cFisicosDev, paEstadoArticuloEntrega, -1, iTipoDocumento, lnDocumento
            MarcoMovimientoStockFisicoEnLocal iTipoLocal, lnLocal, lnArticulo, cFisicosDev, paEstadoArticuloEntrega, -1
        End If
    End If                  '------------------------------------------------------------------------------------------------------------
                
End Sub

Public Function NombreDocumento(Tipo As Integer)
    
Dim aRetorno As String
    
    aRetorno = ""
    
    Select Case Tipo
        Case TipoDocumento.Contado: aRetorno = "Contado"
        Case TipoDocumento.Credito: aRetorno = "Crédito"
        Case TipoDocumento.NotaCredito: aRetorno = "NCrédito"
        Case TipoDocumento.NotaDevolucion: aRetorno = "NDevol."
        Case TipoDocumento.ReciboDePago: aRetorno = "Recibo"
        Case TipoDocumento.Remito: aRetorno = "Remito"
        Case TipoDocumento.ContadoDomicilio, TipoDocumento.CreditoDomicilio: aRetorno = "Vta.Tel"
        
        'Documentos de Compra de Mercadería
        Case TipoDocumento.CompraCarta: aRetorno = "C-Carta"
        Case TipoDocumento.CompraContado: aRetorno = "C-Contado"
        Case TipoDocumento.CompraCredito: aRetorno = "C-Crédito"
        Case TipoDocumento.CompraNotaCredito: aRetorno = "C-NCrédito"
        Case TipoDocumento.CompraNotaDevolucion: aRetorno = "C-NDevol."
        Case TipoDocumento.CompraRemito: aRetorno = "C-Remito"
        Case TipoDocumento.CompraCarpeta: aRetorno = "Importación"
        
        Case TipoDocumento.Envios: aRetorno = "Reparto"
        Case TipoDocumento.Traslados: aRetorno = "Traslado"
        Case TipoDocumento.CambioEstadoMercaderia: aRetorno = "Cambio Estado Merc."
        
    End Select
        
    NombreDocumento = aRetorno
    
End Function

