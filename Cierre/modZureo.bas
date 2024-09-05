Attribute VB_Name = "modZureo"
Option Explicit

Private Type typDIA
    Fecha As Date
    Valor1 As Currency
    Valor2 As Currency
End Type
Dim arrDIAS() As typDIA

Public prmMonedaContabilidad As Integer
Public prmMonedaDisp As Long
Public prmErrorText As String

Public mSQL As String

Public prmCtaCajaPesos As Long
Public prmCtaVtasContado As Long
Public prmCtaDeudoresXVenta As Long
Public prmCtaVtasCredito As Long

Public prmCtaIVA_Venta As Long

Public prmCtaMoraCuotas As Long
Public prmCtaSeñasRecibidas As Long

Dim rsCG As rdoResultset
Dim mNewID As Long
Dim txtError As String, mTXT As String

Private Enum enuRelacion
    Moneda = 1
    TipoDocumento = 2
    Contactos = 3
    Cuentas = 4
    CtaIVA_Compra = 5
    CtaCOFIS_Compra = 6
    CtaAcreedoresVarios = 7
    TDocTransferencia = 8
    TDocSalidaCaja = 9
    TDocVentasContado = 10
    CtaCajaPesos = 11
    CtaVtasContado = 12
    TDocVentasCredito = 13
    CtaDeudoresXVenta = 14
    CtaVtasCredito = 15
    
    CtaIVA_Venta = 16
    CtaCOFIS_Venta = 17
    TDocVentasNotasCredito = 18
    TDocVentasNotasContado = 19
    TDocVentasRecibo = 20
    TDocEntradaCaja = 21
    CtaMoraCuotas = 22
    CtaSeñasRecibidas = 23
End Enum

Public C_KEY_MEMO  As String

Public Function CargoDatosEmpresa() As Boolean
On Error GoTo errCDE
   
    CargoDatosEmpresa = False
    prmMonedaContabilidad = 1

    mSQL = "Select * from ZureoCGSA Where Tipo NOT  IN (1,2,3,4)"
    Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case RsAux!Tipo

            Case enuRelacion.CtaVtasContado: prmCtaVtasContado = RsAux!IDZureo
            
            Case enuRelacion.CtaDeudoresXVenta: prmCtaDeudoresXVenta = RsAux!IDZureo
            Case enuRelacion.CtaVtasCredito: prmCtaVtasCredito = RsAux!IDZureo
                       
            Case enuRelacion.CtaMoraCuotas: prmCtaMoraCuotas = RsAux!IDZureo
            Case enuRelacion.CtaSeñasRecibidas: prmCtaSeñasRecibidas = RsAux!IDZureo
            
            Case enuRelacion.CtaIVA_Venta: prmCtaIVA_Venta = RsAux!IDZureo
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    CargoDatosEmpresa = True
    Exit Function
    
errCDE:
End Function

Public Function errText() As String
    errText = Err.Number & " - " & Err.Description
End Function

Public Function CGSA_VentasContado(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
        
    prmErrorText = "Vtas Cdo: ERROR "
        
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis" & _
                " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.Contado & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.Contado, C_KEY_MEMO & "Ventas Contado", False, _
                                        prmCtaVtasContado, NETO, IVA, COFIS, _
                                        dCtaDisponibilidad, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function CGSA_VentasCredito(dFecha As Date, dSucursales As String) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
        
    prmErrorText = "Vtas Credito: ERROR "
        
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis" & _
                " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.Credito & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.Credito, C_KEY_MEMO & "Ventas Crédito", False, _
                                        prmCtaVtasCredito, NETO, IVA, COFIS, _
                                        prmCtaDeudoresXVenta, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function CGSA_VentasCreditoNotas(dFecha As Date, dSucursales As String) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
        
    prmErrorText = "Vtas Credito: ERROR "
        
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis" & _
                " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaCredito & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.NotaCredito, C_KEY_MEMO & "Notas de Crédito", True, _
                                        prmCtaVtasCredito, NETO, IVA, COFIS, _
                                        prmCtaDeudoresXVenta, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function CGSA_VentasContadoNotas(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
        
    prmErrorText = "Vtas N. Cdo: ERROR "
        
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis" & _
                " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 " & _
                " And DocTipo = " & modComun.TipoDocumento.NotaDevolucion & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.NotaDevolucion, C_KEY_MEMO & "Notas de Devolucion", True, _
                                        prmCtaVtasContado, NETO, IVA, COFIS, _
                                        dCtaDisponibilidad, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function CGSA_VentasContadoNotasE(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
        
    prmErrorText = "Vtas N. Cdo: ERROR "
        
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis" & _
                " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 " & _
                " And DocTipo = " & modComun.TipoDocumento.NotaEspecial & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.NotaEspecial, C_KEY_MEMO & "Notas Especiales", True, _
                                        prmCtaVtasContado, NETO, IVA, COFIS, _
                                        dCtaDisponibilidad, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function


Public Function CGSA_Cobranza(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
    
    prmErrorText = "Cuotas: ERROR "
    
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, " & _
                        " Sum(DPaAmortizacion) Total, 0 as Iva, 0 as Cofis" & _
                " From Documento (Index = iTipoFechaSucursalMoneda), DocumentoPago  " & _
                " Where DocCodigo = DPaDocQSalda " & _
                " And DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.ReciboDePago & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        
        If Not IsNull(rsCG("Cofis").Value) Then COFIS = rsCG("Cofis").Value Else COFIS = 0
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA - COFIS
        mQ = mQ + 1
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.ReciboDePago, C_KEY_MEMO & "Cobranza de Cuotas", False, _
                                        prmCtaDeudoresXVenta, NETO, IVA, COFIS, _
                                        dCtaDisponibilidad, (NETO + IVA + COFIS))
    
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function CGSA_CobranzaMorasAFecha(ByVal dFecha As Date, ByVal dSucursales As String) As clsCantidadImporte
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
    
    prmErrorText = "Moras: ERROR "
    
    Set CGSA_CobranzaMorasAFecha = New clsCantidadImporte
    
    mSQL = " Select SUM(DocTotal) as Total " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaDebito & _
            " And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)"

    mSQL = mSQL & _
                    " Union All " & _
            " Select SUM(DocTotal) * -1 as Total " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaCreditoMora & _
            " And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        CGSA_CobranzaMorasAFecha.Importe = CGSA_CobranzaMorasAFecha.Importe + rsCG("Total").Value
        rsCG.MoveNext
    Loop
    rsCG.Close


    Cons = "Select Count(*) Cantidad from Documento, DocumentoPago " _
                & " Where DocFecha Between '" & Format(dFecha, "mm/dd/yyyy") & " 00:00' And '" & Format(dFecha, "mm/dd/yyyy") & " 23:59'"
        Cons = Cons & " And DocSucursal IN (" & dSucursales & ")"
        Cons = Cons & " And DocTipo = " & modComun.TipoDocumento.ReciboDePago _
                            & " And DocMoneda = " & Moneda _
                            & " And DocCodigo = DPaDocQSalda and DPaMora <> 0 AND DPaDocASaldar NOT IN (SELECT DocCodigo FROM Documento WHERE DocTipo = 45)"
    Set rsCG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CGSA_CobranzaMorasAFecha.Cantidad = rsCG(0)
    rsCG.Close

End Function




Public Function CGSA_CobranzaMoras(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
    
    prmErrorText = "Moras: ERROR "
    
    ReDim arrDIAS(0)
    
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, " & _
                    " 0 AS IVA, SUM(DPaMora) as Total " & _
            " From Documento (Index = iTipoFechaSucursalMoneda), DocumentoPago  " & _
            " Where DocCodigo = DPaDocQSalda " & _
            " And DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.ReciboDePago & _
            " And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " And DPaMora <> 0 And DocIVA <> 0 " & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)" & _
                    " Union All " & _
            " Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, " & _
                        " Sum(DocIVA) AS IVA, SUM(DocTotal) as Total " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaDebito & _
            " And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)"

    mSQL = mSQL & _
                    " Union All " & _
            " Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, " & _
                        " Sum(DocIVA) * -1 AS IVA, SUM(DocTotal) * -1 as Total " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.NotaCreditoMora & _
            " And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)"

    mSQL = mSQL & _
                    " Union All " & _
            " Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano, " & _
                        " Sum(DocIVA) AS IVA, 0 as Total " & _
            " From Documento(Index = iTipoFechaSucursalMoneda) " & _
            " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
            " And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.ReciboDePago & _
            " And DocIVA <> 0 And DocMoneda = " & prmMonedaDisp & " And DocSucursal IN (" & dSucursales & ")" & _
            " Group by Datepart(dd, DocFecha), Datepart(mm, DocFecha), Datepart(yy, DocFecha)"


    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
    
        If Not IsNull(rsCG("IVA").Value) Then IVA = rsCG("IVA").Value Else IVA = 0
        NETO = rsCG("Total").Value - IVA
        
        add_Array fAuxiliar, NETO, IVA
        mQ = mQ + 1
        
        rsCG.MoveNext
    Loop
    rsCG.Close
    
    If mQ <> 0 Then
        Dim idx As Integer
        mQ = UBound(arrDIAS)
        For idx = 0 To mQ
        
        mQOK = mQOK + fnc_AltaComprobante(cBase, arrDIAS(idx).Fecha, modComun.TipoDocumento.Contado, C_KEY_MEMO & "Cobranza de Moras", False, _
                                prmCtaMoraCuotas, arrDIAS(idx).Valor1, arrDIAS(idx).Valor2, 0, _
                                dCtaDisponibilidad, (arrDIAS(idx).Valor1 + arrDIAS(idx).Valor2))
        Next
    End If

End Function

Public Function CGSA_SeñasRecibidas(dFecha As Date, dSucursales As String, dCtaDisponibilidad As Long) As String
Dim mQ As Long, mQOK As Long
Dim fAuxiliar As Date, NETO As Currency, IVA As Currency, COFIS As Currency
    
    prmErrorText = "Señas: ERROR "
    
    mSQL = "Select Datepart(dd, DocFecha) as Dia, Datepart(mm, DocFecha) as Mes, Datepart(yy, DocFecha) as Ano,  Sum(DocIVA) AS IVA, SUM(DocTotal) as Total " & _
                 " From Documento(Index = iTipoFechaSucursalMoneda) " & _
                " Where DocFecha Between " & Format(dFecha, "'yyyy/mm/dd'") & " And " & Format(dFecha, "'yyyy/mm/dd 23:59'") & _
                " And DocIVA = 0 And DocAnulado = 0 And DocTipo = " & modComun.TipoDocumento.ReciboDePago & _
                " And DocMoneda = " & prmMonedaDisp & _
                " And DocSucursal IN (" & dSucursales & ")" & _
                " And DocCodigo NOT IN ( Select DPaDocQSalda from DocumentoPago) " & _
                " Group by Datepart(yy, DocFecha), Datepart(mm, DocFecha), Datepart(dd, DocFecha)"

    Set rsCG = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    Do While Not rsCG.EOF
        
        fAuxiliar = CDate(rsCG!Dia & "/" & rsCG!Mes & "/" & rsCG!Ano)
        NETO = rsCG("Total").Value
        mQ = mQ + 1
        
        mQOK = mQOK + fnc_AltaComprobante(cBase, fAuxiliar, modComun.TipoDocumento.CompraEntradaCaja, C_KEY_MEMO & "Señas Recibidas", False, _
                                        prmCtaSeñasRecibidas, NETO, 0, 0, _
                                        dCtaDisponibilidad, NETO)
        rsCG.MoveNext
    Loop
    rsCG.Close
    
End Function

Public Function fnc_AltaComprobante(ByVal RDOZUREO As rdoConnection, _
                        dFecha As Date, dTipoComp As Integer, dMemo As String, dHaceSalidaCaja As Boolean, _
                        dCuenta1 As Long, dICta1 As Currency, dIIVA As Currency, dICOFIS As Currency, _
                        dContraCuenta As Long, dICCta As Currency, _
                        Optional dICta1ME As Currency = 0, _
                        Optional dTC As Double = 1) As Byte
Dim dMCta1 As Integer

    prmErrorText = "dCuenta1=" & dCuenta1 & " dContraCuenta=" & dContraCuenta & " Memo=" & dMemo
    
    dMCta1 = 0
    If dICta1ME <> 0 Then
        Cons = "Select CueMoneda From ZureoCGSA.dbo.cceCuentas Where CueID = " & dCuenta1
        Set RsAux = RDOZUREO.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CueMoneda) Then dMCta1 = RsAux!CueMoneda
        End If
        RsAux.Close
    End If
    
    '0) Autonumerico en Tabla cceComprobantes      ----------------------------------------------------------------------
    mNewID = -1
    Cons = "Select * from ZureoCGSA.dbo.genAutonumerico Where AutTabla = 'cceComprobantes'"
    Set RsAux = RDOZUREO.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        mNewID = RsAux!AutContador + 1
        RsAux.Edit
        RsAux!AutContador = mNewID
        RsAux.Update
    End If
    RsAux.Close
    If mNewID = -1 Then Err.Raise 8000, "DBFncs", "Resultado de la función get_TableCounter = -1"
        
    '1) Cabezal con los datos del Comprobante   ----------------------------------------------------------------------
    mSQL = "Select * from ZureoCGSA.dbo.cceComprobantes Where ComID = 0"
    Set RsAux = RDOZUREO.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!ComIDEmpresa = 1
    RsAux!ComID = mNewID
    RsAux!ComProveedor = Null
    RsAux!ComFecha = dFecha
    
    RsAux!ComMoneda = prmMonedaDisp 'prmMonedaContabilidad
    RsAux!ComTipo = dTipoComp
            
    RsAux!ComTotal = dICta1 + dIIVA
    RsAux!ComTC = dTC
    RsAux!ComFechaModificacion = Now
    RsAux!ComMemo = IIf(dMemo = "", Null, dMemo)

    RsAux!ComSaldoCero = Null
    RsAux.Update
    RsAux.Close
    
    '2) Paso las cuentas asignadas al comprobante (en CGSA estan separadas) ---------------------------------------------------
    mSQL = "Select * from  ZureoCGSA.dbo.cceComprobanteCuenta Where CCuIDComprobante = " & mNewID
    Set RsAux = RDOZUREO.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)

    RsAux.AddNew
    RsAux!CCuIDComprobante = mNewID
    RsAux!CCuIDCuenta = dCuenta1
    RsAux!CCuIDProyecto = 0: RsAux!CCuIDDepartamento = 0: RsAux!CCuIDReferencia = 0
    RsAux!CCuMoneda = dMCta1 '0
    
    RsAux!CCuImporteCuenta = IIf(dICta1ME <> 0, dICta1ME, dICta1)
    RsAux!CCuDebe = IIf(dHaceSalidaCaja, dICta1, Null)
    RsAux!CCuHaber = IIf(Not dHaceSalidaCaja, dICta1, Null)
    RsAux.Update
    
    If dIIVA <> 0 Then
        RsAux.AddNew
        RsAux!CCuIDComprobante = mNewID
        RsAux!CCuIDCuenta = prmCtaIVA_Venta
        RsAux!CCuIDProyecto = 0: RsAux!CCuIDDepartamento = 0: RsAux!CCuIDReferencia = 0
        RsAux!CCuMoneda = 0
        
        RsAux!CCuImporteCuenta = dIIVA
        RsAux!CCuDebe = IIf(dHaceSalidaCaja, dIIVA, Null)
        RsAux!CCuHaber = IIf(Not dHaceSalidaCaja, dIIVA, Null)
        RsAux.Update
    End If
    
    '3) Contra cuentas  --------------------------------------------------------------------------
    RsAux.AddNew
    RsAux!CCuIDComprobante = mNewID
    RsAux!CCuIDCuenta = dContraCuenta
    RsAux!CCuIDProyecto = 0: RsAux!CCuIDDepartamento = 0: RsAux!CCuIDReferencia = 0
    RsAux!CCuMoneda = prmMonedaDisp
    
    RsAux!CCuImporteCuenta = dICCta
    RsAux!CCuDebe = IIf(Not dHaceSalidaCaja, dICCta, Null)
    RsAux!CCuHaber = IIf(dHaceSalidaCaja, dICCta, Null)
    RsAux.Update
        
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------------------------------
    
End Function

Private Function add_Array(xFecha As Date, xValor1 As Currency, xValor2 As Currency)

Dim idx As Integer, bAddOK As Boolean
    
    If CDate(arrDIAS(0).Fecha) < CDate("01/01/1980") Then
        With arrDIAS(0)
            .Fecha = xFecha
            .Valor1 = xValor1
            .Valor2 = xValor2
        End With
        bAddOK = True
    End If
    If bAddOK Then Exit Function
    
    For idx = LBound(arrDIAS) To UBound(arrDIAS)
        If arrDIAS(idx).Fecha = xFecha Then
            arrDIAS(idx).Valor1 = arrDIAS(idx).Valor1 + xValor1
            arrDIAS(idx).Valor2 = arrDIAS(idx).Valor2 + xValor2
            bAddOK = True
        End If
    Next
    
    If bAddOK Then Exit Function
    idx = UBound(arrDIAS) + 1
    ReDim Preserve arrDIAS(idx)
    With arrDIAS(idx)
        .Fecha = xFecha
        .Valor1 = xValor1
        .Valor2 = xValor2
    End With
    
End Function
