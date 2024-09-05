Attribute VB_Name = "modListados"
Option Explicit

'Definicion de variables para seteo de impresora
'   B: bandeja, N:Nombre
Public paIContadoB As Integer
Public paIContadoN As String

Public paICreditoB As Integer
Public paICreditoN As String

Public paINCreditoB As Integer
Public paINCreditoN As String

Public paINContadoB As Integer
Public paINContadoN As String

Public paINEspecialB As Integer
Public paINEspecialN As String

Public paIReciboB As Integer
Public paIReciboN As String

Public paIConformeB As Integer
Public paIConformeN As String

Public paIRemitoB As Integer
Public paIRemitoN As String

Public paICartaB As Integer
Public paICartaN As String
      
'------------------------------------------------------------------------------------------------------------------------------------
'   Carga los Nombres de los documentos y las impresoras (por defecto) para cada documento.
'   En la BD se guarda solo el nombre de la impresora (Device Name), y aca cargo los demas datos.
'       DriverName, Port y Bandeja
'------------------------------------------------------------------------------------------------------------------------------------
Public Sub CargoParametrosImpresion(Sucursal As Long, Optional VerCtdo As Boolean = True, Optional VerCred As Boolean = True, _
                                                                Optional VerRePa As Boolean = True, Optional VerNDCo As Boolean = True, Optional VerNDCr As Boolean = True, _
                                                                Optional VerNDEs As Boolean = True, Optional VerConf As Boolean = True, Optional VerRemi As Boolean = True, _
                                                                Optional VerCarta As Boolean = True)

Dim X As Printer
Dim mMsgPrinter As String
    On Error GoTo errImp
    'Dado el Código de Sucursal se sacan los nombres de los documentos para cargar los parámetros.
    paDContado = "": paDCredito = "": paDNDevolucion = "": paDNCredito = "": paDRecibo = "": paDNEspecial = ""
    
    paIContadoN = "": paICreditoN = "": paINContadoN = "": paINCreditoN = "": paIReciboN = "": paINEspecialN = "": paIConformeN = "": paIRemitoN = "": paICartaN = ""
    paIContadoB = -1: paICreditoB = -1: paINContadoB = -1: paINCreditoB = -1: paIReciboB = -1: paINEspecialB = -1: paIConformeB = -1: paIRemitoB = -1: paICartaB = -1
    
    Cons = "Select * From Sucursal Where SucCodigo = " & Sucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        'Nombre de Cada Documento--------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then paDCredito = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then paDNCredito = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux!SucDNEspecial) Then paDNEspecial = Trim(RsAux!SucDNEspecial)
        
        If Not IsNull(RsAux!SucDNDebito) Then paDNDebito = Trim(RsAux!SucDNDebito)
        
        'Parametros de las Impresoras------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucICoNombre) Then          'CONTADO
            paIContadoN = Trim(RsAux!SucICoNombre)
            If Not IsNull(RsAux!SucICoBandeja) Then paIContadoB = RsAux!SucICoBandeja
            If Not VerificoQueExistaImpresora(paIContadoN) Then
                mMsgPrinter = mMsgPrinter & Trim(paIContadoN) & " para imprimir documentos Contado." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucICrNombre) Then          'CREDITO
            paICreditoN = Trim(RsAux!SucICrNombre)
            If Not IsNull(RsAux!SucICrBandeja) Then paICreditoB = RsAux!SucICrBandeja
            If Not VerificoQueExistaImpresora(paICreditoN) Then
                mMsgPrinter = mMsgPrinter & Trim(paICreditoN) & " para imprimir documentos Crédito." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucIReNombre) Then          'RECIBO DE PAGO
            paIReciboN = Trim(RsAux!SucIReNombre)
            If Not IsNull(RsAux!SucIReBandeja) Then paIReciboB = RsAux!SucIReBandeja
            If Not VerificoQueExistaImpresora(paIReciboN) Then
                mMsgPrinter = mMsgPrinter & Trim(paIReciboN) & " para imprimir Recibos de Cuotas." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucINdNombre) Then          'NOTA CONTADO
            paINContadoN = Trim(RsAux!SucINdNombre)
            If Not IsNull(RsAux!SucINdBandeja) Then paINContadoB = RsAux!SucINdBandeja
            If Not VerificoQueExistaImpresora(paINContadoN) Then
                mMsgPrinter = mMsgPrinter & Trim(paINContadoN) & " para imprimir Notas de Devolución." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucINcNombre) Then          'NOTA CREDITO
            paINCreditoN = Trim(RsAux!SucINcNombre)
            If Not IsNull(RsAux!SucINcBandeja) Then paINCreditoB = RsAux!SucINcBandeja
            If Not VerificoQueExistaImpresora(paINCreditoN) Then
                mMsgPrinter = mMsgPrinter & Trim(paINCreditoN) & " para imprimir Notas de Crédito." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucINeNombre) Then          'NOTA ESPECIAL
            paINEspecialN = Trim(RsAux!SucINeNombre)
            If Not IsNull(RsAux!SucINeBandeja) Then paINEspecialB = RsAux!SucINeBandeja
            If Not VerificoQueExistaImpresora(paINEspecialN) Then
                mMsgPrinter = mMsgPrinter & Trim(paINEspecialN) & " para imprimir Notas Especiales." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucICnNombre) Then          'CONFORME
            paIConformeN = Trim(RsAux!SucICnNombre)
            If Not IsNull(RsAux!SucICnBandeja) Then paIConformeB = RsAux!SucICnBandeja
            If Not VerificoQueExistaImpresora(paIConformeN) Then
                mMsgPrinter = mMsgPrinter & Trim(paIConformeN) & " para imprimir Conformes." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucIRmNombre) Then          'REMITO
            paIRemitoN = Trim(RsAux!SucIRmNombre)
            If Not IsNull(RsAux!SucIRmBandeja) Then paIRemitoB = RsAux!SucIRmBandeja
            If Not VerificoQueExistaImpresora(paIRemitoN) Then
                mMsgPrinter = mMsgPrinter & Trim(paIRemitoN) & " para imprimir Remitos." & vbCrLf
            End If
        End If
        If Not IsNull(RsAux!SucICaNombre) Then          'Carta.
            paICartaN = Trim(RsAux!SucICaNombre)
            If Not IsNull(RsAux!SucIRmBandeja) Then paICartaB = RsAux!SucICaBandeja
            If Not VerificoQueExistaImpresora(paICartaN) Then
                mMsgPrinter = mMsgPrinter & Trim(paICartaN) & " para imprimir Hojas Carta." & vbCrLf
            End If
        End If
       '------------------------------------------------------------------------------------------------------------------
    End If
    RsAux.Close
    
    If Trim(mMsgPrinter) <> "" Then
        MsgBox "Se deben instalar las siguientes impresoras en su computador: " & vbCrLf & vbCrLf & _
                    mMsgPrinter, vbExclamation, "Instalación de Impresoras"
    End If
    Exit Sub
        
errImp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
End Sub

'------------------------------------------------------------------------------------------------------------------------------------
'   Setea la impresora pasada como parámetro como: por defecto
'------------------------------------------------------------------------------------------------------------------------------------
Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Public Sub VerificoParametrosDocumentos( _
        Optional Contado As Boolean = False, Optional Credito As Boolean = False, Optional NContado As Boolean = False, _
        Optional NCredito As Boolean = False, Optional NEspecial As Boolean = False, Optional Recibo As Boolean = False, _
        Optional Remito As Boolean = False, Optional Conforme As Boolean = False)
    
    If Contado Then
        If paDContado = "" Then
            MsgBox "La sucursal no tiene asociado el documento para facturar Contados." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paIContadoN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para facturar Contados." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If Credito Then
        If paDCredito = "" Then
            MsgBox "La sucursal no tiene asociado el documento para facturar Creditos." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paICreditoN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para facturar Créditos." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If NContado Then
        If paDNDevolucion = "" Then
            MsgBox "La sucursal no tiene asociado el documento para facturar Notas de Devolución." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paINContadoN = "" Then
            MsgBox "La sucursal a no tiene asociada la impresora por defecto para facturar Notas de Devolución." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If NCredito Then
        If paDNCredito = "" Then
            MsgBox "La sucursal no tiene asociado el documento para facturar Notas de Crédito." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paINCreditoN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para facturar Notas de Crédito." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If Recibo Then
        If paDRecibo = "" Then
            MsgBox "La sucursal no tiene asociado el documento para emitir Recibos de Pago." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paIReciboN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para facturar Recibos de Pago." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If

    If NEspecial Then
        If paDNEspecial = "" Then
            MsgBox "La sucursal no tiene asociado el documento para emitir Notas Especiales." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
        If paINEspecialN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para facturar Notas Especiales." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If Remito Then
        If paIRemitoN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para emitir Remitos de Mercadería." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
    If Conforme Then
        If paIConformeN = "" Then
            MsgBox "La sucursal no tiene asociada la impresora por defecto para emitir Conformes." & Chr(vbKeyReturn) _
                & "Comunique el error a su administrador de Base de Datos.", vbCritical, "ATENCIÓN"
        End If
    End If
    
End Sub

Public Sub CargoDatosZoom(cCombo As Control, Optional Maximo As Integer = 120)

    Dim aValor As Integer
    aValor = 10
    Do While aValor <= 120
        cCombo.AddItem aValor
        aValor = aValor + 10
    Loop
    cCombo.ListIndex = 0
    
End Sub

Public Sub EnumeroPiedePagina(vsPrint As Control)

    With vsPrint
        .Columns = 1
        .FontSize = 12
        .Font = "Arial"
        .FontItalic = False
        .TextAlign = 2  'taRight

        For i = 1 To .PageCount
            .StartOverlay i
            
            .CurrentX = .MarginLeft
            .CurrentY = .PageHeight - .MarginBottom + 150
            vsPrint = "Página " & Format(i) & " de " & Format(.PageCount)
            
            .EndOverlay
        Next
        .TextAlign = 0 'taLeft
    End With
    
End Sub

Public Function FormateoString(vsPrint As Control, cLargoCelda As Currency, strTexto As String) As String
Dim largo As Currency

    largo = vsPrint.TextWidth(strTexto)
    Do While vsPrint.TextWidth(strTexto) > cLargoCelda
        strTexto = Mid(strTexto, 1, Len(strTexto) - 1)
    Loop
    FormateoString = strTexto
    
End Function

Public Sub Zoom(vsPrint As Control, Valor As Integer)

     On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    vsPrint.Zoom = Val(Valor)
    Screen.MousePointer = 0
    vsPrint.Visible = True
    
End Sub

Public Sub ZoomIn(vsPrint As Control)
        
    On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    If vsPrint.Zoom > 10 Then vsPrint.Zoom = vsPrint.Zoom - 10
    vsPrint.Visible = True
    Screen.MousePointer = 0
    
End Sub

Public Sub ZoomOut(vsPrint As Control)
        
    On Error Resume Next
    Screen.MousePointer = 11
    vsPrint.Visible = False
    If vsPrint.Zoom < 120 Then vsPrint.Zoom = vsPrint.Zoom + 10
    vsPrint.Visible = True
    Screen.MousePointer = 0
    
End Sub

Public Sub IrAPagina(vsPrint As Control, Pagina As Integer)
    vsPrint.PreviewPage = Pagina
End Sub
Public Sub EncabezadoListado(vsPrint As Control, strTitulo As String, sNombreEmpresa As Boolean)
    
    With vsPrint
        .HdrFont = "Arial"
        .HdrFontSize = 12
        .HdrFontBold = False
    End With
    
    If sNombreEmpresa Then
        vsPrint.Header = strTitulo & "||Carlos Gutiérrez S.A."
    Else
        vsPrint.Header = strTitulo
    End If
    
    vsPrint.Footer = Format(Now, "dd/mm/yy hh:mm")
    
End Sub

Public Function VerificoQueExistaImpresora(NombreImpresora As String) As Boolean
On Error GoTo ErrVQEI
Dim X As Printer
    VerificoQueExistaImpresora = False
    For Each X In Printers
        If Trim(X.DeviceName) = Trim(NombreImpresora) Then
            VerificoQueExistaImpresora = True
            Exit For
        End If
    Next
    Exit Function
ErrVQEI:
    clsGeneral.OcurrioError "Ocurrio un error al verificar que existan las impresoras asignadas.", Err.Description
End Function
