Attribute VB_Name = "Funciones"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim RsAuxiliar As rdoResultset

'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'--------------------------------------------------------------------------------------------------------
'   PROCEDMIENTO CargoCombo: Carga el combo con los datos de la consulta pasada
'   como parámetro.
'
'   PARÁMETROS:
'       Cons: Cosulta seleccionando los datos a cargar - RS(0) = Codigo, RS(1) = Dato.
'       Combo: Combo a cargar.
'       Seleccionado: Dato a seleccionar por defecto (Texto).
'--------------------------------------------------------------------------------------------------------
Public Sub CargoCombo(Consulta As String, Combo As Control, Seleccionado As String)

Dim iSel As Integer     'Guardo el indice del seleccionado
    
    iSel = -1
    Combo.Clear
    
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then
            iSel = Combo.ListCount
        End If
        
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then
        Combo.ListIndex = iSel
    Else
        Combo.ListIndex = iSel - 1
    End If
    
End Sub

'--------------------------------------------------------------------------------------------------------
'   PROCEDMIENTO BuscoCodigoEnCombo: Busca un el codigo pasado como parámetro dentro del itemData del combo.
'
'   PARÁMETROS:
'       lngCodigo: Codigo a buscar.
'
'   RETORNA:
'       Si encuentra el dato, setea automáticamente el combo, sino lo marca en vacio.
'--------------------------------------------------------------------------------------------------------

Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)

    If cCombo.ListCount > 0 And lngCodigo > 0 Then
        
        For I = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(I) = lngCodigo Then
                cCombo.ListIndex = I
                Exit Sub
            End If
        Next I
        cCombo.ListIndex = -1
        
    Else
    
        cCombo.ListIndex = -1
        
    End If

End Sub

Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function

Public Sub Foco(C As Control)
    
    On Error Resume Next
    If C.Enabled Then
        C.SelStart = 0
        C.SelLength = Len(C.Text)
        C.SetFocus
    End If
    
End Sub

Public Function TelefonoATexto(Cliente As Long) As String
Dim RsTel As rdoResultset
Dim aTelefonos As String

    Cons = "Select * from Telefono, TipoTelefono" _
        & " Where TelCliente = " & Cliente _
        & " And TelTipo = TTeCodigo"
    Set RsTel = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsTel.EOF Then
        aTelefonos = ""
        Do While Not RsTel.EOF
            aTelefonos = aTelefonos & Trim(RsTel!TTeNombre) & ": " & Trim(RsTel!TelNumero)
            If Not IsNull(RsTel!TelInterno) Then aTelefonos = aTelefonos & "(" & Trim(RsTel!TelInterno) & ")"
            aTelefonos = aTelefonos & ", "
            RsTel.MoveNext
        Loop
        aTelefonos = Mid(aTelefonos, 1, Len(aTelefonos) - 2)
    Else
        aTelefonos = "S/D"
    End If
    RsTel.Close
    
    TelefonoATexto = aTelefonos

End Function

Public Function FormatoReferencia(Valor As String, Formato As String)

    On Error GoTo errFormato
    
    FormatoReferencia = ""
    Valor = Trim(Valor)
    Select Case UCase(Trim(Formato))
        Case "MONEDA"
            If IsNumeric(Valor) Then FormatoReferencia = Format(Valor, "#,##0.00")
            
        Case "NUMERO"
            If IsNumeric(Valor) Then FormatoReferencia = Valor
        
        Case "TEXTO"
            FormatoReferencia = Valor
            
        Case "FECHA"
            If IsDate(Valor) Then FormatoReferencia = Format(Valor, "d-Mmm yyyy")
            
        Case "CEDULA"
            If IsNumeric(clsGeneral.QuitoFormatoCedula(Valor)) Then
                If clsGeneral.CedulaValida(clsGeneral.QuitoFormatoCedula(Valor)) Then
                    FormatoReferencia = clsGeneral.RetornoFormatoCedula(clsGeneral.QuitoFormatoCedula(Valor))  'Format(Valor, "@.@@@.@@@-@")
                End If
            End If
    End Select
    Exit Function
    
errFormato:
    FormatoReferencia = ""
End Function

Public Function FormatoGrabarReferencia(Valor As String, Formato As String)

On Error GoTo errFormato
    
    FormatoGrabarReferencia = ""
    Select Case UCase(Trim(Formato))
        Case "MONEDA"
            FormatoGrabarReferencia = CCur(Valor)
            
        Case "NUMERO"
            FormatoGrabarReferencia = CLng(Valor)
        
        Case "TEXTO"
            FormatoGrabarReferencia = Valor
            
        Case "FECHA"
            FormatoGrabarReferencia = Format(Valor, FormatoFH)
            
        Case "CEDULA"
            FormatoGrabarReferencia = clsGeneral.QuitoFormatoCedula(Valor)
    End Select
    Exit Function
    
errFormato:
End Function

'Public Function RetornoEstadoEnvio(Estado As Integer) As String

'    Select Case Estado
'        Case EstadoEnvio.AImprimir
'            RetornoEstadoEnvio = cEnvAImprimir
'        Case EstadoEnvio.AConfirmar
'            RetornoEstadoEnvio = cEnvAConfirmar
'        Case EstadoEnvio.Impreso
'            RetornoEstadoEnvio = cEnvImpreso
'        Case EstadoEnvio.Anulado
'            RetornoEstadoEnvio = cEnvAnulado
'        Case EstadoEnvio.Entregado
'            RetornoEstadoEnvio = cEnvEntregado
'        Case EstadoEnvio.Rebotado
'            RetornoEstadoEnvio = cEnvRebotado
'    End Select
    
'End Function

Function BuscoUsuario(Digito As Integer) As Integer
On Error GoTo ErrBU

    Cons = "SELECT * FROM USUARIO WHERE UsuDigito = " & Digito
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        BuscoUsuario = 0
        MsgBox "No existe un usuario para el dígito ingresado.", vbExclamation, "ATENCIÓN"
    Else
        BuscoUsuario = RsAux!UsuCodigo
    End If
    RsAux.Close
    Exit Function
    
ErrBU:
    clsGeneral.OcurrioError "Ocurrió un error inesperado."
    BuscoUsuario = 0
End Function

Function BuscoNombreUsuario(Codigo As Long) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoNombreUsuario = ""

    Cons = "SELECT * FROM USUARIO WHERE UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoNombreUsuario = Trim(Rs!UsuIdentificacion)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoInicialUsuario(Codigo As Integer) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset

    BuscoInicialUsuario = ""

    Cons = "SELECT * FROM USUARIO WHERE UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoInicialUsuario = Trim(Rs!UsuInicial)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoDigitoUsuario(Codigo As Long) As String
On Error GoTo ErrBU
Dim Rs As rdoResultset

    BuscoDigitoUsuario = ""

    Cons = "SELECT * FROM USUARIO WHERE UsuCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoDigitoUsuario = Trim(Rs!UsuDigito)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoSignoMoneda(Codigo As Variant) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoSignoMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoSignoMoneda = Trim(Rs!MonSigno)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoNombreMoneda = Trim(Rs!MonNombre)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Function BuscoNombreSucursal(Codigo As Long) As String
On Error GoTo ErrBU
    
Dim Rs As rdoResultset

    BuscoNombreSucursal = ""

    Cons = "SELECT * FROM Sucursal WHERE SucCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then BuscoNombreSucursal = Trim(Rs!SucAbreviacion)
    
    Rs.Close
    Exit Function
    
ErrBU:
End Function

Public Function TasadeCambio(MOriginal As Integer, MDestino As Integer, fecha As Date, Optional FechaTC As String = "", Optional TipoTC As Integer = -1) As Currency

Dim RsTC As rdoResultset

    On Error GoTo errTC
    If TipoTC = -1 Then TipoTC = 1
    TasadeCambio = 1
    Cons = "Select * from TasaCambio" _
            & " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " _
                                          & " Where TCaFecha < '" & Format(fecha, "mm/dd/yyyy 23:59") & "'" _
                                          & " And TCaOriginal = " & MOriginal _
                                          & " And TCaDestino = " & MDestino _
                                          & " And TCaTipo = " & TipoTC & ")" _
            & " And TCaOriginal = " & MOriginal _
            & " And TCaDestino = " & MDestino _
            & " And TCaTipo = " & TipoTC
            
    Set RsTC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsTC.EOF Then
        TasadeCambio = CCur(Format(RsTC!TCaComprador, "#.000"))
        FechaTC = Format(RsTC!TCaFecha, "dd/mm/yyyy")
    End If
    RsTC.Close
    Exit Function
    
errTC:
End Function

Public Function MovimientoDeCaja(Tipo As Long, fecha As Date, Disponibilidad As Long, Moneda As Long, Importe As Currency, _
                Optional Comentario As String = "", Optional Salida As Boolean = False) As Long

Dim RsMov As rdoResultset, rs1 As rdoResultset
Dim aMovimiento As Long, aMonedaD As Integer
Dim TC As Currency, aImporteD As Currency
    
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
    RsMov!MDiFecha = Format(fecha, "mm/dd/yyyy")
    RsMov!MDiHora = Format(fecha, "hh:mm:ss")
    RsMov!MDiTipo = Tipo
    If Comentario <> "" Then RsMov!MDiComentario = Trim(Comentario)
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    'Saco el Id de movimiento-------------------------------------------------------------------------------
    Cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                " Where MDiTipo = " & Tipo & _
                " And MDiHora = '" & Format(fecha, "hh:mm:ss") & "'"
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
        TC = TasadeCambio(CLng(Moneda), aMonedaD, fecha)
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
            TC = TasadeCambio(CLng(Moneda), CLng(paMonedaPesos), fecha)
            RsMov!MDRImportePesos = Importe * TC
        End If
    End If
    
    RsMov.Update: RsMov.Close
    '------------------------------------------------------------------------------------------------------------
    
    MovimientoDeCaja = aMovimiento
    
End Function


'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Public Function CargoArticulosDeFlete() As String

Dim Fletes As String
    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Fletes = Fletes & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    Fletes = Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function


Public Function CopiarDireccion(lnCodDireccion As Long) As Long

    'Copio la Direccion
    If lnCodDireccion > 0 Then
        
        Screen.MousePointer = 11
        On Error GoTo errorBT
        
        Dim RsDO As rdoResultset
        Dim RsDC As rdoResultset
        
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        
        'Direccion ORIGINAL
        Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
        Set RsDO = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        'Direccion COPIA
        Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
        Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        RsDC.AddNew
        If Not IsNull(RsDO!DirComplejo) Then RsDC!DirComplejo = RsDO!DirComplejo
        RsDC!DirCalle = RsDO!DirCalle
        RsDC!DirPuerta = RsDO!DirPuerta
        RsDC!DirBis = RsDO!DirBis
        If Not IsNull(RsDO!DirLetra) Then RsDC!DirLetra = RsDO!DirLetra
        If Not IsNull(RsDO!DirApartamento) Then RsDC!DirApartamento = RsDO!DirApartamento
        
        If Not IsNull(RsDO!DirCampo1) Then RsDC!DirCampo1 = RsDO!DirCampo1
        If Not IsNull(RsDO!DirSenda) Then RsDC!DirSenda = RsDO!DirSenda
        If Not IsNull(RsDO!DirCampo2) Then RsDC!DirCampo2 = RsDO!DirCampo2
        If Not IsNull(RsDO!DirBloque) Then RsDC!DirBloque = RsDO!DirBloque
        
        If Not IsNull(RsDO!DirEntre1) Then RsDC!DirEntre1 = RsDO!DirEntre1
        If Not IsNull(RsDO!DirEntre2) Then RsDC!DirEntre2 = RsDO!DirEntre2
        If Not IsNull(RsDO!DirAmpliacion) Then RsDC!DirAmpliacion = RsDO!DirAmpliacion
        RsDC!DirConfirmada = RsDO!DirConfirmada
        If Not IsNull(RsDO!DirVive) Then RsDC!DirVive = RsDO!DirVive
        
        RsDC.Update
        RsDC.Close
        RsDO.Close
                    
        Cons = "Select Max(DirCodigo) from Direccion"
        Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        CopiarDireccion = RsDC(0)
        RsDC.Close
        
        cBase.CommitTrans       'FIN TRANSACCION------------------------------------------
        
    Else
        CopiarDireccion = 0
    End If
    Exit Function
    
errorBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción."
    Exit Function

errorET:
    Resume ErrTransaccion

ErrTransaccion:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar copiar la dirección."
End Function

Public Function IVAArticulo(lnCodigo As Long)
Dim RsIva As rdoResultset

On Error GoTo ErrIA
    
    Cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
        & " Where AFaArticulo = " & lnCodigo _
        & " And AFaIVA = IVACodigo"
        
    Set RsIva = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsIva.EOF Then IVAArticulo = 0 Else IVAArticulo = Format(RsIva(0), "#0.00")
    RsIva.Close
    Exit Function
    
ErrIA:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al buscar el tipo de iva del artículo."
End Function

Public Sub BuscoComentariosAlerta(idCliente As Long, _
                                                   Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
                                                   Optional Decision As Boolean = False, Optional Informacion As Boolean = False)
                                                   
Dim RsCom As rdoResultset
Dim aCom As String
Dim sHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    sHay = False
    'Armo el str con los comentarios a consultar-------------------------------------------------
    If Not Alerta And Not Cuota And Not Decision And Not Informacion Then Exit Sub
    aCom = ""
    If Alerta Then aCom = aCom & Accion.Alerta & ", "
    If Cuota Then aCom = aCom & Accion.Cuota & ", "
    If Decision Then aCom = aCom & Accion.Decision & ", "
    If Informacion Then aCom = aCom & Accion.Informacion & ", "
    aCom = Mid(aCom, 1, Len(aCom) - 2)
    '---------------------------------------------------------------------------------------------------
    
    Cons = "Select * From Comentario, TipoComentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo = TCoCodigo " _
            & " And TCoAccion IN (" & aCom & ")"
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCom.EOF Then sHay = True
    RsCom.Close
    
    If Not sHay Then Screen.MousePointer = 0: Exit Sub
    
    Dim aObj As New clsCliente
    aObj.Comentarios idCliente:=idCliente
    DoEvents
    Set aObj = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub
'--------------------------------------------------------------------------------------------------------------
'   Retorna el codigo de barras para imprimir en las facturas (credito o contado).
'   TipoDoc :   Tipo de documento (Contado, Credito, Remito)
'   CodigoDoc: Id del documento.
'--------------------------------------------------------------------------------------------------------------
Public Function CodigoDeBarras(TipoDoc As Integer, CodigoDoc As Long)

    If Len(CStr(CodigoDoc)) < 6 Then
        CodigoDeBarras = TipoDoc & "D" & Format(CodigoDoc, "000000")
    Else
        CodigoDeBarras = TipoDoc & "D" & CodigoDoc
    End If
    CodigoDeBarras = "*" & CodigoDeBarras & "*"
    
End Function

Public Function PropiedadesConnect(Conexion As String, _
                                                    Optional Database As Boolean = True, Optional DSN As Boolean = False, _
                                                    Optional Server As Boolean = True) As String
Dim aRetorno As String

    On Error GoTo errConnect
    PropiedadesConnect = ""
    
    If DSN Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DSN=") + 4, Len(Conexion)))
    If Server Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "SERVER=") + 7, Len(Conexion)))
    If Database Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DATABASE=") + 9, Len(Conexion)))
    
    aRetorno = Trim(Mid(aRetorno, 1, InStr(aRetorno, ";") - 1))
    
    PropiedadesConnect = aRetorno
    Exit Function
    
errConnect:
End Function

Public Function CambioClienteEnvios(Cliente As Long, Envios As String) As Boolean
On Error GoTo ErrCCE
    
    CambioClienteEnvios = True
    If InStr(Envios, ",") = 0 Then
        If Envios = "0" Then Exit Function
    End If
    
    MsgBox "Se cambiará el cliente de los envíos ingresados.", vbInformation, "ATENCIÓN"
    Cons = "Update Envio Set EnvCliente = " & Cliente _
            & " Where EnvCodigo IN (" & Envios & ")"
    cBase.Execute (Cons)
    Exit Function

ErrCCE:
    clsGeneral.OcurrioError "Ocurrió un error al intentar modificar el cliente en los envíos."
    CambioClienteEnvios = False
End Function

Public Sub EjecutarApp(Path As String, Optional Valor As String = "", Optional Modal As Boolean = False)

    On Error GoTo errApp
    Screen.MousePointer = 11
    Dim plngRet As Long
    
    If Valor = "" Then
        Dim aPos As Integer
        aPos = InStr(Path, ".exe")
        If aPos <> 0 Then
            Valor = Mid(Path, aPos + Len(".exe") + 1)
            Path = Mid(Path, 1, aPos - 1)
        End If
    End If
    plngRet = ShellExecute(0, "open", Path, Valor, 0, 1)
    If plngRet = 0 And Err.Description <> "" Then GoTo errApp
    
    Screen.MousePointer = 0
    Exit Sub
    
errApp:
    Screen.MousePointer = 0
    MsgBox "Error al ejecutar la aplicación " & Path & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error de Aplicación"
End Sub
