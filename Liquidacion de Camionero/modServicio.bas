Attribute VB_Name = "modServicio"
Option Explicit

Public Enum EstadoS
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum

Public Enum TipoServicio
    Visita = 1
    Retiro = 2
    Entrega = 3
End Enum

Public Enum EstadoP
    SinCargo = 1
    FueraGarantia = 2
    Abonado = 3
End Enum

Public Enum FacturaServicio
    CGSA = 1
    Camion = 2
    SinFactura = 3
End Enum

Public Enum TipoRenglonS
    Llamado = 1
    Cumplido = 2
    CumplidoPresupuesto = 3
    CumplidoArticulo = 4
End Enum

Public paPrimeraHoraEnvio As Long
Public paUltimaHoraEnvio As Long

Public Function EstadoServicio(Valor As Integer) As String
    
    On Error Resume Next
    Dim retorno As String
    retorno = ""

    Select Case Valor
        Case EstadoS.Anulado: retorno = "Anulado"
        Case EstadoS.Entrega: retorno = "Entrega"
        Case EstadoS.Retiro: retorno = "Retiro"
        Case EstadoS.Taller: retorno = "Taller"
        Case EstadoS.Visita: retorno = "Visita"
        Case EstadoS.Cumplido: retorno = "Cumplido"
    End Select

    EstadoServicio = retorno
End Function

Public Function TelefonoATexto(Cliente As Long) As String

Dim RsTel As rdoResultset
Dim aTelefonos As String

    On Error GoTo errTelefono
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

errTelefono:
End Function

Public Function CalculoEstadoProducto(idProducto As Long) As Integer

Dim rs As rdoResultset
Dim meses As Integer, retorno As Integer
Dim FCompra As Date
Dim haydatos As Boolean

    On Error GoTo Error
    haydatos = False
    
    'Consulto para sacar los datos de la garantia-----------------------------------------------
    meses = 0
    Cons = "Select * from Producto, Articulo, ArticuloFacturacion, Garantia " _
           & " Where ProCodigo = " & idProducto _
           & " And ProArticulo = ArtID " _
           & " And ArtId = AFaArticulo " _
           & " And AFaGarantia = GarCodigo"
    Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not rs.EOF Then
        If Not IsNull(rs!GarMeses) Then meses = rs!GarMeses
        If Not IsNull(rs!ProCompra) Then FCompra = rs!ProCompra Else FCompra = CDate("1/1/1900")
        haydatos = True
    End If
    rs.Close
    '--------------------------------------------------------------------------------------------------
    
    If haydatos Then
        Select Case DateAdd("m", meses, FCompra)
            Case Is < CDate(Format(gFechaServidor, "dd/mm/yyyy")): retorno = EstadoP.FueraGarantia
            Case Is >= CDate(Format(gFechaServidor, "dd/mm/yyyy")): retorno = EstadoP.SinCargo
        End Select
    Else
        retorno = EstadoP.FueraGarantia
    End If
    
    CalculoEstadoProducto = retorno
    
Error:
End Function

Public Function EstadoProducto(Valor As Integer, Optional Abreviacion As Boolean = True) As String
    
    On Error Resume Next
    Dim retorno As String
    retorno = ""

    Select Case Valor
        Case EstadoP.SinCargo: If Abreviacion Then retorno = "SC" Else retorno = "Sin Cargo"
        Case EstadoP.FueraGarantia: If Abreviacion Then retorno = "FG" Else retorno = "Fuera de Garantía"
        Case EstadoP.Abonado: If Abreviacion Then retorno = "AB" Else retorno = "Abonado"
    End Select

    EstadoProducto = retorno
End Function

Public Function TipoFacturaServicio(Valor As Integer) As String
    
    On Error Resume Next
    TipoFacturaServicio = ""
    Select Case Valor
        Case FacturaServicio.Camion: TipoFacturaServicio = "Camión"
        Case FacturaServicio.CGSA: TipoFacturaServicio = "Empresa"
        Case FacturaServicio.SinFactura: TipoFacturaServicio = "Sin Factura"
    End Select
    
End Function

Public Function CopiaDireccion(idDireccion As Long) As Long

    Dim RsDO As rdoResultset, RsDC As rdoResultset
    Dim aCopia As Long
    
    aCopia = 0
    If idDireccion = 0 Then CopiaDireccion = 0: Exit Function
        
    On Error GoTo errorBT
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Cons = "Select * from Direccion Where DirCodigo = " & idDireccion
    Set RsDO = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues) 'Direccion ORIGINAL
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues) 'Direccion COPIA
    
    RsDC.AddNew
    RsDC!DirCalle = RsDO!DirCalle
    RsDC!DirPuerta = RsDO!DirPuerta
    RsDC!DirBis = RsDO!DirBis
    If Not IsNull(RsDO!DirLetra) Then RsDC!DirLetra = RsDO!DirLetra
    If Not IsNull(RsDO!DirApartamento) Then RsDC!DirApartamento = RsDO!DirApartamento
    If Not IsNull(RsDO!DirSenda) Then RsDC!DirSenda = RsDO!DirSenda
    If Not IsNull(RsDO!DirBloque) Then RsDC!DirBloque = RsDO!DirBloque
    If Not IsNull(RsDO!DirEntre1) Then RsDC!DirEntre1 = RsDO!DirEntre1
    If Not IsNull(RsDO!DirEntre2) Then RsDC!DirEntre2 = RsDO!DirEntre2
    If Not IsNull(RsDO!DirAmpliacion) Then RsDC!DirAmpliacion = RsDO!DirAmpliacion
    RsDC!DirConfirmada = RsDO!DirConfirmada
    If Not IsNull(RsDO!DirVive) Then RsDC!DirVive = RsDO!DirVive
    
    RsDC.Update: RsDC.Close
    
    Cons = "Select Max(DirCodigo) from Direccion"
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    aCopia = RsDC(0)
    RsDC.Close
    
    RsDO.Close
    cBase.CommitTrans    'FIN TRANSACCION------------------------------------------
    
    CopiaDireccion = aCopia
    Exit Function
    
errorBT:
    MsgBox "No se ha podido copiar la dirección." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    MsgBox "No se ha podido copiar la dirección." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description
End Function

Public Function RetornoGarantia(IDArticulo As Long) As String
Dim RsGar As rdoResultset
    On Error GoTo ErrRG
    RetornoGarantia = ""
    Cons = "Select Garantia.* from ArticuloFacturacion, Garantia " _
           & " Where AFaArticulo = " & IDArticulo _
           & " And AFaGarantia = GarCodigo"
    Set RsGar = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsGar.EOF Then RetornoGarantia = Trim(RsGar!GarNombre)
    RsGar.Close
    Exit Function
ErrRG:
    MsgBox "Ocurrió un error al buscar la garantía del producto." & Chr(vbKeyReturn) & Trim(Err.Description), vbExclamation, "ATENCIÓN"
End Function

Public Function BuscoZonaDireccion(idDireccion As Long) As Long
On Error GoTo ErrBZD
Dim RsZona As rdoResultset
    
    BuscoZonaDireccion = 0
    
    Cons = "Select CZoZona From Direccion, CalleZona" _
        & " Where DirCodigo = " & idDireccion & " And DirCalle = CZoCalle " _
        & " And CZoDesde <= DirPuerta And CZoHasta >= DirPuerta"
    Set RsZona = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsZona.EOF Then BuscoZonaDireccion = RsZona!CZoZona
    RsZona.Close
    Exit Function
    
ErrBZD:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el código de zona.", Trim(Err.Description)
    BuscoZonaDireccion = 0
End Function

Public Function ValidoRangoHorario(strHora As String) As String
    
    ValidoRangoHorario = ""
    If InStr(1, strHora, "-") > 0 Then
        Select Case Len(strHora)
            Case 9
                If CLng(Mid(strHora, 1, InStr(1, strHora, "-") - 1)) > CLng(Mid(strHora, InStr(1, strHora, "-") + 1, Len(strHora))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN": Exit Function
                End If
                ValidoRangoHorario = strHora
            Case 5
                If InStr(1, strHora, "-") = 1 Then
                    If CLng(Mid(strHora, InStr(1, strHora, "-") + 1, Len(strHora))) < paPrimeraHoraEnvio Then
                        MsgBox "El horario ingresado es menor a la primera hora de entrega.", vbExclamation, "ATENCIÓN"
                        Exit Function
                    Else
                        If paPrimeraHoraEnvio < 1000 Then
                            ValidoRangoHorario = "0" & paPrimeraHoraEnvio & strHora
                        Else
                            ValidoRangoHorario = paPrimeraHoraEnvio & strHora
                        End If
                        Exit Function
                    End If
                Else
                    If InStr(1, strHora, "-") = 5 Then
                        If CLng(Mid(strHora, 1, InStr(1, strHora, "-") - 1)) > paUltimaHoraEnvio Then
                            MsgBox "El horario ingresado es mayor que la última hora de envio.", vbExclamation, "ATENCIÓN"
                            Exit Function
                        Else
                            ValidoRangoHorario = strHora & paUltimaHoraEnvio
                        End If
                    Else
                        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                        Exit Function
                    End If
                End If
            
            Case 8
                If CLng(Mid(strHora, 1, InStr(1, strHora, "-") - 1)) > CLng(Mid(strHora, InStr(1, strHora, "-") + 1, Len(strHora))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    Exit Function
                End If
                If InStr(1, strHora, "-") = 4 Then ValidoRangoHorario = "0" & strHora
            
            Case Else
                    MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                    Exit Function
                    
        End Select
    End If
    
End Function
