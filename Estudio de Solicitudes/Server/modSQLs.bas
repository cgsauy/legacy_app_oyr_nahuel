Attribute VB_Name = "modSQLs"
Option Explicit

Private Type typTrama
    IDTipo As Integer
    IDEstadoTrama As String
    IDUserPara As Long
    DatosTrama As String
End Type

Public arrTramas() As typTrama
Public idxTrama As Integer

Private Type typABM
    IDTipo As Integer
    IDEstado As String          'N- nueva,  M- modificada,   E- elimianda   I- Igual
    IDCodigo As Long
    IDUserPara As Long
    sol_UsuarioR As Integer
    sol_Estado As Integer
    sol_Proceso As Integer
    aux_Modificado As String
End Type

Dim arrABMs() As typABM

Dim mSQL As String
Dim mTrama As String

Dim Idx As Integer
Dim bAdd As Boolean

Dim colUsuarios As New Collection   'Usuarios

Public prmUserLogs As String

Public Function arrInicializoVariables()
    
    ReDim arrABMs(0)
    ReDim arrTramas(0)
    idxTrama = -1
    
End Function

'Private Function ValidoConexionSQL() As Boolean
'
'    Dim RsF As rdoResultset
'    On Error GoTo errFecha
'    Cons = "Select GetDate()"
'    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    gFechaServidor = RsF(0)
'    RsF.Close
'
'    ValidoConexionSQL = True
'    On Error Resume Next
'    Time = gFechaServidor
'    Date = gFechaServidor
'    Exit Function
'
'errFecha:
'    On Error Resume Next
'    gFechaServidor = Now
'    If ValidoConexionSQL = False Then
'        InicioConexionBD miConexion.TextoConexion("comercio")
'        ValidoConexionSQL
'    End If
'
'End Function

Public Function arrInicializoTramas()
    
'    ValidoConexionSQL
    ReDim arrTramas(0)
    idxTrama = -1
    
    '1) Recorro el array de ABM y a todas las tramas marcadas como "E" eliminadas las remuevo x q ya
    '    las envie en la pasada anterior.
    loc_DepuroArrayAMB
    
    '2) Marco los datos q quedaron como Eliminados (al consultar voy a cambiarles el estado)
    For Idx = 1 To UBound(arrABMs)
        arrABMs(Idx).IDEstado = "E"
    Next
    
    '3)  Cargo los datos A Enviar
    SQLSolicitudes
    SQLServicios
    If prmUserLogs <> "" Then
        SQLGastosAAutorizar
        'SQLSucesosAAutorizar
    End If
    'busco si tengo sucesos a autorizar a quienes resuelven
    SQLSucesosAAutorizar
    SQLSolicitudesResueltas
    
    '4) Agrego las tramas que quedaron marcadas como eliminadas
    For Idx = 1 To UBound(arrABMs)
        If arrABMs(Idx).IDEstado = "E" Then
            idxTrama = idxTrama + 1
            ReDim Preserve arrTramas(idxTrama)      'La trama se puede agregar xq es nueva o se modifico !!
            With arrTramas(idxTrama)
                .IDTipo = arrABMs(Idx).IDTipo
                .IDEstadoTrama = "E"
                .IDUserPara = arrABMs(Idx).IDUserPara
                .DatosTrama = arrABMs(Idx).IDCodigo
            End With
        End If
    Next

End Function

Public Function SQLSolicitudes()
On Error GoTo errSQL

    'mSQL = "Select SolCodigo, SolProceso, SolDevuelta, SolFecha, SolUsuarioR, SolUsuarioS, SolEstado, " & _
                        " CliCategoria, CliTipo, " & _
                        " NombreP = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                        " NombreE = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')') " & _
                " From Solicitud, Cliente" & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
                " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" & _
                " And SolProceso IN (2, 5)" & _
                " And SolEstado IN (0, 4)" & _
                " And SolCliente = CliCodigo"
                
                
    mSQL = "SELECT SolCodigo, SolFecha, SolDevuelta, SolProceso, SolEstado,  CliCiRuc, CliCategoria, " & _
                 " IsNull(RTrim(CEmNombre), RTrim(CPeApellido1) + RTrim(' ' + IsNull(CPeApellido2,''))+', ' + RTrim(CPeNombre1) + RTrim(' ' + IsNull(CPeNombre2,''))) Nombre, " & _
                 " RTrim(IsNull(UsrSol.UsuIdentificacion, '')) Solicitante, RTrim(IsNull(UsrRes.UsuIdentificacion, '')) Resolv, IsNull(UsrRes.UsuCodigo, 0) as SolUsuarioR, " & _
                 " Sum((TCuCantidad + (Convert(bit, TCuVencimientoC)-1)) * RSoValorCuota - (Convert(bit, IsNull(TCuVencimientoE,0)) * IsNull(RSoValorEntrega,0))) Monto " & _
           " FROM Solicitud  " & _
                " INNER JOIN RenglonSolicitud on SolCodigo = RSoSolicitud  " & _
                " INNER JOIN TipoCuota on RSoTipoCuota = TCuCodigo  " & _
                " INNER JOIN Cliente on SolCliente = CliCodigo  " & _
                " INNER JOIN Usuario UsrSol ON SolUsuarioS = UsrSol.UsuCodigo  " & _
                " LEFT OUTER JOIN Usuario UsrRes ON SolUsuarioR = UsrRes.UsuCodigo  " & _
                " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente  " & _
                " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " & _
           " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" & _
           " And SolProceso IN (2, 5)" & _
           " And SolEstado IN (0, 4)" & _
           " GROUP BY SolCodigo, SolFecha, SolDevuelta, SolProceso, SolEstado, CliCiRuc, CliCategoria, CEmNombre, CPeApellido1, CPeApellido2,  CPeNombre1, CPeNombre2, UsrSol.UsuIdentificacion, UsrRes.UsuIdentificacion, UsrRes.UsuCodigo " & _
           " ORDER BY SolCodigo ASC "
                
                
    'Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, mSQL, CTE_KeyConnect, rdOpenDynamic, rdConcurValues, True) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        
        bAdd = checkArrayABM(Asuntos.Solicitudes, RsAux!SolCodigo, 0, RsAux!solProceso, RsAux!solEstado, IIf(IsNull(RsAux!solUsuarioR), 0, RsAux!solUsuarioR))
        'Codigo|Proceso|Devuelta|Fecha|Cliente|CliCategoria|IDUsrR|NameUsrR|Estado|NameUsrS|Importe
        If bAdd Then
            mTrama = RsAux!SolCodigo & "|"
            mTrama = mTrama & RsAux!solProceso & "|"
            
            mTrama = mTrama & IIf(IsNull(RsAux!SolDevuelta), "0", "1") & "|"
            
            mTrama = mTrama & Format(RsAux!SolFecha, "dd/mm/yyyy hh:nn:ss") & "|"
            
            mTrama = mTrama & Trim(RsAux!Nombre) & "|"
            
            mTrama = mTrama & IIf(IsNull(RsAux!CliCategoria), "0", RsAux!CliCategoria) & "|"
            
            mTrama = mTrama & IIf(IsNull(RsAux!solUsuarioR), "0", RsAux!solUsuarioR) & "|"
            
            mTrama = mTrama & RsAux!Resolv & "|"
            
            mTrama = mTrama & IIf(IsNull(RsAux!solEstado), "0", RsAux!solEstado) & "|"
            
            If Not IsNull(RsAux!Solicitante) Then mTrama = mTrama & RsAux!Solicitante
            mTrama = mTrama & "|"
                        
            mTrama = mTrama & Format(RsAux!Monto, "#,##0.00")
                        
            arrAddTrama Asuntos.Solicitudes, "N", 0, mTrama
        
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Function
errSQL:
     frmServer.loc_InsertoError "SQLSolicitudes"
End Function

Private Function SQLServicios()
On Error GoTo errSQL
Dim rs2 As rdoResultset

    mSQL = "Select SerCodigo, SerModificacion, SerProducto, TalCostoTecnico, UsuIdentificacion " & _
                " From Servicio, Taller " & _
                        " Left Outer Join Usuario on TalTecnico = UsuCodigo " & _
                " Where SerCodigo = TalServicio " & _
                " And TalFPresupuesto is Not Null " & _
                " And SerCostoFinal is Null " & _
                " Order by SerCodigo"

    'Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, mSQL, CTE_KeyConnect, rdOpenForwardOnly, rdConcurReadOnly, True) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        'Codigo|Producto|Tecnico|CostoT
        bAdd = checkArrayABM(Asuntos.Servicios, RsAux!SerCodigo, 0, auxModificado:=RsAux!SerModificacion)
        
        If bAdd Then
            mTrama = RsAux!SerCodigo & "|"
            
            '---------------------------------------------------------------------------------------------------------
            mSQL = "Select ArtCodigo, ArtNombre " & _
                        " From Producto, Articulo " & _
                        " Where ProArticulo = ArtID" & _
                        " And ProCodigo = " & RsAux!SerProducto
                        
            Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
            If Not rs2.EOF Then
                mTrama = mTrama & Format(rs2!ArtCodigo, "(000,000) ") & Trim(rs2!ArtNombre)
            End If
            rs2.Close
            
            mTrama = mTrama & "|"
            '---------------------------------------------------------------------------------------------------------
            
            If Not IsNull(RsAux!UsuIdentificacion) Then mTrama = mTrama & Trim(RsAux!UsuIdentificacion)
            mTrama = mTrama & "|"
            
            mTrama = mTrama & Format(RsAux!TalCostoTecnico, "#,##0.00")
            
            arrAddTrama Asuntos.Servicios, "N", 0, mTrama
            
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Function
errSQL:
     frmServer.loc_InsertoError "SQLServicios"
End Function

Private Function SQLGastosAAutorizar()
On Error GoTo errSQL

    mSQL = "Select ComFModificacion, ComUsrAutoriza, ComCodigo, ComFecha, ComUsuario, " & _
                          " ComImporte, IsNull(ComIva, 0) as ComIva, IsNull(ComCofis,0) as ComCofis, MonSigno, PClFantasia" & _
                " From Compra, Moneda, ProveedorCliente " & _
                " Where ComUsrAutoriza IN (" & prmUserLogs & ")" & _
                " And ComVerificado IS NULL " & _
                " And ComMoneda = MonCodigo And ComProveedor = PClCodigo"

    'Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, mSQL, CTE_KeyConnect, rdOpenForwardOnly, rdConcurReadOnly, True) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        
        bAdd = checkArrayABM(Asuntos.GastosAAutorizar, RsAux!ComCodigo, 0, auxModificado:=RsAux!ComFModificacion)
        'Codigo|Proveedor|Importe|UsuarioG
        If bAdd Then
            mTrama = RsAux!ComCodigo & "|"
            mTrama = mTrama & Trim(RsAux!PClFantasia) & "|"
            
            mTrama = mTrama & Trim(RsAux!MonSigno) & " " & Format(RsAux!ComImporte + RsAux!ComIVA + RsAux!ComCoFIS, "#,##0.00")
            mTrama = mTrama & "|"
            
            If Not IsNull(RsAux!ComUsuario) Then mTrama = mTrama & loc_ItemUsuarios(RsAux!ComUsuario)
            
            arrAddTrama Asuntos.GastosAAutorizar, "N", RsAux!ComUsrAutoriza, mTrama
            
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Function
errSQL:
     frmServer.loc_InsertoError "SQLGastosAAutorizar"
End Function

Private Function SQLSucesosAAutorizar()
On Error GoTo errSQL

     '2/8/2012 agregué condición SucAutoriza = 0 para levantar los sucesos que debe autorizar cualquiera de los que estén resolviendo.
     mSQL = "Select SucCodigo, TSuNombre ,isNull(SucDescripcion, '') as SucDescripcion, SucValor, SucAutoriza, SucUsuario" & _
                " From Suceso " & _
                " Left Outer Join TipoSuceso On SucTipo = TSuCodigoSistema  " & _
                " Where IsNull(SucVerificado, 0) = 0 " & _
                " And SucAutoriza IN (" & IIf(prmUserLogs = "", "0", prmUserLogs & ", 0") & ")"

    'Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, mSQL, CTE_KeyConnect, rdOpenForwardOnly, rdConcurReadOnly, True) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        
        bAdd = checkArrayABM(Asuntos.SucesosAAutorizar, RsAux!SucCodigo, RsAux!SucAutoriza)
        'Codigo|NombreSuceso|Descripcion|Valor|Usuario
        If bAdd Then
            mTrama = RsAux!SucCodigo & "|"
            mTrama = mTrama & Trim(RsAux!TSuNombre) & "|"
            mTrama = mTrama & Trim(RsAux!SucDescripcion) & "|"
            mTrama = mTrama & Format(RsAux!SucValor, "#,##0.00") & "|"
        
            If Not IsNull(RsAux!SucUsuario) Then mTrama = mTrama & loc_ItemUsuarios(RsAux!SucUsuario)
            
            arrAddTrama Asuntos.SucesosAAutorizar, "N", RsAux!SucAutoriza, mTrama
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Function
errSQL:
     frmServer.loc_InsertoError "SQLSucesosAAutorizar"
End Function

Private Function SQLSolicitudesResueltas()
On Error GoTo errSQL

'   ATENCION: EL "usuario para" es la SUCURSAL, los clientes mandan el id de sucursal en vez del id de usuario

    'mSQL = "Select SolCodigo, SolProceso, SolDevuelta, SolFecha, SolUsuarioR, SolUsuarioS, SolEstado, SolFResolucion, " & _
                        " SolTipo, SolComentarioR, SolSucursal, CliTipo, " & _
                        " NombreP = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                        " NombreE = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')') " & _
                " From Solicitud, Cliente" & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
                " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" & _
                " And SolProceso NOT IN (3, 4)" & _
                " And SolEstado <> 0" & _
                " And SolCliente = CliCodigo" & _
                " And SolVisible Is NULL "
                
    mSQL = "Select SolCodigo, SolProceso, SolDevuelta, SolFecha, SolUsuarioS, SolEstado, SolFResolucion, " & _
                        " SolTipo, SolSucursal, CliTipo, ResComentario, SolUsuarioR, " & _
                        " NombreP = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), " & _
                        " NombreE = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')') " & _
                " From Solicitud, SolicitudResolucion, Cliente" & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente" & _
                " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" & _
                " And SolProceso NOT IN (3, 4)" & _
                " And SolEstado IN ( 1, 2, 3)" & _
                " And SolCliente = CliCodigo" & _
                " And SolVisible Is NULL " & _
                " And SolCodigo = ResSolicitud " & _
                " And ResNumero = (Select MAX(ResNumero) From SolicitudResolucion Where SolCodigo = ResSolicitud)"
           
    'Set RsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, mSQL, CTE_KeyConnect, rdOpenForwardOnly, rdConcurReadOnly, True) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        
        bAdd = checkArrayABM(Asuntos.SolicitudesResueltas, RsAux!SolCodigo, 0, RsAux!solProceso, RsAux!solEstado, IIf(IsNull(RsAux!solUsuarioR), 0, RsAux!solUsuarioR))
        'Codigo|Estado|Proceso||NameUsrR|Cliente|NameUsrS|FResolucion|Tipo|ComentarioR
        If bAdd Then
            mTrama = RsAux!SolCodigo & "|"
            mTrama = mTrama & IIf(IsNull(RsAux!solEstado), "0", RsAux!solEstado) & "|"
            mTrama = mTrama & RsAux!solProceso & "|"
            
            If Not IsNull(RsAux!solUsuarioR) Then mTrama = mTrama & loc_ItemUsuarios(RsAux!solUsuarioR)
            mTrama = mTrama & "|"
            
            If RsAux!CliTipo = 1 Then
                mTrama = mTrama & Trim(RsAux!NombreP) & "|"
            Else
                mTrama = mTrama & Trim(RsAux!NombreE) & "|"
            End If
            
            If Not IsNull(RsAux!SolUsuarioS) Then mTrama = mTrama & loc_ItemUsuarios(RsAux!SolUsuarioS)
            mTrama = mTrama & "|"
            
            mTrama = mTrama & Format(RsAux!SolFResolucion, "dd/mm/yyyy hh:nn:ss") & "|"
            
            mTrama = mTrama & IIf(IsNull(RsAux!SolTipo), "0", RsAux!SolTipo) & "|"
            
            'mTrama = mTrama & IIf(IsNull(rsAux!SolComentarioR), "0", rsAux!SolComentarioR)
            If IsNull(RsAux!ResComentario) Then mTrama = mTrama & "0" Else mTrama = mTrama & Trim(RsAux!ResComentario)
                        
            arrAddTrama Asuntos.SolicitudesResueltas, "N", RsAux!SolSucursal, mTrama
        End If
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Function
errSQL:
     frmServer.loc_InsertoError "SQLSolicitudesResueltas"
End Function

Private Function arrAddTrama(TipoT As Integer, EstadoT As String, UsuarioPara As Long, DatosT As String)

    idxTrama = idxTrama + 1
    ReDim Preserve arrTramas(idxTrama)      'La trama se puede agregar xq es nueva o se modifico !!
    
    With arrTramas(idxTrama)
        .IDTipo = TipoT
        .IDEstadoTrama = Trim(EstadoT)
        .IDUserPara = UsuarioPara
        .DatosTrama = DatosT
    End With
    
End Function

Private Function loc_ItemUsuarios(idUsr As Long) As String
    On Error GoTo usrAgregar
    
    Dim aItem As String
    aItem = "I" & CStr(idUsr)
    
    loc_ItemUsuarios = colUsuarios.Item(aItem)
    Exit Function
    
usrAgregar:
    On Error GoTo usrErrAdd
    loc_ItemUsuarios = z_BuscoUsuario(idUsr, Identificacion:=True)
    colUsuarios.Add loc_ItemUsuarios, CStr(aItem)
    
usrErrAdd:
End Function

Private Function checkArrayABM(Tipo As Integer, Codigo As Long, UsuarioPara As Long, _
                    Optional solProceso As Integer = 0, Optional solEstado As Integer = 0, Optional solUsuarioR As Long = 0, _
                    Optional auxModificado As String = "") As Boolean

'Esta funcion Retorna si debo Agregar la trama al array de Tramas
'Se agrega cuando la Trama no existe o se detecta que fue modificada

    checkArrayABM = True
    Dim bAdd As Boolean
    Dim Idx As Integer
    
    bAdd = True
    
    For Idx = 1 To UBound(arrABMs)
        With arrABMs(Idx)
            If Tipo = .IDTipo And Codigo = .IDCodigo Then
                bAdd = False
                
                Select Case Tipo
                    Case Asuntos.Solicitudes, Asuntos.SolicitudesResueltas
                            If .sol_Estado = solEstado And .sol_Proceso = solProceso And .sol_UsuarioR = solUsuarioR Then
                                .IDEstado = "I"
                                checkArrayABM = False
                            Else
                                .IDEstado = "M"
                                .sol_Estado = solEstado
                                .sol_Proceso = solProceso
                                .sol_UsuarioR = solUsuarioR
                            End If
                    
                    Case Asuntos.Servicios, Asuntos.GastosAAutorizar, Asuntos.SucesosAAutorizar
                            If Trim(.aux_Modificado) = Trim(auxModificado) Then
                                .IDEstado = "I"
                                checkArrayABM = False
                            Else
                                .IDEstado = "M"
                                .aux_Modificado = auxModificado
                            End If
                End Select
                
                Exit For
            End If
        End With
    Next

    If bAdd Then
        Idx = UBound(arrABMs) + 1
        ReDim Preserve arrABMs(Idx)
        With arrABMs(Idx)
            .IDTipo = Tipo
            .IDCodigo = Codigo
            .IDEstado = "N"
            .IDUserPara = UsuarioPara
            .sol_Estado = solEstado
            .sol_Proceso = solProceso
            .sol_UsuarioR = solUsuarioR
            .aux_Modificado = auxModificado
        End With
    End If
    
End Function

Private Function loc_DepuroArrayAMB()

Dim arrAux() As typABM
Dim IdxA As Integer
    
    IdxA = 0
    ReDim arrAux(IdxA)
    
    For Idx = 1 To UBound(arrABMs)
        If arrABMs(Idx).IDEstado <> "E" Then
            IdxA = IdxA + 1
            ReDim Preserve arrAux(IdxA)
            arrAux(IdxA) = arrABMs(Idx)
        End If
    Next
    
    ReDim arrABMs(0)
    arrABMs = arrAux
    ReDim arrAux(0)
        
End Function

