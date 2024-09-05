Attribute VB_Name = "modStart"
Option Explicit

Public Const CTE_KeyConnect = "comercio" 'Constante de Logueo a la base de datos.

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDeTerminal As Long

Public mTiposDocumentos As clsTiposDocIdent
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public prmIPServer As String
Public prmPortServer As Long

Public Const sc_FIN = vbCrLf

Public Enum Asuntos
    Solicitudes = 1
    Servicios = 2
    GastosAAutorizar = 3
    SucesosAAutorizar = 4
End Enum

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public aTexto As String

'Parametros necesarios para el formulario de empleos-------------------------
Public paMonedaEmpleo As Long
Public paTipoIngreso As Long
Public paVigenciaEmpleo As Long
Public paTipoTelefonoE As Long
Public paMonedaFija As Long, paMonedaFijaTexto As String
Public paMonedaPesos As Long
Public paToleranciaMora As Integer
Public paResolucionEstandar As Long
Public paAnosAntecedentes As Integer

Public paECivilConyuge As Long
Public paempNoTrabajaMas As Integer
Public paempSeguroParo As Integer

'---------------------------------------------------------------------------------------
Public paRelPadre As Long

Public paCatsDistribuidor As String      'Categorias de distribuidores
Public paMayorEdad1 As Integer
Public paCatCliFallecido As Long
Public paPlantillasVOpe As String
Public paTipoTelefonoLlamoDe As Long

Public prmPlantillaPuente As Long
Public pathApp As String

Public prmPlantillaAGastos As Integer
Public prmPlantillaACompras As Integer
Public prmPlantillaAProveedores As Integer
Public prmPlantillaASucesos As Integer

Public prmOcupacionesEmp As String

'Datos diccionario para sustituir en comentario resolucion  -----
Private Type typDiccionario
    Incorrecto As String
    Correcto As String
End Type
Private arrDicc() As typDiccionario

Public prmAutorizaCredHasta As Long
Public prmValidezLimiteCredito As Integer

Public prmAlertaPorcCostoClearing As Byte

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If App.PrevInstance Then
        MsgBox "Esta aplicación está activa. " & vbCrLf & "No se puede abrir una nueva instanica.", vbExclamation, "Asuntos Pendientes está activo..."
        End
    End If
    
    If Not ObtenerConexionBD(cBase, CTE_KeyConnect) Then Screen.MousePointer = 0: Exit Sub
    CargoParametrosLocales
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    Set mTiposDocumentos = New clsTiposDocIdent
    mTiposDocumentos.CargoTiposActivos cBase
    
    frmLista.Show
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Function CargoParametrosLocales(Optional SoloServerAP As Boolean = False)

    On Error GoTo errParametro
    paAnosAntecedentes = 3
    prmOcupacionesEmp = "0"
    
    Cons = "Select * from Parametro"
    If SoloServerAP Then Cons = Cons & " Where ParNombre = 'serverasuntos_port_ip'"
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, CTE_KeyConnect, rdOpenDynamic) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "monedaempleo": paMonedaEmpleo = RsAux!ParValor
            
            Case "tipoingreso": paTipoIngreso = RsAux!ParValor
            Case "tipotelefonoe": paTipoTelefonoE = RsAux!ParValor
            Case "vigenciaempleo": paVigenciaEmpleo = RsAux!ParValor
            
            Case "monedafija": paMonedaFija = RsAux!ParValor
            
            Case "toleranciamora": paToleranciaMora = RsAux!ParValor
            Case "resolucionestandar": paResolucionEstandar = RsAux!ParValor
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            
            Case "ecivilconyuge": paECivilConyuge = RsAux!ParValor
            
            Case "empnotrabajamas": paempNoTrabajaMas = RsAux!ParValor
            Case "empseguroparo": paempSeguroParo = RsAux!ParValor
            
            Case "relpadre": paRelPadre = RsAux!ParValor
            
            Case "catclidistribuidor": paCatsDistribuidor = "," & Trim(RsAux!ParTexto) & ","
            
            Case "catclifallecido": paCatCliFallecido = RsAux!ParValor
            Case "mayordeedad": If Not IsNull(RsAux!ParTexto) Then paMayorEdad1 = Val(RsAux!ParTexto) Else paMayorEdad1 = 90
            Case "categoriacliente": paCategoriaCliente = RsAux!ParValor
            
            Case "plantillasvisualizacion": If Not IsNull(RsAux!ParTexto) Then paPlantillasVOpe = Trim(RsAux!ParTexto) Else paPlantillasVOpe = "0"
            Case "tipotelefonollamode": paTipoTelefonoLlamoDe = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            
            Case "pathapp": pathApp = Trim(RsAux!ParTexto) & "\"
            
            Case "plantillaautorizarsucesos": prmPlantillaASucesos = RsAux!ParValor
            Case "plantillaautorizargastos": prmPlantillaAGastos = RsAux!ParValor
            Case "plantillaautorizarcompras": prmPlantillaACompras = RsAux!ParValor
            Case "plantillaautorizarproveedores": prmPlantillaAProveedores = RsAux!ParValor
            
              Case "serverasuntos_port_ip"
                    prmPortServer = RsAux!ParValor
                    prmIPServer = Trim(RsAux!ParTexto)
                    
            Case LCase("vOPEOcupacionesRelaciones"): prmOcupacionesEmp = Trim(RsAux!ParTexto)
            Case LCase("tipocuotacontado"): paTipoCuotaContado = RsAux!ParValor
            
            Case LCase("ValidezLimiteCredito"): prmValidezLimiteCredito = RsAux!ParValor
            Case LCase("AlertaPorcCostoClearing"): prmAlertaPorcCostoClearing = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close

    If paMonedaFija <> 0 Then
        Cons = "Select * from Moneda Where MonCodigo = " & paMonedaFija
        'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsAux, Cons, CTE_KeyConnect, rdOpenDynamic) <> RAQ_SinError Then Exit Function
        If Not RsAux.EOF Then
            paMonedaFijaTexto = Trim(RsAux!MonSigno)
        Else
            MsgBox "El código de moneda fija (parámetro) no existe en la base de datos.", vbCritical, "ERROR"
            paMonedaFija = 0
        End If
        RsAux.Close
    End If
        
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, CTE_KeyConnect, rdOpenDynamic) <> RAQ_SinError Then Exit Function
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplScript) Then prmPlantillaPuente = RsAux!AplScript
    RsAux.Close
    
    dic_LoadDiccionario
    Exit Function
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros locales.", Err.Description
End Function

Public Function dic_LoadDiccionario()
On Error GoTo errDic

    'Cargo el array c/datos de diccionario      ----------------------------------------------------------------------
    Dim J As Integer: J = 0
    ReDim arrDicc(0)
    
    Cons = "Select * from Diccionario Where DicTipo = 4"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, CTE_KeyConnect, rdOpenDynamic) <> RAQ_SinError Then Exit Function
    Do While Not RsAux.EOF
        ReDim Preserve arrDicc(J)
        arrDicc(J).Incorrecto = Trim(RsAux!DicIncorrecto)
        arrDicc(J).Correcto = Trim(RsAux!DicCorrecto)
        J = J + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    Exit Function
    
errDic:
    clsGeneral.OcurrioError "Error al cargar el diccionario de abreviaciones.", Err.Description
End Function

Public Function TelefonoATexto(Cliente As Long, Optional TieneLlamoDe As Boolean = False) As String

Dim RsTel As rdoResultset
Dim aTelefonos As String

    On Error GoTo errTelefono
    TieneLlamoDe = False
    
    Cons = "Select * from Telefono, TipoTelefono" _
        & " Where TelCliente = " & Cliente _
        & " And TelTipo = TTeCodigo"
    'Set RsTel = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsTel, Cons, CTE_KeyConnect) <> RAQ_SinError Then Exit Function
    If Not RsTel.EOF Then
        aTelefonos = ""
        Do While Not RsTel.EOF
            aTelefonos = aTelefonos & Trim(RsTel!TTeNombre) & ": " & Trim(RsTel!TelNumero)
            If Not IsNull(RsTel!TelInterno) Then aTelefonos = aTelefonos & "(" & Trim(RsTel!TelInterno) & ")"
            aTelefonos = aTelefonos & ", "
            
            If RsTel!TelTipo = paTipoTelefonoLlamoDe Then TieneLlamoDe = True
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

Public Function dic_PulirTexto(mTexto As String, mCursorPos As Integer) As String
On Error GoTo errFnc

Dim xId As Integer
Dim mRet As String, mT1 As String, mT2 As String, mWord As String


    mT1 = Mid(mTexto, 1, mCursorPos)
    If InStrRev(mT1, " ") <> 0 Then
        mWord = Mid(mT1, InStrRev(mT1, " ") + 1)
        mT1 = Mid(mT1, 1, InStrRev(mT1, " "))
    Else
        mWord = mT1
        mT1 = ""
    End If
    
    mT2 = Mid(mTexto, mCursorPos + 1)
    
    mWord = " " & mWord & " "
    For xId = LBound(arrDicc) To UBound(arrDicc)
        mWord = Replace(mWord, " " & Trim(arrDicc(xId).Incorrecto) & " ", " " & Trim(arrDicc(xId).Correcto) & " ", Compare:=vbTextCompare)
    Next
    mWord = Mid(mWord, 2, Len(mWord) - 2)
    
    dic_PulirTexto = mT1 & mWord & mT2
    mCursorPos = Len(mT1 & mWord)
    
    Exit Function
errFnc:
    dic_PulirTexto = mTexto
End Function

Public Function z_BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
    z_BuscoUsuario = BuscoUsuario(Codigo, Identificacion, Digito, Iniciales)
End Function

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    'Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsUsr, Cons, CTE_KeyConnect) <> RAQ_SinError Then Exit Function
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    
    BuscoUsuario = aRetorno
    
End Function

