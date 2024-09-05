Attribute VB_Name = "modGastoZureo"
Option Explicit

Public rdoCZureo As rdoConnection
Public objUsers As clsCheckIn
Public iUserZureo As Long
Public prmUsuario As Long, prmPWDZureo As String
Public prmSRMerImpARecibir As Long, prmTipoCompAsiento As Integer

Public Function fnc_ConnectZureo() As Boolean
On Error GoTo errCZ
    Dim oGeneric As New clsDBFncs
    If Not oGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
        MsgBox "Error al conectarse a la base de datos de Zureo.", vbExclamation, "Conexión Zureo"
    Else
        fnc_ConnectZureo = True
    End If
    Exit Function
errCZ:
    MsgBox "No se logró realizar el login a Zureo, verifique si está instalado en el pc.", vbInformation, "ATENCIÓN"
End Function

Public Sub CargoParametrosParaZureo()
Dim sQy As String
Dim rsP As rdoResultset

    sQy = "SELECT ParNombre, ParValor, ParTexto FROM CGSA.dbo.Parametro WHERE ParNombre IN ('usuariozureo', 'SubrubroMercImpARecibir', 'TipoComprobanteAsiento')"
    If ObtenerResultSet(cBase, rsP, sQy, logImportaciones) = RAQ_SinError Then
        Do While Not rsP.EOF
            Select Case LCase(Trim(rsP!ParNombre))
                Case "usuariozureo"
                    prmUsuario = rsP("ParValor")
                    prmPWDZureo = Trim(rsP("ParTexto"))
                Case LCase("SubrubroMercImpARecibir")
                    prmSRMerImpARecibir = rsP("ParValor")
                Case LCase("TipoComprobanteAsiento")
                    prmTipoCompAsiento = rsP("ParValor")
            End Select
            rsP.MoveNext
        Loop
        rsP.Close
    End If
    If prmSRMerImpARecibir = 0 Or prmTipoCompAsiento = 0 Then
        MsgBox "No se cargaron los parámetros para el ingreso del comprobante en ZUREO.", vbCritical, "ATENCIÓN"
    End If
End Sub

Public Function IngresoGastoZureo(ByVal Fecha As Date, ByVal sNroDoc As String, _
            ByVal sMemo As String, ByVal colGastos As Collection) As Long
On Error GoTo errIGZ
    
    If Not fnc_ValidoAcceso Then
        MsgBox "No se tiene acceso a ZUREO, no se ingresaran los gastos.", vbCritical, "Login ZUREO"
        Exit Function
    End If
    Dim m_ReturnID As Long
    Dim objComp As New clsComprobantes
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
    
    Dim cImpTotalComp As Currency
    
    Dim oGasto As clsGastoTotal
    For Each oGasto In colGastos
        Set OBJ_CTA = New clsDCuenta
        With OBJ_CTA
            .VaAlDebe = 0
            .Cuenta = oGasto.idSubro
            .ImporteComp = oGasto.TotalMonComprobante
            .ImporteCta = oGasto.TotalMonCuenta
            .MonedaCta = paMonedaPesos
        End With
        cImpTotalComp = cImpTotalComp + oGasto.TotalMonComprobante
        colCuentas.Add OBJ_CTA
        Set OBJ_CTA = Nothing
    Next
    Set OBJ_COM = New clsDComprobante
    With OBJ_COM
        .doAccion = 1
        .Ente = 0
        .Empresa = 1
        .Numero = sNroDoc
        .Fecha = CDate(Format(Fecha, "dd/mm/yyyy"))
        .Tipo = prmTipoCompAsiento
        .Moneda = paMonedaPesos
        .ImporteTotal = cImpTotalComp
        .TC = 1
        .Memo = sMemo
        .UsuarioAlta = iUserZureo
        .UsuarioAutoriza = iUserZureo
    End With
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 1
        .Cuenta = prmSRMerImpARecibir
        .ImporteComp = cImpTotalComp
        .ImporteCta = cImpTotalComp
        .MonedaCta = paMonedaDolar           '  xCuentaS_M 'prmMonedaContabilidad
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    Set OBJ_COM.Cuentas = colCuentas
    If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then IngresoGastoZureo = objComp.prm_Comprobante
    
    Set OBJ_COM = Nothing
    Set objUsers = Nothing
    Exit Function
errIGZ:
    Set objUsers = Nothing
    clsGeneral.OcurrioError "Error al insertar el gasto.", Err.Description, "Gastos de importaciones"
    Screen.MousePointer = 0
End Function

Public Function fnc_ValidoAcceso() As Boolean
On Error GoTo errVAcc
Dim prmAccesosUserLog As String

    fnc_ValidoAcceso = False
    Set objUsers = New clsCheckIn
    prmAccesosUserLog = objUsers.ValidateAccess(rdoCZureo, "org01", "Comprobantes", 1)

    If prmAccesosUserLog <> "" Then
        fnc_ValidoAcceso = True
    Else
        
        Set RsAux = rdoCZureo.OpenResultset("Select ParValor From genParametros Where ParNombre = 'sis_MultiUsuario'", rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux(0) = 0 Then
                iUserZureo = 0
                RsAux.Close
                fnc_ValidoAcceso = True
                Exit Function
            End If
        End If
        RsAux.Close
        
        Dim mRet As Integer
        'Como valor Q usuarios (-1 error, 0 no hay, 1 hay uno ,2 hay mas de 1)
        iUserZureo = 0

        mRet = objUsers.GetUserData(rdoCZureo, "org01", UserID:=iUserZureo)

        If (mRet <> -1) And (iUserZureo = -1) Then
            '0- No hay acceso;  1- Hay acceso

            mRet = objUsers.doLogIn(rdoCZureo, "org01", CStr(prmUsuario), prmPWDZureo)

            If mRet = 1 Then
                prmAccesosUserLog = objUsers.ValidateAccess(rdoCZureo, "org01", "Comprobantes", 1)
                If prmAccesosUserLog <> "" Then iUserZureo = prmUsuario:  fnc_ValidoAcceso = True
            End If
        End If
    End If
    Exit Function
errVAcc:
    Set objUsers = Nothing
End Function

