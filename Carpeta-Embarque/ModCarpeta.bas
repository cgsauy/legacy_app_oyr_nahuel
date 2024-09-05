Attribute VB_Name = "ModCarpeta"
Option Explicit
Public Const ColorNaranja = &H40FF&
Public Const DateMinValue = "01/01/2000"

Public RsHelp As rdoResultset, RsCom As rdoResultset

Public clsGeneral As New clsorCGSA
'Public paLocalZF  As Long
'Public paLocalPuerto As Long

Public UsuLogueado As Long
Public miconexion As New clsConexion

Public rdoCZureo As rdoConnection

Public objUsers As clsCheckIn
Private iUserZureo As Long


Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        RelojA
        If miconexion.AccesoAlMenu("MaEmbarque") Then
            InicioConexionBD miconexion.TextoConexion(logImportaciones)
            CargoParametrosImportaciones
            UsuLogueado = miconexion.UsuarioLogueado(True)
            
            If fnc_ConnectZureo Then
                If Trim(Command()) <> "" Then
                    MaEmbarque.pSeleccionado = CLng(Command())
                Else
                    MaEmbarque.pSeleccionado = 0
                End If
                MaEmbarque.pModal = False
                MaEmbarque.Show vbModeless
            End If
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            RelojD
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miconexion.AccesoAlMenu ("MaEmbarque")
        InicioConexionBD miconexion.TextoConexion(logImportaciones)
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
End Sub

Private Function fnc_ConnectZureo() As Boolean
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

Public Sub RelojA()
    Screen.MousePointer = 11
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub
Private Function IngresoGastoZureo(ByVal cImporte As Currency, ByVal Fecha As Date, ByVal idProveedor As Integer, _
        ByVal sNroDoc As String, ByVal IDTipoDocumento As Integer, ByVal sMemo As String, ByVal iSRubro As Long, _
        ByVal iCta2 As Long) As Long
On Error GoTo errIGZ
    
    If Not fnc_ValidoAcceso Then
        MsgBox "No se tiene acceso a ZUREO, no se ingresaran los gastos.", vbCritical, "Login ZUREO"
        Exit Function
    End If
    
    Dim m_ReturnID As Long
    Dim objComp As New clsComprobantes
    
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
    
    Dim iTC As Currency
    iTC = TasadeCambio(paMonedaDolar, paMonedaPesos, PrimerDia(Fecha) - 1)
        
    Set OBJ_COM = New clsDComprobante
    With OBJ_COM
        .doAccion = 1
        .Ente = idProveedor
        .Empresa = 1
        .Numero = sNroDoc
        .Fecha = CDate(Format(Fecha, "dd/mm/yyyy"))
        .Tipo = IDTipoDocumento
        .Moneda = paMonedaDolar
        .ImporteTotal = cImporte
        .TC = iTC
        .Memo = sMemo
        .UsuarioAlta = iUserZureo
        .UsuarioAutoriza = iUserZureo
    End With
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 1
        .Cuenta = iSRubro
        .ImporteComp = cImporte
        .ImporteCta = cImporte '* iTC
        .MonedaCta = paMonedaDolar           '  xCuentaS_M 'prmMonedaContabilidad
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 0
        .Cuenta = iCta2
        .ImporteComp = cImporte
        .ImporteCta = cImporte '* iTC
        .MonedaCta = paMonedaDolar ' xCuentaE_M 'prmMonedaContabilidad
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

Public Function InsertoGastoImportacionZureo(idEmbarque As Long, cImporte As Currency, Fecha As Date, idProveedor As Integer, strCarpeta As String, CodEmbarque As String, sSerieNumero As String, _
                                    IDTipoDocumento As Integer, Arbitraje As Double, BcoEmisor As String, LC As String, FPago As String, SubRubro As Long, ByVal iCodBanco As Long, Optional SaldoCero As Boolean = False) As Long
Dim aCompra As Long

    On Error GoTo errIGI
    
    Dim aTexto As String
    aTexto = "C: " & Trim(strCarpeta)
    If Trim(BcoEmisor) <> "" Then aTexto = aTexto & ", " & Trim(BcoEmisor)
    If Trim(LC) <> "" Then aTexto = aTexto & ", LC: " & Trim(LC)
    If Trim(FPago) <> "" Then aTexto = aTexto & ", " & Trim(FPago)
    aTexto = aTexto & ", Arb. U$S = " & Arbitraje
    
    Dim iCuenta2 As Long
    If iCodBanco = 0 Then
        iCuenta2 = paCtaAnticipoDivisa
    Else
        'Saco el valor de la cuenta de la tabla BancoLocal.
        Cons = "Select BLoCuenta From BancoLocal Where BLoCodigo = " & iCodBanco & " AND BLoCuenta > 0 "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            iCuenta2 = RsAux("BLoCuenta")
        Else
            iCuenta2 = paCtaAnticipoDivisa
        End If
        RsAux.Close
    End If
    
    'Si es nota invierto los valoers.
    If IDTipoDocumento = TipoDocumento.CompraNotaCredito Then
        aCompra = IngresoGastoZureo(Abs(cImporte), Fecha, idProveedor, Trim(sSerieNumero), IDTipoDocumento, aTexto, iCuenta2, SubRubro)
    Else
        aCompra = IngresoGastoZureo(Abs(cImporte), Fecha, idProveedor, Trim(sSerieNumero), IDTipoDocumento, aTexto, SubRubro, iCuenta2)
    End If
        
    'Si inserte el gasto inserto en esta tabla.
    If aCompra > 0 Then
    
        On Error GoTo errGI
        InsertoGastoImportacionZureo = aCompra
        
        Cons = "Select * from GastoImportacion Where GImIDCompra = " & aCompra
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        RsAux!GImIDCompra = aCompra
        RsAux!GImIDSubrubro = SubRubro
        RsAux!GImImporte = cImporte
        RsAux!GImCostear = cImporte
        RsAux!GImNivelFolder = Folder.cFEmbarque
        RsAux!GImFolder = idEmbarque
        RsAux.Update
        RsAux.Close
    End If
    Exit Function
    
errIGI:
    clsGeneral.OcurrioError "Error al invocar a comprobantes para insertar el gasto.", Err.Description, "Insertar gasto de importación"
    Exit Function
errGI:
    clsGeneral.OcurrioError "Error al insertar en la tabla gasto importación, comuniquese con el administrador." & vbCrLf & vbCrLf & "ATENCIÓN Se detala el insert", "Dato a insertar: " & "INSERT INTO GastoImportacion (GImIDCompra, GImIDSubRubro, GImImporte, GImCostear, GimNivelFolder, GImFolder) Values(" & _
            aCompra & ", " & SubRubro & ", " & cImporte & ", " & cImporte & ", " & Folder.cFEmbarque & ", " & idEmbarque & ")" & vbCrLf & vbCrLf & "Error: " & Err.Description, "Insertar gasto"
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

