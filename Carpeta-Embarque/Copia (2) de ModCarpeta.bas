Attribute VB_Name = "ModCarpeta"
Option Explicit
Public RsHelp As rdoResultset, RsCom As rdoResultset

Public clsGeneral As New clsorCGSA
Public paLocalZF  As Long
Public paLocalPuerto As Long

Public UsuLogueado As Long
Public miconexion As New clsConexion

Public rdoCZureo As rdoConnection

Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        RelojA
        If miconexion.AccesoAlMenu("MaEmbarque") Then
            InicioConexionBD miconexion.TextoConexion(logImportaciones)
            CargoParametrosImportaciones
            UsuLogueado = miconexion.UsuarioLogueado(True)
            
            fnc_ConnectZureo
            
            If Trim(Command()) <> "" Then
                MaEmbarque.pSeleccionado = CLng(Command())
            Else
                MaEmbarque.pSeleccionado = 0
            End If
            MaEmbarque.pModal = False
            MaEmbarque.Show vbModeless
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
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
End Sub

Private Function fnc_ConnectZureo() As Boolean
    Dim oGeneric As New clsDBFncs
    If Not oGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
        MsgBox "Error al conectarse a la base de datos de Zureo.", vbExclamation, "Conexión Zureo"
    Else
        fnc_ConnectZureo = True
    End If
End Function

Public Sub RelojA()
    Screen.MousePointer = 11
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub
Private Function IngresoGastoZureo(ByVal cImporte As Currency, ByVal Fecha As Date, ByVal idProveedor As Long, _
        ByVal sNroDoc As String, ByVal IDTipoDocumento As Integer, ByVal sMemo As String, ByVal iSRubro As Long, _
        ByVal iBcoEmisor As Long) As Long
On Error GoTo errIGZ
    
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
        .UsuarioAlta = paCodigoDeUsuario
        .UsuarioAutoriza = paCodigoDeUsuario
    End With
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 0
        .Cuenta = iSRubro
        .ImporteComp = cImporte
        .ImporteCta = cImporte * iTC
        .MonedaCta = paMonedaDolar           '  xCuentaS_M 'prmMonedaContabilidad
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_CTA = New clsDCuenta
    With OBJ_CTA
        .VaAlDebe = 1
        .Cuenta = iBcoEmisor
        .ImporteComp = cImporte
        .ImporteCta = cImporte * iTC
        .MonedaCta = paMonedaDolar ' xCuentaE_M 'prmMonedaContabilidad
    End With
    colCuentas.Add OBJ_CTA
    Set OBJ_CTA = Nothing
    
    Set OBJ_COM.Cuentas = colCuentas
    If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
    
    Set OBJ_COM = Nothing
    Exit Function
errIGZ:
    clsGeneral.OcurrioError "Error al insertar el gasto.", Err.Description, "Gastos de importaciones"
    Screen.MousePointer = 0
End Function

Public Function InsertoGastoImportacionZureo(idEmbarque As Long, cImporte As Currency, Fecha As Date, idProveedor As Long, strCarpeta As String, CodEmbarque As String, sSerieNumero As String, _
                                    IDTipoDocumento As Integer, Arbitraje As Double, BcoEmisor As String, LC As String, FPago As String, SubRubro As Long, ByVal idBcoEmisor As Long, Optional SaldoCero As Boolean = False) As Long
Dim aCompra As Long

    On Error GoTo errIGI
    
    Dim aTexto As String
    aTexto = "C: " & Trim(strCarpeta)
    If Trim(BcoEmisor) <> "" Then aTexto = aTexto & ", " & Trim(BcoEmisor)
    If Trim(LC) <> "" Then aTexto = aTexto & ", LC: " & Trim(LC)
    If Trim(FPago) <> "" Then aTexto = aTexto & ", " & Trim(FPago)
    aTexto = aTexto & ", Arb. U$S = " & Arbitraje
    
    aCompra = IngresoGastoZureo(cImporte, Fecha, idProveedor, Trim(sSerieNumero), IDTipoDocumento, aTexto, SubRubro, idBcoEmisor)
        
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

Private Sub CargoCamposBDGastos(IdCompra As Long, cMonto As Currency, idEmbarque As Long, SubRubro As Long)

    Cons = "Select * from GastoSubrubro Where GSrIDCompra = " & IdCompra
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!GSrIDCompra = IdCompra
    RsAux!GSrIDSubrubro = SubRubro
    RsAux!GSrImporte = cMonto
    RsAux.Update
    RsAux.Close
    
    Cons = "Select * from GastoImportacion Where GImIDCompra = " & IdCompra
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!GImIDCompra = IdCompra
    RsAux!GImIDSubrubro = SubRubro
    RsAux!GImImporte = cMonto
    RsAux!GImCostear = cMonto
    RsAux!GImNivelFolder = Folder.cFEmbarque
    RsAux!GImFolder = idEmbarque
    RsAux.Update
    RsAux.Close
    
End Sub

Public Sub CargoCamposBDComprobante(cMonto As Currency, Fecha As Date, idProveedor As Long, Serie As String, Numero As Long, strCarpeta As String, IDTipoDocumento As Integer, Arbitraje As Double, strBcoEmisor As String, LC As String, FormaPago As String)
    
     RsCom!ComSaldo = cMonto    'Divisa / Arbitraje
    
    RsCom!ComTipoDocumento = IDTipoDocumento
    RsCom!ComFecha = Format(Fecha, sqlFormatoF)
    RsCom!ComProveedor = idProveedor
    
    RsCom!ComMoneda = paMonedaDolar
    
    RsCom!ComSerie = Serie
    RsCom!ComNumero = Numero
    
    RsCom!ComImporte = cMonto
    
    RsCom!ComTC = TasadeCambio(paMonedaDolar, paMonedaPesos, PrimerDia(Fecha) - 1)
    
    Dim aTexto As String
    aTexto = "C: " & Trim(strCarpeta)
    If Trim(strBcoEmisor) <> "" Then aTexto = aTexto & ", " & Trim(strBcoEmisor)
    If Trim(LC) <> "" Then aTexto = aTexto & ", LC: " & Trim(LC)
    If Trim(FormaPago) <> "" Then aTexto = aTexto & ", " & Trim(FormaPago)
    aTexto = aTexto & ", Arb. U$S = " & Arbitraje
    
    RsCom!ComComentario = aTexto '"Divisa de Carpeta: " & strCarpeta & ", Arb. U$S = " & Arbitraje
    
    RsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
    
    
End Sub

