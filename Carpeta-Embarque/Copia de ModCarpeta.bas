Attribute VB_Name = "ModCarpeta"
Option Explicit
Public RsHelp As rdoResultset, RsCom As rdoResultset

Public clsGeneral As New clsorCGSA
Public paLocalZF  As Long
Public UsuLogueado As Long
Public miconexion As New clsConexion
'Constantes.----------------------------------------------

Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        RelojA
        If miconexion.AccesoAlMenu("MaEmbarque") Then
            InicioConexionBD miconexion.TextoConexion(logImportaciones)
            CargoParametrosImportaciones
            UsuLogueado = 6 ' miconexion.UsuarioLogueado(True)
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

Public Sub RelojA()
    Screen.MousePointer = 11
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub
Public Function IngresoGastoAutomatico(idEmbarque As Long, cImporte As Currency, fecha As Date, idProveedor As Long, strCarpeta As String, CodEmbarque As String, Serie As String, Numero As Long, IDTipoDocumento As Integer, Arbitraje As Double, BcoEmisor As String, LC As String, FPago As String, SubRubro As Long, Optional SaldoCero As Boolean = False) As Long
Dim aCompra As Long

    If idProveedor = 0 Then
        Cons = "Select * From ProveedorMercaderia Where PMeNombre = 'n/d'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        idProveedor = RsAux!PMeCodigo
        RsAux.Close
    End If
    
    'Cargo tabla: Compra----------------------------------------------------------------
    Cons = "Select * from Compra Where ComCodigo = 0"
    Set RsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsCom.AddNew
    CargoCamposBDComprobante cImporte, fecha, idProveedor, Serie, Numero, strCarpeta, IDTipoDocumento, Arbitraje, BcoEmisor, LC, FPago
    
    '26-6-2000 : ultimo ajuste para no pasar el parametro lo hago aca, sino tengo que alterar todos
    If SaldoCero Then RsCom!ComSaldo = 0
    
    RsCom.Update: RsCom.Close
    
    Cons = "Select Max(ComCodigo) from Compra"
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCompra = RsCom(0)
    RsCom.Close
    IngresoGastoAutomatico = aCompra
    
    'Cargo tabla: GastosSubRubro
    CargoCamposBDGastos aCompra, cImporte, idEmbarque, SubRubro
    
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

Public Sub CargoCamposBDComprobante(cMonto As Currency, fecha As Date, idProveedor As Long, Serie As String, Numero As Long, strCarpeta As String, IDTipoDocumento As Integer, Arbitraje As Double, strBcoEmisor As String, LC As String, FormaPago As String)
    
     RsCom!ComSaldo = cMonto    'Divisa / Arbitraje
    
    RsCom!ComTipoDocumento = IDTipoDocumento
    RsCom!ComFecha = Format(fecha, sqlFormatoF)
    RsCom!ComProveedor = idProveedor
    
    RsCom!ComMoneda = paMonedaDolar
    
    RsCom!ComSerie = Serie
    RsCom!ComNumero = Numero
    
    RsCom!ComImporte = cMonto
    
    RsCom!ComTC = TasadeCambio(paMonedaDolar, paMonedaPesos, PrimerDia(fecha) - 1)
    
    Dim aTexto As String
    aTexto = "C: " & Trim(strCarpeta)
    If Trim(strBcoEmisor) <> "" Then aTexto = aTexto & ", " & Trim(strBcoEmisor)
    If Trim(LC) <> "" Then aTexto = aTexto & ", LC: " & Trim(LC)
    If Trim(FormaPago) <> "" Then aTexto = aTexto & ", " & Trim(FormaPago)
    aTexto = aTexto & ", Arb. U$S = " & Arbitraje
    
    RsCom!ComComentario = aTexto '"Divisa de Carpeta: " & strCarpeta & ", Arb. U$S = " & Arbitraje
    
    RsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
    
    
End Sub

