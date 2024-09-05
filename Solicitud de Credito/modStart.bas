Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String
Public gPathListados As String

Public paLocalZF As Long    'Por Cambio Oct/2000
Public paLocalPuerto As Long

Public paQTelefonos As Integer

Public paTCDolar As Currency

Public Const prmPathSound = "C:\AA Aplicaciones\Sonidos\"
Public Const sndGrabar = "EmitirFactura.wav"
Public Const sndArtNoHabilitado = "ArtNoHabilitado.wav"
Public Const sndArtFueraUso = "ArtFueraUso.wav"

Public prmPlantillaPuente As Long
Public prmArticuloFleteVenta As Long
Public prmArticulosDeFletes As String
Public prmCategoriaDistribuidor As String

Public prmQMaxArticulosPlan As Integer

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        
        CargoParametrosComercio
        CargoParametrosSucursal
        
        CargoParametrosLocal
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        If Trim(Command()) <> "" Then
            Dim sParams() As String
            sParams = Split(Trim(Command()), "|")
            
            'Cxxx|Lxxxx         (Cliente, Llamada)
            For i = LBound(sParams) To UBound(sParams)
                Select Case UCase(Mid(sParams(i), 1, 1))
                    Case "C": frmSolicitud.prmIDCliente = Val(Mid(sParams(i), 2))
                    
                    Case "L": frmSolicitud.prmIDLlamada = Val(Mid(sParams(i), 2))
                End Select
            Next
        End If
        
        frmSolicitud.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()
    On Error GoTo errCPL
    
    prmArticuloFleteVenta = 0
    
    Cons = "Select * from Parametro " & _
               " Where ParNombre like '%telefono%' " & _
               " OR ParNombre IN ('catcliDistribuidor', 'ArticuloFleteParaVentas', 'articulopisoagencia', 'articulodiferenciaenvio', 'SolMaxCantArticulos') "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        Select Case Trim(LCase(RsAux!ParNombre))
            Case "qtelefonosobligatoria": paQTelefonos = RsAux!ParValor
            
            Case LCase("ArticuloFleteParaVentas"): prmArticuloFleteVenta = RsAux!ParValor
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            
            Case LCase("SolMaxCantArticulos"): prmQMaxArticulosPlan = RsAux!ParValor
            
            Case LCase("catcliDistribuidor")
                If Not IsNull(RsAux("ParTexto")) Then prmCategoriaDistribuidor = Trim(RsAux("ParTexto"))
            
        End Select
        
        RsAux.MoveNext
    Loop
    
    RsAux.Close
    
    If prmCategoriaDistribuidor <> "" Then prmCategoriaDistribuidor = "," & Replace(prmCategoriaDistribuidor, " ", "") & ","
    
    FechaDelServidor
    paTCDolar = TasadeCambio(paMonedaDolar, paMonedaPesos, gFechaServidor)
    
    
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplScript) Then prmPlantillaPuente = RsAux!AplScript
    RsAux.Close
    
    Exit Sub
    
errCPL:
    clsGeneral.OcurrioError "Error al cargar los parámetros locales al módulo.", Err.Description
End Sub
'------------------------------------------------------------------------------------------------
'   Busca comentarios con Accion de Alerta para el cliente pasado como parametro
'------------------------------------------------------------------------------------------------
Public Sub BuscoComentariosAlerta(idCliente As Long, Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
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


Public Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function


Public Function fnc_EsDelTipoFlete(IDArticulo As Long) As Boolean
On Error GoTo errFnc
    fnc_EsDelTipoFlete = False
    
    If prmArticulosDeFletes = "" Then
        Dim miRs As rdoResultset
        
        prmArticulosDeFletes = "|"
        
        Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
        Set miRs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not miRs.EOF
            prmArticulosDeFletes = prmArticulosDeFletes & CStr(miRs!TFlArticulo) & "|"
            miRs.MoveNext
        Loop
        miRs.Close
        
        prmArticulosDeFletes = prmArticulosDeFletes & paArticuloPisoAgencia & "|" & paArticuloDiferenciaEnvio & "|"
    End If
    
    fnc_EsDelTipoFlete = (InStr(prmArticulosDeFletes, "|" & CStr(IDArticulo) & "|") <> 0)
    Exit Function

errFnc:
    clsGeneral.OcurrioError "Error al validar artículos de fletes.", Err.Description
End Function

