Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA


'Public paEstadoArticuloEntrega As Integer
Public prmPathApp As String
Public prmPlBalance  As String
Public prmCBase As String

'Datos de la base de datos II (para tambien generar movimientos)
Public cBaseMov As rdoConnection
Public bHayBDMov As Boolean
Public prmCBaseMov As String

Public prmFileName As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    bHayBDMov = False
    
    If miConexion.AccesoAlMenu("Diferencias Balance") Then
    
        prmCBase = "balance"
        prmCBase = Trim(InputBox("Se hacen consultas y correcciones de stock sobre la base de datos:", "BD de Consultas (Datos Balance)", prmCBase))
        
        If Not InicioConexionBD(miConexion.TextoConexion(prmCBase), 45) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        prmCBaseMov = "comercio"
        prmCBaseMov = Trim(InputBox("Se hacen correcciones de stock sobre la base de datos:", "Doble Corrección de Stock (BD Comercio)", prmCBaseMov))
        
        If prmCBaseMov <> "" Then
            If InicioConexionBDMov(miConexion.TextoConexion(prmCBaseMov)) Then bHayBDMov = True
        End If
        
        If prmCBaseMov = prmCBase Then bHayBDMov = False
        
        CargoParametrosLocales
        CargoFileName
        
        frmListado.Show vbModeless
        
        prmCBase = propConnect(cBase.Connect, "database")
        If bHayBDMov Then
            prmCBaseMov = propConnect(cBaseMov.Connect, "database")
            cons = " (Consultas y Correciones sobre '" & prmCBase & "'   Hay doble corrección a '" & prmCBaseMov & "')"
        Else
            
            cons = " (Consultas y Correciones sobre '" & prmCBase & "')"
        End If
        frmListado.Caption = Trim(frmListado.Caption) & cons
        
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

Private Function propConnect(mTexto As String, Key As String) As String
On Error GoTo errProp
    
    Dim arrCampos() As String, arrData() As String
    arrCampos = Split(mTexto, ";")
    
    For I = LBound(arrCampos) To UBound(arrCampos)
        arrData = Split(arrCampos(I), "=")
        If UCase(arrData(0)) = UCase(Key) Then
            propConnect = arrData(1)
            Exit For
        End If
    Next

errProp:
End Function

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro Where ParNombre IN ( 'EstadoArticuloEntrega', 'PathApp', 'PlBalance' )"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto)
            
            Case "plbalance": prmPlBalance = Trim(rsAux!ParTexto)
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select * from Terminal Where TerNombre = '" & miConexion.NombreTerminal & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then paCodigoDeTerminal = rsAux!TerCodigo
    rsAux.Close
    
    Exit Sub

errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


'****     FUNCIONES DOBLES PARA LA BASE DE DATOS SECUNDARIA              cBaseMov    ---------------------****************
Public Function InicioConexionBDMov(strConexion, Optional sqlTimeOut As Integer = 15) As Boolean
    
    On Error GoTo ErrICBD
    InicioConexionBDMov = False
    
    'Conexion a la base de datos----------------------------------------
    Set cBaseMov = eBase.OpenConnection("", , , strConexion)
    cBaseMov.QueryTimeout = sqlTimeOut
    
    InicioConexionBDMov = True
    Exit Function
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al intentar comunicarse con la Base de Datos." & vbCrLf & _
                "Error: " & Err.Description, vbExclamation, "Error de Conexión"
End Function

Public Function CierroConexionBDMov()
    On Error Resume Next
    cBaseMov.Close
End Function


Public Sub bdMov_MarcoMovimientoStockFisicoEnLocal(TipoLocal As Integer, CodigoLocal As Long, Articulo As Long, Cantidad As Currency, Estado As Integer, AltaOBaja As Integer)

Dim RsSLo As rdoResultset

    cons = "Select * From StockLocal " _
            & " Where StLTipoLocal = " & TipoLocal _
            & " And StlLocal = " & CodigoLocal _
            & " And StLArticulo = " & Articulo _
            & " And StLEstado = " & Estado
    Set RsSLo = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If RsSLo.EOF Then
        RsSLo.AddNew
        RsSLo!StLTipoLocal = TipoLocal
        RsSLo!StlLocal = CodigoLocal
        RsSLo!StLArticulo = Articulo
        RsSLo!StLEstado = Estado
        RsSLo!StLCantidad = Cantidad * AltaOBaja
        RsSLo.Update
    Else
        RsSLo.Edit
        RsSLo!StLCantidad = RsSLo!StLCantidad + (Cantidad * AltaOBaja)
        RsSLo.Update
    End If
    
    RsSLo.Close
    
End Sub

Public Sub bdMov_MarcoMovimientoStockFisico(lnUsuario As Long, iTipoLocal As Integer, iLocal As Long, lnArticulo As Long, cCantidad As Currency, iEstadoMercaderia As Integer, iAltaoBaja As Integer, Optional iTipoDocumento As Integer = -1, Optional lnDocumento As Long = -1)
        
Dim rsFis As rdoResultset

    cons = "Select * from MovimientoStockFisico Where MSFCodigo = 0"
    Set rsFis = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsFis.AddNew
    
    rsFis!MSFFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
    rsFis!MSFTipoLocal = iTipoLocal
    rsFis!MSFLocal = iLocal
    
    rsFis!MSFArticulo = lnArticulo
    rsFis!MSFCantidad = cCantidad * iAltaoBaja
    rsFis!MSFEstado = iEstadoMercaderia
    
    If iTipoDocumento <> -1 Then
        rsFis!MSFTipoDocumento = iTipoDocumento
        If lnDocumento <> -1 Then rsFis!MSFDocumento = lnDocumento Else rsFis!MSFDocumento = Null
    Else
        rsFis!MSFTipoDocumento = Null
        rsFis!MSFDocumento = Null
    End If
    
    rsFis!MSFUsuario = lnUsuario
    
    If paCodigoDeTerminal > 0 Then rsFis!MSFTerminal = paCodigoDeTerminal Else rsFis!MSFTerminal = Null
    
    rsFis.Update
    rsFis.Close
    
End Sub

Public Sub bdMov_MarcoMovimientoStockTotal(Articulo As Long, TipoEstado As Integer, Estado As Integer, Cantidad As Currency, AltaOBaja As Integer)
 
 Dim RsSTo As rdoResultset
 
    cons = "Select * From StockTotal" _
            & " Where StTArticulo = " & Articulo _
            & " And StTTipoEstado = " & TipoEstado _
            & " And StTEstado = " & Estado
            
    Set RsSTo = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If RsSTo.EOF Then
        RsSTo.AddNew
        RsSTo!StTArticulo = Articulo
        RsSTo!StTTipoEstado = TipoEstado
        RsSTo!StTEstado = Estado
        RsSTo!StTCantidad = Cantidad * AltaOBaja
        RsSTo.Update
    Else
        RsSTo.Edit
        RsSTo!StTCantidad = RsSTo!StTCantidad + (Cantidad * AltaOBaja)
        RsSTo.Update
    End If
    RsSTo.Close
    
End Sub


Private Function CargoFileName()
On Error GoTo errCF

    prmFileName = "StockBalance" & Format(Now, "yyyy") & ".txt"
    
    cons = "Select Max(CabMesCosteo) as CabMesCosteo from CMCabezal"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then prmFileName = "StockBalance" & Format(rsAux!CabMesCosteo, "yyyy") & ".txt"
    rsAux.Close
    
    prmFileName = App.Path & "\" & prmFileName
    
    Exit Function
errCF:
    prmFileName = ""
End Function
