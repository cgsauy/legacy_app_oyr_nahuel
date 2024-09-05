Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paEstadoSano As Integer
Public paEstadoRoto As Integer

Public prmPathApp As String
Public prmCCDistribuidor As String
Public prmTCCAmbio As Integer
Public prmMenUsuarioCambioArticulo As String
Public prmMenUsuarioSistema As Long

Public Const prmSucesoMovStock = 11

'----------------------------------------------------------------
Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum
Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum
Public Enum TipoControlMercaderia
    CambioEstado = 1
    EntregaMercaderia = 2
End Enum
'Definiciones de Tipos de Locales
Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

'Public Enum TipoCliente
'    Cliente = 1
'    Empresa = 2
'End Enum

Public Enum TipoService
    CGSA = 1
    Rydesul = 2
End Enum

'Parámetros de Entrada
Dim prmE_IDCambio As Long, prmE_IDServiceCGSA As Long
    
Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean
    bAccesoOK = False
    bAccesoOK = miConexion.AccesoAlMenu("Cambio en Garantia")
    
    If bAccesoOK Then
    
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 45) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        InicioConexionBDRydesul miConexion.TextoConexion("rydesul"), 30
        
        CargoParametrosLocales
        
        CargoParametrosEntrada Trim(Command())
        'prmE_IDServiceCGSA = 230
        frmMain.Show vbModeless
                
        frmMain.gbl_CargaConParametros prmE_IDCambio, prmE_IDServiceCGSA
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Autorización"
        End If
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

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'EstadoArticuloEntrega', 'EstadoRoto', 'PathApp', 'catcliDistribuidor', 'TComentarioCambioProd', 'menUsuarioCambioArticulo', 'UsuarioSistema')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "estadoarticuloentrega": paEstadoSano = rsAux!ParValor
            Case "estadoroto": paEstadoRoto = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto)
            
            Case "tcomentariocambioprod": prmTCCAmbio = rsAux!ParValor
            
            Case "catclidistribuidor"
                If Not IsNull(rsAux!ParTexto) Then
                    prmCCDistribuidor = Trim(rsAux!ParTexto)
                Else
                    prmCCDistribuidor = 0
                End If
            
            Case "menusuariocambioarticulo"
                If Not IsNull(rsAux!ParTexto) Then
                    prmMenUsuarioCambioArticulo = Trim(rsAux!ParTexto)
                Else
                    prmMenUsuarioCambioArticulo = ""
                End If
            
            Case "usuariosistema": prmMenUsuarioSistema = rsAux!ParValor
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cons = "Select * from Terminal Where TerNombre = '" & miConexion.NombreTerminal & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeTerminal = rsAux!TerCodigo
        paCodigoDeSucursal = rsAux!TerSucursal
    End If
    rsAux.Close
    
    Exit Sub

errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Public Function EndMain()
    On Error Resume Next
    CierroConexionBDRydesul
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

Private Function CargoParametrosEntrada(mPrms As String)
    On Error GoTo errCPE
    'ID ID Cambio |S ID Service CGSA
    prmE_IDServiceCGSA = 0: prmE_IDCambio = 0
    
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For I = LBound(arrPrms) To UBound(arrPrms)
        arrValues = Split(arrPrms(I), " ")
        Select Case UCase(arrValues(0))
            
            Case "ID"
                prmE_IDCambio = Val(arrValues(1))
                        
            Case "S": prmE_IDServiceCGSA = Val(arrValues(1))
            
        End Select
        
    Next
    
    Exit Function
    
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function

