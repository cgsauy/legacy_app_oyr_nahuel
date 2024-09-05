VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm InicioEMercaderia 
   BackColor       =   &H8000000C&
   Caption         =   "Entrega de Mercadería (Módulo SS.FF.) - Versión 1.00"
   ClientHeight    =   2505
   ClientLeft      =   1215
   ClientTop       =   3795
   ClientWidth     =   9180
   Icon            =   "InicioEMercaderia.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del Sistema"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nueva sección"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "devolucion"
            Object.ToolTipText     =   "Devolución de entregas"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sInicio 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2250
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Sucursal: "
            TextSave        =   "Sucursal: "
            Key             =   "Sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Teminal: "
            TextSave        =   "Teminal: "
            Key             =   "Terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8864
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1640
            MinWidth        =   4
            TextSave        =   "01/06/2012"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InicioEMercaderia.frx":0ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InicioEMercaderia.frx":0DD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "InicioEMercaderia.frx":10EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevoUsuario 
         Caption         =   "&Nueva sección"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuDevolucion 
         Caption         =   "&Devolución de Entregas"
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEntregaDevolucion 
         Caption         =   "Ingreso de Mercadería por Devolución"
      End
   End
   Begin VB.Menu MnuFunciones 
      Caption         =   "Se&cciones"
      Begin VB.Menu MnuF1 
         Caption         =   "Vacío........"
         Shortcut        =   {F1}
      End
      Begin VB.Menu MnuF2 
         Caption         =   "Vacío........"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuF3 
         Caption         =   "Vacío........"
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuF4 
         Caption         =   "Vacío........"
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuF5 
         Caption         =   "Vacío........"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuF6 
         Caption         =   "Vacío........"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MnuF7 
         Caption         =   "Vacío........"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuF8 
         Caption         =   "Vacío........"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MnuF9 
         Caption         =   "Vacío........"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MnuF11 
         Caption         =   "Vacío........"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MnuF12 
         Caption         =   "Vacío........"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu MnuLector 
      Caption         =   "&Lector"
      Begin VB.Menu MnuLeHabilitar 
         Caption         =   "&Habilitar Lector de barras"
      End
      Begin VB.Menu MnuLeDeshabilitar 
         Caption         =   "&Deshabilitar Lector de barras"
      End
   End
   Begin VB.Menu MnuOtros 
      Caption         =   "O&tros"
      Begin VB.Menu MnuArticulo 
         Caption         =   "Modificar Artículos"
      End
      Begin VB.Menu MnuImpEntregas 
         Caption         =   "Imprimir Entregas"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      NegotiatePosition=   3  'Right
      Begin VB.Menu MnuSalirSist 
         Caption         =   "&Salir del Sistema"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "InicioEMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPidoLogin As Boolean
Public bSinLector As Boolean

Private Sub MDIForm_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub MDIForm_Load()
    
    On Error GoTo errInicio
    ObtengoSeteoForm Me, 10, 10, Screen.Width - 20, Screen.Height - 600
    
    SetToolBarFlat Toolbar1, True
    bSinLector = True
    
    'Conexion a la base de datos----------------------------------------
    Set eBase = rdoCreateEnvironment("", "", "")
    eBase.CursorDriver = rdUseServer
        
    Cons = miConexion.TextoConexion("comercio")
    txtConexion = Cons
    'Cons = "dsn=SSFF;uid=sa;pwd=;server=PROLIANT;database=CGSAII;"
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , Cons)
    cBase.QueryTimeout = 15
    '------------------------------------------------------------------------
    
    FechaDelServidor
    
    Me.Show
    
    gPathListados = App.Path & "\REPORTES\"
    pathApp = App.Path & "\Aplicaciones"
    
    CargoParametros
    CargoParametrosSucursal
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    Exit Sub

errInicio:
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la conexión con el servidor."
    On Error Resume Next
    If Me.Visible Then Unload Me
    End
End Sub

Private Sub VerificoConfiguracion()

    'VERIFICO FORMATO DE MONEDA Y FECHA---------------------------------------
    If CCur("2,222.22") <> 2222.22 Or CDate("03/10/97") <> "03/10/97" Then
        MsgBox "La configuración de la terminal es incompatible con la configuración del sistema. " & Chr(13) _
        & "Por mayor información ver Requerimientos del Sistema en el manual de usuario.", vbCritical, "ERROR CRÍTICO"
        cBase.Close
        eBase.Close
        End
    End If
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If FormActivo("SeEntrega") Then
        MsgBox "Hay secciones de entrega abiertas. Ciérrelas antes de salir del sistema.", vbInformation, "ATENCIÓN"
        Cancel = 1
    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    On Error Resume Next
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    cBase.Close
    eBase.Close
    GuardoSeteoForm Me
    End
    
End Sub

Private Sub CargoParametrosSucursal()

Dim aNombreTerminal As String

    aNombreTerminal = miConexion.NombreTerminal
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Local" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = LocCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        If Not IsNull(RsAux!LocDisponibilidad) Then paDisponibilidad = RsAux!LocDisponibilidad
        sInicio.Panels("Terminal").Text = "Terminal: " & aNombreTerminal & " "
        sInicio.Panels("Sucursal").Text = "Sucursal: " & Trim(RsAux!LocNombre) & " "
        prmNombreLocal = Trim(RsAux!LocNombre)
        If RsAux!LocSinLector Then bSinLector = True Else bSinLector = False
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
            & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Unload Me
        Exit Sub
    End If
    '-------------------------------------------------------------------------------------------------------------------------

    
End Sub

Private Sub CargoParametros()

    'Parametros a cero-----------------
    paECivilConyuge = 0
    paDepartamento = 0
    paLocalidad = 0
    paMonedaEmpleo = 0
    paTipoIngreso = 0
    paTipoTelefonoP = 0
    paTipoTelefonoE = 0
    paCategoriaCliente = 0
    paVigenciaEmpleo = 0
    paTipoCuotaContado = 0
    paMonedaFacturacion = 0
    paArticuloPisoAgencia = 0
    paArticuloDiferenciaEnvio = 0
    paPrimeraHoraEnvio = 0
    paUltimaHoraEnvio = 0
    paEnvioFechaPrometida = 0
    paMonedaFija = 0
    paMonedaFijaTexto = ""
    paVaToleranciaMonedaPorc = 0
    paVaToleranciaDiasExh = 0
    paVaToleranciaDiasExhTit = 0
    paToleranciaMora = 0
    paDiasCobranzaCuota = 0
    paCoeficienteMora = 1
    paIvaMora = 1
    paEstadoArticuloEntrega = 1
    paLlamadaAMoroso = 0
    
    paMonedaDeuda = 0
    paCartaAGarantia = 0
    paCartaATitular = 0
    
    'Dias para niveles de iconos
    paIconoPendienteN2Dias = 60
    paIconoVencimientoN2Dias = 60

    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case Trim(RsAux!ParNombre)
            Case "ECivilConyuge": paECivilConyuge = RsAux!ParValor
            
            Case "Departamento": paDepartamento = RsAux!ParValor
            Case "Localidad": paLocalidad = RsAux!ParValor
            
            Case "MonedaEmpleo": paMonedaEmpleo = RsAux!ParValor
            Case "MonedaFacturacion": paMonedaFacturacion = RsAux!ParValor
            Case "TipoIngreso": paTipoIngreso = RsAux!ParValor
            Case "TipoTelefonoP": paTipoTelefonoP = RsAux!ParValor
            Case "TipoTelefonoE": paTipoTelefonoE = RsAux!ParValor
            Case "CategoriaCliente": paCategoriaCliente = RsAux!ParValor
            Case "VigenciaEmpleo": paVigenciaEmpleo = RsAux!ParValor
            
            Case "TipoCuotaContado": paTipoCuotaContado = RsAux!ParValor
                
            Case "ArticuloPisoAgencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "ArticuloDiferenciaEnvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            Case "PrimeraHoraEnvio": paPrimeraHoraEnvio = RsAux!ParValor
            Case "UltimaHoraEnvio": paUltimaHoraEnvio = RsAux!ParValor
            Case "EnvioFechaPrometida": paEnvioFechaPrometida = RsAux!ParValor
            Case "MonedaFija": paMonedaFija = RsAux!ParValor
            
            Case "VaToleranciaMonedaPorc": paVaToleranciaMonedaPorc = RsAux!ParValor
            Case "VaToleranciaDiasExh": paVaToleranciaDiasExh = RsAux!ParValor
            Case "VaToleranciaDiasExhTit": paVaToleranciaDiasExhTit = RsAux!ParValor
            Case "ToleranciaMora": paToleranciaMora = RsAux!ParValor
            Case "CoeficienteMora": paCoeficienteMora = ((RsAux!ParValor / 100) + 1) ^ (1 / 30)                        'Como es mensual calculo el diario
            Case "IvaMora": paIvaMora = RsAux!ParValor
                
            Case "DiasCobranzaCuota": paDiasCobranzaCuota = RsAux!ParValor
            Case "EstadoArticuloEntrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "IconoPendienteN2Dias": paIconoPendienteN2Dias = RsAux!ParValor
            Case "IconoVencimientoN2Dias": paIconoVencimientoN2Dias = RsAux!ParValor
            Case "LlamadaAMoroso": paLlamadaAMoroso = RsAux!ParValor
            Case "MonedaDeuda": paMonedaDeuda = RsAux!ParValor
                
            Case "CartaAGarantia": paCartaAGarantia = RsAux!ParValor
            Case "CartaATitular": paCartaATitular = RsAux!ParValor
            
            Case "TipoArticuloServicio": paTipoArticuloServicio = RsAux!ParValor
            
            Case "TComentarioEMercaderia": prmTCAlEntregar = Trim(RsAux!ParTexto)
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paMonedaFija <> 0 Then
        Cons = "Select * from Moneda Where MonCodigo = " & paMonedaFija
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            paMonedaFijaTexto = Trim(RsAux!MonSigno)
        Else
            MsgBox "El código de moneda fija (parámetro) no existe en la base de datos.", vbCritical, "ERROR"
            paMonedaFija = 0
        End If
        RsAux.Close
    End If
    
    'Habria que validar que parametros no fueron cargados
    
    
End Sub


Private Sub MnuArticulo_Click()
    'EjecutarApp pathApp & "\Articulos Entrega.exe"
    EjecutarApp pathApp & "\Articulos_Entrega.exe"
End Sub

Private Sub MnuDevolucion_Click()
    AccionDevolucion
End Sub

Private Sub ActivoForm(Tecla As Long)

Dim f As Form
    
    For Each f In Forms
        On Error GoTo Continuo
        If f.pTecla = Tecla Then
            If f.WindowState = vbMinimized Then f.WindowState = vbNormal
            On Error Resume Next
            f.SetFocus
            Exit Sub
        End If
Continuo:
    Next
End Sub

Private Sub MnuEntregaDevolucion_Click()
    EjecutarApp pathApp & "\Entrada por Devolucion"
End Sub

Private Sub MnuF1_Click()
    ActivoForm vbKeyF1
End Sub

Private Sub MnuF11_Click()
ActivoForm vbKeyF11
End Sub
Private Sub MnuF12_Click()
    ActivoForm vbKeyF12
End Sub
Private Sub MnuF2_Click()
    ActivoForm vbKeyF2
End Sub
Private Sub MnuF3_Click()
    ActivoForm vbKeyF3
End Sub
Private Sub MnuF4_Click()
    ActivoForm vbKeyF4
End Sub
Private Sub MnuF5_Click()
    ActivoForm vbKeyF5
End Sub
Private Sub MnuF6_Click()
    ActivoForm vbKeyF6
End Sub
Private Sub MnuF7_Click()
    ActivoForm vbKeyF7
End Sub
Private Sub MnuF8_Click()
    ActivoForm vbKeyF8
End Sub
Private Sub MnuF9_Click()
    ActivoForm vbKeyF9
End Sub

Private Sub MnuImpEntregas_Click()
    frmImprimirEntrega.Show vbModeless, Me
End Sub

Private Sub MnuLeDeshabilitar_Click()
    AccionLector No:=True
End Sub

Private Sub MnuLeHabilitar_Click()
    AccionLector Si:=True
End Sub

Private Sub MnuNuevoUsuario_Click()
    AbrirSeccion
End Sub

Private Sub MnuSalirSist_Click()
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "salir": Unload Me
        Case "nuevo": AbrirSeccion
        Case "devolucion": AccionDevolucion
    End Select
    
End Sub

Private Sub AccionLector(Optional Si As Boolean = False, Optional No As Boolean = False)

Dim rsLec As rdoResultset

    On Error GoTo errLector
    If Si Then
        If Not bSinLector Then
            MsgBox "El lector de barras ya está habilitado.", vbExclamation, "Lector Habilitado"
            Exit Sub
        End If
        
        Screen.MousePointer = 11
        Cons = "Select * from Local Where LocCodigo = " & paCodigoDeSucursal
        Set rsLec = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsLec.Edit
        rsLec!LocSinLector = 0
        rsLec.Update: rsLec.Close
        
        bSinLector = False
        
        FechaDelServidor
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Varios, paCodigoDeTerminal, miConexion.UsuarioLogueado(Codigo:=True), 0, , _
                Trim(sInicio.Panels("Sucursal").Text) & ". Habilita lector de barras."
        
        MsgBox "El lector de barras fue habilitado con éxito.", vbInformation, "Lector Habilitado"
    
    Else
    
        If bSinLector Then
            MsgBox "El lector de barras ya está deshabilitado.", vbExclamation, "Lector Deshabilitado"
            Exit Sub
        End If
        Screen.MousePointer = 11
        Dim objSuceso As New clsSuceso, gSucesoUsr As Long, gSucesoDef As String
        
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Deshabilitar Lector de Barras", cBase
        DoEvents
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
        
        Screen.MousePointer = 11
        Cons = "Select * from Local Where LocCodigo = " & paCodigoDeSucursal
        Set rsLec = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsLec.Edit
        rsLec!LocSinLector = 1
        rsLec.Update: rsLec.Close
        
        bSinLector = True
        
        FechaDelServidor
        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Varios, paCodigoDeTerminal, gSucesoUsr, 0, , _
            Trim(sInicio.Panels("Sucursal").Text) & ". Deshabilita lector de barras.", gSucesoDef
        
        MsgBox "El lector de barras fue deshabilitado con éxito.", vbInformation, "Lector Deshabilitado"
    
    End If
    Screen.MousePointer = 0
    Exit Sub

errLector:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AbrirSeccion()

    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu("Seccion Entrega") Then Exit Sub
    
    InUsuario.pUsuarioCodigo = 0
    InUsuario.pUsuarioNombre = ""
    InUsuario.pUsuarioTecla = 0
    InUsuario.Show vbModal, Me
    
    If InUsuario.pUsuarioCodigo <> 0 Then
        On Error GoTo errAbrir
        Dim fEntrega As New SeEntrega
        Screen.MousePointer = 11
        
        fEntrega.pTecla = InUsuario.pUsuarioTecla
        fEntrega.Caption = UCase(InUsuario.pUsuarioNombre)
        fEntrega.Tag = InUsuario.pUsuarioCodigo
        
        Select Case InUsuario.pUsuarioTecla
            Case vbKeyF1: MnuF1.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF2: MnuF2.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF3: MnuF3.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF4: MnuF4.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF5: MnuF5.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF6: MnuF6.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF7: MnuF7.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF8: MnuF8.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF9: MnuF9.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF11: MnuF11.Caption = UCase(InUsuario.pUsuarioNombre)
            Case vbKeyF12: MnuF12.Caption = UCase(InUsuario.pUsuarioNombre)
        End Select
        
        fEntrega.Show vbModeless, Me
         
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errAbrir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al abrir la nueva sección de entrega."
End Sub

Private Sub AccionDevolucion()

    Screen.MousePointer = 11
    If Not miConexion.AccesoAlMenu("Devolucion Entrega") Then Exit Sub
    DeEntrega.Show vbModeless, Me
    
End Sub
